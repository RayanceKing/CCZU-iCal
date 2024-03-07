import json
import sys
import datetime
import time
import re
import uuid
from icalendar import Calendar, Event, Alarm
from lxml import etree
from typing import Optional


def getDomOffline(filePath: str) -> Optional[str]:
    try:
        with open(filePath, 'r', encoding='utf-8') as file:
            content = file.read()
        return content
    except Exception as e:
        print(f"读取本地文件失败: {e}")
        return None


def classHandler(text):
    # 结构文本
    textDom = etree.HTML(text)
    tables = textDom.xpath("//div/table")
    tableup, tabledown = tables[1], tables[2]
    # 提取所有课程名
    classNameList = tableup.xpath(
        './tr[@class="dg1-item"]/td[position()=2]/text()')
    # 从表格中提取课程信息
    classmatrix = [
        tr.xpath("./td[position()>1]/text()")
        for tr in tabledown.xpath("tr[position()>1]")
    ]
    classmatrixT = [each for each in zip(*classmatrix)]
    oeDict = {"单": 1, "双": 2}
    courseInfo = dict()
    courseList = dict()
    global courseInfoRes

    # day: 一周中的某一天 / courses: 一天内的所有课程

    for day, courses in enumerate(classmatrixT):
        for course_time, course_cb in enumerate(courses):
            course_list = list(filter(None, course_cb.split("/")))
            targetlen = len(course_list)
            index = 0
            while index < targetlen:
                course = course_list[index]
                if course != "\xa0":  # 检查课程是否为空
                    id = uuid.uuid3(uuid.NAMESPACE_DNS, course + str(day)).hex
                    if not course_time or id not in courseInfo.keys():
                        nl = list(filter(lambda x: course.startswith(x), classNameList))
                        if not nl:  # 如果没有匹配的课程名称
                            if index < targetlen - 1:  # 如果不是最后一个元素
                                # 将当前课程和下一个课程合并，并替换下一个课程
                                course_list[index + 1] = f"{course}/{course_list[index + 1]}"
                                index += 1  # 跳过下一个元素
                                continue
                            else:
                                raise ValueError("无法正确解析课程名称")
                        # 如果找到匹配的课程名称
                        assert len(nl) == 1, "多个课程名称匹配，无法正确解析"
                        classname = nl[0]
                        course = course.replace(classname, "").strip()
                        res = re.match(r"(\S+)? *([单双]?) *((\d+-\d+,?)+)", course)
                        assert res, "Course information parsing exception"
                        info = {
                            "classname": classname,
                            "classtime": [course_time + 1],
                            "day": day + 1,
                            "week": list(filter(None, res.group(3).split(","))),
                            "oe": oeDict.get(res.group(2), 3),
                            "classroom": [res.group(1)],
                        }
                        courseInfo[id] = info
                    elif id in courseInfo.keys():
                        courseInfo[id]["classtime"].append(course_time + 1)
                index += 1

    # 合并同一课程的不同上课时间
    for course in courseInfo.values():
        purecourse = {key: value for key,
                      value in course.items() if key != "classroom"}
        # 如果课程已经存在，将教室信息添加到课程信息中
        if str(purecourse) in courseList:
            courseList[str(purecourse)]["classroom"].append(
                course["classroom"][0])
        # 如果课程不存在，将课程信息添加到课程列表中
        else:
            courseList[str(purecourse)] = course
    # 将课程列表转换为课程信息列表
    courseInfoRes = [course for course in courseList.values()]
    print(courseInfoRes)
    print("课表格式化成功")


# 定义函数，传入课表，返回ics文件


def setReminder(reminder):
    # reminder: 课前提醒时间
    global timeReminder
    reminder = 15 if reminder == "" else reminder
    # 将分钟转换为ics文件中的时间格式
    time_tuple = re.match(
        r"(([\d ]+) days, )*(\d+):(\d+):(\d+)",
        str(datetime.timedelta(minutes=int(reminder))),
    ).groups()[1:]
    # 将时间格式转换为ics文件中的时间格式
    time_map = map(lambda x: x if x else "0", time_tuple)
    timeReminder = "-P{}DT{}H{}M{}S".format(*list(time_map))
    print("SetReminder:", timeReminder)


# 定义函数，传入课表，返回ics文件


def setClassTime():
    # 从配置文件中读取上课时间
    data = []
    with open("/Users/wangyuliang/文件-本地/200-Code/CCZU-iCal/conf_classTime.json", "r") as f:
        data = json.load(f)
    global classTimeList
    classTimeList = data["classTime"]
    print("上课时间配置成功")


def save(string):
    f = open("class.ics", "wb")
    f.write(string.encode("utf-8"))
    f.close()

# 定义类，传入课表，返回ics文件


class ICal(object):
    def __init__(self, firstWeekDate, schedule, courseInfo):
        self.firstWeekDate = firstWeekDate
        self.schedule = schedule
        self.courseInfo = courseInfo

    # 传入字符串日期，返回类实例

    @classmethod
    def withStrDate(cls, strdate, *args):
        firstWeekDate = time.strptime(strdate, "%Y%m%d")
        return cls(firstWeekDate, *args)

    # 传入时间戳，返回类实例

    def handler(self, info):
        weekday = info["day"]
        oe = info["oe"]
        firstDate = datetime.datetime.fromtimestamp(
            int(time.mktime(self.firstWeekDate))
        )
        info["daylist"] = list()
        # 将课程的周数转换为日期
        for weeks in info["week"]:
            startWeek, endWeek = map(int, weeks.split("-"))
            startDate, endDate = (
                firstDate
                + datetime.timedelta(days=(float((startWeek - 1) * 7) + weekday - 1)),
                firstDate
                + datetime.timedelta(days=(float((endWeek - 1) * 7) + weekday - 1)),
            )

            # 如果课程为单周或双周，将其添加到课程信息中
            while True:
                if (
                    oe == 3
                    or (oe == 1)
                    and (startWeek % 2 == 1)
                    or (oe == 2)
                    and (startWeek % 2 == 0)
                ):
                    info["daylist"].append(startDate.strftime("%Y%m%d"))
                startDate = startDate + datetime.timedelta(days=7.0)
                startWeek = startWeek + 1
                if startDate > endDate:
                    break
        return info

    # 传入课表，返回ics文件

    def to_ical(self):
        prop = {
            "PRODID": "-//Google Inc//Google Calendar 70.9054//EN",
            "VERSION": "2.0",
            "CALSCALE": "GREGORIAN",
            "METHOD": "PUBLISH",
            "X-WR-CALNAME": "课程表",
            "X-WR-TIMEZONE": "Asia/Shanghai",
        }
        # 将课表信息添加到ics文件中
        cal = Calendar()
        for key, value in prop.items():
            cal.add(key, value)

        courseInfo = map(self.handler, self.courseInfo)
        for course in courseInfo:
            startTime = self.schedule[course["classtime"][0] - 1]["startTime"]
            endTime = self.schedule[course["classtime"][-1] - 1]["endTime"]
            classroom = list(filter(None, course["classroom"]))
            createTime = datetime.datetime.now()
            for day in course["daylist"]:
                sub_prop = {
                    "CREATED": createTime,
                    "SUMMARY": "{0} | {1}".format(
                        course["classname"], "/".join(classroom)
                    ),
                    "UID": uuid.uuid4().hex + "@google.com",
                    "DTSTART": datetime.datetime.strptime(
                        day + startTime, "%Y%m%d%H%M"
                    ),
                    "DTEND": datetime.datetime.strptime(day + endTime, "%Y%m%d%H%M"),
                    "DTSTAMP": createTime,
                    "LAST-MODIFIED": createTime,
                    "SEQUENCE": "0",
                    "TRANSP": "OPAQUE",
                    "X-APPLE-TRAVEL-ADVISORY-BEHAVIOR": "AUTOMATIC",
                }
                # 如果课前提醒时间不为空，将其添加到课程信息中
                sub_prop_alarm = {
                    "ACTION": "DISPLAY",
                    "DESCRIPTION": "This is an event reminder",
                    "TRIGGER": timeReminder,
                }
                event = Event()
                for key, value in sub_prop.items():
                    event.add(key, value)
                alarm = Alarm()
                for key, value in sub_prop_alarm.items():
                    alarm[key] = value
                event.add_component(alarm)
                cal.add_component(event)

        # 每周信息
        fweek = datetime.datetime.fromtimestamp(
            int(time.mktime(self.firstWeekDate))) - datetime.timedelta(days=1.0)
        createTime = datetime.datetime.now()
        for _ in range(18):
            sub_prop = {
                "CREATED": createTime,
                "SUMMARY": "学期第 {} 周".format(_ + 1),
                "UID": uuid.uuid4().hex + "@google.com",
                "DTSTART": fweek.date(),
                "DTEND": (fweek + datetime.timedelta(days=7.0)).date(),
                "DTSTAMP": createTime,
                "LAST-MODIFIED": createTime,
                "SEQUENCE": "0",
                "TRANSP": "OPAQUE",
                "X-APPLE-TRAVEL-ADVISORY-BEHAVIOR": "AUTOMATIC",
            }
            fweek += datetime.timedelta(days=7.0)
            event = Event()
            for key, value in sub_prop.items():
                event.add(key, value)
            cal.add_component(event)

        return (
            bytes.decode(
                cal.to_ical(), encoding="utf-8").replace("\r\n", "\n").strip()
        )


# 主函数
if __name__ == "__main__":
    firstWeekDate = "20240226"
    classTimeList = None
    courseInfoRes = None
    timeReminder = "15"

    # 使用本地HTML文件进行离线测试
    filePath = "/Users/wangyuliang/文件-本地/200-Code/教务管理信息系统 2.html"  # 保存课表页面的本地文件路径
    textDom = getDomOffline(filePath)
    if not textDom:
        print("遇到错误，请检查本地文件路径是否正确")
        sys.exit(0)
    else:
        print("从本地文件获取课表成功")

    print("开始课表格式化...")
    classHandler(textDom)
    print("正在配置上课时间...")
    setClassTime()

    '''firstWeekDate =  input(    "请输入此学期第一周的星期一日期(eg 20230904):")  # 周一第一周的开始数据
    print("正在配置第一周周一日期...")'''
    # print("SetFirstWeekDate:", firstWeekDate)

    '''reminder = input("正在配置提醒功能,请以分钟为单位设定课前提醒时间(默认值为15):")
    print("正在配置课前提醒...")
    setReminder(reminder)'''

    print("正在生成ics文件...")
    iCal = ICal.withStrDate(firstWeekDate, classTimeList, courseInfoRes)
    with open("/Users/wangyuliang/文件-本地/200-Code/class.ics", "w", encoding="utf-8") as f:
        f.write(iCal.to_ical())
    print("文件保存成功")
