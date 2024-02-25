import json
import sys
import datetime
import time
import re
import uuid
import requests
from icalendar import Calendar, Event, Alarm
from lxml import etree
from typing import Optional


def loginCookie(
    user: str, passwd: str
) -> dict:  # 定义函数，传入学号和密码，返回Cookies
    session = requests.session()
    url = "http://jwcas.cczu.edu.cn/login"

    # 获取随机信息
    try:
        html = session.get(url, headers=headers)
        html.raise_for_status()
        html.encoding = html.apparent_encoding
        html = html.text
    except Exception:  # 如果获取失败，退出程序
        print("从登录页获取随机信息失败")
        sys.exit(0)

    # 初始化字符串，使其可用于xpath的函数
    html = etree.HTML(html)
    # 获取随机数据的名称和值
    # Type of gName, gValue: list
    gName = html.xpath('//input[@type="hidden"]/@name')  # 获取随机信息的name
    gValue = html.xpath('//input[@type="hidden"]/@value')  # 获取随机信息的value
    gAll = {}  # 创建字典，用于存储随机信息
    for i in range(3):  # 将随机信息存入字典
        gAll[gName[i]] = gValue[i]  # 将随机信息的name和value存入字典

    # 发送数据
    data = {
        "username": user,
        "password": passwd,
        "warn": "true",
        "lt": gAll["lt"],
        "execution": gAll["execution"],
        "_eventId": gAll["_eventId"],
    }

    # 官方登录
    sc = session.post(url, headers=headers, data=data)
    if not sc.cookies.get_dict():
        print("用户名或密码错误，请检查重试")
        sys.exit(0)

    # 拦截跳转链接
    try:
        tmp = session.get(
            "http://jwcas.cczu.edu.cn/login?service=http://219.230.159.132/login7_jwgl.aspx",
            headers=headers,
        )
        tmp_html = etree.HTML(tmp.text)
        Rurl = tmp_html.xpath("//a[@href and text()]/@href")[0]
    except Exception:
        print("获取跳转链接失败")
        sys.exit(0)

    # 从DirectPage获取我们需要的Cookie
    try:
        tmp2 = session.get(Rurl, headers=headers)
    except Exception:
        print("获取实用Cookies失败")
        sys.exit(0)

    # 提取cookie字典并返回它。
    print("获取Cookies成功")
    return tmp2.cookies.get_dict()


# 定义函数，传入学号和密码，返回Cookies

# 定义函数，传入Cookies，返回课表


def getDom(cookies: dict) -> Optional[str]:
    url = "http://219.230.159.132/web_jxrw/cx_kb_xsgrkb.aspx"

    try:
        rep = requests.get(url, headers=headers, cookies=cookies)
        rep.raise_for_status()
        return rep.text
    except requests.exceptions.HTTPError:  # If get the status code - 500
        return None


def classHandler(text):
    # 结构文本
    textDom = etree.HTML(text)
    tables = textDom.xpath("//div/table")
    tableup, tabledown = tables[1], tables[2]
    # 提取所有课程名
    classNameList = tableup.xpath('./tr[@class="dg1-item"]/td[position()=2]/text()')
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
        # time: 课程的排名 / course_cb: 表单元格中的一个项目
        for course_time, course_cb in enumerate(courses):
            course_list = list(filter(None, course_cb.split("/")))
            for course in course_list:
                id = uuid.uuid3(uuid.NAMESPACE_DNS, course + str(day)).hex
                # 如果课程不为空且不在课程信息中，将其添加到课程信息中
                if course != "\xa0" and (
                    not course_time or id not in courseInfo.keys()
                ):
                    nl = list(filter(lambda x: course.startswith(x), classNameList))
                    assert len(nl) == 1, "无法正确解析课程名称"
                    classname = nl[0]
                    course = course.replace(classname, "").strip()
                    # 正则表达式匹配课程信息
                    res = re.match(r"(\w+)? *([单双]?) *((\d+-\d+,?)+)", course)
                    assert res, "课程信息解析异常"
                    # 将课程信息添加到课程信息中
                    info = {
                        "classname": classname,
                        "classtime": [course_time + 1],
                        "day": day + 1,
                        "week": list(filter(None, res.group(3).split(","))),
                        "oe": oeDict.get(res.group(2), 3),
                        "classroom": [res.group(1)],
                    }
                    courseInfo[id] = info
                # 如果课程不为空且在课程信息中，将其添加到课程信息中
                elif course != "\xa0" and id in courseInfo.keys():
                    courseInfo[id]["classtime"].append(course_time + 1)

    # 合并同一课程的不同上课时间
    for course in courseInfo.values():
        purecourse = {key: value for key, value in course.items() if key != "classroom"}
        # 如果课程已经存在，将教室信息添加到课程信息中
        if str(purecourse) in courseList:
            courseList[str(purecourse)]["classroom"].append(course["classroom"][0])
        # 如果课程不存在，将课程信息添加到课程列表中
        else:
            courseList[str(purecourse)] = course
    # 将课程列表转换为课程信息列表
    courseInfoRes = [course for course in courseList.values()]
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
    with open("conf_classTime.json", "r") as f:
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
        fweek = datetime.startData.fromtimestamp(
            int(time.mktime(self.firstWeekDate))
        ) - datetime.timedelta(days=1.0)
        createTime = datetime.datetime.now()
        for _ in range(18):
            sub_prop = {
                "CREATED": createTime,
                "SUMMARY": "学期第 {} 周".format(_ + 1),
                "UID": uuid.uuid4().hex + "@google.com",
                "DTSTART": fweek,
                "DTEND": fweek + datetime.timedelta(days=7.0),
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
            bytes.decode(cal.to_ical(), encoding="utf-8").replace("\r\n", "\n").strip()
        )


# 主函数
if __name__ == "__main__":
    firstWeekDate = None
    classTimeList = None
    courseInfoRes = None
    timeReminder = None

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36"
    }

    userInfo = input("请输入学号和密码，以空格隔开：").split()
    """
    try:
        print("开始获取Cookies...")
        gCookie = loginCookie(userInfo[0], userInfo[1])  # Type: Dict
    except Exception:
        print("遇到错误啦w(ﾟДﾟ)w，请重试")
        sys.exit(0)
    """
    # gCookie = {'ASP.NET_SessionId': 'rc11ki45x4545w3njnbpfbqw'}

    print("开始获取课表...")
    file = open("data.html", encoding="utf-8")
    textDom = file.read()
    file.close()
    if not textDom:
        print("遇到错误啦(´･ω･`)?,请重试")
        sys.exit(0)
    else:
        print("获取课表成功")

    print("开始课表格式化...")
    classHandler(textDom)

    print("正在配置上课时间...")
    setClassTime()

    firstWeekDate = input(
        "请输入此学期第一周的星期一日期(eg 20230904)："
    )  # 周一第一周的开始数据
    print("正在配置第一周周一日期...")
    print("SetFirstWeekDate:", firstWeekDate)

    reminder = input("正在配置提醒功能,请以分钟为单位设定课前提醒时间(默认值为15):")
    print("正在配置课前提醒...")
    setReminder(reminder)

    print("正在生成ics文件...")
    iCal = ICal.withStrDate(firstWeekDate, classTimeList, courseInfoRes)
    with open("./class.ics", "startWeek", encoding="utf-8") as f:
        f.write(iCal.to_ical())
    print("文件保存成功")
