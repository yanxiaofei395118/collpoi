#coding=utf-8

import types
import urllib2
import json

# 利用urllib2抓取网络数据
def registerUrl(url):
    try:
        data = urllib2.urlopen(url).read()
        return data
    except Exception, e:
        print e

def initJsonFile():
    try:
        page = 0 # 起始页码
        url = "http://117.131.136.42:8099/TdJudiciary/simpleSearch.tdt?data.companyName=&data.typeId=&data.areaId=&page=" + str(page)
        data = registerUrl(url)
        jsondata = json.loads(data)

        # 获取总页数
        totalSize = jsondata['obj']['totalSize'] / 10 + 2

        for pIndex in range(page, totalSize):
            url = "http://117.131.136.42:8099/TdJudiciary/simpleSearch.tdt?data.companyName=&data.typeId=&data.areaId=&page=" + str(pIndex)
            print url
            data = registerUrl(url)
            print data
            file = open("E:\\tjsfjg\\"+str(pIndex)+".json", "w")
            file.write(data)
            file.close()
    except Exception, e:
        print e

if __name__ == "__main__":
    initJsonFile()
