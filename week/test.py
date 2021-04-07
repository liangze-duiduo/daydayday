import datetime
titleTime = (datetime.datetime.now() - datetime.timedelta(days = 4)).strftime("%Y-%m-%d") + '——' + datetime.datetime.now().strftime("%Y-%m-%d")
print(titleTime)