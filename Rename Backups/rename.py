import os, datetime

yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
date_str = yesterday.strftime("%m.%d.%y")

os.rename('D:\Backups\Mozaik and Production Files(1)', 'D:\Backups\Mozaik and Production Files ' + date_str)
os.rename('D:\Backups\Mozaik and Production Files(1).fkc', 'D:\Backups\Mozaik and Production Files ' + date_str + '.fkc')

