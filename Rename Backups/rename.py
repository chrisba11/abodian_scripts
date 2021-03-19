import os, datetime

date_str = datetime.datetime.now().strftime("%m.%d.%y")

os.rename('Mozaik and Production Files(1)', 'Mozaik and Production Files ' + date_str)
os.rename('Mozaik and Production Files(1).fkc', 'Mozaik and Production Files ' + date_str + '.fkc')

