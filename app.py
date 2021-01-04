from dimigo import document, cube, office, exit
from datetime import date
import os

seoryu = document.get('자퇴')
seoryu.set(document.fields.NAME, os.getenv('name'))
seoryu.set(document.fields.DATE, date.today())
seoryu.set(document.fields.REASON, "탈디미 너무달아")

jatweseo = cube.print(seoryu)

if office.request(jatweseo) == office.response.accepted:
    print("자퇴성공! 탈디미!")
    exit()
else:
    print("자퇴실패")
    life.exit()
