# *_*coding:utf-8 *_*
#导入excel文件将文字写入图片

from PIL import Image, ImageDraw, ImageFont
import xlrd


def readExcel(filename):
    df = xlrd.open_workbook(filename)
    table = df.sheet_by_index(0)
    nrow = table.nrows
    label1 = table.row_values(1)[0] #标签名
    label2 = table.row_values(1)[1] #标签名
    elist = []
    for i in range(nrow-2):
        elist.append({label1:int(table.row_values(i+2)[0]),label2:table.row_values(i+2)[1].encode('utf-8')})

    return elist


def writeFont(id,txt):
    #print id,txt
    filename = "img/%d.jpg"%id
    #print filename
    font_img = Image.open(filename)
    fw,fh = font_img.size
    draw = ImageDraw.Draw(font_img)

    draw.rectangle((0,fh-70,fw,fh-10),'yellow')
    ttfront = ImageFont.truetype("FZSEJW.TTF",55,encoding='unic')
    tw,th = ttfront.getsize(txt)
    posx = (fw-tw/3)/2
    draw.text((posx,fh-70),unicode(txt,'utf-8'),font=ttfront,fill="#ff0000")

    font_img.save("out/%d.jpg"%id)

def main(exfile):
    diclist = readExcel(exfile)
    for i in diclist:
        writeFont(i['imgId'],i['desc'])

    print '全部写入完成'

print "注意需先输入excel文件路径，以字符串的形式（如'文件路径'）"
exfile = input("请输入excel文件路径：")
print exfile
main(exfile)
input('Press Enter to exit')