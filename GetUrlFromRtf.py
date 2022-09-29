from asyncio.windows_events import NULL
import re
import sys

def main(file):
    if file == '':
        print("usage: GetUrlFromRtf.py filename")
    url = ''
    with open(file,'r',encoding='UTF-8') as f:
        datas = f.read()

    patter = re.compile(r'u-65(.*?)?}')  #解析不到第一个，所以手动添加\\u-65
    tmpioc = re.findall(patter,datas)
    if tmpioc == []:
        return False
    ioc = '\\u-65' + str(tmpioc)

    try:
        ioc = ioc.replace("['",'').replace("']",'').replace('?\\\\u-',',')
        ioc = ioc.replace('\\u-','').replace('?','')
        ioc = ioc.split(',')
        for i in ioc:
            if i != '':
                dt = int(i)
                url += chr(65536 - dt)
        print(url)
    except:
        return False
    return True







if __name__ == "__main__":
    file = sys.argv[1]
    #file = "eee"
    main(file)



