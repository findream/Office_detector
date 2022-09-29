#coding = utf-8
import re
import common
import xml.etree.ElementTree as ET


def detect_remotemacro(filepath):
    zipper = common.is_contain_targetxmlfile(filepath,"word/_rels/settings.xml.rels")
    if zipper is None:
        common._info("remote macro not found")
        return None
    
    tree = ET.parse(zipper.open("word/_rels/settings.xml.rels"))
    relationships = False
    for elem in tree.iter():
        if elem.tag.split('}')[1] == "Relationships":
            relationships = True
            continue
        if relationships == True:
            if "Target" in elem.attrib.keys():
                if elem.attrib["Target"]:
                    common._info("remote macro \n\t %s" % elem.attrib["Target"])
                    return True
    common._info("remote macro not found \n")
    return None
            
            





def main():
    filepath = "E:\\Example\\DDE\\远程模板注入.rar"
    detect_remotemacro(filepath)

if __name__ == '__main__':
    main()
    
