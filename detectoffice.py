'''
detectoffice:detect office file
   [*] dynamic data exchange 
   [*] remate macro
   [*] rtf object

author:HaCky

version:detectoffice V0.1

python version:3.x
'''
# coding = utf-8
import sys
import common
import detectdde
import detectremotemacro
import rtfparse


def main():
    # get file type
    if len(sys.argv) != 2:
        common._error("argv error:[python detectoffice.py officefilename]")
        exit()
    
    filepath = sys.argv[1]
    #filepath = "E:\\Example\\rtf样本\\9dc4d8eb8243e48218668dbfc4565a893e74a25c23d2bb38b720a282c24fbe02"
    filetype = common.get_file_type(filepath)

    # detecttargetmacro
    if filetype == common.FILETYPE_DOCX:
        detectremotemacro.detect_remotemacro(filepath)

    # detect rtf
    if filetype == common.FILETYPE_RTF:
        if rtfparse.detect_rtf(filepath) == True:
            common._info("suspicious rtf object")

    # detectdde
    DetectDDE = detectdde.DETECTDDE(filepath,filetype)
    DetectDDE.detect_dde()

    # detectmacro



if __name__ == '__main__':
    main()
    