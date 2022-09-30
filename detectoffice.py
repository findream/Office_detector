'''
detectoffice:detect office file
   [*] dynamic data exchange 
   [*] remate macro
   [*] rtf object
   [*] vba stomping
   [*] Excel 4.0

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
import GetUrlFromRtf
import pcode2code
import excel4_0


def main():
    # get file type
    if len(sys.argv) != 2:
        common._error("argv error:[python detectoffice.py officefilename]")
        exit()
    
    filepath = sys.argv[1]
    #filepath = "E:\\Example\\rtf样本\\9dc4d8eb8243e48218668dbfc4565a893e74a25c23d2bb38b720a282c24fbe02"
    filetype = common.get_file_type(filepath)

    # detecttargetmacro
    if filetype == common.FILETYPE_DOCX or filetype == common.FILETYPE_XLSX or filetype == common.FILETYPE_PPTX:
        print("[+] -> remotemacro:")
        detectremotemacro.detect_remotemacro(filepath)

    # detect rtf
    if filetype == common.FILETYPE_RTF:
        print("[+] -> rtf object:")
        if rtfparse.detect_rtf(filepath) == True:
            common._info("[*] suspicious rtf object")
            # by tianfeng@360.cn
        print("[+] url in rtf:")
        GetUrlFromRtf.main(filepath)

    # detectdde
    print("[+] -> dde:")
    DetectDDE = detectdde.DETECTDDE(filepath,filetype)
    DetectDDE.detect_dde()

    # TODO:detectmacro

    # vba stomping
    print("[+] -> Vba Stomping:")
    if common.isVbaStomping(filepath) == True:
        try:
            pcode2code.process(filepath)
        except:
            common._error("pcode2code.process except")


    # Excel4.0
    print("[+] -> Excel 4.0")
    if excel4_0.detect_excel4_0(filepath) == True:
        common._info("may be Excel 4.0")
    else:
        common._info("may be not Excel 4.0")




if __name__ == '__main__':
    main()
    