#from unicodedata import decimal
from plugin_biff import cBIFF
import common
import zipfile
import oletools.olevba as olevba

def detect_excel4_0_ole(filename):
    #from oletools.thirdparty.oledump.plugin_biff import cBIFF
    # plugin_biff plugin of new version oletools can except
    flag_excel4_0 = False
    vba_parser = olevba.VBA_Parser(filename, data=None, container=None,relaxed=False)
    xlm_macros = []
    if vba_parser.ole_file is None:
        return False
    for excel_stream in ('Workbook', 'Book'):
        if vba_parser.ole_file.exists(excel_stream):
            data = vba_parser.ole_file.openstream(excel_stream).read()
            try:
                biff_plugin = cBIFF(name=[excel_stream], stream=data, options='-x')
                xlm_macros = biff_plugin.Analyze()
                if len(xlm_macros)>0:
                    for eachsheet in xlm_macros:
                        if "Excel 4.0 macro" in eachsheet:
                            flag_excel4_0 = True
                            break
            except:
                common._error("detect_excel4_0_ole")
        if flag_excel4_0 == True:
            break
    return flag_excel4_0

def detect_excel4_0_xlm(filename):
    filepath = filename
    flag_excel4_0 = False
    if not zipfile.is_zipfile(filepath):
        return False
    zipper = zipfile.ZipFile(filepath)
    for subfilename in zipper.namelist():
        if "xl/macrosheets" in subfilename:
            flag_excel4_0 = True
            break
    return flag_excel4_0
        




def detect_excel4_0(filename):
    # ole
    flag_excel4 = detect_excel4_0_ole(filename)
    # xlm
    if flag_excel4 == False:
        flag_excel4 = detect_excel4_0_xlm(filename)
    return flag_excel4


def main():
    pass

if __name__ == "__main__":
    main()