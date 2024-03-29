#coding =  utf-8
import zipfile
import ooxml
import olefile
from oletools.xls_parser import is_xls
from oletools.ppt_record_parser import is_ppt 
from oletools.olevba import VBA_Parser
#import xls_parser
#import ppt_record_parser
import logger
import os
import sys

FILETYPE_XLS = "xls"
FILETYPE_PPT = "ppt"
FILETYPE_DOC = "doc"
FILETYPE_RTF = "rtf"
FILETYPE_XLSX = "xlsx"
FILETYPE_PPTX = "pptx"
FILETYPE_DOCX = "docx"

# contain target xml file
def is_contain_targetxmlfile(filename,xmlfilepath):
    if not zipfile.is_zipfile(filename):
        _error("file is not a zip")
        return None

    zipper = zipfile.ZipFile(filename)
    subfiles = zipper.namelist()
    if subfiles.__contains__(xmlfilepath):
        return zipper
    zipper.close()
    return None

def is_contain_targetdir(filename,targetdir):
    if not zipfile.is_zipfile(filename):
        _error("file is not a zip")
        return None
    zipper = zipfile.ZipFile(filename)
    subfiles = zipper.namelist()
    for subfile in subfiles:
        if targetdir in subfile:
            zipper.extract(subfile,"C:/Users/Public/",None)
            zipper.close()
            print(os.path.join("C:/Users/Public/",subfile))
            return os.path.join("C:/Users/Public/",subfile)
    return None



# print error
def _error(err_string):
    print("[!] %s" % err_string)


# print info
def _info(info_string):
    print("[*] %s" % info_string)


# get the file type of target file
def get_file_type(filepath):
    # ole xml?
    if olefile.isOleFile(filepath):
        # xls
        if is_xls(filepath):
            return FILETYPE_XLS
        # ppt
        if is_ppt(filepath):
            return FILETYPE_PPT
        # doc
        else:
            return FILETYPE_DOC
    
    # rtf
    with open(filepath,"rb") as fp:
        if fp.read(4) == b'{\\rt':
            return FILETYPE_RTF
        
    # docx xlsx pptx
    try:
        doctype = ooxml.get_type(filepath)
    except Exception:
        return None

    if doctype == ooxml.DOCTYPE_EXCEL:
        return FILETYPE_XLSX
    # docx
    if doctype in (ooxml.DOCTYPE_WORD,ooxml.DOCTYPE_WORD_XML,ooxml.DOCTYPE_WORD_XML2003):
        return FILETYPE_DOCX
    # pptx
    if doctype in (ooxml.DOCTYPE_POWERPOINT,ooxml.DOCTYPE_MIXED):
        return FILETYPE_PPTX

def isVbaStomping(filename):
    flag = False
    vbaParser = None
    try:
        vbaParser = VBA_Parser(filename)
        if vbaParser.ole_file is None:
            for subFile in vbaParser.ole_subfiles:
                vbaProjects = subFile.find_vba_projects()
                if vbaProjects is not None:
                    flag = True
        else:
            vbaProjects = vbaParser.find_vba_projects()
            if vbaProjects is not None:
                flag = True
    except Exception as e:
        print('Error: {}.'.format(e), file=sys.stderr)
    return flag


def main():
    pass


if __name__ == '__main__':
    main()
    