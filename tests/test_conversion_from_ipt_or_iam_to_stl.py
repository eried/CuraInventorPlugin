# -*- coding: utf-8 -*-

'''
Created on 07.11.2017

@author: Thomas Pietrowski
@api: http://help.autodesk.com/view/INVNTOR/2018/ENU/?guid=GUID-AA811AF0-2494-4574-8C43-4C22E608252F
'''

import os

import ctypes

import win32com
win32com.__gen_path__ = os.path.join(os.path.split(__file__)[0], "gen_dir") 

import win32com.client
import pythoncom


#ThisApplication = win32com.client.gencache.EnsureDispatch("Inventor.Application")
try:
    ThisApplication = win32com.client.GetActiveObject("Inventor.Application")
except:
    ThisApplication = win32com.client.Dispatch("Inventor.Application")
    #ThisApplication.Visible=True

#foreign_filename = "C:\\Users\\t.pietrowskie\\AppData\\Roaming\\cura\\3.0\\plugins\\CuraInventorPlugin\\CuraInventorPlugin\\test\\Inventor Professional 2018\\test_cube.ipt"
foreign_filename = "C:\\Users\\t.pietrowskie\\AppData\\Roaming\\cura\\3.0\\plugins\\CuraInventorPlugin\\CuraInventorPlugin\\test\\Inventor Professional 2018\\test_cube.iam"

#Document = ThisApplication.ActiveDocument

# Open files
open_documents = []
if ThisApplication.Documents.Count:
    for i in range(ThisApplication.Documents.Count):
        open_document = ThisApplication.Documents.Item(i+1)
        print(open_document.FullFileName)
        open_documents.append(open_document)

if foreign_filename not in open_documents:
    document = ThisApplication.Documents.Open(foreign_filename)
    document_opened = True
else:
    document_opened = False


STLTranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{533E9A98-FC3B-11D4-8E7E-0010B541CD80}")
Context = ThisApplication.TransientObjects.CreateTranslationContext()
Options = ThisApplication.TransientObjects.CreateNameValueMap()
#    Save Copy As Options:
#       Name Value Map:
#               ExportUnits = 4
#               Resolution = 1
#               AllowMoveMeshNode = False
#               SurfaceDeviation = 60
#               NormalDeviation = 14
#               MaxEdgeLength = 100
#               AspectRatio = 40
#               ExportFileStructure = 0
#               OutputFileType = 0
#               ExportColor = True

if STLTranslatorAddIn.HasSaveCopyAsOptions(document, Context, Options):
    # Set accuracy.
    #   2 = High,  1 = Medium,  0 = Low
    # was 4
    
    # http://help.autodesk.com/view/INVNTOR/2018/ENU/?guid=GUID-5FDFF606-1D15-4FA0-9ED1-1BF4A3BCEBF8
    Options.Remove("Resolution")
    Options.Insert("Resolution", 2)
    
    # Set output file type:
    #   0 - binary,  1 - ASCII
    Options.Remove("OutputFileType")
    Options.Insert("OutputFileType", 1)

    # IOMechanismEnum Enumerator - http://help.autodesk.com/view/INVNTOR/2018/ENU/?guid=GUID-A3660CD6-8B11-48CE-9FA5-E51DCC6F8DEB
    Context.Type = 13059 #kFileBrowseIOMechanism

    Data = ThisApplication.TransientObjects.CreateDataMedium()
    filedir, filename = os.path.split(foreign_filename)
    # if "." in filename
    real_filename = os.path.splitext(filename)[0]
    stl_filename = real_filename + ".stl"
    stl_fullpath = os.path.join(filedir, stl_filename)
    Data.FileName = stl_fullpath
    print("saved as: {}".format(stl_fullpath))

    STLTranslatorAddIn.SaveCopyAs(document, Context, Options, Data)

if document_opened:
    document.Close(True)