from glob import glob
import shutil
import os
import warnings
warnings.filterwarnings(action='ignore',category=UserWarning,module='gensim')
import win32com
import win32con
import win32gui
import codecs
from win32com.client import Dispatch
import pythoncom
import os
import sys


dataset_loca=sys.argv[1]
dataset_loca3=sys.argv[2]+'/'

class MSOffice2txt():
  def __init__(self, fileType=['doc','ppt']):
    self.docCom = None
    self.pptCom = None
    pythoncom.CoInitialize()
    if type(fileType) is not list:
      return 'Error, please check the fileType, it must be list[]'
    for ft in fileType:
      if ft == 'doc':
        self.docCom = self.docApplicationOpen()
  def close(self):
    self.docApplicationClose(self.docCom)
 
  def docApplicationOpen(self):
    docCom = win32com.client.Dispatch('Word.Application')
    docCom.Visible = 1
    docCom.DisplayAlerts = 0
    docHwnd = win32gui.FindWindow(None, 'Microsoft Word')
    win32gui.ShowWindow(docHwnd, win32con.SW_HIDE)
    return docCom
 
  def docApplicationClose(self,docCom):
    if docCom is not None:
      docCom.Quit()
 
  def doc2Txt(self, docCom, docFile, txtFile):
    doc = docCom.Documents.Open(FileName=docFile,ReadOnly=1)
    doc.SaveAs(txtFile, 2)
    doc.Close()
 
 
    

  def translate(self, filename, txtFilename):
    if filename.endswith('doc') or filename.endswith('docx'):
      if self.docCom is None:
        self.docCom = self.docApplicationOpen()
      self.doc2Txt(self.docCom, filename, txtFilename)
      return True
    else:
      return False


files = glob(dataset_loca) 
oldfile=[]
count=0
msoffice = MSOffice2txt()
for file_name in files:
  tmpp=file_name
  tmpp=tmpp.split('\\',2)[-1]
  tmpp=tmpp.split('.',1)[0]
  tmpp=tmpp.replace(' ','_')
  shutil.copyfile(file_name,dataset_loca3+tmpp+'.doc')
  count=count+1
  msoffice.translate(dataset_loca3+tmpp+'.doc', dataset_loca3+tmpp+'.txt')
 


 
