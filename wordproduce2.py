# -*- coding: UTF-8 -*-
import win32com
import os
from win32com.client import Dispatch, constants

class WordWrap:

  def __init__(self,templatefile=None):
    self.wordApp = Dispatch('Word.Application')
    if templatefile == None:
      self.wordDoc = self.wordApp.Documents.Add()
    else:
      self.wordDoc = self.wordApp.Documents.Add(Template=templatefile)

  def quit(self):
      self.wordApp.ActiveDocument.Close()
      #self.wordApp.Quit()

  def saveAs(self,filename,delete_existing=True):
    if delete_existing and os.path.exists(filename):
         os.remove(filename)
    self.wordApp.ActiveDocument.SaveAs(FileName=filename)

  def textReplace(self,oldStr,newStr):
    find = self.wordApp.Selection.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Execute(oldStr, False, False, False, False, False, True, 1, True, newStr, 2)
