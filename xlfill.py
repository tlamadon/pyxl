# simple class that opens 
# and edits content of xlsx file
# this will just change values of existing cells
# and not touch the rest (style, formulas)

import zipfile
import xml.etree.ElementTree as ET
import os
import tempfile
import re
import xlutil
import warnings

class XlFill:
  """ Object to open an excel work book and modify it's content

      wb = XlFill('ccew.xlsx')
      wb.open('sheet1')
      wb['A1']
      wb['A1'] = '-48'
      wb.close()
  """
  def __init__ (self,filename):
    self.filename    = filename
    self.zip_content = zipfile.ZipFile(filename, mode='a') 
    content = self.zip_content.namelist()
    self.sheets= set()
    # get the list of worksheets
    for files in self.zip_content.namelist():
      res = re.search('worksheets/(.*)\\.xml',files)
      if (res is not None):
        self.sheets.add(str(res.group(1)))
    self.ws_tree = None
    self.ws_name = None
    self.ns = None
    self.strings = list()
    self.cell_hash = {} # creating a dictionary to point to the cells

  def open(self,name):
    ''' opens worksheet with given name'''
    if (name in self.sheets):
      self.ws_tree = ET.fromstring(self.zip_content.read('xl/worksheets/' + name + '.xml')) 
      self.ws_name = name
      self.ns = self.ws_tree.tag.split('}')[0].replace('{','')

      # caching the cells
      nodes = self.ws_tree.findall(".//{" + self.ns + "}c")
      for n in nodes:
        self.cell_hash[n.attrib['r']]=n

    # get the list of strings
      tmp_tree = ET.fromstring(self.zip_content.read('xl/sharedStrings.xml'))
      nodes = tmp_tree.findall(".//{" + self.ns + "}t")
      self.strings = list()
      for n in nodes:
        self.strings.append(n.text)


  def getXmlCell(self,ref):
    ''' find a cell using the excel reference letter/number
	  '''
    if (ref in self.cell_hash.keys()):
      return(self.cell_hash[ref])
    else:
      return None

  def __getitem__(self, ar):
    ''' get the value of the cell '''
    if type(ar) is str:
      k = ar
    else:
      k = self.coord(ar[0],ar[1])
    node = self.getXmlCell(k)
    if node is not None:
      if (node.attrib['t']=='n'):
        return(node.findtext('{' + self.ns +'}v'))
      elif (node.attrib['t']=='s'):
        return(self.strings[ int(node.findtext('{' + self.ns +'}v')) ] )
    return None

  def __setitem__(self, ar, v):
    ''' rewrite the content of the open sheet to the atchive'''
    if (isinstance(ar,str)):
      k = ar
    else:
      k = self.coord(ar[0],ar[1])
    node = self.getXmlCell(k)
    if node is not None:
      node.find('{' + self.ns +'}v').text = str(v)
    else:
      warnings.warn("cell " + ar +  " does not exist")

  def close(self):
    f = tempfile.NamedTemporaryFile(delete=False,mode='w')
    ET.ElementTree(self.ws_tree).write(f) # need to make a tree before writing
    f.close()
    self.zip_content.write(f.name,arcname = 'xl/worksheets/' + self.ws_name + '.xml')
    self.zip_content.close()
    os.unlink(f.name)

  def coord(self,r,c=None):
    if (isinstance(r,int)):
      return(xlutil.int2letter(c) + str(r+1))



