# @todo
# I should implement a better indexing system. I should be able to create
# an excel range in mnay ways like "A1B5" or A1, (4,5) or else, and then
# shift it, iterate through it, etc....

from pandas.io.parsers import ExcelFile,DataFrame
import numpy as np
from sets import Set
import re
import itertools

def sord(s):
  val = 0
  for i in range(0,len(s)):
	val = val*26
	val += ord(s[i]) - 64
  return(val-1)

def getExcelChunck(file,ws,drange,rid=-1,cid=-1):
  xls = ExcelFile(file)
  df = xls.parse(ws)
  
  # get the range from expression
  # B4H4
  m  = re.search('([A-Z]+)([0-9]+)([A-Z]+)([0-9]+)', drange)
  c1 = sord(m.group(1)) 
  c2 = sord(m.group(3))+1
  r1 = int(m.group(2))-2
  r2 = int(m.group(4))-2
  
  df2 = df.ix[r1:r2,c1:c2]
  
  if (rid>=0):
    rh = int(rid) -2
    df2.columns = df.ix[rh,c1:c2]
    df2.columns = df2.columns.map(lambda x: str(x).strip().replace('.0',''))
  
  if (cid>=0):  
    ch = sord(cid)
    df2.index   = df.ix[r1:r2,ch]
    df2.index = df2.index.map(lambda x: str(x).strip().replace('.0',''))   
  return(df2)

def fillArrayXl(wb,drange,ri,ci,df):
  # first we get the array
  m  = re.search('([A-Z]+)([0-9]+)([A-Z]+)([0-9]+)', drange)
  c1 = sord(m.group(1)) 
  c2 = sord(m.group(3))
  r1 = int(m.group(2))-1
  r2 = int(m.group(4))-1
  # we start in the corner, we loop over height and 
  # width, every time we get the keys and fill in the value
  for ix in range(r1,r2+1):
    for iy in range(c1,c2+1):
      # get the col key from excle
      col_key = wb[ri,iy]
      row_key = wb[ix,ci]
      # make sure key exists
      if (col_key in df.columns) & (row_key in df.index):
        wb[ix,iy] = str(df.ix[row_key,col_key])
      else:
        print "could not find " + row_key + ":" + col_key
  
# iterates through the cells
# in the given drange for example 'A1D4'
def iterxl(drange,num=False):
  # first we extract the bounds
  m  = re.search('([A-Z]+)([0-9]+)([A-Z]+)([0-9]+)', drange)
  c1 = sord(m.group(1)) 
  c2 = sord(m.group(3))
  r1 = int(m.group(2))
  r2 = int(m.group(4))

  if num:
    arange = range(c1,c2+1)
    brange = range(r1-1,r2)
    return itertools.product(brange,arange)

  # then we create the ranges
  arange = map(lambda x: int2letter(x)  , range(c1,c2+1))
  brange = map(lambda x: str(x)         , range(r1,r2+1))
  return itertools.product(arange,brange)

def coord2int(s):
  m  = re.search('([A-Z]+)([0-9]+)', s)
  c1 = sord(m.group(1)) 
  r1 = int(m.group(2))-1
  return r1,c1
 
def letter2int(s):
  val = 0
  for i in range(0,len(s)):
    val = val*26
    val += ord(s[i]) - 64
  return(val-1)
def int2letter(n):
  if (n==0):
    return('A')
  val = ''
  while n>=0:
    nr = n % 26;
    val = chr(65+nr) + val
    n = (n - nr)/26 - 1
  return(val)

def proj(val,start,step):
  rem = (val - start)
  return rem - (rem % step) + start

class XLCell(object):
  def __init__(self, ar='A1'):
    ''' the cell is defined by either a couple coordinates with two integers or by a string with a letter followed by a number'''
    if type(ar) is str:
      self.x,self.y = coord2int(ar)
    else:
      self.x = ar[0]
      self.y=  ar[1]
  def setCol(self,y):
    if type(y) is str:
      self.y = letter2int(y)
    else:
      self.y = y
  def setRow(self,x):
    self.x = x
  def shift(self,x,y):
    tmp = self.copy()
    tmp.x+=x
    tmp.y+=y
    assert tmp.x>=0
    assert tmp.y>=0
    return tmp
  def spreadBy(self,x,y):
    pass
  def spreadTo(self,x,y):
    pass
  def str(self):
    return int2letter(self.y) + str(self.x+1)
  def copy(self):
    return(XLCell(self.str()))
  def __str__(self):
    return int2letter(self.y) + str(self.x+1)

class XLRange(object):
  """creates excel ranges from either """
  def __init__(self, cell_tl,cell_br):
    self.cell_tl = cell_tl.copy()
    self.cell_br = cell_br.copy()

  def __str__(self):
    return self.cell_tl.str() + self.cell_br.str()

