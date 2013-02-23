pyxl
====

Python class to open and edit an existing excel file. The class is very basic but it does the job. In particular it
keeps the existing style and formulas of the file which makes it convenient to seperate the designing of the document
from the filling of the actual values

here is a simple example on filling in an array of values:

      import xlutil,xlFill

      # open excel file example.xlsx sheet 1
      wb = XlFill('example.xlsx')
      wb.open('sheet1')

      # using the iterator from xlutil, go through
      # the cells of the array defined by P9V9
      for (r,c) in xlutil.iterxl('P9V9',num=True):

        year = wb[5,c]          # extract the value in columm c, row 5
        wb[r,c] = str(year + r) # fill the value in the current cell

      wb.close() # write the document back to disk

I think you get the picture.

Catches:

 - it seems to be an error in getting a cell when the first lines of the worksheet are completely empty. 
 - you have to refer to worksheet by number sheet1, sheet2 .... not by their actual name, weird.        

