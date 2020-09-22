Graph Any Data

Summary

Will graph almost any data and limits that you supply
Can review each point and display the data in excel.
Also calculates summary data.

This program will graph almost any data. The graph
will display limits that you supply, let you review
each point and can copy the data into excel. The graph 
will change to accommodate the values you send to it. 
The graph will auto scale and is has red, yellow
and green sections to show what data is outside your limits.
The graph also calculates summary data (Ave, Range and SD) 
The form can easily be added to any project.

How the program works
The program uses two classes to store the graph information.
This program will graph 40 points and expects that you have
created the ClsData in the same order as you want to have the
points plotted

ClsData - is used to store all of your data
   
  Clsdata.YValue     = The y value to be ploted
  Clsdata.XValue     = Calculated for each point (point location * 3)
  Clsdata.PDate      = A place to store the date
  Clsdata.PName      = A place to store a point name
  Clsdata.PError     = Calculated for each point (above UCL, UWL or below LCL, LWL)
  Clsdata.PErrorType = Calculated for each point (blank,Warning or Error)
  Clsdata.PColor     = Calculated for each point (vbWhite =no error  Yellow= Outside WL   RED= Outside CL)
  
ClsInfo - is used to draw Limits
 
  ClsInfo.InfoName = TAR    (the chart target)
  ClsInfo.InfoName = UCL    (the chart UCL)
  ClsInfo.InfoName = LCL    (the chart LCL)
  ClsInfo.InfoName = UWL    (the chart UWL)
  ClsInfo.InfoName = LWL    (the chart LWL)
  ClsInfo.InfoName = MAX    (the chart Y MAX)
  ClsInfo.InfoName = MIN    (the chart Y MIN)
  ClsInfo.InfoName = INC    (consecutive points increasing or decreasing)
  ClsInfo.InfoName = ABOVE  (consecutive points above or below warning)
  ClsInfo.InfoName = Title  (the chart TITLE)
  
  ClsInfo.YValue = The y value of the limit 

  ClsInfo.YValue = for Title = the title of the graph

  ClsInfo.YValue = for INC or ABOVE = the number of points


Copy Button
The copy button will allow display the data in Excel.

Viewing data

As you move the mouse across the graph, the point nearest to
cursor will turn white and the information about the point 
will be displayed at the bottom of the screen.
You can also hit the Previous or Next button to go
from point to point and the information about the point 
will be displayed at the bottom of the screen.
   
As you change the values along the side of the
chart, you can generate different graphs.