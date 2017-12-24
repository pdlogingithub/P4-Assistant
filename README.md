# P4-Assistant
Standalone application based on python and pyqt,for quering information from perforce server,just like "Find files" in p4v,but with more control and can export as xls file.

# usage  
Main ui:  
![1](ui.png)   

Input connection information,if you are not querying informations specific to user or workspace then you only need to type in server port.  
Press "save connection" will save current connection information to text file and program will read this file next you lauch.  
![2](run.png)  

After pressing "run" button,the results will show in the bottom section,you can sort them by clicking on the any colum titile.  
![3](result.png)   

In the filter section,youcan aplly filter setting,add "-" before keyword for exclude filter,use ";" to seperate multiple keywords,filter will be applied in order.filter can also be saved  
![4](filter.png)   

Filter results:  
![5](filter_result.png)   

You can export the result to xls file and edit them in excel for more advanced control,you can also load in a existing one.
![6](xls.png)   
