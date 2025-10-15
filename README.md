# Gsheet-importrange-by-appscript

Copy data from one google sheet to another App Script without fail #ref error.

#Steps
1. Open a google sheet.
2. Goto tools>app script
3. Add script id to add library section : v3 **1IV7T2BPRuka-NIgvqh0ksSA_f9HrlPy_dmKZnoF9QlcIv6bx6D19CG5A**
                                          v4 AKfycbzVPQcET6-t_5UleB4woKZiQkBVDr-6y-b7hfXbO8efbfWnFZauNZc_eSN1YfuKAyOh
5. Embed sheets & namerange location similarly to Sample excel or Set manually variables of sheet ids & their name ranges
6. Run code.

v3: [https://script.google.com/macros/library/d/1IV7T2BPRuka-NIgvqh0ksSA_f9HrlPy_dmKZnoF9QlcIv6bx6D19CG5A/1](url)
v4: https://script.google.com/macros/library/d/1IV7T2BPRuka-NIgvqh0ksSA_f9HrlPy_dmKZnoF9QlcIv6bx6D19CG5A/4

Use function as per requirement.
1. Data by Gsheet ID:
2. Data by Gsheet name:
3. Mail excel files write:
4. Mail CSV file:
                function importSheet2Sheet(sourceID,sourceRange, destinationID, destinationRange,how ,filterColumnNum,filterBy,filterKey)
5. Get data:
                function getdata(sourceID,sourceRange)
6. Data by filtering:
                function filterbykey(data,filterColumnNum,filterBy,filterKey)
7. Write data by Range:
                 function exportData2Range(destinationID, destinationRange,data,how)
   
 * destinationID - Sheet id of Destination file on where you want to copy.
 * destinationRange - Destination data namedRange or range start reference ex.'Sheet1!A1:F17'.
 * data - Source data to copy as arrey not blob.
 * how - How you want to copy valid entries are "Replace" or "Append"
 * sourceID - Sheet id of source source file from where you wan to copy.
 * sourceRange - Source data namedRange or range. 
