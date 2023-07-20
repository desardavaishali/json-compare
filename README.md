# json-compare
 Compare 2 json files and lists differences in excel. </n></n>

<b>What:</B>
 This is a javascript code to compare 2 json files and list the differences in an excel file.
 
</n>

 <b>How: </b>
 1. Download all the files
 2. Install node with npm install node command & dependancies - npm install exceljs xlsx
 3. Open command prompt navigate to the directory where "json-compare-excel.js" file is located.
 4. Run the command <i><b> node json-compare-excel.js [file1] [file2] [filter.txt] </b></i>
   <i> file1 </i>- Main file </n></n>
   <i> file2 </i>- file to be compared against file1</n></n>
   <i> filter.txt </i>- is an optional field. </n></n>
    If given, the output will exclude the fields mentioned in filter.txt</n></n>
 5. Excel file is generated in the same folder as the code with the filename same as [file2]. 
