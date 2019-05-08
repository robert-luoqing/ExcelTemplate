# Excel Template
The lib is binding excel template file and c# object to generate new excel stream or file.

For example we have excel template which include a lot of cell. The lib will parse the cell content and use c# object to replace it.
the format of cell content as below
{{object.propertyName}} or {{object[0].propertyName}}

If you want loop rows using IList object. you need follow below rules:
The first cell of row that need start loop must have {!1-n!} before data of cell
The value surround "{!" and "!}"
"1-n", the "1" mean only 1 row will do loop, the "n" is mean we will replace "n" as index base 0
"2-m", the "2" mean the 2 row will do loop, the "m" is mean we will replace "n" as index base 0
for example, the first cell of row has "2-m", the other cell has path "{{test[m].name}}"
if the test has 3 item, the first two row will be replace as test[0].name, the second will be test[1].name

# Code Sample
```c#
var myClass = new Class();
myClass.ClassName = "Class Four";
myClass.Teacher = "Robert.R";
myClass.Students = new List<Student>();
myClass.Students.Add(new Student()
{
    Name="Sharon",
    Age=12,
    Gender="Female",
    Score = 70
});

myClass.Students.Add(new Student()
{
    Name = "Robert",
    Age = 13,
    Gender = "Male",
    Score = 65
});

var installExecuteFile = Assembly.GetExecutingAssembly().Location;
var fileInfo = new FileInfo(installExecuteFile);
var installPath = fileInfo.Directory.FullName;

var stream = ExcelTemplateHelper.HandleExcel(installPath+@"\ExcelSimple\Simple.xlsx", myClass);
using (var fileStream = File.Create(installPath + @"\ExcelSimple\1.xlsx"))
{
    stream.Seek(0, SeekOrigin.Begin);
    stream.CopyTo(fileStream);
}
```
#Excel Template
![alt text](https://raw.githubusercontent.com/robert-luoqing/ExcelTemplate/master/Images/excel-template-source.png)
#Convert Excel
![alt text](https://raw.githubusercontent.com/robert-luoqing/ExcelTemplate/master/Images/excel-template-dist.png)