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

