# xlsximport
**xlsximport is a [SQLite extension](https://sqlite.org/loadext.html) that allows direct importing of XLSX Excel 2007 files (a.k.a. Open Office XML)**

### FEATURES:
* This extension defines the `importxlsx()` [application-defined SQL function](https://sqlite.org/appfunc.html) to import XLSX Excel 2007 files (a.k.a. Open Office XML) data into an SQLite table. 
* Only depends on [zipfile SQLite extension](https://www.sqlite.org/zipfile.html) and [expat](http://expat.sourceforge.net/).

The XLSX format is just a glorified ZIP (that I open thanks to [zipfile](https://www.sqlite.org/zipfile.html)), containing a set of XML files (that I parse thanks to [expat](http://expat.sourceforge.net/)). The [SQLite shell](https://www.sqlite.org/cli.html) already imports/exports a variety of text formats.
The direct importing of XLSX files removes the need of intermediate XLSX to CSV converters (like [cxlsx_to_csv](https://github.com/vpaesa/cxlsx_to_csv)).

### DATES BEHAVIOUR:
Notice that Excel stores dates as the number of days that have elapsed since the Excel Epoch (1-January-1900 for Windows Excel, 2-January-1904 for Mac Excel), and this extension imports dates simply as the floating point value they are stored. Be aware the dates in SQLite are stored as [Julian days](https://en.wikipedia.org/wiki/Julian_date) since another Epoch (noon Jan 1st, 4713 BC).

After importing, you can perform by yourself the conversion you need by using simple SQL UPDATE commands: 
```
  UPDATE some_table SET some_date = some_date + 2415018.5; -- Convert from WinExcel to Julian Days
  UPDATE some_table SET some_date = some_date + 2416480.5; -- Convert from MacExcel to Julian Days
```

### USAGE:
```
.load importxlsx

select importxlsx(FILE, TABLE SHEET);

--  FILE  input spreadsheet in XLSX Excel 2007 format
--  TABLE name of the table where data is imported
--  SHEET number of the sheet within the workbook (default is first one)
```
### importxlsx():

The `importxlsx()` function takes two (optionally three) arguments which are: the name of the disk file from which XLSX data is to be read, the name of the SQLite table into which the XLSX data is to be inserted, and (optionally) the number of the sheet being imported.

Example:
```
.load importxlsx
select importxlsx('C:/work/somedata.xlsx', 'tab1');
select importxlsx('C:/work/somedata.xlsx', 'tab2', 2);
```
There are two cases to consider: (1) Table `tab1` does not previously exist and (2) table `tab1` does already exist.

In the first case, when the table does not previously exist, the table is automatically created and the content of the first row of the input XLSX file is used to determine the name of all the columns in the table. In other words, if the table does not previously exist, the first row of the XLSX file is interpreted to be column names and the actual data starts on the second row of the XLSX file.

For the second case, when the table already exists, every row of the XLSX file, including the first row, is assumed to be actual content. If the XLSX file contains an initial row of column labels, that row will be read as data and inserted into the table. To avoid this, make sure that table does not previously exist.

### COMPILATION:
`cc -shared -lexpat xlsximport.c`

### MEMORY USAGE:
Be aware the XML data inside your .XLSX file is fully decompressed in memory. The command `unzip -l` will give you a good estimate of the memory needed.
So for example this large 19MB spreadsheet will use temporally about 135MB of memory:
```
$ ls -ks owid-covid-data.xlsx
19228 owid-covid-data.xlsx

$ unzip -l owid-covid-data.xlsx
Archive:  owid-covid-data.xlsx
  Length      Date    Time    Name
---------  ---------- -----   ----
     1167  01-31-1980 00:00   [Content_Types].xml
      587  01-31-1980 00:00   _rels/.rels
      697  01-31-1980 00:00   xl/_rels/workbook.xml.rels
135734695  01-31-1980 00:00   xl/worksheets/sheet1.xml
      550  01-31-1980 00:00   xl/workbook.xml
    28949  01-31-1980 00:00   xl/sharedStrings.xml
     1341  01-31-1980 00:00   xl/styles.xml
     6994  01-31-1980 00:00   xl/theme/theme1.xml
      592  01-31-1980 00:00   docProps/core.xml
      784  01-31-1980 00:00   docProps/app.xml
---------                     -------
135776356                     10 files
```

### LICENSE:
* My importxlsx code is Public Domain. Same as [SQLite](https://www.sqlite.org/). Notice though that [Expat](http://expat.sourceforge.net/) license is MIT.
