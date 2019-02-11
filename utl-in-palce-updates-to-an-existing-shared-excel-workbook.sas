In palce updates to an existing shared excel workbook

Alothough you don't need a primary keys, inplace modifications
of excel works best with a primary key.

I kept the workbook open while updating from SAS.
When I closed and then opend the workbook alfreds age was
changed to '99'.

github
https://tinyurl.com/y5qq56m7
https://github.com/rogerjdeangelis/utl-in-palce-updates-to-an-existing-shared-excel-workbook

other excel repositories
https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=

Problem: Change Alfreds age to 99 in named range class Multiple solutions
(you need classic 64bit SAS for this?) Updating named ranged cells in Excel using SAS 9.4

Problem: Change Alfreds age to 99 in named range class


INPUT  Shared Workbook
======================

    %utlfkil(d:/xls/class.xlsx);

    libname xel "d:/xls/class.xlsx";
    data xel.class;
     set sashelp.class;
    run;quit;
    libname xel clear;

    Click on Review->Share Workbook->Tick the box next to
    'Allow changes by more than one user at the same time.
    This also allows workbook merging'


    d:/xls/class.xlsx  with sheet name class and named range class


    WORKBOOK d:/xls/class.xlsx with sheet class (you can use [sheet1])

   d:/xls/class.xlsx
      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
   2  | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
   3  | BARBARA    |    F       |    13      |    58      |  101.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
   20 | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+

   [CLASS]


EXAMPLE OUTPUT
--------------

    d:/xls/class.xlsx  with sheet name class and named range class

    WORKBOOK d:/xls/class.xlsx with sheet class (you can use [sheet1])

   d:/xls/class.xlsx
      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
   2  | ALFRED     |    M       |    99      |    69      |  112.5     |   ** note age changed to 99;
      +------------+------------+------------+------------+------------+
   3  | BARBARA    |    F       |    13      |    58      |  101.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
   20 | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+

   [CLASS]


PROCESS
=======

 SAS  ('99' and 'Alfred' can be in macro vars
        can be macro variable or in anothe SAS or excel named range)

   libname xel "d:/xls/class.xlsx" scan_text=no;
   data xel.class;
      modify xel.class;
      age=99;
      where name= 'Alfred';
   run;quit;
   libname xel clear;

NOTE: The data set XEL.class has been updated.
There were 1 observations rewritten, 0 observations added and 0 observations deleted.

 SAS/Passthru

    * this does it in place;
    proc sql dquote=ansi;
       connect to excel as excel(Path="d:/xls/class.xlsx");
       execute(
         update [class]
         set age=88
         where name="Alfred"
       ) by excel;
       disconnect from excel;
    Quit;

NOTE: The data set XEL.class has been updated.

here are many ways to do this

 1. Using macro variables and explicit passthru to excel
 2. Creating a temp table with updates and using passthru to jupdate using the two tables
 3. Bring the nameded range into sas update and send back
 4. Using Python openpyxl ( I like this the best see below )



