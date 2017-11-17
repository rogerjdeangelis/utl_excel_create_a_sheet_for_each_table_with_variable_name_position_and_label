Creating an excel sheet for each table with variable name, position and label.

see
https://stackoverflow.com/questions/47200335/proc-contents-in-macro-varnum

  Solutions avoided

   a. The #byval solution using excelXP is not included because I do not want a XML workbook.
      see https://amadeus.co.uk/sas-tips/naming-excel-worksheets-with-by-group-values/
   b. Proc eexport is not used because I feel it should be deprecated.


 Three solutions  ( proc export is not needed - SAS forum favorite)

   1. proc contents and libname solution
   2. proc contents and a hash
   3. proc contents, ods excel and proc print


INPUT
=====

   SASHELP.IRIS
                 Variables in Creation Order
    #    Variable       Type    Len    Label

    1    SPECIES        Char     10    Iris Species
    2    SEPALLENGTH    Num       8    Sepal Length (mm)
    3    SEPALWIDTH     Num       8    Sepal Width (mm)
    4    PETALLENGTH    Num       8    Petal Length (mm)
    5    PETALWIDTH     Num       8    Petal Width (mm)



  SASHELP.COMET

                               Variables in Creation Order
    #    Variable    Type    Len    Label

    1    DOSE        Num       8    1,2 Dimethylhydrazine dihydrochloride Dose Level
    2    RAT         Num       8    Rat Index
    3    SAMPLE      Num       8    Slide Index of Grouped Cells from a Rat


  SASHELP.CITIYR

                                Variables in Creation Order
    #    Variable    Type    Len    Format    Label

    1    DATE        Num       6    YEAR4.    Date of Observation
    2    PAN         Num       7              POPULATION EST.: ALL AGES, INC.ARMED F.
    3    PAN17       Num       7              POPULATION EST.: 16 YRS AND OVER,INC ARM
    4    PAN18       Num       7              POPULATION EST.: 18-64 YRS,INC.ARMED F.O
    5    PANF        Num       7              POPULATION EST.: FEMALES,ALL AGES,INC.AR
    6    PANM        Num       7              POPULATION EST.: MALES, ALL AGES, INC.AR


WORKING CODE SOLUTION ONE (see below for other solutions)
==========================================================

  COMPILE TIME DOSUBL
  ===================
      proc contents data=work._all_ position mt=data;
        ods output position=position;

  MAINLINE
  ========
     set position;
         by member notsorted;

    if first.member then do;
      call symputx("member",member);

      rc=dosubl('
            data xel.%scan(&member,2,%str(.));
               set position(where=(member="&member"));
            run;quit;
      ');
    end;


OUTPUT  (excel workbook d:/xls/dsns.xlsx with sheets iris, citiyr and comet)
============================================================================
  d:/xls/dsns.xlsx  with three sheets

   SHEET CITIYR

      ------------------------------------------------------------------------
      |     A        |  B   |         C            |           D             |
      |----------------------------------------------------------------------+
   1  |MEMBER        |#     |VARIABLE              |LABEL                    |
      |--------------+------+----------------------+-------------------------|
   2  |WORK.CITIYR   |1     |DATE                  |Date of Observation      |
      |              |------+----------------------+-------------------------+
   3  |              |2     |PAN                   |POPULATION EST.: ALL     |
      |              |------+----------------------+-------------------------+
   4  |              |3     |PAN17                 |POPULATION EST.: 16 YRS  |
      |              |------+----------------------+-------------------------+
   5  |              |4     |PAN18                 |POPULATION EST.: 18-64   |
      |              |------+----------------------+-------------------------+
   6  |              |5     |PANF                  |POPULATION EST.:         |
      |              |------+----------------------+-------------------------+
   7  |              |6     |PANM                  |POPULATION EST.: MALES,  |
      |--------------+------+----------------------+-------------------------+

   [CITIYR]

      ------------------------------------------------------------------------
      |     A                |  B   |         C    |           D             |
      |----------------------+------+--------------+-------------------------+
   1  |MEMBER                |#     |VARIABLE      |LABEL                    |
      |--------------+------+----------------------+-------------------------|
   2  |WORK.COMET            |1     |DOSE          |1,2 Dimethylhydrazine    |
      |                      |------+--------------+-------------------------+
   3  |                      |2     |RAT           |Rat Index                |
      |                      |------+--------------+-------------------------+
   4  |                      |3     |SAMPLE        |Slide Index of Grouped   |
      |                      |------+--------------+-------------------------+
   5  |                      |4     |LENGTH        |Tail Length of the Comet |
      ------------------------------------------------------------------------

   [COMET]

      ------------------------------------------------------------------------
      |     A        |  B    |         C           |           D             |
      |----------------------+------+--------------+-------------------------+
   1  |MEMBER                |#     |VARIABLE      |LABEL                    |
      |--------------+------+----------------------+-------------------------|
   2  |WORK.IRIS             |1     |SPECIES       |Iris Species             |
      |                      |------+--------------+-------------------------+
   3  |                      |2     |SEPALLENGTH   |Sepal Length (mm)        |
      |                      |------+--------------+-------------------------+
   4  |                      |3     |SEPALWIDTH    |Sepal Width (mm)         |
      |                      |------+--------------+-------------------------+
   5  |                      |4     |PETALLENGTH   |Petal Length (mm)        |
      |                      |------+--------------+-------------------------+
   6  |                      |5     |PETALWIDTH    |Petal Width (mm)         |
      |----------------------+------+--------------+-------------------------+

   [IRIS]


*                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;

* WORK.IRIS, WORK.COMET and WORK.CITIYR;

%symdel dsn / nowarn;

proc datasets lib=work kill;
run;quit;

libname xel clear;

%utlfkil(d:/xls/dsns.xlsx);    * delete if exists;

libname xel "d:/xls/dsns.xlsx";

data _null_;

  do dsns='citiyr','iris','comet';
    call symputx('dsn',dsns);

    rc=dosubl('
       data &dsn;
          set sashelp.&dsn (obs=3);
       run;quit;
    ');
  end;
  stop;

run;quit;


*_ _ _
| (_) |__  _ __   __ _ _ __ ___   ___
| | | '_ \| '_ \ / _` | '_ ` _ \ / _ \
| | | |_) | | | | (_| | | | | | |  __/
|_|_|_.__/|_| |_|\__,_|_| |_| |_|\___|

;

data _null_;

   if _n_=0 then do;
      %let rc=%sysfunc(dosubl('
          proc contents data=work._all_ position mt=data;
            ods output position=position;
          run;quit;
      '));
    end;

    set position;
    by member notsorted;

    if first.member then do;
      call symputx("member",member);

      rc=dosubl('
            data xel.%scan(&member,2,%str(.));
               set position(where=(member="&member"));
            run;quit;
      ');

    end;

run;quit;

libname xel clear; * important;

see
https://stackoverflow.com/questions/47200335/proc-contents-in-macro-varnum


*_               _
| |__   __ _ ___| |__
| '_ \ / _` / __| '_ \
| | | | (_| \__ \ | | |
|_| |_|\__,_|___/_| |_|

;

* repeated code to clean up previous solution;

%symdel dsn / nowarn;

proc datasets lib=work kill;
run;quit;

libname xel clear;
%utlfkil(d:/xls/dsns.xlsx);
libname xel "d:/xls/dsns.xlsx";

data _null_;

  do dsns='citiyr','iris','comet';
    call symputx('dsn',dsns);

    rc=dosubl('
       data &dsn;
          set sashelp.&dsn (obs=3);
       run;quit;
    ');
  end;
  stop;

run;quit;


* SOLUTION;

proc contents data=work._all_ position mt=data;
  ods output position=position;
run;quit;

data _null_;
    dcl hash  a  (ordered: "a");
    a.definekey("key");
    a.definedata("member","variable","num","label");
    a.definedone();
  do until (last.member);
    set position;
    by member notsorted;
    key+1;
    a.add();
  end;
  outdsn=cats("xel.",scan(member,2,"."));
  a.output(dataset:outdsn);
  a.delete();
run;quit;

libname xel clear;


*          _                          _
  ___   __| |___     _____  _____ ___| |
 / _ \ / _` / __|   / _ \ \/ / __/ _ \ |
| (_) | (_| \__ \  |  __/>  < (_|  __/ |
 \___/ \__,_|___/   \___/_/\_\___\___|_|

;

* just in case;

%symdel name / nowarn;
proc datasets lib=work kill;
run;quit;

libname xel clear;
%utlfkil(d:/xls/dsns.xlsx); * delete if exists;

data _null_;

  do dsns='citiyr','iris','comet';
    call symputx('dsn',dsns);

    rc=dosubl('
       data &dsn;
          set sashelp.&dsn (obs=3);
       run;quit;
    ');
  end;
  stop;

run;quit;


ods excel file="d:/xls/dsns.xlsx";

data _null_;

  if _n_=0 then do;
     %let rc=%sysfunc(dosubl('
         ods exclude all;
            proc contents data=work._all_  mt=data directory;
            ods output members=meta;
         run;quit;
         ods select all;
     '));
  end;

  set meta;
  call symputx('name',name);

  rc=dosubl('
         ods exclude all;
           proc contents data=work.&name position mt=data;
             ods output position=position;
         run;quit;
         ods select all;
         ods excel options(sheet_name="&name");
         proc print data=position;
         run;quit;
  ');
run;quit;
ods excel close;




