# utl-altair-slc-excel-using-undocumented-excel-vba-and-sql-functions-to-prepare-a-sheet-for-import
Altair slc excel using undocumented excel vba and sql functions to prepare a sheet for import
    %let pgm=utl-altair-slc-excel-using-undocumented-excel-vba-and-sql-functions-to-prepare-a-sheet-for-import;

    %stop submission;

    Altair slc excel using undocumented excel vba and sql functions to prepare a sheet for import

    Too long to post here, see github
    https://github.com/rogerjdeangelis/utl-altair-slc-excel-using-undocumented-excel-vba-and-sql-functions-to-prepare-a-sheet-for-import


    PROBLEM (Remove '%' characters around name variable using just excel vba and sql

      CONTENTS

        1 slc remove tab charcacter
        2 slc proc r (best solution handles length issues automatically)
        3 slc list UNDOCUMENTED vba and sql functions?
        4 could not get pyodbc to work
          NOTE: pyodbc.Error: ('HYC00', '[HYC00] [Microsoft][ODBC Excel Driver]Optional feature not implemented
          (106) (SQLSetConnectAttr(SQL_ATTR_AUTOCOMMIT))')
          (kept suggesting python excel packages.

     input

     d:/xls/tabs/xlsx

       +------------------------------------------------
       |     A               |    B       |     C      |
       +------------------------------------------------
    1  |   NAME              |   SEX      |    AGE     |
       +---------------------+------------+------------+
    2  |     %Alfred%        |    M       |    14      |
       +---------------------+------------+------------+
        ...
       +---------------------+------------+------------+
    7  |     %JAMES%         |    M       |    12      |
       +---------------------+------------+------------+

     [CLASS]

      CONTENTS

        1 slc remove tab charcacter
        2 slc proc r (best solution handles length issues automatically)
        3 slc list vba and sql functions
        4 could not get pyodbc to work
          NOTE: pyodbc.Error: ('HYC00', '[HYC00] [Microsoft][ODBC Excel Driver]Optional feature not implemented
          (106) (SQLSetConnectAttr(SQL_ATTR_AUTOCOMMIT))')
          (kept suggesting python excel packages.


    METHOD

      Use slc passthru to Microsift access sql and vba  functions like isnumeric. trim and length to prepare a sheet for import.
      It is better to use MS native VBA and SQL functions to prepare a worksheet then post processing the slc imported dataset.
      NOTE The excel odbc driver returns all charater variables as 255 bye strings, you need to change the lenghts ouside
      the odbc sql.
      It is undocumented that the large number of vba functions can be used with odbs sql.
      This may be more powerfull the excel powerquery?

      Little off and undocumented?
       SQL CHAR function does not work
       VBA CHR does the same thing and works (suggest use of vba functons in excel queries)


    RELATED REPO

    https://github.com/rogerjdeangelis/utl-altair-slc-determining-type-and-length-of-excel-columns-before-importing

    As a side note
      You should be able to run the sql query in any language that supports ODBC,
      r, python, octave(matlab). poweshell. pspp(spss)...

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    proc datasets lib=workx kill nodetails nolist;
    run;

    %utlfkil(d:/xls/tabs.xlsx);

    libname xel excel "d:/xls/tabs.xlsx";

    data xel.class;
      length name $8;
      input
        name$
        sex$ age;
        name=cats('%',name,'%');
    cards4;
    Alfred  M 14
    Alice   F 13
    Barbara F 13
    Carol   F 14
    Henry   M 14
    James   M 12
    ;;;;
    run;quit;

    /**************************************************************************************************************************/
    /*   d:/xls/tabs/xlsx                                                                                                     */
    /*                                                                                                                        */
    /*    +------------------------------------------------                                                                   */
    /*    |     A               |    B       |     C      |                                                                   */
    /*    +------------------------------------------------                                                                   */
    /* 1  |   NAME              |   SEX      |    AGE     |                                                                   */
    /*    +---------------------+------------+------------+                                                                   */
    /* 2  |     %Alfred%        |    M       |    14      |                                                                   */
    /*    +---------------------+------------+------------+                                                                   */
    /*     ...                                                                                                                */
    /*    +---------------------+------------+------------+                                                                   */
    /* 7  |     %JAMES%         |    M       |    12      |                                                                   */
    /*    +---------------------+------------+------------+                                                                   */
    /*                                                                                                                        */
    /*  [CLASS]                                                                                                               */
    /**************************************************************************************************************************/

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */
    1                                          Altair SLC       14:09 Sunday, February 15, 2026

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿ods _all_ close;
               ^
    ERROR: Expected a statement keyword : found "?"
    NOTE: Library workx assigned as follows:
          Engine:        SAS7BDAT
          Physical Name: d:\wpswrkx

    NOTE: Library slchelp assigned as follows:
          Engine:        WPD
          Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

    NOTE: Library worksas assigned as follows:
          Engine:        SAS7BDAT
          Physical Name: d:\worksas

    NOTE: Library workwpd assigned as follows:
          Engine:        WPD
          Physical Name: d:\workwpd


    LOG:  14:09:58
    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.018
          cpu time  : 0.000


    NOTE: AUTOEXEC processing completed

    1         %utlfkil(d:/xls/tabs.xlsx);
    2
    3         libname xel excel "d:/xls/tabs.xlsx";
    NOTE: Library xel assigned as follows:
          Engine:        OLEDB
          Physical Name: d:/xls/tabs.xlsx

    4
    5         data xel.class workx.classchk;
    6           length name $10;
    7           input
    8             name$
    9             sex$ age;
    10            name=cats('%',name,'%');
    11        cards4;

    NOTE: Data set "XEL.class" has an unknown number of observation(s) and 3 variable(s)
    NOTE: Data set "WORKX.classchk" has 6 observation(s) and 3 variable(s)
    NOTE: The data step took :
          real time : 0.105
          cpu time  : 0.093


    12        Alfred  M 14
    13        Alice   F 13
    14        Barbara F 13
    15        Carol   F 14
    16        Henry   M 14
    17        James   M 12
    18        ;;;;
    19        run;quit;
    20
    21
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 0.769
          cpu time  : 0.828

    /*                   _                   _ _
      ___ _ __ ___  __ _| |_ ___    ___   __| | |__   ___
     / __| `__/ _ \/ _` | __/ _ \  / _ \ / _` | `_ \ / __|
    | (__| | |  __/ (_| | ||  __/ | (_) | (_| | |_) | (__
     \___|_|  \___|\__,_|\__\___|  \___/ \__,_|_.__/ \___|
                         _                                             _           _
     _ __ ___  _   _ ___| |_   _ __ _   _ _ __     __ _ ___   __ _  __| |_ __ ___ (_)_ __
    | `_ ` _ \| | | / __| __| | `__| | | | `_ \   / _` / __| / _` |/ _` | `_ ` _ \| | `_ \
    | | | | | | |_| \__ \ |_  | |  | |_| | | | | | (_| \__ \  (_| | (_| | | | | | | | | | |
    |_| |_| |_|\__,_|___/\__| |_|   \__,_|_| |_|  \__,_|___/ \__,_|\__,_|_| |_| |_|_|_| |_|
    */

    /*--- has to be on one line, backtick not available                                                                            ---*/
    /*--- drop down to ADMIN powershell or paste into poweshell                                                                    ---*/

    %utl_submit_ps64('
    Add-OdbcDsn -Name "tab" -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -SetPropertyValue "Dbq=d:\xls\tabs.xlsx";
    Get-OdbcDsn;
    ');

    OR (can use backtick)
    Open poweshell in ADMIN mode just paste this script (note use of backtick)

    Add-OdbcDsn -Name "tab" `
        -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" `
        -DsnType "User" `
        -SetPropertyValue "Dbq=d:\xls\tabs.xlsx"
    Get-OdbcDsn

    /**************************************************************************************************************************/
    /*   Name       : tab                                                                                                     */
    /*  DsnType    : User                                                                                                     */
    /*  Platform   : 64-bit                                                                                                   */
    /*  DriverName : Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)                                                   */
    /*  Attribute  : {DBQ, DriverId, ImplicitCommitSync, Threads...}                                                          */
    /**************************************************************************************************************************/

    /*         _ _            _
      ___   __| | |__   ___  | | ___   __ _
     / _ \ / _` | `_ \ / __| | |/ _ \ / _` |
    | (_) | (_| | |_) | (__  | | (_) | (_| |
     \___/ \__,_|_.__/ \___| |_|\___/ \__, |
                                      |___/
    */
    1                                          Altair SLC     08:50 Saturday, February 14, 2026

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿ods _all_ close;
               ^
    ERROR: Expected a statement keyword : found "?"
    NOTE: Library workx assigned as follows:
          Engine:        SAS7BDAT
          Physical Name: d:\wpswrkx

    NOTE: Library slchelp assigned as follows:
          Engine:        WPD
          Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

    NOTE: Library worksas assigned as follows:
          Engine:        SAS7BDAT
          Physical Name: d:\worksas

    NOTE: Library workwpd assigned as follows:
          Engine:        WPD
          Physical Name: d:\workwpd


    LOG:  8:50:36
    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.047
          cpu time  : 0.000


    NOTE: AUTOEXEC processing completed

    1         %utl_submit_ps64('
    2         Add-OdbcDsn -Name "tab" -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -SetPropertyValue "Dbq=d:\xls\tabs.xlsx";
    3         Get-OdbcDsn;
    4         ');

    NOTE: The file py_pgm is:
          Filename='d:\wpswrk\_TD10328\py_pgm.ps1',
          Owner Name=BUILTIN\Administrators,
          File size (bytes)=0,
          Create Time=08:50:35 Feb 14 2026,
          Last Accessed=08:50:35 Feb 14 2026,
          Last Modified=08:50:35 Feb 14 2026,
          Lrecl=32766, Recfm=V

    Add-OdbcDsn -Name "tab" -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -SetPropertyValue "Dbq=d:\xls\tabs.xlsx"

    Get-OdbcDsn

    NOTE: 2 records were written to file py_pgm
          The minimum record length was 384
          The maximum record length was 384
    NOTE: The data step took :
          real time : 0.000
          cpu time  : 0.000


    d:\wpswrk\_TD10328\py_pgm.ps1

    NOTE: The infile rut is:
          Unnamed Pipe Access Device,
          Process=powershell.exe -executionpolicy bypass -file d:\wpswrk\_TD10328\py_pgm.ps1 ,
          Lrecl=32767, Recfm=V



    Name       : dBASE Files
    DsnType    : User
    Platform   : 64-bit
    DriverName : Microsoft Access dBASE Driver (*.dbf, *.ndx, *.mdx)
    Attribute  : {Threads, SafeTransactions, ImplicitCommitSync, DriverId...}

    Name       : Excel Files
    DsnType    : User
    Platform   : 64-bit
    DriverName : Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)
    Attribute  : {SafeTransactions, DriverId, ImplicitCommitSync, Threads...}

    Name       : MS Access Database
    DsnType    : User
    Platform   : 64-bit
    DriverName : Microsoft Access Driver (*.mdb, *.accdb)
    Attribute  : {Threads, SafeTransactions, ImplicitCommitSync, DriverId...}



    Name       : tab
    DsnType    : User
    Platform   : 64-bit
    DriverName : Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)
    Attribute  : {DBQ, DriverId, ImplicitCommitSync, Threads...}



    Name       : sqlitedsn
    DsnType    : System
    Platform   : 64-bit
    DriverName : GM-Software SQLite3 ODBC Driver
    Attribute  : {EnableViews, Exclusive, UseTriggers, NoFollow...}



    NOTE: 58 records were written to file PRINT

    NOTE: 58 records were read from file rut
          The minimum record length was 0
          The maximum record length was 73

    NOTE: The data step took :
          real time : 0.888
          cpu time  : 0.031


    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 1.014
          cpu time  : 0.093

    Open poweshell in admin mode and paste this script

    Get-OdbcDsn
    Add-OdbcDsn -Name "tab" `
        -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" `
        -DsnType "User" `
        -SetPropertyValue "Dbq=d:\xls\tabs.xlsx"
    Get-OdbcDsn

    * works ;
    proc sql;
      connect to odbc (dsn="tab");
      create table workx.tab as
      select * from connection to odbc (
        select
           *
        from [class$]
      );
      disconnect from odbc;
    quit;


    %utl_submit_ps64('
      Add-OdbcDsn -Name "tabx"
          -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
          -DsnType "User"
          -SetPropertyValue "Dbq=d:\xls\tabs.xlsx";
      Get-OdbcDsn;
    ');


    %utl_submit_ps64('
    Add-OdbcDsn -Name "tabx" -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -SetPropertyValue "Dbq=d:\xls\tabs.xlsx";
    Get-OdbcDsn;
    ');

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    proc sql;
      connect to odbc (dsn="tab");
      create table workx.tab as
      select
         name length=8
        ,put(name,$hex18.) as name_hex
        ,sex length=1
        ,put(clean_name,$hex14.) as clean_name_hex
        ,age
        ,clean_name length=8
      from
        connection
           to odbc (
        select
           name
          ,replace(name,'%','') as clean_name
          ,sex
          ,age
        from
           [class$]
      );
      disconnect from odbc;
    quit;

    /**************************************************************************************************************************/
    /* Whats happening                                                                                                        */
    /*                      INPUT            NOTE MISSING HEX 25              CLEAN_                                          */
    /*     NAME            NAME_HEX           CLEAN_NAME_HEX      SEX  AGE     NAME                                           */
    /*                 %               %                                                                                      */
    /*   %Alfred%     25 416C66726564 25 20   416C6672656420       M    14    Alfred                                          */
    /*                                                                                                                        */
    /* WORKX.TAB total obs=6 15FEB2026:14:25:13                                                                               */
    /*                                                                          CLEAN_                                        */
    /* Obs      NAME           NAME_HEX         SEX    CLEAN_NAME_HEX    AGE     NAME                                         */
    /*                                                                                                                        */
    /*  1     %Alfred%    25416C667265642520     M     416C6672656420     14    Alfred                                        */
    /*  2     %Alice%     25416C696365252020     F     416C6963652020     13    Alice                                         */
    /*  3     %Barbara    254261726261726125     F     42617262617261     13    Barbara                                       */
    /*  4     %Carol%     254361726F6C252020     F     4361726F6C2020     14    Carol                                         */
    /*  5     %Henry%     2548656E7279252020     M     48656E72792020     14    Henry                                         */
    /*  6     %James%     254A616D6573252020     M     4A616D65732020     12    James                                         */
    /**************************************************************************************************************************/


    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       14:29 Sunday, February 15, 2026

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿ods _all_ close;
               ^
    ERROR: Expected a statement keyword : found "?"
    NOTE: Library workx assigned as follows:
          Engine:        SAS7BDAT
          Physical Name: d:\wpswrkx

    NOTE: Library slchelp assigned as follows:
          Engine:        WPD
          Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

    NOTE: Library worksas assigned as follows:
          Engine:        SAS7BDAT
          Physical Name: d:\worksas

    NOTE: Library workwpd assigned as follows:
          Engine:        WPD
          Physical Name: d:\workwpd


    LOG:  14:29:01
    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.033
          cpu time  : 0.000


    NOTE: AUTOEXEC processing completed

    1         proc sql;
    2           connect to odbc (dsn="tab");
    NOTE: Connected to DB: tab (EXCEL version 12.00.0000)
    NOTE: Connected to DB: tab (EXCEL version 12.00.0000)
    NOTE: Successfully connected to database ODBC as alias ODBC.
    3           create table workx.tab as
    4           select
    5              name length=8
    6             ,put(name,$hex18.) as name_hex
    7             ,sex length=1
    8             ,put(clean_name,$hex14.) as clean_name_hex
    9             ,age
    10            ,clean_name length=8
    11          from
    12            connection
    13               to odbc (
    14            select
    15               name
    16              ,replace(name,'%','') as clean_name
    17              ,sex
    18              ,age
    19            from
    20               [class$]
    21          );
    NOTE: Data set "WORKX.tab" has 6 observation(s) and 6 variable(s)
    22          disconnect from odbc;
    NOTE: Successfully disconnected from database ODBC.
    23        quit;
    NOTE: Procedure sql step took :
          real time : 0.705
          cpu time  : 0.843


    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 0.792
          cpu time  : 0.890

    /*___        _
    |___ \   ___| | ___   _ __  _ __ ___   ___   _ __
      __) | / __| |/ __| | `_ \| `__/ _ \ / __| | `__|
     / __/  \__ \ | (__  | |_) | | | (_) | (__  | |
    |_____| |___/_|\___| | .__/|_|  \___/ \___| |_|
                         |_|
    */

    proc delete data=work.final;
    run;quit;

    options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
    proc r;
    submit;
    library(RODBC);
    ch <- odbcConnect("tab");
    want <- sqlQuery(ch,"
      select
         name
        ,replace(name,'%','') as clean_name
        ,sex
        ,age
      from
         [class$]

        ");
    want
    endsubmit;
    import data=workx.final r=want;
    run;

    proc print data=workx.final;
    run;

    /**************************************************************************************************************************/
    /*   R                             |       Back ro parent SLC                           |                                 */
    /*                                 |                                                    |                                 */
    /* Altair SLC                      |   Altair SLC                                       | Variables in Creation Order     */
    /*                                 |                                                    |                                 */
    /*        name clean_name sex age  |   Obs      NAME       CLEAN_NAME    SEX    AGE   # |    Variable      Type    Len    */
    /*                                 |                                                    |                                 */
    /* 1  %Alfred%     Alfred   M  14  |    1     %Alfred%      Alfred        M      14   1 |    NAME          Char      9    */
    /* 2   %Alice%      Alice   F  13  |    2     %Alice%       Alice         F      13   2 |    CLEAN_NAME    Char      7    */
    /* 3 %Barbara%    Barbara   F  13  |    3     %Barbara%     Barbara       F      13   3 |    SEX           Char      1    */
    /* 4   %Carol%      Carol   F  14  |    4     %Carol%       Carol         F      14   4 |    AGE           Num       8    */
    /* 5   %Henry%      Henry   M  14  |    5     %Henry%       Henry         M      14     |                                 */
    /* 6   %James%      James   M  12  |    6     %James%       James         M      12     |                                 */
    /**************************************************************************************************************************/

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       15:59 Sunday, February 15, 2026

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿ods _all_ close;
               ^
    ERROR: Expected a statement keyword : found "?"
    NOTE: Library workx assigned as follows:
          Engine:        SAS7BDAT
          Physical Name: d:\wpswrkx

    NOTE: Library slchelp assigned as follows:
          Engine:        WPD
          Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

    NOTE: Library worksas assigned as follows:
          Engine:        SAS7BDAT
          Physical Name: d:\worksas

    NOTE: Library workwpd assigned as follows:
          Engine:        WPD
          Physical Name: d:\workwpd


    LOG:  15:59:53
    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.031
          cpu time  : 0.015


    NOTE: AUTOEXEC processing completed

    1         options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
    2         proc r;
    3         submit;
    4         library(RODBC);
    5         ch <- odbcConnect("tab");
    6         want <- sqlQuery(ch,"
    7           select
    8              name
    9             ,replace(name,'%','') as clean_name
    10            ,sex
    11            ,age
    12          from
    13             [class$]
    14
    15            ");
    16        want
    17        endsubmit;
    NOTE: Using R version 4.5.2 (2025-10-31 ucrt) from C:\Program Files\R\R-4.5.2

    NOTE: Submitting statements to R:

    > library(RODBC);
    > ch <- odbcConnect("tab");
    > want <- sqlQuery(ch,"
    +   select
    +      name
    +     ,replace(name,'%','') as clean_name
    +     ,sex
    +     ,age
    +   from
    +      [class$]
    +
    +     ");
    > want

    NOTE: Processing of R statements complete

    18        import data=workx.final r=want;
    NOTE: Creating data set 'WORKX.final' from R data frame 'want'
    NOTE: Column names modified during import of 'want'
    NOTE: Data set "WORKX.final" has 6 observation(s) and 4 variable(s)

    19        run;
    NOTE: Procedure r step took :
          real time : 1.139
          cpu time  : 0.031


    20
    21        proc print data=workx.final;
    22        run;
    NOTE: 6 observations were read from "WORKX.final"
    NOTE: Procedure print step took :
          real time : 0.000
          cpu time  : 0.000


    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 1.217
          cpu time  : 0.109

    /*____        _                       _   __                  _   _
    |___ / __   _| |__   __ _   ___  __ _| | / _|_   _ _ __   ___| |_(_) ___  _ __  ___
      |_ \ \ \ / / `_ \ / _` | / __|/ _` | || |_| | | | `_ \ / __| __| |/ _ \| `_ \/ __|
     ___) | \ V /| |_) | (_| | \__ \ (_| | ||  _| |_| | | | | (__| |_| | (_) | | | \__ \
    |____/   \_/ |_.__/ \__,_| |___/\__, |_||_|  \__,_|_| |_|\___|\__|_|\___/|_| |_|___/
                                |_|
    */

    I did check most of the vba character functions and they worked.
    I did not test them all. Also note the same function may appear in both VBA and SQL
    I suggest you tray VBA first?


    All obs from bothsrt total obs=147

    Obs    TYPE    FUNCTION                        ARGS               DESCRIPTION

      1    VBA     ABS                  number                        Absolute value
      2    SQL     ABS                  number                        Absolute value
      3    VBA     ASC                  string                        Returns ASCII code of first character
      4    SQL     ASCII                string                        Returns ASCII code of leftmost character
      5    VBA     ATN                  number                        Arc tangent
      6    SQL     AVG                  column                        Average of values
      7    VBA     CBOOL                expression                    Converts to boolean
      8    VBA     CBYTE                expression                    Converts to byte
      9    VBA     CCUR                 expression                    Converts to currency
     10    VBA     CDATE                expression                    Converts to date
     11    VBA     CDBL                 expression                    Converts to double (number)
     12    VBA     CDEC                 expression                    Converts to decimal
     13    SQL     CEILING              number                        Smallest integer = number
     14    SQL     CHAR                 code                          Returns character for given ASCII code
     15    VBA     CHOOSE               index, value1, value2...      Returns value based on index
     16    VBA     CHR                  code                          Returns character for given ASCII code (CHR(9) for tab)
     18    VBA     CINT                 expression                    Converts to integer
     19    VBA     CLNG                 expression                    Converts to long integer
     20    SQL     CONCAT               string1 string2               Concatenates two strings
     21    VBA     COS                  angle                         Cosine of angle
     22    SQL     COS                  angle                         Cosine of angle
     23    SQL     COUNT                *                             Count of rows/values
     24    VBA     CSNG                 expression                    Converts to single precision
     25    VBA     CSTR                 expression                    Converts to string
     26    SQL     CURDATE               / CURRENT_DATE               Current date
     27    SQL     CURRENT_TIMESTAMP                                  Current date and time
     28    SQL     CURTIME               / CURRENT_TIME               Current time
     29    VBA     CVAR                 expression                    Converts to variant
     30    VBA     CVErr                errornumber                   Returns error variant
     31    SQL     DATABASE                                           Current database name
     32    VBA     DATEADD              interval, number, date        Adds interval to date
     33    VBA     DATEDIFF             interval, date1, date2        Difference between dates
     34    VBA     DATEPART             interval, date                Returns specified part of date
     35    VBA     DATESERIAL           year, month, day              Creates date from components
     36    VBA     DATEVALUE            string                        Converts string to date
     37    VBA     DAVG                 expr, domain [,criteria]      Average of values (domain aggregate)
     38    VBA     DAY                  date                          Extracts day of month (1-31)
     39    SQL     DAYNAME              date                          Name of weekday
     40    SQL     DAYOFMONTH           date                          Day of month (1-31)
     41    SQL     DAYOFWEEK            date                          Day of week (1-7)
     42    SQL     DAYOFYEAR            date                          Day of year (1-366)
     43    VBA     DCOUNT               expr, domain [,criteria]      Count of records (domain aggregate)
     44    VBA     DMAX                 expr, domain [,criteria]      Maximum value (domain aggregate)
     45    VBA     DMIN                 expr, domain [,criteria]      Minimum value (domain aggregate)
     46    VBA     DSTDEV               expr, domain [,criteria]      Standard deviation (domain aggregate)
     47    VBA     DSTDEVP              expr, domain [,criteria       Population std dev (domain aggregate)
     48    VBA     DSUM                 expr, domain [,criteria]      Sum of values (domain aggregate)
     49    VBA     DVAR                 expr, domain [,criteria]      Variance (domain aggregate)
     50    VBA     DVARP                expr, domain [,criteria]      Population variance (domain aggregate)
     51    VBA     ERROR$               [errornumber]                 Returns error message
     52    VBA     EXP                  number                        e raised to power
     53    SQL     EXP                  power                          raised to specified power
     54    SQL     EXTRACT              part FROM date                Extracts date part (YEAR, MONTH, etc.)
     55    VBA     FIRST                expression                    First value in group/domain
     56    VBA     FIX                  number                        Returns integer portion (no rounding)
     57    SQL     FLOOR                number                        Largest integer = number
     58    VBA     HOUR                 time                          Extracts hour (0-23)
     59    SQL     HOUR                 time                          Hour (0-23)
     60    VBA     IF                   condition, true, false        Alias for IIF
     61    SQL     IFNULL               expr value                    Returns value if expr is NULL
     62    VBA     IIF                  condition, true, false        If-Then-Else logic (PREFERRED)
     63    SQL     IIF                  true false                    If then else
     64    VBA     INSTR                [start,]string,substring      Finds position of substring
     65    VBA     INT                  number                        Returns integer portion
     66    VBA     ISDATE               expression                    Tests if expression is a date
     67    VBA     ISEMPTY              expression                    Tests if variant is uninitialized
     68    VBA     ISERROR              expression                    Tests if expression is an error
     69    VBA     ISNULL               expression                    Tests if expression is NULL
     70    VBA     ISNUMERIC            expression                    Tests if expression is numeric (your test)
     72    SQL     ISNUMERIC            expression                    Returns -1 if numeric, 0 if not numeric
     73    VBA     ISTEXT               expression                    Tests if expression is text
     74    VBA     LAST                 expression                    Last value in group/domain
     75    VBA     LCASE                string                        Converts string to lowercase
     76    SQL     LCASE                string                        Converts string to lowercase
     77    VBA     LEFT                 string, count                 Extracts leftmost characters
     78    SQL     LEFT                 string count                  Extracts leftmost characters
     79    VBA     LEN                  string                        Returns number of characters
     80    SQL     LEN                  string                        Returns number of characters in string
     81    SQL     LENGTH               string                        Returns string length (alias for LEN)
     82    SQL     LOCATE               search target[ start]         Finds position of substring
     83    VBA     LOG                  number                        Natural logarithm
     84    SQL     LOG                  number                        Natural logarithm
     85    VBA     LOG10                number                        Base-10 logarithm
     86    VBA     LTRIM                string                        Removes leading spaces
     87    SQL     LTRIM                string                        Removes leading spaces
     88    SQL     MAX                  column                        Maximum value
     89    VBA     MID                  string, start[, length]       Extracts substring from position
     90    SQL     MIN                  column                        Minimum value
     91    VBA     MINUTE               time                          Extracts minute (0-59)
     92    SQL     MINUTE               time                          Minute (0-59)
     93    SQL     MOD                  dividend divisor              Remainder of division
     94    VBA     MONTH                date                          Extracts month (1-12)
     95    SQL     MONTH                date                          Month (1-12)
     96    VBA     MONTHNAME            month                         Name of month
     97    SQL     MONTHNAME            date                          Name of month
     98    SQL     NOW                                                Current date and time
     99    SQL     NULLIF               expr1 expr2                   Returns NULL if expressions equal
    100    VBA     NZ                   variant [, value_if_null]     Converts NULL to zero or specified value
    101    SQL     POSITION             target IN search              Finds substring position
    102    SQL     POWER                base exponent                 Number raised to power
    103    VBA     QBColor              color                         Returns QuickBasic color code
    104    SQL     QUARTER              date                          Quarter (1-4)
    105    SQL     RAND                 [seed]                        Random number (0-1)
    106    VBA     REPLACE              string, find, replace         Replaces occurrences of substring
    107    VBA     RGB                  red, green, blue              Returns RGB color code
    108    VBA     RIGHT                string, count                 Extracts rightmost characters
    109    SQL     RIGHT                string count                  Extracts rightmost characters
    110    VBA     RND                  [seed]                        Random number (0-1)
    111    VBA     ROUND                number [,places]              Rounds to specified decimals
    112    SQL     ROUND                number places                 Rounds to specified decimals
    113    VBA     RTRIM                string                        Removes trailing spaces
    114    SQL     RTRIM                string                        Removes trailing spaces
    115    VBA     SECOND               time                          Extracts second (0-59)
    116    SQL     SECOND               time                          Second (0-59)
    117    VBA     SGN                  number                        Returns sign (-1, 0, 1)
    118    SQL     SIGN                 number                        Returns sign (-1, 0, 1)
    119    VBA     SIN                  angle                         Sine of angle
    120    SQL     SIN                  angle                         Sine of angle
    121    VBA     SPACE                count                         Returns string of spaces
    122    VBA     SQR                  number                        Square root
    123    SQL     SQRT                 number                        Square root
    124    VBA     STR                  number                        Converts number to string
    125    VBA     STRING               count, character              Returns character repeated count times
    126    SQL     SUBSTRING            string start length           Extracts portion of string
    127    SQL     SUM                  column                        Sum of values
    128    VBA     SWITCH               expr1, value1, expr2, va      Multiple condition evaluation
    129    VBA     TAN                  angle                         Tangent of angle
    130    SQL     TAN                  angle                         Tangent of angle
    131    VBA     TIMESERIAL           hour, minute, second          Creates time from components
    132    SQL     TIMESTAMPADD         interval, count, timestamp    Adds interval to timestamp
    133    SQL     TIMESTAMPDIFF        interval, start, end          Difference between timestamps
    134    VBA     TIMEVALUE            string                        Converts string to time
    135    VBA     TRIM                 string                        Removes leading and trailing spaces
    136    SQL     TRIM                 string                        Removes leading and trailing spaces
    137    SQL     TRUNCATE             number places                 Truncates to decimal places
    138    VBA     TYPENAME             variable                      Returns data type as string
    139    VBA     UCASE                string                        Converts string to uppercase
    140    SQL     UCASE                string                        Converts string to uppercase
    141    SQL     USER                                               Current database user
    142    VBA     VAL                  string                        Returns numeric value from string
    143    SQL     WEEK                 date                          Week number
    144    VBA     WEEKDAY              date                          Day of week (1-7, Sunday=1)
    145    VBA     WEEKDAYNAME          weekday                       Name of weekday
    146    VBA     YEAR                 date                          Extracts year
    147    SQL     YEAR                 date                          Year

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
