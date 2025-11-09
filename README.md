# utl-altair-slc-summarizing-excel-bank-statements
Altair slc summarizing excel bank statements
    %let pgm=utl-altair-slc-summarizing-excel-bank-statements;

    %stop_submission;

    Altair slc summarizing excel bank statements

    Too long to post on listserve, see github for full solution

    github
    https://github.com/rogerjdeangelis/utl-altair-slc-summarizing-excel-bank-statements

    community.altair.com
    https://community.altair.com/discussion/19740

    PROBLEM
    -------

      select bank statement
        group by: bank, type, industry
        sum: amount, trans
        count unique values: id, plastic

      Given bank statements in an excel workbook add a summarization excel sheet
      using the slc proc sql (same code works in R and Python)

      proc sql;
        create
          table xls.sumary as
        select
           bank
          ,type
          ,industry
          ,sum(amount)             as sum_amount
          ,sum(trans)              as sum_trans
          ,count(distinct id)      as unq_id
          ,count(distinct plastic) as unq_plastic
        from
           xls.banks
        group
           by bank, type, industry
      ;quit;


    group by: BANK, TYPE, INDUSTRY
    sum: AMOUNT, TRANS
    count unique values: ID, PLASTIC

    /*               _     _
     _ __  _ __ ___ | |__ | | ___ _ __ ___
    | `_ \| `__/ _ \| `_ \| |/ _ \ `_ ` _ \
    | |_) | | | (_) | |_) | |  __/ | | | | |
    | .__/|_|  \___/|_.__/|_|\___|_| |_| |_|
    |_|
    */

    /***********************************************************************************************/
    /*                                                                                             */
    /*  INPUT                                                                                      */
    /*  -----                                                                                      */
    /*                                                                                             */
    /*   WORKBOOK: d:/xls/banks.xlsz SHEET=BANKS                                                   */
    /*   -----------------------+                                                                  */
    /*   | A1| fx    | ID       |                                                                  */
    /*   ---------------------- --------------------------------------------------------------|    */
    /*   [_] |    A     |  B    |    C         |    D  |         E               |   F  |  G  |    */
    /*   -------------------------------------------------------------------------------------+    */
    /*    1  |    ID    |PLASTIC|     BANK     | TYPE  |  INDUSTRY               |AMOUNT|TRANS|    */
    /*    -- |----------+-------+--------------+-------+-------------------------|------+-----|    */
    /*    2  |103140997 1000021 | Banco Amicana| Debito| RETROS_EFFECTO          |1000  |14   |    */
    /*    -- |----------+-------+--------------+-------+-------------------------|------+-----|    */
    /*    3  |103140997 1000021 | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |100   |14   |    */
    /*    -- |----------+-------+--------------+-------+-------------------------|------+-----|    */
    /*    4  |103140997 1000022 | Banco Amicana| Debito| RETROS_EFFECTO          |91.67 |1    |    */
    /*    -- |----------+-------+--------------+-------+-------------------------|------+-----|    */
    /*    5  |103140997 1000022 | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |91.67 |1    |    */
    /*    -- |----------+-------+--------------+-------+-------------------------|------+-----|    */
    /*    6  |103140997 1000021 | Baco LANTAM  | Debito| COMPRAS_TOTA            |41.67 |1    |    */
    /*    -- |----------+-------+--------------+-------+-------------------------|------+-----|    */
    /*    7  |103140997 1000021 | Baco LANTAM  | Debito| OTROS                   | 1.67 |1    |    */
    /*    -- |----------+-------+--------------+-------+-------------------------|------+-----|    */
    /*    8  |103140997 1000021 | Baco LANTAM  | Debito| RETROS_EFFECTO          |816.67|10   |    */
    /*    -- |----------+-------+--------------+-------+-------------------------|------+-----|    */
    /*    9  |103140997 1000021 | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |816.67|11   |    */
    /*    -- |----------+-------+--------------+-------+-------------------------|------+-----+    */
    /*   [BANKS]                                                                                   */
    /*                                                                                             */
    /*---------------------------------------------------------------------------------------------*/
    /*                                                                                             */
    /*  OUTPUT                                                                                     */
    /*  ------                                                                                     */
    /*                                                                                             */
    /*   SAME WORKBOOK: SHEET=SUMARY                                                               */
    /*   -----------------------+                                                                  */
    /*   | A1| fx    | BANK     |                                                                  */
    /*   ---------------------- ------------------------------------------------------------+      */
    /*   [_] |    A         |  B    |    C                   |    D   |  E  |   F   |  G    |      */
    /*   -----------------------------------------------------------------------------------+      */
    /*    1  |              |       |                         | SUM   | SUM | UNIQUE|UNIQUE |      */
    /*    1  |     BANK     | TYPE  |  INDUSTRY               |AMOUNT |TRANS|  ID   |PLASTIC|      */
    /*    -- |--------------+-------+-------------------------|-------+-----|-------+-------|      */
    /*    2  | Baco LANTAM  | Debito| COMPRAS_TOTAL           |  41.67|  1  |   1   |  1    |      */
    /*    -- |--------------+-------+-------------------------|-------+-----|-------+-------|      */
    /*    3  | Baco LANTAM  | Debito| OTROS                   |  41.67|  1  |   1   |  1    |      */
    /*    -- |--------------+-------+-------------------------|-------+-----|-------+-------|      */
    /*    4  | Baco LANTAM  | Debito| RETROS_EFFECTO          | 816.67| 10  |   1   |  1    |      */
    /*    -- |--------------+-------+-------------------------|-------+-----|-------+-------|      */
    /*    5  | Banco Amicana| Debito| RETROS_EFFECTO          |1091.67| 15  |   1   |  2    |      */
    /*    -- |--------------+-------+-------------------------|-------+-----|-------+-------|      */
    /*    6  | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |1008.34| 26  |   1   |  2    |      */
    /*    -- |--------------+-------+-------------------------|-------+-----|-------+-------+      */
    /*   [SUMARY]                                                                                  */
    /*                                                                                             */
    /***********************************************************************************************/

    INPUT
    -----

    WORKBOOK: d:/xls/banks.xlsz SHEET=BANKS
    -----------------------+
    | A1| fx    | ID       |
    ---------------------- --------------------------------------------------------------|
    [_] |    A     |  B    |      c       |    D  |         E               |   F  |  G  |
    -------------------------------------------------------------------------------------+
     1  |    ID    |PLASTIC|     BANK     | TYPE  |  INDUSTRY               |AMOUNT|TRANS|
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     2  |103140997 1000021 | Banco Amicana| Debito| RETROS_EFFECTO          |1000  |14   |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     3  |103140997 1000021 | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |100   |14   |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     4  |103140997 1000022 | Banco Amicana| Debito| RETROS_EFFECTO          |91.67 |1    |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     5  |103140997 1000022 | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |91.67 |1    |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     6  |103140997 1000021 | Baco LANTAM  | Debito| COMPRAS_TOTA            |41.67 |1    |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     7  |103140997 1000021 | Baco LANTAM  | Debito| OTROS                   | 1.67 |1    |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     8  |103140997 1000021 | Baco LANTAM  | Debito| RETROS_EFFECTO          |816.67|10   |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     9  |103140997 1000021 | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |816.67|11   |
     -- |----------+-------+--------------+-------+-------------------------|------+-----+
    [BANKS]


    OUTPUT
    ------

    SAME WORKBOOK: SHEET=SUMARY
    -----------------------+
    | A1| fx    | BANK     |
    ---------------------- ------------------------------------------------------------+
    [_] |    A         |  B    |          c             |    D   |  E  |   F   |  G    |
    -----------------------------------------------------------------------------------+
     1  |              |       |                         | SUM   | SUM | UNIQUE|UNIQUE |
     1  |     BANK     | TYPE  |  INDUSTRY               |AMOUNT |TRANS|  ID   |PLASTIC|
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     2  | Baco LANTAM  | Debito| COMPRAS_TOTAL           |  41.67|  1  |   1   |  1    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     3  | Baco LANTAM  | Debito| OTROS                   |  41.67|  1  |   1   |  1    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     4  | Baco LANTAM  | Debito| RETROS_EFFECTO          | 816.67| 10  |   1   |  1    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     5  | Banco Amicana| Debito| RETROS_EFFECTO          |1091.67| 15  |   1   |  2    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     6  | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |1008.34| 26  |   1   |  2    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------+
    [SUMARY]



    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    &_init;
    libname xls excel "d:/xls/banks.xlsx";

    data xls.banks;

      informat bank $24. type $8. industry $32.;
      input id plastic bank & type industry amount trans;
    cards4;
    103140997 1000021 Banco Amicana  Debito  RETROS_EFFECTO           1000    14
    103140997 1000021 Banco Amicana  Debito  RETROS_EFFECTO,COMPRAS    100    14
    103140997 1000022 Banco Amicana  Debito  RETROS_EFFECTO           91.67    1
    103140997 1000022 Banco Amicana  Debito  RETROS_EFFECTO,COMPRAS   91.67    1
    103140997 1000021 Baco LANTAM    Debito  COMPRAS_TOTAL            41.67    1
    103140997 1000021 Baco LANTAM    Debito  OTROS                    41.67    1
    103140997 1000021 Baco LANTAM    Debito  RETROS_EFFECTO          816.67   10
    103140997 1000021 Banco Amicana  Debito  RETROS_EFFECTO,COMPRAS  816.67   11
    ;;;;
    run;quit;

    libname xls clear;

    WORKBOOK: d:/xls/banks.xlsz SHEET=BANKS
    -----------------------+
    | A1| fx    | ID       |
    ---------------------- --------------------------------------------------------------|
    [_] |    A     |  B    |    C         |    D  |         E               |   F  |  G  |
    -------------------------------------------------------------------------------------+
     1  |    ID    |PLASTIC|     BANK     | TYPE  |  INDUSTRY               |AMOUNT|TRANS|
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     2  |103140997 1000021 | Banco Amicana| Debito| RETROS_EFFECTO          |1000  |14   |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     3  |103140997 1000021 | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |100   |14   |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     4  |103140997 1000022 | Banco Amicana| Debito| RETROS_EFFECTO          |91.67 |1    |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     5  |103140997 1000022 | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |91.67 |1    |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     6  |103140997 1000021 | Baco LANTAM  | Debito| COMPRAS_TOTA            |41.67 |1    |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     7  |103140997 1000021 | Baco LANTAM  | Debito| OTROS                   | 1.67 |1    |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     8  |103140997 1000021 | Baco LANTAM  | Debito| RETROS_EFFECTO          |816.67|10   |
     -- |----------+-------+--------------+-------+-------------------------|------+-----|
     9  |103140997 1000021 | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |816.67|11   |
     -- |----------+-------+--------------+-------+-------------------------|------+-----+
    [BANKS]

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       09:21 Sunday, November  9, 2025

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿;;;;
               ^
    ERROR: Expected a statement keyword : found "ï"

    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.030
          cpu time  : 0.015


    NOTE: AUTOEXEC processing completed

    1         libname xls excel "d:/xls/banks.xlsx";
    NOTE: Library xls assigned as follows:
          Engine:        OLEDB
          Physical Name: d:/xls/banks.xlsx

    2
    3         &_init;
    WARNING: Macro variable "&_init" was not resolved
    ERROR: Expected a statement keyword : found "&"
    4
    5         data xls.banks;
    6
    7           informat bank $24. type $8. industry $32.;
    8           input id plastic bank & type industry amount trans;
    9         cards4;

    ERROR: A database error occurred. The database specific error follows:
           DATABASE error: Table 'banks' already exists.
    NOTE: The data step took :
          real time : 0.559
          cpu time  : 0.421


    10        103140997 1000021 Banco Amicana  Debito  RETROS_EFFECTO           1000    14
    11        103140997 1000021 Banco Amicana  Debito  RETROS_EFFECTO,COMPRAS    100    14
    12        103140997 1000022 Banco Amicana  Debito  RETROS_EFFECTO           91.67    1
    13        103140997 1000022 Banco Amicana  Debito  RETROS_EFFECTO,COMPRAS   91.67    1
    14        103140997 1000021 Baco LANTAM    Debito  COMPRAS_TOTAL            41.67    1
    15        103140997 1000021 Baco LANTAM    Debito  OTROS                    41.67    1
    16        103140997 1000021 Baco LANTAM    Debito  RETROS_EFFECTO          816.67   10
    17        103140997 1000021 Banco Amicana  Debito  RETROS_EFFECTO,COMPRAS  816.67   11
    18        ;;;;
    19        run;quit;

    2                                          Altair SLC       09:21 Sunday, November  9, 2025

    20
    21
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 1.224
          cpu time  : 1.015

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    /*--- only needed for testing. sas and slc cannot drop a sheet ----*/

    options set=RHOME "D:\d451";
    &_init_;
    proc r;
    submit;
    library(openxlsx)
    wkb="d:/xls/banks.xlsx"
    wb <- loadWorkbook(wkb)
    removeWorksheet(wb, "SUMARY"")
    saveWorkbook(wb, wkb, overwrite = TRUE)
    endsubmit;
    run;quit;

    libname xls excel "d:/xls/banks.xlsx";

    proc sql;
      create
        table xls.sumary as
      select
         bank
        ,type
        ,industry
        ,sum(amount)             as sum_amount
        ,sum(trans)              as sum_trans
        ,count(distinct id)      as unq_id
        ,count(distinct plastic) as unq_plastic
      from
         xls.banks
      group
         by bank, type, industry
    ;quit;

    libname xls clear;

    /*           _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| `_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    */

    SAME WORKBOOK: SHEET=SUMARY
    -----------------------+
    | A1| fx    | BANK     |
    ---------------------- ------------------------------------------------------------+
    [_] |    A         |  B    |          c             |    D   |  E  |   F   |  G    |
    -----------------------------------------------------------------------------------+
     1  |              |       |                         | SUM   | SUM | UNIQUE|UNIQUE |
     1  |     BANK     | TYPE  |  INDUSTRY               |AMOUNT |TRANS|  ID   |PLASTIC|
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     2  | Baco LANTAM  | Debito| COMPRAS_TOTAL           |  41.67|  1  |   1   |  1    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     3  | Baco LANTAM  | Debito| OTROS                   |  41.67|  1  |   1   |  1    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     4  | Baco LANTAM  | Debito| RETROS_EFFECTO          | 816.67| 10  |   1   |  1    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     5  | Banco Amicana| Debito| RETROS_EFFECTO          |1091.67| 15  |   1   |  2    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------|
     6  | Banco Amicana| Debito| RETROS_EFFECTO,COMPRAS  |1008.34| 26  |   1   |  2    |
     -- |--------------+-------+-------------------------|-------+-----|-------+-------+
    [SUMARY]

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
