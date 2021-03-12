&ANALYZE-SUSPEND _VERSION-NUMBER AB_v10r12 GUI
&ANALYZE-RESUME
&Scoped-define WINDOW-NAME C-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS C-Win 
/*------------------------------------------------------------------------

  File: 

  Description: 

  Input Parameters:
      <none>

  Output Parameters:
      <none>

  Author: 

  Created: 

------------------------------------------------------------------------*/
/*          This .W file was created with the Progress AppBuilder.      */
/*----------------------------------------------------------------------*/

/* Create an unnamed pool to store all the widgets created 
     by this procedure. This is a good default which assures
     that this procedure's triggers and internal procedures 
     will execute in this procedure's storage, and that proper
     cleanup will occur on deletion of the procedure. */

CREATE WIDGET-POOL.

/* ***************************  Definitions  ************************** */

/* Parameters Definitions ---                                           */

/* Global Variable Definitions ---                                      */

DEFINE NEW GLOBAL SHARED VAR data-ini AS DATE NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR data-fim AS DATE NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR centro AS CHAR NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR oee-ano AS INT NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR oee-mes AS CHAR NO-UNDO.

/* Local Variable Definitions ---                                       */
DEF VAR limpa-grafico-ap AS LOGICAL INIT FALSE.
DEF VAR limpa-grafico-par AS LOGICAL INIT FALSE.
DEF VAR data-atu-oee AS DATE.
DEF VAR mes-atu-oee AS CHAR FORMAT "x(10)".
DEF VAR i-mes AS INT.
DEF VAR impressao AS INTEGER INIT 1.
DEF VAR cFile AS CHAR FORMAT "x(50)".

/* Temp Tables Definitions ---                                          */

DEFINE NEW GLOBAL SHARED TEMP-TABLE tt-oee NO-UNDO
    FIELD data AS DATE LABEL "Data Apontamento"
    FIELD turno AS DECIMAL LABEL "Tempo Programado"
    FIELD cod-item LIKE split-operac.it-codigo LABEL "Item" 
    FIELD temp-prod AS DECIMAL LABEL "Tempo Produzindo"
    FIELD prod-teorica AS DECIMAL LABEL "Produá∆o Teorica"
    FIELD prod-real AS INT LABEL "Produá∆o Real"
    FIELD pc-boas AS INT LABEL "Peáas Boas"
    FIELD pc-ruins AS INT LABEL "Peáas Ruins"
    FIELD pc-boas-ruins AS INT LABEL "Peáas Boas + Ruins"
    FIELD temp-par-alt AS DECIMAL
    FIELD temp-par-nao-alt AS DECIMAL
    FIELD oee-disp AS DECIMAL LABEL "Disponibilidade"
    FIELD oee-perf AS DECIMAL LABEL "Performance"
    FIELD oee-quali AS DECIMAL LABEL "Qualidade"
    FIELD oee-percent AS DECIMAL LABEL "Percentual OEE"
    FIELD mes AS INT.

DEFINE NEW GLOBAL SHARED TEMP-TABLE tt-parada NO-UNDO
    FIELD cod-ctrab LIKE rep-parada-ctrab.cod-ctrab LABEL "teste"
    FIELD cod-parada AS CHAR LABEL "Cod"
    FIELD des-parada LIKE motiv-parada.des-parada
    FIELD log-alter-eficien LIKE motiv-parada.log-alter-eficien
    FIELD dat-inic-parada AS DATE LABEL "Data"
    FIELD qtd-segs-inic LIKE rep-parada-ctrab.qtd-segs-inic
    FIELD dat-fim-parada LIKE rep-parada-ctrab.dat-fim-parada
    FIELD qtd-segs-fim LIKE rep-parada-ctrab.qtd-segs-fim
    FIELD it-codigo LIKE split-operac.it-codigo
    FIELD inicio AS DATETIME
    FIELD fim AS DATETIME
    FIELD tempo AS DECIMAL LABEL "Tempo".

DEFINE NEW GLOBAL SHARED TEMP-TABLE tt-oee-men NO-UNDO
    FIELD mes-nome AS CHAR FORMAT "x(10)" LABEL "Mes"
    FIELD turno AS DECIMAL LABEL "Tempo Programado"
    FIELD cod-item LIKE split-operac.it-codigo LABEL "Item" 
    FIELD temp-prod AS DECIMAL LABEL "Tempo Produzindo"
    FIELD prod-teorica AS DECIMAL LABEL "Produá∆o Teorica"
    FIELD prod-real AS INT LABEL "Produá∆o Real"
    FIELD pc-boas AS INT LABEL "Peáas Boas"
    FIELD pc-ruins AS INT LABEL "Peáas Ruins"
    FIELD pc-boas-ruins AS INT LABEL "Peáas Boas + Ruins"
    FIELD temp-par-alt AS DECIMAL
    FIELD temp-par-nao-alt AS DECIMAL
    FIELD oee-disp AS DECIMAL LABEL "Disponibilidade"
    FIELD oee-perf AS DECIMAL LABEL "Performance"
    FIELD oee-quali AS DECIMAL LABEL "Qualidade"
    FIELD oee-percent AS DECIMAL LABEL "Percentual OEE"
    FIELD mes AS INT.

DEFINE NEW GLOBAL SHARED TEMP-TABLE tt-parada-men NO-UNDO
    FIELD cod-ctrab LIKE rep-parada-ctrab.cod-ctrab LABEL "teste"
    FIELD cod-parada AS CHAR //LIKE rep-parada-ctrab.cod-parada
    FIELD des-parada LIKE motiv-parada.des-parada
    FIELD log-alter-eficien LIKE motiv-parada.log-alter-eficien
    FIELD dat-inic-parada LIKE rep-parada-ctrab.dat-inic-parada
    FIELD qtd-segs-inic LIKE rep-parada-ctrab.qtd-segs-inic
    FIELD dat-fim-parada LIKE rep-parada-ctrab.dat-fim-parada
    FIELD qtd-segs-fim LIKE rep-parada-ctrab.qtd-segs-fim
    FIELD it-codigo LIKE split-operac.it-codigo
    FIELD inicio AS DATETIME
    FIELD fim AS DATETIME
    FIELD tempo AS DECIMAL
    FIELD mes AS INT.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 

/* ********************  Preprocessor Definitions  ******************** */

&Scoped-define PROCEDURE-TYPE Window
&Scoped-define DB-AWARE no

/* Name of designated FRAME-NAME and/or first browse and/or first query */
&Scoped-define FRAME-NAME DEFAULT-FRAME
&Scoped-define BROWSE-NAME APONTAMENTOS

/* Internal Tables (found by Frame, Query & Browse Queries)             */
&Scoped-define INTERNAL-TABLES tt-oee tt-oee-men tt-parada tt-parada-men

/* Definitions for BROWSE APONTAMENTOS                                  */
&Scoped-define FIELDS-IN-QUERY-APONTAMENTOS tt-oee.data tt-oee.turno tt-oee.temp-prod INT(truncate(tt-oee.prod-teorica,0)) tt-oee.prod-real tt-oee.pc-boas tt-oee.pc-boas-ruins   
&Scoped-define ENABLED-FIELDS-IN-QUERY-APONTAMENTOS   
&Scoped-define SELF-NAME APONTAMENTOS
&Scoped-define OPEN-QUERY-APONTAMENTOS //OPEN QUERY {&SELF-NAME} FOR EACH tt-oee.
&Scoped-define TABLES-IN-QUERY-APONTAMENTOS tt-oee
&Scoped-define FIRST-TABLE-IN-QUERY-APONTAMENTOS tt-oee


/* Definitions for BROWSE APONTAMENTOS-MES                              */
&Scoped-define FIELDS-IN-QUERY-APONTAMENTOS-MES tt-oee-men.mes-nome tt-oee-men.turno tt-oee-men.temp-prod int(TRUNCATE(tt-oee-men.prod-teorica,0)) tt-oee-men.prod-real tt-oee-men.pc-boas tt-oee-men.pc-boas-ruins   
&Scoped-define ENABLED-FIELDS-IN-QUERY-APONTAMENTOS-MES   
&Scoped-define SELF-NAME APONTAMENTOS-MES
&Scoped-define OPEN-QUERY-APONTAMENTOS-MES //OPEN QUERY {&SELF-NAME} FOR EACH tt-oee-men.
&Scoped-define TABLES-IN-QUERY-APONTAMENTOS-MES tt-oee-men
&Scoped-define FIRST-TABLE-IN-QUERY-APONTAMENTOS-MES tt-oee-men


/* Definitions for BROWSE Paradas                                       */
&Scoped-define FIELDS-IN-QUERY-Paradas tt-parada.dat-inic-parada tt-parada.cod-parada tt-parada.des-parada tt-parada.tempo   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Paradas   
&Scoped-define SELF-NAME Paradas
&Scoped-define OPEN-QUERY-Paradas //OPEN QUERY {&SELF-NAME} FOR EACH tt-parada.
&Scoped-define TABLES-IN-QUERY-Paradas tt-parada
&Scoped-define FIRST-TABLE-IN-QUERY-Paradas tt-parada


/* Definitions for BROWSE Paradas-mes                                   */
&Scoped-define FIELDS-IN-QUERY-Paradas-mes tt-parada-men.dat-inic-parada tt-parada-men.cod-parada tt-parada-men.des-parada tt-parada-men.tempo   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Paradas-mes   
&Scoped-define SELF-NAME Paradas-mes
&Scoped-define OPEN-QUERY-Paradas-mes //OPEN QUERY {&SELF-NAME} FOR EACH tt-parada-men.
&Scoped-define TABLES-IN-QUERY-Paradas-mes tt-parada-men
&Scoped-define FIRST-TABLE-IN-QUERY-Paradas-mes tt-parada-men


/* Definitions for FRAME DEFAULT-FRAME                                  */
&Scoped-define OPEN-BROWSERS-IN-QUERY-DEFAULT-FRAME ~
    ~{&OPEN-QUERY-APONTAMENTOS}~
    ~{&OPEN-QUERY-APONTAMENTOS-MES}~
    ~{&OPEN-QUERY-Paradas}~
    ~{&OPEN-QUERY-Paradas-mes}

/* Standard List Definitions                                            */
&Scoped-Define ENABLED-OBJECTS Paradas BUTTON-6 edt-centro btn-plano-acao ~
APONTAMENTOS cb-modo edt-data-ini edt-data-fim res-disp res-oee bt-con-dir ~
res-perf edt-ano res-quali bt-con-men bt-imprimir BUTTON-4 BUTTON-5 ~
edt-pc-ruins edt-pc-boa edt-item RECT-16 RECT-17 RECT-18 RECT-19 RECT-20 ~
RECT-23 RECT-24 RECT-25 RECT-26 RECT-27 RECT-28 RECT-29 RECT-30 img-oee ~
RECT-32 img-disp img-perf img-quali img-seta-e img-seta-d APONTAMENTOS-MES ~
Paradas-mes 
&Scoped-Define DISPLAYED-OBJECTS edt-centro cb-modo edt-data-ini ~
edt-data-fim edt-ano edt-pc-ruins edt-pc-boa edt-item 

/* Custom List Definitions                                              */
/* List-1,List-2,List-3,List-4,List-5,List-6                            */

/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME



/* ***********************  Control Definitions  ********************** */

/* Define the widget handle for the window                              */
DEFINE VAR C-Win AS WIDGET-HANDLE NO-UNDO.

/* Definitions of handles for OCX Containers                            */
DEFINE VARIABLE g AS WIDGET-HANDLE NO-UNDO.
DEFINE VARIABLE chg AS COMPONENT-HANDLE NO-UNDO.
DEFINE VARIABLE g2 AS WIDGET-HANDLE NO-UNDO.
DEFINE VARIABLE chg2 AS COMPONENT-HANDLE NO-UNDO.

/* Definitions of the field level widgets                               */
DEFINE BUTTON bt-con-dir 
     LABEL "Consultar" 
     SIZE 15 BY .96.

DEFINE BUTTON bt-con-men 
     LABEL "Consultar" 
     SIZE 15 BY .96.

DEFINE BUTTON bt-imprimir 
     IMAGE-UP FILE "adeicon/print.bmp":U
     LABEL "Imprimir" 
     SIZE 6.29 BY 1.13 TOOLTIP "Imprimir Relatorio".

DEFINE BUTTON btn-plano-acao 
     LABEL "Plano de Aá∆o" 
     SIZE 23 BY 1.13.

DEFINE BUTTON BUTTON-4 
     LABEL "ATUALIZA" 
     SIZE 15 BY .88.

DEFINE BUTTON BUTTON-5 
     LABEL "LIMPA" 
     SIZE 15 BY .88.

DEFINE BUTTON BUTTON-6 
     LABEL "Button 6" 
     SIZE 15 BY 1.13.

DEFINE BUTTON res-disp  NO-FOCUS FLAT-BUTTON NO-CONVERT-3D-COLORS
     LABEL "0 %" 
     SIZE 10.14 BY 1.33
     FGCOLOR 0 FONT 21.

DEFINE BUTTON res-oee  NO-FOCUS FLAT-BUTTON NO-CONVERT-3D-COLORS
     LABEL "0 %" 
     SIZE 24 BY 2
     FGCOLOR 0 FONT 20.

DEFINE BUTTON res-perf  NO-FOCUS FLAT-BUTTON NO-CONVERT-3D-COLORS
     LABEL "0 %" 
     SIZE 10.14 BY 1.33
     FGCOLOR 0 FONT 21.

DEFINE BUTTON res-quali  NO-FOCUS FLAT-BUTTON NO-CONVERT-3D-COLORS
     LABEL "0 %" 
     SIZE 10.14 BY 1.33
     FGCOLOR 0 FONT 21.

DEFINE VARIABLE cb-modo AS CHARACTER FORMAT "X(256)":U INITIAL "Di†rio" 
     VIEW-AS COMBO-BOX INNER-LINES 2
     LIST-ITEMS "Di†rio","Mensal" 
     DROP-DOWN-LIST
     SIZE 16 BY 1 NO-UNDO.

DEFINE VARIABLE edt-centro AS CHARACTER FORMAT "X(256)":U 
     LABEL "Centro de Trabalho" 
     VIEW-AS COMBO-BOX INNER-LINES 10
     DROP-DOWN-LIST
     SIZE 16 BY 1
     FONT 4 NO-UNDO.

DEFINE VARIABLE edt-item AS CHARACTER 
     VIEW-AS EDITOR
     SIZE 51 BY 2
     FONT 4 NO-UNDO.

DEFINE VARIABLE edt-pc-boa AS CHARACTER INITIAL "0" 
     VIEW-AS EDITOR NO-WORD-WRAP NO-BOX
     SIZE 22 BY 1.46
     BGCOLOR 15 FGCOLOR 10 FONT 22 NO-UNDO.

DEFINE VARIABLE edt-pc-ruins AS CHARACTER INITIAL "0" 
     VIEW-AS EDITOR NO-WORD-WRAP NO-BOX
     SIZE 22 BY 1.5
     BGCOLOR 15 FGCOLOR 12 FONT 22 NO-UNDO.

DEFINE VARIABLE edt-ano AS CHARACTER FORMAT "X(256)":U INITIAL "0" 
     LABEL "Ano" 
     VIEW-AS FILL-IN 
     SIZE 14 BY .88 NO-UNDO.

DEFINE VARIABLE edt-data-fim AS DATE FORMAT "99/99/99":U 
     VIEW-AS FILL-IN 
     SIZE 14 BY .88 NO-UNDO.

DEFINE VARIABLE edt-data-ini AS DATE FORMAT "99/99/99":U 
     LABEL "Data" 
     VIEW-AS FILL-IN 
     SIZE 13.57 BY .88 NO-UNDO.

DEFINE IMAGE img-disp
     FILENAME "adeicon/seta-verde.png":U
     STRETCH-TO-FIT
     SIZE 4 BY 1.25.

DEFINE IMAGE img-oee
     FILENAME "adeicon/seta-verde.png":U
     STRETCH-TO-FIT
     SIZE 5.57 BY 1.75.

DEFINE IMAGE img-perf
     FILENAME "adeicon/seta-verde.png":U
     STRETCH-TO-FIT
     SIZE 4 BY 1.25.

DEFINE IMAGE img-quali
     FILENAME "adeicon/seta-verde.png":U
     STRETCH-TO-FIT
     SIZE 4 BY 1.25.

DEFINE IMAGE img-seta-d
     FILENAME "adeicon/next-au.bmp":U
     SIZE 3 BY .75.

DEFINE IMAGE img-seta-e
     FILENAME "adeicon/prev-au.bmp":U
     SIZE 3 BY .75.

DEFINE RECTANGLE RECT-16
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 128.57 BY 13.

DEFINE RECTANGLE RECT-17
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 180.57 BY 1.5.

DEFINE RECTANGLE RECT-18
     EDGE-PIXELS 2 GRAPHIC-EDGE    
     SIZE 51 BY .96
     BGCOLOR 21 .

DEFINE RECTANGLE RECT-19
     EDGE-PIXELS 2 GRAPHIC-EDGE    
     SIZE 51 BY 2.38
     BGCOLOR 15 .

DEFINE RECTANGLE RECT-20
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 51 BY 1.67.

DEFINE RECTANGLE RECT-23
     EDGE-PIXELS 2 GRAPHIC-EDGE    
     SIZE 24 BY 1
     BGCOLOR 21 .

DEFINE RECTANGLE RECT-24
     EDGE-PIXELS 2 GRAPHIC-EDGE    
     SIZE 25 BY 1
     BGCOLOR 21 .

DEFINE RECTANGLE RECT-25
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 24 BY 1.88.

DEFINE RECTANGLE RECT-26
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 25 BY 1.88.

DEFINE RECTANGLE RECT-27
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 38 BY 12.25.

DEFINE RECTANGLE RECT-28
     EDGE-PIXELS 2 GRAPHIC-EDGE    
     SIZE 38 BY .83
     BGCOLOR 21 .

DEFINE RECTANGLE RECT-29
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 51 BY 1.67.

DEFINE RECTANGLE RECT-30
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 51 BY 1.67.

DEFINE RECTANGLE RECT-32
     EDGE-PIXELS 2 GRAPHIC-EDGE    
     SIZE 51 BY .83
     BGCOLOR 21 .

/* Query definitions                                                    */
&ANALYZE-SUSPEND
DEFINE QUERY APONTAMENTOS FOR 
      tt-oee SCROLLING.

DEFINE QUERY APONTAMENTOS-MES FOR 
      tt-oee-men SCROLLING.

DEFINE QUERY Paradas FOR 
      tt-parada SCROLLING.

DEFINE QUERY Paradas-mes FOR 
      tt-parada-men SCROLLING.
&ANALYZE-RESUME

/* Browse definitions                                                   */
DEFINE BROWSE APONTAMENTOS
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS APONTAMENTOS C-Win _FREEFORM
  QUERY APONTAMENTOS DISPLAY
      tt-oee.data
    tt-oee.turno
    tt-oee.temp-prod
    INT(truncate(tt-oee.prod-teorica,0)) LABEL "Produá∆o Te¢rica"
    tt-oee.prod-real
    tt-oee.pc-boas
    tt-oee.pc-boas-ruins
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS NO-TAB-STOP SIZE 89.43 BY 13
         FONT 4
         TITLE "Apontamentos de Produá∆o - Agrupados por Dia" FIT-LAST-COLUMN.

DEFINE BROWSE APONTAMENTOS-MES
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS APONTAMENTOS-MES C-Win _FREEFORM
  QUERY APONTAMENTOS-MES DISPLAY
      tt-oee-men.mes-nome
    tt-oee-men.turno
    tt-oee-men.temp-prod
    int(TRUNCATE(tt-oee-men.prod-teorica,0)) LABEL "Produá∆o Te¢rica" 
    tt-oee-men.prod-real
    tt-oee-men.pc-boas
    tt-oee-men.pc-boas-ruins
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS NO-TAB-STOP SIZE 89.43 BY 13
         FONT 4
         TITLE "Apontamentos de Produá∆o - Agrupados por Màs" FIT-LAST-COLUMN.

DEFINE BROWSE Paradas
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Paradas C-Win _FREEFORM
  QUERY Paradas DISPLAY
      tt-parada.dat-inic-parada
     tt-parada.cod-parada
     tt-parada.des-parada
     tt-parada.tempo
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS NO-TAB-STOP SIZE 51 BY 12.25
         FONT 4 ROW-HEIGHT-CHARS .46 FIT-LAST-COLUMN.

DEFINE BROWSE Paradas-mes
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Paradas-mes C-Win _FREEFORM
  QUERY Paradas-mes DISPLAY
      tt-parada-men.dat-inic-parada
     tt-parada-men.cod-parada
     tt-parada-men.des-parada
     tt-parada-men.tempo
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS NO-TAB-STOP SIZE 51 BY 12.25
         FONT 4 ROW-HEIGHT-CHARS .46 FIT-LAST-COLUMN.


/* ************************  Frame Definitions  *********************** */

DEFINE FRAME DEFAULT-FRAME
     Paradas AT ROW 14.25 COL 2 WIDGET-ID 300
     BUTTON-6 AT ROW 1.25 COL 115 WIDGET-ID 132
     edt-centro AT ROW 1.38 COL 14.14 COLON-ALIGNED WIDGET-ID 128
     btn-plano-acao AT ROW 1.25 COL 151 WIDGET-ID 130 NO-TAB-STOP 
     APONTAMENTOS AT ROW 15.75 COL 93 WIDGET-ID 200
     cb-modo AT ROW 1.38 COL 31.29 COLON-ALIGNED NO-LABEL WIDGET-ID 116
     edt-data-ini AT ROW 1.38 COL 52.14 COLON-ALIGNED WIDGET-ID 108
     edt-data-fim AT ROW 1.38 COL 71.72 COLON-ALIGNED NO-LABEL WIDGET-ID 110
     res-disp AT ROW 5.71 COL 41 WIDGET-ID 46 NO-TAB-STOP 
     res-oee AT ROW 3.46 COL 27.14 WIDGET-ID 42 NO-TAB-STOP 
     bt-con-dir AT ROW 1.33 COL 88.57 WIDGET-ID 4
     res-perf AT ROW 7.25 COL 41 WIDGET-ID 48 NO-TAB-STOP 
     edt-ano AT ROW 1.38 COL 51.72 COLON-ALIGNED WIDGET-ID 122
     res-quali AT ROW 8.79 COL 41 WIDGET-ID 54 NO-TAB-STOP 
     bt-con-men AT ROW 1.33 COL 69.29 WIDGET-ID 120
     bt-imprimir AT ROW 1.25 COL 174.72 WIDGET-ID 126 NO-TAB-STOP 
     BUTTON-4 AT ROW 1.38 COL 138 WIDGET-ID 6 NO-TAB-STOP 
     BUTTON-5 AT ROW 1.38 COL 154.43 WIDGET-ID 8 NO-TAB-STOP 
     edt-pc-ruins AT ROW 11.5 COL 29 NO-LABEL WIDGET-ID 72 NO-TAB-STOP 
     edt-pc-boa AT ROW 11.54 COL 3 NO-LABEL WIDGET-ID 70 NO-TAB-STOP 
     edt-item AT ROW 26.71 COL 2 NO-LABEL WIDGET-ID 74 NO-TAB-STOP 
     APONTAMENTOS-MES AT ROW 15.75 COL 93 WIDGET-ID 400
     Paradas-mes AT ROW 14.25 COL 2 WIDGET-ID 500
     "Percentuais" VIEW-AS TEXT
          SIZE 13.72 BY .58 AT ROW 2.67 COL 21.29 WIDGET-ID 78
          BGCOLOR 21 FONT 23
     "OEE" VIEW-AS TEXT
          SIZE 14 BY 1.75 AT ROW 3.58 COL 12.72 WIDGET-ID 36
          FGCOLOR 21 FONT 20
     "Peáas Boas" VIEW-AS TEXT
          SIZE 13 BY .54 AT ROW 10.71 COL 8.29 WIDGET-ID 60
          BGCOLOR 21 FONT 23
     "PERFORMANCE" VIEW-AS TEXT
          SIZE 28.29 BY 1.13 AT ROW 7.38 COL 12.72 WIDGET-ID 52
          FGCOLOR 21 FONT 21
     "QUALIDADE" VIEW-AS TEXT
          SIZE 28.29 BY 1.13 AT ROW 8.92 COL 12.72 WIDGET-ID 58
          FGCOLOR 21 FONT 21
     "DISPONIBILIDADE" VIEW-AS TEXT
          SIZE 28.29 BY 1.13 AT ROW 5.83 COL 12.72 WIDGET-ID 44
          FGCOLOR 21 FONT 21
     "Paradas do dia" VIEW-AS TEXT
          SIZE 13.72 BY .58 AT ROW 13.58 COL 19 WIDGET-ID 86
          BGCOLOR 21 FONT 23
     "Gr†fico Paradas do dia" VIEW-AS TEXT
          SIZE 21.43 BY .58 AT ROW 15.83 COL 62.14 WIDGET-ID 90
          BGCOLOR 21 FONT 23
     "Peáas Refugadas" VIEW-AS TEXT
          SIZE 18.14 BY .58 AT ROW 10.67 COL 31.72 WIDGET-ID 62
          BGCOLOR 21 FONT 23
     RECT-16 AT ROW 2.5 COL 54 WIDGET-ID 10
     RECT-17 AT ROW 1.08 COL 2 WIDGET-ID 12
     RECT-18 AT ROW 2.5 COL 2 WIDGET-ID 14
     RECT-19 AT ROW 3.33 COL 2 WIDGET-ID 16
     RECT-20 AT ROW 5.58 COL 2 WIDGET-ID 18
     RECT-23 AT ROW 10.5 COL 2 WIDGET-ID 24
     RECT-24 AT ROW 10.5 COL 28 WIDGET-ID 26
     RECT-25 AT ROW 11.38 COL 2 WIDGET-ID 28
     RECT-26 AT ROW 11.38 COL 28 WIDGET-ID 30
     RECT-27 AT ROW 16.5 COL 54 WIDGET-ID 32
     RECT-28 AT ROW 15.75 COL 54 WIDGET-ID 34
     RECT-29 AT ROW 7.13 COL 2 WIDGET-ID 50
     RECT-30 AT ROW 8.67 COL 2 WIDGET-ID 56
     img-oee AT ROW 3.58 COL 4.29 WIDGET-ID 92
     RECT-32 AT ROW 13.5 COL 2 WIDGET-ID 84
     img-disp AT ROW 5.75 COL 5 WIDGET-ID 94
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 1 ROW 1
         SIZE 182 BY 31.5
         BGCOLOR 15 FONT 4 WIDGET-ID 100.

/* DEFINE FRAME statement is approaching 4K Bytes.  Breaking it up   */
DEFINE FRAME DEFAULT-FRAME
     img-perf AT ROW 7.29 COL 5 WIDGET-ID 96
     img-quali AT ROW 8.79 COL 5 WIDGET-ID 98
     img-seta-e AT ROW 1.5 COL 68.43 WIDGET-ID 112
     img-seta-d AT ROW 1.5 COL 70.72 WIDGET-ID 114
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 1 ROW 1
         SIZE 182 BY 31.5
         BGCOLOR 15 FONT 4 WIDGET-ID 100.


/* *********************** Procedure Settings ************************ */

&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: Window
   Allow: Basic,Browse,DB-Fields,Window,Query
   Other Settings: COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS

/* *************************  Create Window  ************************** */

&ANALYZE-SUSPEND _CREATE-WINDOW
IF SESSION:DISPLAY-TYPE = "GUI":U THEN
  CREATE WINDOW C-Win ASSIGN
         HIDDEN             = YES
         TITLE              = "OEE"
         HEIGHT             = 28
         WIDTH              = 182.29
         MAX-HEIGHT         = 31.5
         MAX-WIDTH          = 182.29
         VIRTUAL-HEIGHT     = 31.5
         VIRTUAL-WIDTH      = 182.29
         RESIZE             = yes
         SCROLL-BARS        = no
         STATUS-AREA        = no
         BGCOLOR            = ?
         FGCOLOR            = ?
         KEEP-FRAME-Z-ORDER = yes
         THREE-D            = yes
         MESSAGE-AREA       = no
         SENSITIVE          = yes.
ELSE {&WINDOW-NAME} = CURRENT-WINDOW.
/* END WINDOW DEFINITION                                                */
&ANALYZE-RESUME



/* ***********  Runtime Attributes and AppBuilder Settings  *********** */

&ANALYZE-SUSPEND _RUN-TIME-ATTRIBUTES
/* SETTINGS FOR WINDOW C-Win
  VISIBLE,,RUN-PERSISTENT                                               */
/* SETTINGS FOR FRAME DEFAULT-FRAME
   FRAME-NAME Custom                                                    */
/* BROWSE-TAB Paradas 1 DEFAULT-FRAME */
/* BROWSE-TAB APONTAMENTOS btn-plano-acao DEFAULT-FRAME */
/* BROWSE-TAB APONTAMENTOS-MES img-seta-d DEFAULT-FRAME */
/* BROWSE-TAB Paradas-mes APONTAMENTOS-MES DEFAULT-FRAME */
ASSIGN 
       edt-item:READ-ONLY IN FRAME DEFAULT-FRAME        = TRUE.

ASSIGN 
       edt-pc-boa:READ-ONLY IN FRAME DEFAULT-FRAME        = TRUE.

ASSIGN 
       edt-pc-ruins:READ-ONLY IN FRAME DEFAULT-FRAME        = TRUE.

ASSIGN 
       res-disp:AUTO-RESIZE IN FRAME DEFAULT-FRAME      = TRUE.

ASSIGN 
       res-oee:AUTO-RESIZE IN FRAME DEFAULT-FRAME      = TRUE.

ASSIGN 
       res-perf:AUTO-RESIZE IN FRAME DEFAULT-FRAME      = TRUE.

ASSIGN 
       res-quali:AUTO-RESIZE IN FRAME DEFAULT-FRAME      = TRUE.

IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(C-Win)
THEN C-Win:HIDDEN = no.

/* _RUN-TIME-ATTRIBUTES-END */
&ANALYZE-RESUME


/* Setting information for Queries and Browse Widgets fields            */

&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE APONTAMENTOS
/* Query rebuild information for BROWSE APONTAMENTOS
     _START_FREEFORM
//OPEN QUERY {&SELF-NAME} FOR EACH tt-oee.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE APONTAMENTOS */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE APONTAMENTOS-MES
/* Query rebuild information for BROWSE APONTAMENTOS-MES
     _START_FREEFORM
//OPEN QUERY {&SELF-NAME} FOR EACH tt-oee-men.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE APONTAMENTOS-MES */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Paradas
/* Query rebuild information for BROWSE Paradas
     _START_FREEFORM
//OPEN QUERY {&SELF-NAME} FOR EACH tt-parada.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Paradas */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Paradas-mes
/* Query rebuild information for BROWSE Paradas-mes
     _START_FREEFORM
//OPEN QUERY {&SELF-NAME} FOR EACH tt-parada-men.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Paradas-mes */
&ANALYZE-RESUME

 


/* **********************  Create OCX Containers  ********************** */

&ANALYZE-SUSPEND _CREATE-DYNAMIC

&IF "{&OPSYS}" = "WIN32":U AND "{&WINDOW-SYSTEM}" NE "TTY":U &THEN

CREATE CONTROL-FRAME g ASSIGN
       FRAME           = FRAME DEFAULT-FRAME:HANDLE
       ROW             = 2.75
       COLUMN          = 55
       HEIGHT          = 12.42
       WIDTH           = 126
       TAB-STOP        = no
       WIDGET-ID       = 80
       HIDDEN          = no
       SENSITIVE       = yes.

CREATE CONTROL-FRAME g2 ASSIGN
       FRAME           = FRAME DEFAULT-FRAME:HANDLE
       ROW             = 17.5
       COLUMN          = 55
       HEIGHT          = 9.75
       WIDTH           = 36
       TAB-STOP        = no
       WIDGET-ID       = 104
       HIDDEN          = no
       SENSITIVE       = yes.
/* g OCXINFO:CREATE-CONTROL from: {0002E55D-0000-0000-C000-000000000046} type: ChartSpace */
/* g2 OCXINFO:CREATE-CONTROL from: {0002E55D-0000-0000-C000-000000000046} type: ChartSpace */

&ENDIF

&ANALYZE-RESUME /* End of _CREATE-DYNAMIC */


/* ************************  Control Triggers  ************************ */

&Scoped-define SELF-NAME C-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL C-Win C-Win
ON END-ERROR OF C-Win /* OEE */
OR ENDKEY OF {&WINDOW-NAME} ANYWHERE DO:
  /* This case occurs when the user presses the "Esc" key.
     In a persistently run window, just ignore this.  If we did not, the
     application would exit. */
  IF THIS-PROCEDURE:PERSISTENT THEN RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL C-Win C-Win
ON WINDOW-CLOSE OF C-Win /* OEE */
DO:
  /* This event will close the window and terminate the procedure.  */
  APPLY "CLOSE":U TO THIS-PROCEDURE.
  RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define BROWSE-NAME APONTAMENTOS
&Scoped-define SELF-NAME APONTAMENTOS
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL APONTAMENTOS C-Win
ON VALUE-CHANGED OF APONTAMENTOS IN FRAME DEFAULT-FRAME /* Apontamentos de Produá∆o - Agrupados por Dia */
DO:
  ASSIGN data-atu-oee = date(APONTAMENTOS:GET-BROWSE-COLUMN(1):SCREEN-VALUE IN FRAME {&FRAME-NAME}).

  RUN carrega-parada.
  RUN grafico_par.

  FIND FIRST tt-oee WHERE tt-oee.data = data-atu-oee NO-LOCK NO-ERROR.

  IF AVAIL tt-oee THEN DO:
      
      //PERCENTUAL OEE
      res-oee:WIDTH = 24.
      res-oee:HEIGHT = 2.
      res-oee:LABEL = string(tt-oee.oee-percent) + " %".

      IF tt-oee.oee-percent >= 75 THEN
          img-oee:LOAD-IMAGE("adeicon/seta-verde.png").
      ELSE
          img-oee:LOAD-IMAGE("adeicon/seta-vermelha.png").

      //DISPONIBILIDADE
      res-disp:WIDTH = 10.14.
      res-disp:HEIGHT = 1.33.
      res-disp:LABEL = string(tt-oee.oee-disp) + " %".

      IF tt-oee.oee-disp >= 91 THEN
          img-disp:LOAD-IMAGE("adeicon/seta-verde.png").
      ELSE
          img-disp:LOAD-IMAGE("adeicon/seta-vermelha.png").

      //PERFORMANCE
      res-perf:WIDTH = 10.14.
      res-perf:HEIGHT = 1.33.
      res-perf:LABEL = string(tt-oee.oee-perf) + " %".

      IF tt-oee.oee-perf >= 91 THEN
          img-perf:LOAD-IMAGE("adeicon/seta-verde.png").
      ELSE
          img-perf:LOAD-IMAGE("adeicon/seta-vermelha.png").

      //QUALIDADE
      res-quali:WIDTH = 10.14.
      res-quali:HEIGHT = 1.33.
      res-quali:LABEL = string(tt-oee.oee-quali) + " %".

      IF tt-oee.oee-quali >= 91 THEN
          img-quali:LOAD-IMAGE("adeicon/seta-verde.png").
      ELSE
          img-quali:LOAD-IMAGE("adeicon/seta-vermelha.png").


      edt-pc-boa:SCREEN-VALUE = STRING(tt-oee.pc-boas).
      edt-pc-ruins:SCREEN-VALUE = STRING(tt-oee.pc-ruins).
      

  END.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define BROWSE-NAME APONTAMENTOS-MES
&Scoped-define SELF-NAME APONTAMENTOS-MES
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL APONTAMENTOS-MES C-Win
ON VALUE-CHANGED OF APONTAMENTOS-MES IN FRAME DEFAULT-FRAME /* Apontamentos de Produá∆o - Agrupados por Màs */
DO:
  ASSIGN mes-atu-oee = STRING(APONTAMENTOS-MES:GET-BROWSE-COLUMN(1):SCREEN-VALUE IN FRAME {&FRAME-NAME}).
  
  IF mes-atu-oee = "Janeiro" THEN DO:
    ASSIGN i-mes = 1.
  END.

  IF mes-atu-oee = "Fevereiro" THEN DO:
    ASSIGN i-mes = 2.
  END.

  IF mes-atu-oee = "Maráo" THEN DO:
    ASSIGN i-mes = 3.
  END.

  IF mes-atu-oee = "Abril" THEN DO:
    ASSIGN i-mes = 4.
  END.

  IF mes-atu-oee = "Maio" THEN DO:
    ASSIGN i-mes = 5.
  END.

  IF mes-atu-oee = "Junho" THEN DO:
    ASSIGN i-mes = 6.
  END.

  IF mes-atu-oee = "Julho" THEN DO:
    ASSIGN i-mes = 7.
  END.

  IF mes-atu-oee = "Agosto" THEN DO:
    ASSIGN i-mes = 8.
  END.

  IF mes-atu-oee = "Setembro" THEN DO:
    ASSIGN i-mes = 9.
  END.

  IF mes-atu-oee = "Outubro" THEN DO:
    ASSIGN i-mes = 10.
  END.

  IF mes-atu-oee = "Novembro" THEN DO:
    ASSIGN i-mes = 11.
  END.

  IF mes-atu-oee = "Dezembro" THEN DO:
    ASSIGN i-mes = 12.
  END.

  RUN carrega-parada.
  RUN grafico_par_men.

  FIND FIRST tt-oee-men WHERE tt-oee-men.mes-nome = mes-atu-oee NO-LOCK NO-ERROR.

  IF AVAIL tt-oee-men THEN DO:

      //PERCENTUAL OEE
      res-oee:WIDTH = 24.
      res-oee:HEIGHT = 2.
      res-oee:LABEL = string(tt-oee-men.oee-percent) + " %".

      IF tt-oee-men.oee-percent >= 75 THEN
          img-oee:LOAD-IMAGE("adeicon/seta-verde.png").
      ELSE
          img-oee:LOAD-IMAGE("adeicon/seta-vermelha.png").

      //DISPONIBILIDADE
      res-disp:WIDTH = 10.14.
      res-disp:HEIGHT = 1.33.
      res-disp:LABEL = string(tt-oee-men.oee-disp) + " %".

      IF tt-oee-men.oee-disp >= 91 THEN
          img-disp:LOAD-IMAGE("adeicon/seta-verde.png").
      ELSE
          img-disp:LOAD-IMAGE("adeicon/seta-vermelha.png").

      //PERFORMANCE
      res-perf:WIDTH = 10.14.
      res-perf:HEIGHT = 1.33.
      res-perf:LABEL = string(tt-oee-men.oee-perf) + " %".

      IF tt-oee-men.oee-perf >= 91 THEN
          img-perf:LOAD-IMAGE("adeicon/seta-verde.png").
      ELSE
          img-perf:LOAD-IMAGE("adeicon/seta-vermelha.png").

      //QUALIDADE
      res-quali:WIDTH = 10.14.
      res-quali:HEIGHT = 1.33.
      res-quali:LABEL = string(tt-oee-men.oee-quali) + " %".

      IF tt-oee-men.oee-quali >= 91 THEN
          img-quali:LOAD-IMAGE("adeicon/seta-verde.png").
      ELSE
          img-quali:LOAD-IMAGE("adeicon/seta-vermelha.png").


      edt-pc-boa:SCREEN-VALUE = STRING(tt-oee-men.pc-boas).
      edt-pc-ruins:SCREEN-VALUE = STRING(tt-oee-men.pc-ruins).

  END.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME bt-con-dir
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-con-dir C-Win
ON CHOOSE OF bt-con-dir IN FRAME DEFAULT-FRAME /* Consultar */
DO:
  RUN LIMPA_TEMP-TABLE.
  IF edt-centro:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "SELECIONE" THEN DO:
      MESSAGE "Selecione um centro de trabalho.".
  END.

  ELSE DO:
      ASSIGN centro = string(edt-centro:SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      ASSIGN data-ini = date(edt-data-ini:SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      ASSIGN data-fim = date(edt-data-fim:SCREEN-VALUE IN FRAME {&FRAME-NAME}).

      FIND FIRST rep-oper-ctrab NO-LOCK 
          WHERE rep-oper-ctrab.cod-ctrab = centro
          AND rep-oper-ctrab.dat-inic-report >= data-ini
          AND rep-oper-ctrab.dat-fim-report <= data-fim NO-ERROR.

      IF AVAIL rep-oper-ctrab THEN DO:

          RUN "_brandl/Oee - Diario.p".
          OPEN QUERY APONTAMENTOS FOR EACH tt-oee WHERE tt-oee.temp-prod > 0 BY tt-oee.data.
          RUN grafico_ap.
          RUN grafico_par.
          RUN item-parada.
          RUN carrega-parada.

      END.

      IF NOT AVAIL rep-oper-ctrab THEN DO:
        MESSAGE "Nenhum registro encontrado." VIEW-AS ALERT-BOX.
      END.
      

  END.
  
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME bt-con-men
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-con-men C-Win
ON CHOOSE OF bt-con-men IN FRAME DEFAULT-FRAME /* Consultar */
DO:
  RUN LIMPA_TEMP-TABLE.

  IF edt-centro:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "SELECIONE" THEN DO:
      MESSAGE "Selecione um centro de trabalho.".
  END.

  ELSE DO:
      ASSIGN centro = string(edt-centro:SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      ASSIGN oee-ano = INT(edt-ano:SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      
      RUN "_brandl/Oee - Mensal.p".
      OPEN QUERY APONTAMENTOS-MES FOR EACH tt-oee-men WHERE tt-oee-men.temp-prod > 0.
      RUN grafico_ap_men.
      RUN grafico_par_men.
      //RUN item-parada.
      RUN carrega-parada.
    
      //MESSAGE "Em desenvolvimento..." VIEW-AS ALERT-BOX.
  END.
    
   
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME bt-imprimir
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-imprimir C-Win
ON CHOOSE OF bt-imprimir IN FRAME DEFAULT-FRAME /* Imprimir */
DO:

    ASSIGN centro = string(edt-centro:SCREEN-VALUE IN FRAME {&FRAME-NAME}).
    ASSIGN data-ini = date(edt-data-ini:SCREEN-VALUE IN FRAME {&FRAME-NAME}).
    ASSIGN data-fim = date(edt-data-fim:SCREEN-VALUE IN FRAME {&FRAME-NAME}).
    
    IF impressao = 1 THEN DO:
        RUN "sfc/sfcp0010.w".
    END.

    IF impressao = 2 THEN DO:
        SYSTEM-DIALOG GET-DIR cFile
        TITLE "Escolha uma pasta de destino"
        INITIAL-DIR "C:\Temp".

        RUN relatorio-mensal.
    END.
    
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME btn-plano-acao
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL btn-plano-acao C-Win
ON CHOOSE OF btn-plano-acao IN FRAME DEFAULT-FRAME /* Plano de Aá∆o */
DO:
  RUN "_brandl/plano-acao.r".
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME BUTTON-4
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL BUTTON-4 C-Win
ON CHOOSE OF BUTTON-4 IN FRAME DEFAULT-FRAME /* ATUALIZA */
DO:
  OPEN QUERY APONTAMENTOS FOR EACH tt-oee WHERE tt-oee.temp-prod > 0.
  RUN carrega-parada.
  RUN grafico_ap.
  RUN grafico_par.
  RUN item-parada.
  
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME BUTTON-5
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL BUTTON-5 C-Win
ON CHOOSE OF BUTTON-5 IN FRAME DEFAULT-FRAME /* LIMPA */
DO:
  RUN LIMPA_TEMP-TABLE.
  OPEN QUERY APONTAMENTOS FOR EACH tt-oee WHERE tt-oee.temp-prod > 0.
  RUN carrega-parada.
  RUN grafico_ap.
  RUN grafico_par.
  RUN item-parada.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME BUTTON-6
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL BUTTON-6 C-Win
ON CHOOSE OF BUTTON-6 IN FRAME DEFAULT-FRAME /* Button 6 */
DO:
  
    SYSTEM-DIALOG GET-DIR cFile
    TITLE "Escolha uma pasta de destino"
    INITIAL-DIR "C:\Temp".

    RUN relatorio-diario.
  
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME cb-modo
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL cb-modo C-Win
ON VALUE-CHANGED OF cb-modo IN FRAME DEFAULT-FRAME
DO:
  RUN Oculta-itens.

END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME edt-centro
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL edt-centro C-Win
ON VALUE-CHANGED OF edt-centro IN FRAME DEFAULT-FRAME /* Centro de Trabalho */
DO:
  BUTTON-4:VISIBLE IN FRAME {&FRAME-NAME}= FALSE.
  BUTTON-5:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.

   IF cb-modo:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Mensal" THEN DO:
      ASSIGN impressao = 2.
      edt-data-ini:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      edt-data-fim:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      bt-con-dir:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      img-seta-e:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      img-seta-d:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      APONTAMENTOS:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      Paradas:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.

      edt-ano:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      bt-con-men:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      APONTAMENTOS-MES:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      Paradas-mes:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.

      edt-ano:SCREEN-VALUE IN FRAME {&FRAME-NAME} = string(YEAR(DATE(TODAY))).

      RUN LIMPA_TEMP-TABLE.
      OPEN QUERY APONTAMENTOS FOR EACH tt-oee WHERE tt-oee.temp-prod > 0.
      RUN carrega-parada.
      RUN grafico_ap.
      RUN grafico_par.
      RUN item-parada.

  END.

  IF cb-modo:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Di†rio" THEN DO:
      ASSIGN impressao = 1.
      edt-data-ini:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      edt-data-fim:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      bt-con-dir:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      img-seta-e:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      img-seta-d:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      APONTAMENTOS-MES:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      Paradas-mes:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.

      edt-ano:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      bt-con-men:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      APONTAMENTOS:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      Paradas:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.

      RUN LIMPA_TEMP-TABLE.
      OPEN QUERY APONTAMENTOS FOR EACH tt-oee WHERE tt-oee.temp-prod > 0.
      RUN carrega-parada.
      RUN grafico_ap.
      RUN grafico_par.
      RUN item-parada.

  END.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define BROWSE-NAME Paradas
&Scoped-define SELF-NAME Paradas
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Paradas C-Win
ON VALUE-CHANGED OF Paradas IN FRAME DEFAULT-FRAME
DO:
    RUN item-parada.

END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define BROWSE-NAME Paradas-mes
&Scoped-define SELF-NAME Paradas-mes
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Paradas-mes C-Win
ON VALUE-CHANGED OF Paradas-mes IN FRAME DEFAULT-FRAME
DO:
    RUN item-parada.

END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define BROWSE-NAME APONTAMENTOS
&UNDEFINE SELF-NAME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK C-Win 


/* ***************************  Main Block  *************************** */

/* Set CURRENT-WINDOW: this will parent dialog-boxes and frames.        */
ASSIGN CURRENT-WINDOW                = {&WINDOW-NAME} 
       THIS-PROCEDURE:CURRENT-WINDOW = {&WINDOW-NAME}.

/* The CLOSE event can be used from inside or outside the procedure to  */
/* terminate it.                                                        */
ON CLOSE OF THIS-PROCEDURE 
   RUN disable_UI.

/* Best default for GUI applications is...                              */
PAUSE 0 BEFORE-HIDE.

/* Now enable the interface and wait for the exit condition.            */
/* (NOTE: handle ERROR and END-KEY so cleanup code will always fire.    */
MAIN-BLOCK:
DO ON ERROR   UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK
   ON END-KEY UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK:
  RUN enable_UI.
  IF NOT THIS-PROCEDURE:PERSISTENT THEN
    WAIT-FOR CLOSE OF THIS-PROCEDURE.
END.

RUN LIMPA_TEMP-TABLE.
RUN Oculta-itens.
RUN carrega-ctrab.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


/* **********************  Internal Procedures  *********************** */

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE carrega-ctrab C-Win 
PROCEDURE carrega-ctrab :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
  DEF VAR lista AS CHAR.
  DEF VAR controle AS INT.

  ASSIGN lista = "SELECIONE"
         controle = 1.

  FOR EACH ctrab.
    
/*       IF controle = 1 THEN DO:                                  */
/*           ASSIGN lista = STRING(ctrab.cod-ctrab).               */
/*           ASSIGN controle = 2.                                  */
/*       END.                                                      */
/*                                                                 */
/*       IF controle = 2 THEN DO:                                  */
/*          ASSIGN  lista = lista + "," + STRING(ctrab.cod-ctrab). */
/*       END.                                                      */

      ASSIGN  lista = lista + "," + STRING(ctrab.cod-ctrab).
  END.

  edt-centro:LIST-ITEMS IN FRAME {&FRAME-NAME} = lista.
  edt-centro:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "SELECIONE".
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE carrega-parada C-Win 
PROCEDURE carrega-parada :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
IF cb-modo:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Di†rio" THEN DO:
    OPEN QUERY Paradas FOR EACH tt-parada WHERE tt-parada.dat-inic-parada = data-atu-oee.
END.

ELSE IF cb-modo:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Mensal" THEN DO:
    OPEN QUERY Paradas-mes FOR EACH tt-parada-men WHERE tt-parada-men.mes = i-mes.
END.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE control_load C-Win  _CONTROL-LOAD
PROCEDURE control_load :
/*------------------------------------------------------------------------------
  Purpose:     Load the OCXs    
  Parameters:  <none>
  Notes:       Here we load, initialize and make visible the 
               OCXs in the interface.                        
------------------------------------------------------------------------------*/

&IF "{&OPSYS}" = "WIN32":U AND "{&WINDOW-SYSTEM}" NE "TTY":U &THEN
DEFINE VARIABLE UIB_S    AS LOGICAL    NO-UNDO.
DEFINE VARIABLE OCXFile  AS CHARACTER  NO-UNDO.

OCXFile = SEARCH( "Oee.wrx":U ).
IF OCXFile = ? THEN
  OCXFile = SEARCH(SUBSTRING(THIS-PROCEDURE:FILE-NAME, 1,
                     R-INDEX(THIS-PROCEDURE:FILE-NAME, ".":U), "CHARACTER":U) + "wrx":U).

IF OCXFile <> ? THEN
DO:
  ASSIGN
    chg = g:COM-HANDLE
    UIB_S = chg:LoadControls( OCXFile, "g":U)
    g:NAME = "g":U
    chg2 = g2:COM-HANDLE
    UIB_S = chg2:LoadControls( OCXFile, "g2":U)
    g2:NAME = "g2":U
  .
  RUN initialize-controls IN THIS-PROCEDURE NO-ERROR.
END.
ELSE MESSAGE "Oee.wrx":U SKIP(1)
             "The binary control file could not be found. The controls cannot be loaded."
             VIEW-AS ALERT-BOX TITLE "Controls Not Loaded".

&ENDIF

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE disable_UI C-Win  _DEFAULT-DISABLE
PROCEDURE disable_UI :
/*------------------------------------------------------------------------------
  Purpose:     DISABLE the User Interface
  Parameters:  <none>
  Notes:       Here we clean-up the user-interface by deleting
               dynamic widgets we have created and/or hide 
               frames.  This procedure is usually called when
               we are ready to "clean-up" after running.
------------------------------------------------------------------------------*/
  /* Delete the WINDOW we created */
  IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(C-Win)
  THEN DELETE WIDGET C-Win.
  IF THIS-PROCEDURE:PERSISTENT THEN DELETE PROCEDURE THIS-PROCEDURE.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE enable_UI C-Win  _DEFAULT-ENABLE
PROCEDURE enable_UI :
/*------------------------------------------------------------------------------
  Purpose:     ENABLE the User Interface
  Parameters:  <none>
  Notes:       Here we display/view/enable the widgets in the
               user-interface.  In addition, OPEN all queries
               associated with each FRAME and BROWSE.
               These statements here are based on the "Other 
               Settings" section of the widget Property Sheets.
------------------------------------------------------------------------------*/
  RUN control_load.
  DISPLAY edt-centro cb-modo edt-data-ini edt-data-fim edt-ano edt-pc-ruins 
          edt-pc-boa edt-item 
      WITH FRAME DEFAULT-FRAME IN WINDOW C-Win.
  ENABLE Paradas BUTTON-6 edt-centro btn-plano-acao APONTAMENTOS cb-modo 
         edt-data-ini edt-data-fim res-disp res-oee bt-con-dir res-perf edt-ano 
         res-quali bt-con-men bt-imprimir BUTTON-4 BUTTON-5 edt-pc-ruins 
         edt-pc-boa edt-item RECT-16 RECT-17 RECT-18 RECT-19 RECT-20 RECT-23 
         RECT-24 RECT-25 RECT-26 RECT-27 RECT-28 RECT-29 RECT-30 img-oee 
         RECT-32 img-disp img-perf img-quali img-seta-e img-seta-d 
         APONTAMENTOS-MES Paradas-mes 
      WITH FRAME DEFAULT-FRAME IN WINDOW C-Win.
  {&OPEN-BROWSERS-IN-QUERY-DEFAULT-FRAME}
  VIEW C-Win.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE grafico_ap C-Win 
PROCEDURE grafico_ap :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
def var c_nome as c.
def var i_val  as c.
DEF VAR controle_conteudo AS INT.

IF limpa-grafico-ap = TRUE THEN DO:
    chg:chartspace:charts:DELETE(0). //Limpa Grafico anterior
END.

chg:chartspace:haschartspacetitle = TRUE. //Exibe o Titulo do Grafico

chg:chartspace:HasChartSpaceLegend = FALSE. //Exibe as Legendas

chg:chartspace:chartspacetitle:caption = "PERCENTUAL OEE". //Define Titulo do grafico

chg:chartspace:charts:add(0). //Adiciona uma area de grafico

chg:chartspace:charts(0):SeriesCollection:add(0). //Cria uma Serie

chg:chartspace:Charts(0):SeriesCollection(0):DataLabelsCollection:ADD. //Adiciona Labels em cada item do grafico

chg:chartspace:charts(0):SeriesCollection(0):Border:Color = "BLACK". //Muda a cor da borda da barra
chg:chartspace:charts(0):SeriesCollection(0):Interior:Color = "RED". //Muda a cor do preenchimento da barra

chg:chartspace:charts(0):SeriesCollection(0):Interior:SetTwoColorGradient(chg:chartspace:constants:chGradientHorizontal, chg:chartspace:constants:chGradientVariantStart, "Blue", "black").
    
FOR EACH tt-oee WHERE tt-oee.temp-prod > 0 NO-LOCK.

    //Passando Valores do Banco para o Grafico
    IF controle_conteudo = 0 THEN DO:

        ASSIGN data-atu-oee = tt-oee.data.
        assign c_nome = string(tt-oee.data).
        assign i_val = string(tt-oee.oee-percent).

        res-oee:WIDTH IN FRAME {&FRAME-NAME} = 24.
        res-oee:HEIGHT IN FRAME {&FRAME-NAME} = 2.
        res-oee:LABEL IN FRAME {&FRAME-NAME} = string(tt-oee.oee-percent) + " %".
    
        IF tt-oee.oee-percent >= 75 THEN
            img-oee:LOAD-IMAGE("adeicon/seta-verde.png").
        ELSE
            img-oee:LOAD-IMAGE("adeicon/seta-vermelha.png").

        //DISPONIBILIDADE
        res-disp:WIDTH = 10.14.
        res-disp:HEIGHT = 1.33.
        res-disp:LABEL = string(tt-oee.oee-disp) + " %".
    
        IF tt-oee.oee-disp >= 91 THEN
            img-disp:LOAD-IMAGE("adeicon/seta-verde.png").
        ELSE
            img-disp:LOAD-IMAGE("adeicon/seta-vermelha.png").
    
        //PERFORMANCE
        res-perf:WIDTH = 10.14.
        res-perf:HEIGHT = 1.33.
        res-perf:LABEL = string(tt-oee.oee-perf) + " %".
    
        IF tt-oee.oee-perf >= 91 THEN
            img-perf:LOAD-IMAGE("adeicon/seta-verde.png").
        ELSE
            img-perf:LOAD-IMAGE("adeicon/seta-vermelha.png").
    
        //QUALIDADE
        res-quali:WIDTH = 10.14.
        res-quali:HEIGHT = 1.33.
        res-quali:LABEL = string(tt-oee.oee-quali) + " %".
    
        IF tt-oee.oee-quali >= 91 THEN
            img-quali:LOAD-IMAGE("adeicon/seta-verde.png").
        ELSE
            img-quali:LOAD-IMAGE("adeicon/seta-vermelha.png").
    
    
        edt-pc-boa:SCREEN-VALUE = STRING(tt-oee.pc-boas).
        edt-pc-ruins:SCREEN-VALUE = STRING(tt-oee.pc-ruins).

        ASSIGN controle_conteudo = 1.

    END.

    ELSE DO:
        assign c_nome = c_nome + "," + string(tt-oee.data).
        assign i_val = i_val + "," + string(tt-oee.oee-percent).
    END.
   
end.

assign c_nome = c_nome + ",".
assign i_val = i_val + ",".

ASSIGN controle_conteudo = 0.
chg:chartspace:Charts(0):SeriesCollection(0):SetData(chg:chartspace:constants:chDimCategories,chg:chartspace:constants:chDataLiteral,c_nome).
chg:chartspace:Charts(0):SeriesCollection(0):SetData(chg:chartspace:constants:chDimValues,chg:chartspace:constants:chDataLiteral,i_val).

//Tipo de exibicao do grafico, consultar modelos pelo VB usando crtl=space.
//chg:chartspace:Charts(0):TYPE = chg:chartspace:constants:chChartTypePie3D.
chg:chartspace:Charts(0):TYPE = chg:chartspace:constants:chChartTypeColumnClustered.

ASSIGN limpa-grafico-ap = TRUE.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE grafico_ap_men C-Win 
PROCEDURE grafico_ap_men :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
def var c_nome as c.
def var i_val  as c.
DEF VAR controle_conteudo AS INT.

IF limpa-grafico-ap = TRUE THEN DO:
    chg:chartspace:charts:DELETE(0). //Limpa Grafico anterior
END.

chg:chartspace:haschartspacetitle = TRUE. //Exibe o Titulo do Grafico

chg:chartspace:HasChartSpaceLegend = FALSE. //Exibe as Legendas

chg:chartspace:chartspacetitle:caption = "PERCENTUAL OEE". //Define Titulo do grafico

chg:chartspace:charts:add(0). //Adiciona uma area de grafico

chg:chartspace:charts(0):SeriesCollection:add(0). //Cria uma Serie

chg:chartspace:Charts(0):SeriesCollection(0):DataLabelsCollection:ADD. //Adiciona Labels em cada item do grafico

chg:chartspace:charts(0):SeriesCollection(0):Border:Color = "BLACK". //Muda a cor da borda da barra
chg:chartspace:charts(0):SeriesCollection(0):Interior:Color = "RED". //Muda a cor do preenchimento da barra

chg:chartspace:charts(0):SeriesCollection(0):Interior:SetTwoColorGradient(chg:chartspace:constants:chGradientHorizontal, chg:chartspace:constants:chGradientVariantStart, "Blue", "black").
    
FOR EACH tt-oee-men WHERE tt-oee-men.temp-prod > 0 NO-LOCK.

    //Passando Valores do Banco para o Grafico
    IF controle_conteudo = 0 THEN DO:

        ASSIGN oee-mes = tt-oee-men.mes-nome.
        ASSIGN i-mes = tt-oee-men.mes.
        assign c_nome = string(tt-oee-men.mes-nome).
        assign i_val = string(tt-oee-men.oee-percent).

        res-oee:WIDTH IN FRAME {&FRAME-NAME} = 24.
        res-oee:HEIGHT IN FRAME {&FRAME-NAME} = 2.
        res-oee:LABEL IN FRAME {&FRAME-NAME} = string(tt-oee-men.oee-percent) + " %".
    
        IF tt-oee-men.oee-percent >= 75 THEN
            img-oee:LOAD-IMAGE("adeicon/seta-verde.png").
        ELSE
            img-oee:LOAD-IMAGE("adeicon/seta-vermelha.png").

        //DISPONIBILIDADE
        res-disp:WIDTH = 10.14.
        res-disp:HEIGHT = 1.33.
        res-disp:LABEL = string(tt-oee-men.oee-disp) + " %".
    
        IF tt-oee-men.oee-disp >= 91 THEN
            img-disp:LOAD-IMAGE("adeicon/seta-verde.png").
        ELSE
            img-disp:LOAD-IMAGE("adeicon/seta-vermelha.png").
    
        //PERFORMANCE
        res-perf:WIDTH = 10.14.
        res-perf:HEIGHT = 1.33.
        res-perf:LABEL = string(tt-oee-men.oee-perf) + " %".
    
        IF tt-oee-men.oee-perf >= 91 THEN
            img-perf:LOAD-IMAGE("adeicon/seta-verde.png").
        ELSE
            img-perf:LOAD-IMAGE("adeicon/seta-vermelha.png").
    
        //QUALIDADE
        res-quali:WIDTH = 10.14.
        res-quali:HEIGHT = 1.33.
        res-quali:LABEL = string(tt-oee-men.oee-quali) + " %".
    
        IF tt-oee-men.oee-quali >= 91 THEN
            img-quali:LOAD-IMAGE("adeicon/seta-verde.png").
        ELSE
            img-quali:LOAD-IMAGE("adeicon/seta-vermelha.png").
    
    
        edt-pc-boa:SCREEN-VALUE = STRING(tt-oee-men.pc-boas).
        edt-pc-ruins:SCREEN-VALUE = STRING(tt-oee-men.pc-ruins).

        ASSIGN controle_conteudo = 1.

    END.

    ELSE DO:
        assign c_nome = c_nome + "," + string(tt-oee-men.mes-nome).
        assign i_val = i_val + "," + string(tt-oee-men.oee-percent).
    END.
   
end.

assign c_nome = c_nome + ",".
assign i_val = i_val + ",".

ASSIGN controle_conteudo = 0.
chg:chartspace:Charts(0):SeriesCollection(0):SetData(chg:chartspace:constants:chDimCategories,chg:chartspace:constants:chDataLiteral,c_nome).
chg:chartspace:Charts(0):SeriesCollection(0):SetData(chg:chartspace:constants:chDimValues,chg:chartspace:constants:chDataLiteral,i_val).

//Tipo de exibicao do grafico, consultar modelos pelo VB usando crtl=space.
//chg:chartspace:Charts(0):TYPE = chg:chartspace:constants:chChartTypePie3D.
chg:chartspace:Charts(0):TYPE = chg:chartspace:constants:chChartTypeColumnClustered.

ASSIGN limpa-grafico-ap = TRUE.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE grafico_par C-Win 
PROCEDURE grafico_par :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
def var c_nome as c.
def var i_val  as c.
DEF VAR soma-paradas AS DEC.
DEF VAR controle_conteudo AS INT.

IF limpa-grafico-par = TRUE THEN DO:
    chg2:chartspace:charts:DELETE(0). //Limpa Grafico anterior
END.

chg2:chartspace:haschartspacetitle = TRUE. //Exibe o Titulo do Grafico

chg2:chartspace:HasChartSpaceLegend = TRUE. //Exibe as Legendas


chg2:chartspace:chartspacetitle:caption = "Percentual Diario de Paradas". //Define Titulo do grafico

chg2:chartspace:charts:add(0). //Adiciona uma area de grafico

chg2:chartspace:charts(0):SeriesCollection:add(0). //Cria uma Serie

chg2:chartspace:Charts(0):SeriesCollection(0):DataLabelsCollection:ADD. //Adiciona Labels em cada item do grafico

DEF VAR c_nome2 AS CHAR.
DEF VAR i_val2 AS CHAR.
DEF VAR tempo-somado AS DEC.

FOR EACH tt-parada NO-LOCK WHERE tt-parada.dat-inic-parada = data-atu-oee.

    ASSIGN soma-paradas = soma-paradas + tt-parada.tempo.

END.


FOR EACH tt-parada NO-LOCK WHERE tt-parada.dat-inic-parada = data-atu-oee BREAK BY tt-parada.cod-parada.


    tempo-somado = tempo-somado + tt-parada.tempo.

    IF LAST-OF(tt-parada.cod-parada) THEN DO:
        //Passando Valores do Banco para o Grafico
        IF controle_conteudo = 0 THEN DO:

            assign c_nome = string(tt-parada.cod-parada).
            assign i_val = string(ROUND(((tempo-somado / soma-paradas) * 100),0)).

            ASSIGN controle_conteudo = 1.
        END.

        ELSE DO:
            assign c_nome = c_nome + "," + string(tt-parada.cod-parada).
            assign i_val = i_val + "," + string(ROUND(((tempo-somado / soma-paradas) * 100),0)).
        END.


        ASSIGN tempo-somado = 0.

    END.
   
END.


ASSIGN controle_conteudo = 0.


chg2:chartspace:Charts(0):SeriesCollection(0):SetData(chg2:chartspace:constants:chDimCategories,chg2:chartspace:constants:chDataLiteral,c_nome).
chg2:chartspace:Charts(0):SeriesCollection(0):SetData(chg2:chartspace:constants:chDimValues,chg2:chartspace:constants:chDataLiteral,i_val).

//Tipo de exibicao do grafico, consultar modelos pelo VB usando crtl=space.
//chg:chartspace:Charts(0):TYPE = chg:chartspace:constants:chChartTypePie3D.
chg2:chartspace:Charts(0):TYPE = chg2:chartspace:constants:chChartTypePie3D.

ASSIGN limpa-grafico-par = TRUE.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE grafico_par_men C-Win 
PROCEDURE grafico_par_men :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
def var c_nome as c.
def var i_val  as c.
DEF VAR soma-paradas AS DEC.
DEF VAR controle_conteudo AS INT.

IF limpa-grafico-par = TRUE THEN DO:
    chg2:chartspace:charts:DELETE(0). //Limpa Grafico anterior
END.

chg2:chartspace:haschartspacetitle = TRUE. //Exibe o Titulo do Grafico

chg2:chartspace:HasChartSpaceLegend = TRUE. //Exibe as Legendas


chg2:chartspace:chartspacetitle:caption = "Percentual Diario de Paradas". //Define Titulo do grafico

chg2:chartspace:charts:add(0). //Adiciona uma area de grafico

chg2:chartspace:charts(0):SeriesCollection:add(0). //Cria uma Serie

chg2:chartspace:Charts(0):SeriesCollection(0):DataLabelsCollection:ADD. //Adiciona Labels em cada item do grafico

DEF VAR c_nome2 AS CHAR.
DEF VAR i_val2 AS CHAR.
DEF VAR tempo-somado AS DEC.

FOR EACH tt-parada-men NO-LOCK WHERE tt-parada-men.mes = i-mes.

    ASSIGN soma-paradas = soma-paradas + tt-parada-men.tempo.

END.


FOR EACH tt-parada-men NO-LOCK WHERE tt-parada-men.mes = i-mes BREAK BY tt-parada-men.cod-parada.


    tempo-somado = tempo-somado + tt-parada-men.tempo.

    IF LAST-OF(tt-parada-men.cod-parada) THEN DO:
        //Passando Valores do Banco para o Grafico
        IF controle_conteudo = 0 THEN DO:

            assign c_nome = string(tt-parada-men.cod-parada).
            assign i_val = string(ROUND(((tempo-somado / soma-paradas) * 100),0)).

            ASSIGN controle_conteudo = 1.
        END.

        ELSE DO:
            assign c_nome = c_nome + "," + string(tt-parada-men.cod-parada).
            assign i_val = i_val + "," + string(ROUND(((tempo-somado / soma-paradas) * 100),0)).
        END.


        ASSIGN tempo-somado = 0.

    END.
   
END.


ASSIGN controle_conteudo = 0.


chg2:chartspace:Charts(0):SeriesCollection(0):SetData(chg2:chartspace:constants:chDimCategories,chg2:chartspace:constants:chDataLiteral,c_nome).
chg2:chartspace:Charts(0):SeriesCollection(0):SetData(chg2:chartspace:constants:chDimValues,chg2:chartspace:constants:chDataLiteral,i_val).

//Tipo de exibicao do grafico, consultar modelos pelo VB usando crtl=space.
//chg:chartspace:Charts(0):TYPE = chg:chartspace:constants:chChartTypePie3D.
chg2:chartspace:Charts(0):TYPE = chg2:chartspace:constants:chChartTypePie3D.

ASSIGN limpa-grafico-par = TRUE.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE item-parada C-Win 
PROCEDURE item-parada :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
  edt-item:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "".

  DEF VAR data AS DATE.
  DEF VAR cod AS CHAR.
  DEF VAR descricao AS CHAR.
  DEF VAR tempo-total AS DECIMAL.

  IF cb-modo:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Di†rio" THEN DO:
      ASSIGN data = date(PARADAS:GET-BROWSE-COLUMN(1):SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      ASSIGN cod = STRING(PARADAS:GET-BROWSE-COLUMN(2):SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      ASSIGN descricao = STRING(PARADAS:GET-BROWSE-COLUMN(3):SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      ASSIGN tempo-total = DEC(PARADAS:GET-BROWSE-COLUMN(4):SCREEN-VALUE IN FRAME {&FRAME-NAME}).
    
      FOR FIRST tt-parada 
          WHERE tt-parada.dat-inic-parada = data
            AND tt-parada.cod-parada = cod
            AND tt-parada.des-parada = descricao
            AND tt-parada.tempo = tempo-total,
    
          FIRST motiv-parada WHERE motiv-parada.cod-parada = tt-parada.cod-parada,
          FIRST ITEM WHERE ITEM.it-codigo = tt-parada.it-codigo NO-LOCK.
    
          IF AVAIL tt-parada THEN DO:
              IF motiv-parada.log-alter-eficien = TRUE THEN DO:
                  edt-item:SCREEN-VALUE IN FRAME {&FRAME-NAME} = tt-parada.it-codigo + " - " + ITEM.desc-item.
              END.
        
              ELSE DO:
                  edt-item:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Paradas que n∆o alteram eficiància n∆o possuem item.".
              END.
              
              
          END.
    
      END.
  END.

  IF cb-modo:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Mensal" THEN DO:
      ASSIGN data = date(PARADAS-MES:GET-BROWSE-COLUMN(1):SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      ASSIGN cod = STRING(PARADAS-MES:GET-BROWSE-COLUMN(2):SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      ASSIGN descricao = STRING(PARADAS-MES:GET-BROWSE-COLUMN(3):SCREEN-VALUE IN FRAME {&FRAME-NAME}).
      ASSIGN tempo-total = DEC(PARADAS-MES:GET-BROWSE-COLUMN(4):SCREEN-VALUE IN FRAME {&FRAME-NAME}).
    
      FOR FIRST tt-parada-men 
          WHERE tt-parada-men.dat-inic-parada = data
            AND tt-parada-men.cod-parada = cod
            AND tt-parada-men.des-parada = descricao
            AND tt-parada-men.tempo = tempo-total,
    
          FIRST motiv-parada WHERE motiv-parada.cod-parada = tt-parada-men.cod-parada,
          FIRST ITEM WHERE ITEM.it-codigo = tt-parada-men.it-codigo NO-LOCK.
    
          IF AVAIL tt-parada-men THEN DO:
              IF motiv-parada.log-alter-eficien = TRUE THEN DO:
                  edt-item:SCREEN-VALUE IN FRAME {&FRAME-NAME} = tt-parada-men.it-codigo + " - " + ITEM.desc-item.
              END.
        
              ELSE DO:
                  edt-item:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Paradas que n∆o alteram eficiància n∆o possuem item.".
              END.
              
              
          END.
    
      END.
  END.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE LIMPA_TEMP-TABLE C-Win 
PROCEDURE LIMPA_TEMP-TABLE :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/

EMPTY TEMP-TABLE tt-oee.
EMPTY TEMP-TABLE tt-parada.
EMPTY TEMP-TABLE tt-oee-men.

res-oee:LABEL IN FRAME {&FRAME-NAME} = "0 %".
res-disp:LABEL IN FRAME {&FRAME-NAME} = "0 %".
res-perf:LABEL IN FRAME {&FRAME-NAME} = "0 %".
res-quali:LABEL IN FRAME {&FRAME-NAME} = "0 %".

res-oee:WIDTH IN FRAME {&FRAME-NAME} = 24.
res-oee:HEIGHT IN FRAME {&FRAME-NAME} = 2.

res-disp:WIDTH = 10.14.
res-disp:HEIGHT = 1.33.

res-perf:WIDTH = 10.14.
res-perf:HEIGHT = 1.33.

res-quali:WIDTH = 10.14.
res-quali:HEIGHT = 1.33.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE Oculta-itens C-Win 
PROCEDURE Oculta-itens :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
BUTTON-4:VISIBLE IN FRAME {&FRAME-NAME}= FALSE.
BUTTON-5:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.

   IF cb-modo:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Mensal" THEN DO:
      ASSIGN impressao = 2.
      edt-data-ini:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      edt-data-fim:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      bt-con-dir:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      img-seta-e:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      img-seta-d:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      APONTAMENTOS:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      Paradas:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.

      edt-ano:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      bt-con-men:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      APONTAMENTOS-MES:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      Paradas-mes:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.

      edt-ano:SCREEN-VALUE IN FRAME {&FRAME-NAME} = string(YEAR(DATE(TODAY))).

      RUN LIMPA_TEMP-TABLE.
      OPEN QUERY APONTAMENTOS FOR EACH tt-oee WHERE tt-oee.temp-prod > 0.
      RUN carrega-parada.
      RUN grafico_ap.
      RUN grafico_par.
      RUN item-parada.

  END.

  IF cb-modo:SCREEN-VALUE IN FRAME {&FRAME-NAME} = "Di†rio" THEN DO:
      ASSIGN impressao = 1.
      edt-data-ini:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      edt-data-fim:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      bt-con-dir:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      img-seta-e:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      img-seta-d:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      APONTAMENTOS-MES:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      Paradas-mes:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.

      edt-ano:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      bt-con-men:VISIBLE IN FRAME {&FRAME-NAME} = FALSE.
      APONTAMENTOS:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.
      Paradas:VISIBLE IN FRAME {&FRAME-NAME} = TRUE.

      RUN LIMPA_TEMP-TABLE.
      OPEN QUERY APONTAMENTOS FOR EACH tt-oee WHERE tt-oee.temp-prod > 0.
      RUN carrega-parada.
      RUN grafico_ap.
      RUN grafico_par.
      RUN item-parada.

      ASSIGN data-ini = DATE(MONTH(TODAY),01,YEAR(TODAY)).
      IF MONTH(TODAY) = 12 THEN
        ASSIGN data-fim =  DATE(MONTH(TODAY) + 1 ,01,YEAR(TODAY) + 1).
      ELSE
          ASSIGN data-fim =  DATE(MONTH(TODAY) + 1 ,01,YEAR(TODAY)).
      
      ASSIGN data-fim = data-fim - 1.
          ASSIGN edt-data-ini:SCREEN-VALUE IN FRAME {&FRAME-NAME} = STRING(data-ini)
                 edt-data-fim:SCREEN-VALUE IN FRAME {&FRAME-NAME} = STRING(data-fim).

  END.


END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE relatorio-diario C-Win 
PROCEDURE relatorio-diario :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEF VAR ex AS COM-HANDLE NO-UNDO.
DEF VAR ChWorkSheet AS COM-HANDLE NO-UNDO.
DEF VAR contador AS INTEGER.
DEF VAR contador-par AS INTEGER.
DEF VAR cont-par-alt-efic AS INTEGER.
DEF VAR cont-par-n-alt-efic AS INTEGER.
DEF VAR qtd-parada AS INTEGER.
DEF VAR tempo-parada AS DECIMAL.
DEF VAR extensao AS CHAR FORMAT "x(4)".
DEF VAR ARQUIVO AS CHAR FORMAT "X(50)".

ASSIGN contador = 12
       contador-par = 12
       cont-par-alt-efic = 63
       cont-par-n-alt-efic = 59.

CREATE "excel.application" ex.

ex:VISIBLE = FALSE.
ex:workbooks:ADD("C:\Users\renan.cano\Desktop\Progress-20210126T140217Z-001\Progress\00 - Desenvolvimento\00 - Programas\Gerador de OEE\Interface Grafica\layout-diario.xlt").

//RELATORIO APONTAMENTO DE PRODUÄ«O E GRAFICOS
ex:worksheets:ITEM(1):SELECT.

FOR EACH tt-oee WHERE tt-oee.temp-prod > 0 NO-LOCK.
    //ANO
    ex:range("B2"):VALUE = "PER÷ODO: " + string(data-ini) + " - " + STRING(data-fim).
    
    //Centro de Trabalho - Nome da Prensa
    ex:range("B3"):VALUE = "CENTRO:  " + centro.
    
    //Mes do apontamento
    ex:range("A" + STRING(contador)):VALUE = tt-oee.data.
    
    //Tempo Programado
    ex:range("B" + STRING(contador)):VALUE = tt-oee.turno.
    
    //Tempo Produzindo
    ex:range("C" + STRING(contador)):VALUE = tt-oee.temp-prod.
    
    //Produá∆o Teorica
    ex:range("D" + STRING(contador)):VALUE = tt-oee.prod-teorica.
    
    //Produá∆o Real
    ex:range("E" + STRING(contador)):VALUE = tt-oee.prod-real.
    
    //Peáas Boas
    ex:range("F" + STRING(contador)):VALUE = tt-oee.pc-boas.
    
    //Peáas Boas + Ruins
    ex:range("G" + STRING(contador)):VALUE = tt-oee.pc-boas-ruins.
    
    //Disponibilidade
    ex:range("H" + STRING(contador)):VALUE = tt-oee.oee-disp.
    
    //Performance
    ex:range("I" + STRING(contador)):VALUE = tt-oee.oee-perf.
    
    //Qualidade
    ex:range("J" + STRING(contador)):VALUE = tt-oee.oee-quali.
    
    //Oee
    ex:range("K" + STRING(contador)):VALUE = tt-oee.oee-percent.

    ASSIGN contador = contador + 1.

END.


//RELATORIO DE PARADAS
ex:worksheets:ITEM(3):SELECT.

FOR EACH tt-parada.
        
    ex:range(string(contador-par) + ":" + string(contador-par)):INSERT().

    //DATA
    ex:range("A" + STRING(contador-par)):VALUE = tt-parada.dat-inic-parada.

    //ITEM DA PARADA
    ex:range("B" + STRING(contador-par)):VALUE = tt-parada.it-codigo.

    //CODIGO PARADA
    ex:range("D" + STRING(contador-par)):VALUE = tt-parada.cod-parada.

    //DESCRIÄ«O DA PARADA
    ex:range("E" + STRING(contador-par)):VALUE = tt-parada.des-parada.

    //TEMPO DE PARADA
    ex:range("F" + STRING(contador-par)):VALUE = tt-parada.tempo.
    
    ASSIGN contador-par = contador-par + 1.

END.

ex:range(string(11) + ":" + string(11)):DELETE().
ex:range(string(contador-par - 1) + ":" + string(contador-par - 1)):DELETE().



//GRAFICO PARADAS N«O ALTERAM EFICIENCIA
ex:worksheets:ITEM(4):SELECT.
DEF VAR tempo-par-n-alt-efic AS DECIMAL.
ASSIGN tempo-par-n-alt-efic = 0.

FOR EACH tt-parada WHERE tt-parada.log-alter-eficien = NO BREAK BY tt-parada.cod-parada.


    IF FIRST-OF(tt-parada.cod-parada) THEN DO:
        ASSIGN cont-par-n-alt-efic = cont-par-n-alt-efic + 1.
        ex:range(string(cont-par-n-alt-efic) + ":" + string(cont-par-n-alt-efic)):INSERT().
    END.
    
    //SOMA PARADAS QUE N«O ALTERAM EFICIENCIA
    tempo-par-n-alt-efic = tempo-par-n-alt-efic + tt-parada.tempo.
    
    IF LAST-OF(tt-parada.cod-parada) THEN DO:
        ex:range("A" + STRING(cont-par-n-alt-efic)):VALUE = tt-parada.des-parada.
        ex:range("B" + STRING(cont-par-n-alt-efic)):VALUE = tempo-par-n-alt-efic.

        ASSIGN tempo-par-n-alt-efic = 0.

    END.

END.

ex:range(string(59) + ":" + string(59)):DELETE().
ex:range(string(cont-par-n-alt-efic) + ":" + string(cont-par-n-alt-efic)):DELETE().



//GRAFICO PARADAS ALTERAM EFICIENCIA
ex:worksheets:ITEM(4):SELECT.
DEF VAR tempo-par-alt-efic AS DECIMAL.
ASSIGN tempo-par-alt-efic = 0.

FOR EACH tt-parada WHERE tt-parada.log-alter-eficien = YES BREAK BY tt-parada.cod-parada.


    IF FIRST-OF(tt-parada.cod-parada) THEN DO:
        ASSIGN cont-par-alt-efic = cont-par-alt-efic + 1.
        ex:range(string(cont-par-alt-efic) + ":" + string(cont-par-alt-efic)):INSERT().
    END.
    
    //SOMA PARADAS QUE N«O ALTERAM EFICIENCIA
    tempo-par-alt-efic = tempo-par-alt-efic + tt-parada.tempo.
    
    IF LAST-OF(tt-parada.cod-parada) THEN DO:
        ex:range("A" + STRING(cont-par-alt-efic)):VALUE = tt-parada.des-parada.
        ex:range("B" + STRING(cont-par-alt-efic)):VALUE = tempo-par-alt-efic.

        ASSIGN tempo-par-alt-efic = 0.

    END.

END.

ex:range(string(63) + ":" + string(63)):DELETE().
ex:range(string(cont-par-alt-efic) + ":" + string(cont-par-alt-efic)):DELETE().


//SALVANDO ARQUIVO
ASSIGN extensao = ".pdf"
       arquivo = "Relatorio Diario OEE - " + string(oee-ano).

ex:ActiveWorkbook:ExportAsFixedFormat(0,cFile + "\" + arquivo + extensao,,,,,,,).

MESSAGE "Arquivo salvo em " + cFile + "\" + arquivo + extensao VIEW-AS ALERT-BOX.

ex:APPLICATION:DisplayAlerts = FALSE.

//ex:VISIBLE = TRUE.
ex:QUIT().

RELEASE OBJECT ex NO-ERROR.
RELEASE OBJECT chWorksheet NO-ERROR.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE relatorio-mensal C-Win 
PROCEDURE relatorio-mensal :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEF VAR ex AS COM-HANDLE NO-UNDO.
DEF VAR ChWorkSheet AS COM-HANDLE NO-UNDO.
DEF VAR contador AS INTEGER.
DEF VAR extensao AS CHAR FORMAT "x(4)".
DEF VAR ARQUIVO AS CHAR FORMAT "X(50)".

ASSIGN contador = 12.

CREATE "excel.application" ex.

ex:VISIBLE = FALSE.
ex:workbooks:ADD("\\server03\ERP\_custom\_brandl\layout\layout-mensal.xlt").


FOR EACH tt-oee-men WHERE tt-oee-men.temp-prod > 0 NO-LOCK.
    //ANO
    ex:range("B2"):VALUE = "ANO: " + string(oee-ano).
    
    //Centro de Trabalho - Nome da Prensa
    ex:range("B3"):VALUE = "CENTRO:  " + centro.
    
    //Mes do apontamento
    ex:range("A" + STRING(contador)):VALUE = tt-oee-men.mes-nome.
    
    //Tempo Programado
    ex:range("B" + STRING(contador)):VALUE = tt-oee-men.turno.
    
    //Tempo Produzindo
    ex:range("C" + STRING(contador)):VALUE = tt-oee-men.temp-prod.
    
    //Produá∆o Teorica
    ex:range("D" + STRING(contador)):VALUE = tt-oee-men.prod-teorica.
    
    //Produá∆o Real
    ex:range("E" + STRING(contador)):VALUE = tt-oee-men.prod-real.
    
    //Peáas Boas
    ex:range("F" + STRING(contador)):VALUE = tt-oee-men.pc-boas.
    
    //Peáas Boas + Ruins
    ex:range("G" + STRING(contador)):VALUE = tt-oee-men.pc-boas-ruins.
    
    //Disponibilidade
    ex:range("H" + STRING(contador)):VALUE = tt-oee-men.oee-disp.
    
    //Performance
    ex:range("I" + STRING(contador)):VALUE = tt-oee-men.oee-perf.
    
    //Qualidade
    ex:range("J" + STRING(contador)):VALUE = tt-oee-men.oee-quali.
    
    //Oee
    ex:range("K" + STRING(contador)):VALUE = tt-oee-men.oee-percent.

    ASSIGN contador = contador + 1.

END.

ASSIGN extensao = ".pdf"
       arquivo = "Relatorio Mensal OEE - " + string(oee-ano).

ex:ActiveWorkbook:ExportAsFixedFormat(0,cFile + "\" + arquivo + extensao,,,,,,,).

MESSAGE "Arquivo salvo em " + cFile + "\" + arquivo + extensao VIEW-AS ALERT-BOX.

ex:APPLICATION:DisplayAlerts = FALSE.

ex:QUIT().

RELEASE OBJECT ex NO-ERROR.
RELEASE OBJECT chWorksheet NO-ERROR.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

