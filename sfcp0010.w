&ANALYZE-SUSPEND _VERSION-NUMBER UIB_v8r12 GUI ADM1
&ANALYZE-RESUME
/* Connected Databases 
          mgcad            PROGRESS
*/
&Scoped-define WINDOW-NAME w-relat
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS w-relat 
/*:T*******************************************************************************
** Copyright TOTVS S.A. (2009)
** Todos os Direitos Reservados.
**
** Este fonte e de propriedade exclusiva da TOTVS, sua reproducao
** parcial ou total por qualquer meio, so podera ser feita mediante
** autorizacao expressa.
*******************************************************************************/
{include/i-prgvrs.i SFCP0010 12.1.23.001}
{utp/ut-glob.i}
/* Chamada a include do gerenciador de licenáas. Necessario alterar os parametros */
/*                                                                                */
/* <programa>:  Informar qual o nome do programa.                                 */
/* <m¢dulo>:  Informar qual o m¢dulo a qual o programa pertence.                  */

&IF "{&EMSFND_VERSION}" >= "1.00" &THEN
    {include/i-license-manager.i <programa> <m¢dulo>}
&ENDIF

/* Create an unnamed pool to store all the widgets created 
     by this procedure. This is a good default which assures
     that this procedure's triggers and internal procedures 
     will execute in this procedure's storage, and that proper
     cleanup will occur on deletion of the procedure. */

CREATE WIDGET-POOL.

DEF VAR c-nom-prog-upc-mg97 AS c.
DEF VAR c-nom-prog-appc-mg97 AS c.
DEF VAR c-programa-mg97 AS c.
DEF VAR c-nom-prog-dpc-mg97 AS c.
DEF VAR c-versao-mg97 AS c.

/* ***************************  Definitions  ************************** */

/* Usuario */
DEF NEW GLOBAL SHARED VAR v_cod_usuar_corren  AS CHAR NO-UNDO.

/*:T Preprocessadores do Template de Relat¢rio                            */
/*:T Obs: Retirar o valor do preprocessador para as p†ginas que n∆o existirem  */

&GLOBAL-DEFINE PGSEL f-pg-sel
&GLOBAL-DEFINE PGCLA 
&GLOBAL-DEFINE PGPAR f-pg-par
&GLOBAL-DEFINE PGDIG f-pg-dig
&GLOBAL-DEFINE PGIMP f-pg-imp

&GLOBAL-DEFINE RTF   NO
  
/* Parameters Definitions ---                                           */

DEF VAR organiza AS INT INIT 1.

/* Temporary Table Definitions ---                                      */

define temp-table tt-param no-undo
    field destino          as integer
    field arquivo          as char format "x(35)"
    field usuario          as char format "x(12)"
    field data-exec        as date
    field hora-exec        as integer
    field classifica       as integer
    field dt_dt_ini        as date
    field dt_dt_fim        as date
    field c_estab_ini      as CHAR
    field c_estab_fim      as CHAR
    field c_cod-ctrab_ini  as CHAR
    field c_cod-ctrab_fim  as CHAR
    field c_gm-codigo_ini  as CHAR
    field c_gm-codigo_fim  as CHAR
    field c_item_ini       as CHAR
    field c_item_fim       as CHAR
    field l-diario         AS LOG   
    field l-mensal         AS LOG
    field desc-classifica  as char format "x(40)"
    field modelo-rtf       as char format "x(35)"
    field l-habilitaRtf    as LOG.

define temp-table tt-digita no-undo
    field CENTRO               as character format "x(16)"
    field DESC-CENTRO          as character format "x(30)"
    field GRUPO                as character format "x(09)"
    field DESC-GRUPO           as character format "x(30)"
    index id CENTRO GRUPO.

define buffer b-tt-digita for tt-digita.

/* Transfer Definitions */

def var raw-param        as raw no-undo.

def temp-table tt-raw-digita
   field raw-digita      as raw.
                    
/* Local Variable Definitions ---                                       */

DEFINE NEW GLOBAL SHARED VAR data-ini AS DATE NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR data-fim AS DATE NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR centro AS CHAR NO-UNDO.

def var l-ok               as logical no-undo.
def var c-arq-digita       as char    no-undo.
def var c-terminal         as char    no-undo.
def var c-rtf              as char    no-undo.
def var c-modelo-default   as char    no-undo.

DEFINE VARIABLE dt-ini AS DATE     NO-UNDO.
DEFINE VARIABLE dt-fim AS DATE     NO-UNDO.
/*15/02/2005 - tech1007 - Variavel definida para tratar se o programa est† rodando no WebEnabler*/
DEFINE SHARED VARIABLE hWenController AS HANDLE NO-UNDO.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 

/* ********************  Preprocessor Definitions  ******************** */

&Scoped-define PROCEDURE-TYPE w-relat
&Scoped-define DB-AWARE no

/* Name of designated FRAME-NAME and/or first browse and/or first query */
&Scoped-define FRAME-NAME f-pg-dig
&Scoped-define BROWSE-NAME br-centros

/* Internal Tables (found by Frame, Query & Browse Queries)             */
&Scoped-define INTERNAL-TABLES ctrab tt-digita

/* Definitions for BROWSE br-centros                                    */
&Scoped-define FIELDS-IN-QUERY-br-centros ctrab.cod-ctrab 
&Scoped-define ENABLED-FIELDS-IN-QUERY-br-centros 
&Scoped-define QUERY-STRING-br-centros FOR EACH ctrab NO-LOCK INDEXED-REPOSITION
&Scoped-define OPEN-QUERY-br-centros OPEN QUERY br-centros FOR EACH ctrab NO-LOCK INDEXED-REPOSITION.
&Scoped-define TABLES-IN-QUERY-br-centros ctrab
&Scoped-define FIRST-TABLE-IN-QUERY-br-centros ctrab


/* Definitions for BROWSE br-digita                                     */
&Scoped-define FIELDS-IN-QUERY-br-digita tt-digita.CENTRO tt-digita.DESC-CENTRO tt-digita.GRUPO tt-digita.DESC-GRUPO   
&Scoped-define ENABLED-FIELDS-IN-QUERY-br-digita tt-digita.CENTRO tt-digita.GRUPO   
&Scoped-define ENABLED-TABLES-IN-QUERY-br-digita tt-digita
&Scoped-define FIRST-ENABLED-TABLE-IN-QUERY-br-digita tt-digita
&Scoped-define SELF-NAME br-digita
&Scoped-define QUERY-STRING-br-digita FOR EACH tt-digita
&Scoped-define OPEN-QUERY-br-digita OPEN QUERY br-digita FOR EACH tt-digita.
&Scoped-define TABLES-IN-QUERY-br-digita tt-digita
&Scoped-define FIRST-TABLE-IN-QUERY-br-digita tt-digita


/* Definitions for FRAME f-ctrab                                        */
&Scoped-define OPEN-BROWSERS-IN-QUERY-f-ctrab ~
    ~{&OPEN-QUERY-br-centros}

/* Definitions for FRAME f-pg-dig                                       */
&Scoped-define OPEN-BROWSERS-IN-QUERY-f-pg-dig ~
    ~{&OPEN-QUERY-br-digita}
&Scoped-define SELF-NAME f-pg-dig
&Scoped-define OPEN-QUERY-f-pg-dig FRAME f-pg-dig:VISIBLE = FALSE.

/* Definitions for FRAME f-pg-imp                                       */
&Scoped-define SELF-NAME f-pg-imp
&Scoped-define OPEN-QUERY-f-pg-imp FRAME f-pg-imp:VISIBLE = FALSE.

/* Definitions for FRAME f-pg-par                                       */
&Scoped-define SELF-NAME f-pg-par
&Scoped-define OPEN-QUERY-f-pg-par FRAME f-pg-par:VISIBLE = FALSE.

/* Standard List Definitions                                            */
&Scoped-Define ENABLED-OBJECTS br-digita bt-inserir bt-recuperar 

/* Custom List Definitions                                              */
/* List-1,List-2,List-3,List-4,List-5,List-6                            */

/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME



/* ***********************  Control Definitions  ********************** */

/* Define the widget handle for the window                              */
DEFINE VAR w-relat AS WIDGET-HANDLE NO-UNDO.

/* Definitions of handles for OCX Containers                            */
DEFINE VARIABLE CtrlFrame AS WIDGET-HANDLE NO-UNDO.
DEFINE VARIABLE chCtrlFrame AS COMPONENT-HANDLE NO-UNDO.

/* Definitions of the field level widgets                               */
DEFINE BUTTON bt-alterar 
     LABEL "Alterar" 
     SIZE 15 BY 1.

DEFINE BUTTON bt-inserir 
     LABEL "Inserir" 
     SIZE 15 BY 1.

DEFINE BUTTON bt-recuperar 
     LABEL "Recuperar" 
     SIZE 13 BY 1.

DEFINE BUTTON bt-retirar 
     LABEL "Retirar" 
     SIZE 15 BY 1.

DEFINE BUTTON bt-salvar 
     LABEL "Salvar" 
     SIZE 15 BY 1.

DEFINE BUTTON bt-arquivo 
     IMAGE-UP FILE "image\im-sea":U
     IMAGE-INSENSITIVE FILE "image\ii-sea":U
     LABEL "" 
     SIZE 4 BY 1.

DEFINE BUTTON bt-config-impr 
     IMAGE-UP FILE "image\im-cfprt":U
     LABEL "" 
     SIZE 4 BY 1.

DEFINE BUTTON bt-modelo-rtf 
     IMAGE-UP FILE "image\im-sea":U
     IMAGE-INSENSITIVE FILE "image\ii-sea":U
     LABEL "" 
     SIZE 4 BY 1.

DEFINE VARIABLE c-arquivo AS CHARACTER 
     VIEW-AS EDITOR MAX-CHARS 256
     SIZE 40 BY .88
     BGCOLOR 15  NO-UNDO.

DEFINE VARIABLE c-modelo-rtf AS CHARACTER 
     VIEW-AS EDITOR MAX-CHARS 256
     SIZE 40 BY .88
     BGCOLOR 15  NO-UNDO.

DEFINE VARIABLE text-destino AS CHARACTER FORMAT "X(256)":U INITIAL " Destino" 
      VIEW-AS TEXT 
     SIZE 8.57 BY .63 NO-UNDO.

DEFINE VARIABLE text-modelo-rtf AS CHARACTER FORMAT "X(256)":U INITIAL "Modelo:" 
      VIEW-AS TEXT 
     SIZE 10.86 BY .63 NO-UNDO.

DEFINE VARIABLE text-modo AS CHARACTER FORMAT "X(256)":U INITIAL "Execuá∆o" 
      VIEW-AS TEXT 
     SIZE 10.86 BY .63 NO-UNDO.

DEFINE VARIABLE text-rtf AS CHARACTER FORMAT "X(256)":U INITIAL "Rich Text Format(RTF)" 
      VIEW-AS TEXT 
     SIZE 20.86 BY .63 NO-UNDO.

DEFINE VARIABLE rs-destino AS INTEGER INITIAL 2 
     VIEW-AS RADIO-SET HORIZONTAL
     RADIO-BUTTONS 
          "Impressora", 1,
"Arquivo", 2,
"Terminal", 3
     SIZE 44 BY 1.08 NO-UNDO.

DEFINE VARIABLE rs-execucao AS INTEGER INITIAL 1 
     VIEW-AS RADIO-SET HORIZONTAL
     RADIO-BUTTONS 
          "On-Line", 1,
"Batch", 2
     SIZE 27.86 BY .92 NO-UNDO.

DEFINE RECTANGLE RECT-7
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 46.14 BY 2.79.

DEFINE RECTANGLE RECT-9
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 46.14 BY 1.71.

DEFINE RECTANGLE rect-rtf
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 46.14 BY 3.5.

DEFINE VARIABLE l-habilitaRtf AS LOGICAL INITIAL no 
     LABEL "RTF" 
     VIEW-AS TOGGLE-BOX
     SIZE 44 BY 1.08 NO-UNDO.

DEFINE VARIABLE tg-1 AS LOGICAL INITIAL yes 
     LABEL "Por per°odo di†rio" 
     VIEW-AS TOGGLE-BOX
     SIZE 16 BY 1.08 NO-UNDO.

DEFINE VARIABLE tg-2 AS LOGICAL INITIAL no 
     LABEL "Por per°odo mensal" 
     VIEW-AS TOGGLE-BOX
     SIZE 16 BY 1.08 NO-UNDO.

DEFINE BUTTON btn-plano-acao 
     LABEL "Plano de Aá∆o" 
     SIZE 26 BY 1.13.

DEFINE VARIABLE fi-datetime AS CHARACTER FORMAT "X(256)":U 
     VIEW-AS FILL-IN 
     SIZE 77 BY .79
     BGCOLOR 8 FGCOLOR 2  NO-UNDO.

DEFINE VARIABLE fi_cod-ctrab_fim AS CHARACTER FORMAT "x(16)" INITIAL "ZZZZZZZZZZZZZZZZ" 
     VIEW-AS FILL-IN 
     SIZE 18.86 BY .88.

DEFINE VARIABLE fi_cod-ctrab_ini AS CHARACTER FORMAT "x(16)" 
     LABEL "Centro de Trabalho":R18 
     VIEW-AS FILL-IN 
     SIZE 24 BY .88.

DEFINE VARIABLE fi_dt_fim AS DATE FORMAT "99/99/9999":U 
     VIEW-AS FILL-IN 
     SIZE 18.29 BY .88 NO-UNDO.

DEFINE VARIABLE fi_dt_ini AS DATE FORMAT "99/99/9999":U 
     LABEL "Data Inicial" 
     VIEW-AS FILL-IN 
     SIZE 17 BY .88 NO-UNDO.

DEFINE VARIABLE fi_estab_fim AS CHARACTER FORMAT "X(05)":U INITIAL "101" 
     VIEW-AS FILL-IN 
     SIZE 8.43 BY .88 NO-UNDO.

DEFINE VARIABLE fi_estab_ini AS CHARACTER FORMAT "X(05)":U INITIAL "101" 
     LABEL "Estabelecimento" 
     VIEW-AS FILL-IN 
     SIZE 8.43 BY .88 NO-UNDO.

DEFINE VARIABLE fi_gm-codigo_fim AS CHARACTER FORMAT "x(9)" INITIAL "ZZZZZZZZZ" 
     VIEW-AS FILL-IN 
     SIZE 11.86 BY .88.

DEFINE VARIABLE fi_gm-codigo_ini AS CHARACTER FORMAT "x(9)" 
     LABEL "Grupo M†quina":R16 
     VIEW-AS FILL-IN 
     SIZE 11.86 BY .88.

DEFINE VARIABLE FI_ITE-fim AS CHARACTER FORMAT "X(256)":U INITIAL "ZZZZZZZZZZZZZZZZZ" 
     VIEW-AS FILL-IN 
     SIZE 17.86 BY .88 NO-UNDO.

DEFINE VARIABLE FI_ITE-INI AS CHARACTER FORMAT "X(256)":U 
     LABEL "Item" 
     VIEW-AS FILL-IN 
     SIZE 16 BY .88 NO-UNDO.

DEFINE IMAGE IMAGE-1
     FILENAME "image\im-fir":U
     SIZE 3 BY .88.

DEFINE IMAGE IMAGE-11
     FILENAME "image\im-fir":U
     SIZE 3 BY .88.

DEFINE IMAGE IMAGE-12
     FILENAME "image\im-las":U
     SIZE 3 BY .88.

DEFINE IMAGE IMAGE-2
     FILENAME "image\im-las":U
     SIZE 3 BY .88.

DEFINE IMAGE IMAGE-3
     FILENAME "image\im-fir":U
     SIZE 3 BY .88.

DEFINE IMAGE IMAGE-4
     FILENAME "image\im-las":U
     SIZE 3 BY .88.

DEFINE IMAGE IMAGE-5
     FILENAME "image\im-fir":U
     SIZE 3 BY .88.

DEFINE IMAGE IMAGE-6
     FILENAME "image\im-las":U
     SIZE 3 BY .88.

DEFINE IMAGE IMAGE-8
     FILENAME "image\im-las":U
     SIZE 3 BY .88.

DEFINE IMAGE IMAGE-9
     FILENAME "image\im-fir":U
     SIZE 3 BY .88.

DEFINE RECTANGLE RECT-14
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 77 BY 1.5.

DEFINE RECTANGLE RECT-15
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 77 BY 3.

DEFINE VARIABLE ch-1 AS LOGICAL INITIAL yes 
     LABEL "Por per°odo di†rio" 
     VIEW-AS TOGGLE-BOX
     SIZE 15 BY .83 NO-UNDO.

DEFINE VARIABLE ch-2 AS LOGICAL INITIAL no 
     LABEL "Por per°odo mensal" 
     VIEW-AS TOGGLE-BOX
     SIZE 17.57 BY .83 NO-UNDO.

DEFINE BUTTON bt-ajuda 
     LABEL "Ajuda" 
     SIZE 10 BY 1.

DEFINE BUTTON bt-cancelar AUTO-END-KEY 
     LABEL "Fechar" 
     SIZE 10 BY 1.

DEFINE BUTTON bt-executar 
     LABEL "Executar" 
     SIZE 10 BY 1.

DEFINE IMAGE im-pg-dig
     FILENAME "image\im-fldup":U
     SIZE 15.86 BY 1.21.

DEFINE IMAGE im-pg-imp
     FILENAME "image\im-fldup":U
     SIZE 15.86 BY 1.21.

DEFINE IMAGE im-pg-par
     FILENAME "image\im-fldup":U
     SIZE 15.86 BY 1.21.

DEFINE IMAGE im-pg-sel
     FILENAME "image\im-fldup":U
     SIZE 15.86 BY 1.21.

DEFINE RECTANGLE RECT-1
     EDGE-PIXELS 2 GRAPHIC-EDGE    
     SIZE 79 BY 1.42
     BGCOLOR 7 .

DEFINE RECTANGLE RECT-6
     EDGE-PIXELS 0    
     SIZE 78.86 BY .13
     BGCOLOR 7 .

DEFINE RECTANGLE rt-folder
     EDGE-PIXELS 1 GRAPHIC-EDGE  NO-FILL   
     SIZE 79 BY 11.38
     FGCOLOR 0 .

DEFINE RECTANGLE rt-folder-left
     EDGE-PIXELS 0    
     SIZE .43 BY 11.21
     BGCOLOR 15 .

DEFINE RECTANGLE rt-folder-right
     EDGE-PIXELS 0    
     SIZE .43 BY 11.21
     BGCOLOR 7 .

DEFINE RECTANGLE rt-folder-top
     EDGE-PIXELS 0    
     SIZE 78.86 BY .13
     BGCOLOR 15 .

/* Query definitions                                                    */
&ANALYZE-SUSPEND
DEFINE QUERY br-centros FOR 
      ctrab SCROLLING.

DEFINE QUERY br-digita FOR 
      tt-digita SCROLLING.
&ANALYZE-RESUME

/* Browse definitions                                                   */
DEFINE BROWSE br-centros
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS br-centros w-relat _STRUCTURED
  QUERY br-centros NO-LOCK DISPLAY
      ctrab.cod-ctrab FORMAT "x(16)":U WIDTH 15.72
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 19 BY 8.75
         FONT 4 FIT-LAST-COLUMN.

DEFINE BROWSE br-digita
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS br-digita w-relat _FREEFORM
  QUERY br-digita DISPLAY
      tt-digita.CENTRO       COLUMN-LABEL 'Ctro Trabalho'
tt-digita.DESC-CENTRO  COLUMN-LABEL 'Descriá∆o'
tt-digita.GRUPO        COLUMN-LABEL 'Grp. M†quina'
tt-digita.DESC-GRUPO   COLUMN-LABEL 'Descriá∆o'
ENABLE
tt-digita.CENTRO
tt-digita.GRUPO
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH SEPARATORS SIZE 73 BY 9
         BGCOLOR 15 FONT 1.


/* ************************  Frame Definitions  *********************** */

DEFINE FRAME f-pg-imp
     rs-destino AT ROW 1.63 COL 3.14 HELP
          "Destino de Impress∆o do Relat¢rio" NO-LABEL
     bt-arquivo AT ROW 2.71 COL 43.14 HELP
          "Escolha do nome do arquivo"
     bt-config-impr AT ROW 2.71 COL 43.14 HELP
          "Configuraá∆o da impressora"
     c-arquivo AT ROW 2.75 COL 3.14 HELP
          "Nome do arquivo de destino do relat¢rio" NO-LABEL
     l-habilitaRtf AT ROW 4.79 COL 3.14
     c-modelo-rtf AT ROW 6.63 COL 3 HELP
          "Nome do arquivo de modelo do relat¢rio" NO-LABEL
     bt-modelo-rtf AT ROW 6.63 COL 43 HELP
          "Escolha do nome do arquivo"
     rs-execucao AT ROW 8.88 COL 2.86 HELP
          "Modo de Execuá∆o" NO-LABEL
     text-destino AT ROW 1.04 COL 3.86 NO-LABEL
     text-rtf AT ROW 4.21 COL 1.14 COLON-ALIGNED NO-LABEL
     text-modelo-rtf AT ROW 5.96 COL 1.14 COLON-ALIGNED NO-LABEL
     text-modo AT ROW 8.13 COL 1.14 COLON-ALIGNED NO-LABEL
     rect-rtf AT ROW 4.5 COL 2
     RECT-7 AT ROW 1.33 COL 2.14
     RECT-9 AT ROW 8.33 COL 2
    WITH 1 DOWN NO-BOX 
         SIDE-LABELS THREE-D 
         AT COL 3 ROW 3
         SIZE 73.72 BY 10.5
         FONT 1 WIDGET-ID 100.

DEFINE FRAME f-relat
     bt-executar AT ROW 14.5 COL 3 HELP
          "Dispara a execuá∆o do relat¢rio"
     bt-cancelar AT ROW 14.5 COL 14 HELP
          "Fechar"
     bt-ajuda AT ROW 14.5 COL 70 HELP
          "Ajuda"
     RECT-1 AT ROW 14.29 COL 2
     RECT-6 AT ROW 13.75 COL 2.14
     rt-folder-top AT ROW 2.5 COL 2.14
     rt-folder-right AT ROW 2.67 COL 80.43
     rt-folder-left AT ROW 2.5 COL 2.14
     rt-folder AT ROW 2.5 COL 2
     im-pg-dig AT ROW 1.5 COL 33.86
     im-pg-par AT ROW 1.5 COL 18
     im-pg-sel AT ROW 1.5 COL 2.14
     im-pg-imp AT ROW 1.5 COL 49.43
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 1 ROW 1
         SIZE 81 BY 15
         DEFAULT-BUTTON bt-executar WIDGET-ID 100.

DEFINE FRAME FRAME-C
    WITH 1 DOWN KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 2 ROW 1.25
         SIZE 79 BY 12.75 WIDGET-ID 200.

DEFINE FRAME f-pg-dig
     br-digita AT ROW 1 COL 1
     bt-inserir AT ROW 10 COL 1
     bt-alterar AT ROW 10 COL 16
     bt-retirar AT ROW 10 COL 31
     bt-salvar AT ROW 10 COL 46
     bt-recuperar AT ROW 10 COL 61
    WITH 1 DOWN NO-BOX 
         THREE-D 
         AT COL 3 ROW 3.31
         SIZE 76.86 BY 10.15
         FONT 1 WIDGET-ID 100.

DEFINE FRAME f-pg-sel
     fi_dt_ini AT ROW 1.54 COL 33 COLON-ALIGNED WIDGET-ID 2
     fi_dt_fim AT ROW 1.54 COL 57.72 COLON-ALIGNED NO-LABEL WIDGET-ID 4
     fi_cod-ctrab_ini AT ROW 3.5 COL 15.86 COLON-ALIGNED HELP
          "Centro de Trabalho" WIDGET-ID 14
     fi_estab_ini AT ROW 1.54 COL 14 COLON-ALIGNED WIDGET-ID 6
     ch-1 AT ROW 4.75 COL 4 WIDGET-ID 64
     ch-2 AT ROW 4.75 COL 19.43 WIDGET-ID 66
     fi_gm-codigo_ini AT ROW 11.5 COL 19 COLON-ALIGNED HELP
          "Grupo de M†quinas onde Ç realizada a operaá∆o" WIDGET-ID 16
     fi_gm-codigo_fim AT ROW 11.5 COL 38.86 COLON-ALIGNED HELP
          "Grupo de M†quinas onde Ç realizada a operaá∆o" NO-LABEL WIDGET-ID 28
     btn-plano-acao AT ROW 11.5 COL 53 WIDGET-ID 60
     fi_cod-ctrab_fim AT ROW 11.63 COL 57.72 COLON-ALIGNED HELP
          "Centro de Trabalho" NO-LABEL WIDGET-ID 26
     FI_ITE-INI AT ROW 12.5 COL 20 COLON-ALIGNED WIDGET-ID 30
     FI_ITE-fim AT ROW 12.5 COL 43.14 COLON-ALIGNED NO-LABEL WIDGET-ID 40
     fi_estab_fim AT ROW 12.5 COL 68.43 COLON-ALIGNED NO-LABEL WIDGET-ID 8
     fi-datetime AT ROW 12.75 COL 2 NO-LABEL WIDGET-ID 62
     "<- Clique duas vezes para escolher." VIEW-AS TEXT
          SIZE 29 BY .79 AT ROW 3.54 COL 43.14 WIDGET-ID 58
     IMAGE-1 AT ROW 1.54 COL 52.72
     IMAGE-2 AT ROW 1.54 COL 56.14
     IMAGE-3 AT ROW 12.5 COL 64 WIDGET-ID 10
     IMAGE-4 AT ROW 12.5 COL 67.14 WIDGET-ID 12
     IMAGE-5 AT ROW 11.5 COL 54 WIDGET-ID 18
     IMAGE-6 AT ROW 11.5 COL 57.14 WIDGET-ID 20
     IMAGE-8 AT ROW 11.5 COL 37.43 WIDGET-ID 24
     IMAGE-9 AT ROW 11.5 COL 34.29 WIDGET-ID 32
     IMAGE-11 AT ROW 12.5 COL 38.43 WIDGET-ID 38
     IMAGE-12 AT ROW 12.54 COL 41.57 WIDGET-ID 36
     RECT-14 AT ROW 1.25 COL 2 WIDGET-ID 52
     RECT-15 AT ROW 3 COL 2 WIDGET-ID 54
    WITH 1 DOWN KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 2 ROW 1.25
         SIZE 79.14 BY 12.79
         FONT 1 WIDGET-ID 100.

DEFINE FRAME f-ctrab
     br-centros AT ROW 1 COL 1 WIDGET-ID 400
    WITH 1 DOWN KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 17.86 ROW 4.5
         SIZE 20 BY 9 WIDGET-ID 300.

DEFINE FRAME f-pg-par
     tg-1 AT ROW 1.29 COL 2
     tg-2 AT ROW 1.29 COL 18.57 WIDGET-ID 2
    WITH 1 DOWN NO-BOX 
         SIDE-LABELS THREE-D 
         AT COL 38.72 ROW 10.13
         SIZE 41.29 BY 1.54
         FONT 1 WIDGET-ID 100.


/* *********************** Procedure Settings ************************ */

&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: w-relat
   Allow: Basic,Browse,DB-Fields,Window,Query
   Add Fields to: Neither
   Other Settings: COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS

/* *************************  Create Window  ************************** */

&ANALYZE-SUSPEND _CREATE-WINDOW
IF SESSION:DISPLAY-TYPE = "GUI":U THEN
  CREATE WINDOW w-relat ASSIGN
         HIDDEN             = YES
         TITLE              = "Gerador OEE"
         HEIGHT             = 15
         WIDTH              = 81.14
         MAX-HEIGHT         = 22.33
         MAX-WIDTH          = 114.14
         VIRTUAL-HEIGHT     = 22.33
         VIRTUAL-WIDTH      = 114.14
         RESIZE             = yes
         SCROLL-BARS        = no
         STATUS-AREA        = yes
         BGCOLOR            = ?
         FGCOLOR            = ?
         KEEP-FRAME-Z-ORDER = yes
         THREE-D            = yes
         MESSAGE-AREA       = no
         SENSITIVE          = yes.
ELSE {&WINDOW-NAME} = CURRENT-WINDOW.
/* END WINDOW DEFINITION                                                */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _INCLUDED-LIB w-relat 
/* ************************* Included-Libraries *********************** */

{src/adm/method/containr.i}
{include/w-relat.i}
{utp/ut-glob.i}

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME




/* ***********  Runtime Attributes and AppBuilder Settings  *********** */

&ANALYZE-SUSPEND _RUN-TIME-ATTRIBUTES
/* SETTINGS FOR WINDOW w-relat
  VISIBLE,,RUN-PERSISTENT                                               */
/* REPARENT FRAME */
ASSIGN FRAME f-ctrab:FRAME = FRAME f-pg-sel:HANDLE
       FRAME FRAME-C:FRAME = FRAME f-relat:HANDLE.

/* SETTINGS FOR FRAME f-ctrab
                                                                        */
/* BROWSE-TAB br-centros 1 f-ctrab */
/* SETTINGS FOR FRAME f-pg-dig
   NOT-VISIBLE FRAME-NAME UNDERLINE                                     */
/* BROWSE-TAB br-digita 1 f-pg-dig */
ASSIGN 
       FRAME f-pg-dig:HIDDEN           = TRUE.

/* SETTINGS FOR BUTTON bt-alterar IN FRAME f-pg-dig
   NO-ENABLE                                                            */
/* SETTINGS FOR BUTTON bt-retirar IN FRAME f-pg-dig
   NO-ENABLE                                                            */
/* SETTINGS FOR BUTTON bt-salvar IN FRAME f-pg-dig
   NO-ENABLE                                                            */
/* SETTINGS FOR FRAME f-pg-imp
   NOT-VISIBLE UNDERLINE                                                */
ASSIGN 
       FRAME f-pg-imp:HIDDEN           = TRUE.

/* SETTINGS FOR EDITOR c-modelo-rtf IN FRAME f-pg-imp
   NO-ENABLE                                                            */
/* SETTINGS FOR FILL-IN text-destino IN FRAME f-pg-imp
   NO-DISPLAY NO-ENABLE ALIGN-L                                         */
ASSIGN 
       text-destino:PRIVATE-DATA IN FRAME f-pg-imp     = 
                "Destino".

ASSIGN 
       text-modelo-rtf:PRIVATE-DATA IN FRAME f-pg-imp     = 
                "Modelo:".

/* SETTINGS FOR FILL-IN text-modo IN FRAME f-pg-imp
   NO-DISPLAY NO-ENABLE                                                 */
ASSIGN 
       text-modo:PRIVATE-DATA IN FRAME f-pg-imp     = 
                "Execuá∆o".

ASSIGN 
       text-rtf:PRIVATE-DATA IN FRAME f-pg-imp     = 
                "Rich Text Format(RTF)".

/* SETTINGS FOR FRAME f-pg-par
   NOT-VISIBLE UNDERLINE                                                */
ASSIGN 
       FRAME f-pg-par:HIDDEN           = TRUE.

ASSIGN 
       tg-1:HIDDEN IN FRAME f-pg-par           = TRUE.

ASSIGN 
       tg-2:HIDDEN IN FRAME f-pg-par           = TRUE.

/* SETTINGS FOR FRAME f-pg-sel
   Custom                                                               */
/* SETTINGS FOR TOGGLE-BOX ch-2 IN FRAME f-pg-sel
   NO-ENABLE                                                            */
/* SETTINGS FOR FILL-IN fi-datetime IN FRAME f-pg-sel
   NO-ENABLE ALIGN-L                                                    */
ASSIGN 
       fi-datetime:READ-ONLY IN FRAME f-pg-sel        = TRUE.

/* SETTINGS FOR FILL-IN fi_estab_fim IN FRAME f-pg-sel
   NO-ENABLE                                                            */
/* SETTINGS FOR FILL-IN fi_estab_ini IN FRAME f-pg-sel
   NO-ENABLE                                                            */
/* SETTINGS FOR FILL-IN fi_gm-codigo_fim IN FRAME f-pg-sel
   NO-ENABLE                                                            */
/* SETTINGS FOR FILL-IN fi_gm-codigo_ini IN FRAME f-pg-sel
   NO-ENABLE                                                            */
/* SETTINGS FOR FILL-IN FI_ITE-fim IN FRAME f-pg-sel
   NO-ENABLE                                                            */
/* SETTINGS FOR FILL-IN FI_ITE-INI IN FRAME f-pg-sel
   NO-ENABLE                                                            */
ASSIGN 
       IMAGE-11:HIDDEN IN FRAME f-pg-sel           = TRUE.

ASSIGN 
       IMAGE-12:HIDDEN IN FRAME f-pg-sel           = TRUE.

ASSIGN 
       IMAGE-8:HIDDEN IN FRAME f-pg-sel           = TRUE.

/* SETTINGS FOR FRAME f-relat
                                                                        */
/* SETTINGS FOR RECTANGLE RECT-1 IN FRAME f-relat
   NO-ENABLE                                                            */
/* SETTINGS FOR RECTANGLE RECT-6 IN FRAME f-relat
   NO-ENABLE                                                            */
/* SETTINGS FOR RECTANGLE rt-folder IN FRAME f-relat
   NO-ENABLE                                                            */
/* SETTINGS FOR RECTANGLE rt-folder-left IN FRAME f-relat
   NO-ENABLE                                                            */
/* SETTINGS FOR RECTANGLE rt-folder-right IN FRAME f-relat
   NO-ENABLE                                                            */
/* SETTINGS FOR RECTANGLE rt-folder-top IN FRAME f-relat
   NO-ENABLE                                                            */
/* SETTINGS FOR FRAME FRAME-C
                                                                        */
IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(w-relat)
THEN w-relat:HIDDEN = no.

/* _RUN-TIME-ATTRIBUTES-END */
&ANALYZE-RESUME


/* Setting information for Queries and Browse Widgets fields            */

&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE br-centros
/* Query rebuild information for BROWSE br-centros
     _TblList          = "mgcad.ctrab"
     _Options          = "NO-LOCK INDEXED-REPOSITION"
     _FldNameList[1]   > mgcad.ctrab.cod-ctrab
"ctrab.cod-ctrab" ? ? "character" ? ? ? ? ? ? no ? no no "15.72" yes no no "U" "" "" "" "" "" "" 0 no 0 no no
     _Query            is OPENED
*/  /* BROWSE br-centros */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE br-digita
/* Query rebuild information for BROWSE br-digita
     _START_FREEFORM
OPEN QUERY br-digita FOR EACH tt-digita.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE br-digita */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _QUERY-BLOCK FRAME f-pg-dig
/* Query rebuild information for FRAME f-pg-dig
     _START_FREEFORM
FRAME f-pg-dig:VISIBLE = FALSE.
     _END_FREEFORM
     _Query            is NOT OPENED
*/  /* FRAME f-pg-dig */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _QUERY-BLOCK FRAME f-pg-imp
/* Query rebuild information for FRAME f-pg-imp
     _START_FREEFORM
FRAME f-pg-imp:VISIBLE = FALSE.
     _END_FREEFORM
     _Query            is NOT OPENED
*/  /* FRAME f-pg-imp */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _QUERY-BLOCK FRAME f-pg-par
/* Query rebuild information for FRAME f-pg-par
     _START_FREEFORM
FRAME f-pg-par:VISIBLE = FALSE.
     _END_FREEFORM
     _Query            is NOT OPENED
*/  /* FRAME f-pg-par */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _QUERY-BLOCK FRAME f-pg-sel
/* Query rebuild information for FRAME f-pg-sel
     _Query            is NOT OPENED
*/  /* FRAME f-pg-sel */
&ANALYZE-RESUME

 


/* **********************  Create OCX Containers  ********************** */

&ANALYZE-SUSPEND _CREATE-DYNAMIC

&IF "{&OPSYS}" = "WIN32":U AND "{&WINDOW-SYSTEM}" NE "TTY":U &THEN

CREATE CONTROL-FRAME CtrlFrame ASSIGN
       FRAME           = FRAME f-pg-par:HANDLE
       ROW             = 1.25
       COLUMN          = 35
       HEIGHT          = 1.25
       WIDTH           = 5
       WIDGET-ID       = 160
       HIDDEN          = yes
       SENSITIVE       = yes.
      CtrlFrame:NAME = "CtrlFrame":U .
/* CtrlFrame OCXINFO:CREATE-CONTROL from: {F0B88A90-F5DA-11CF-B545-0020AF6ED35A} type: PSTimer */
      CtrlFrame:MOVE-BEFORE(tg-1:HANDLE IN FRAME f-pg-par).

&ENDIF

&ANALYZE-RESUME /* End of _CREATE-DYNAMIC */


/* ************************  Control Triggers  ************************ */

&Scoped-define SELF-NAME w-relat
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL w-relat w-relat
ON END-ERROR OF w-relat /* Gerador OEE */
OR ENDKEY OF {&WINDOW-NAME} ANYWHERE DO:
  /* This case occurs when the user presses the "Esc" key.
     In a persistently run window, just ignore this.  If we did not, the
     application would exit. */
   RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL w-relat w-relat
ON WINDOW-CLOSE OF w-relat /* Gerador OEE */
DO:
  /* This event will close the window and terminate the procedure.  */
  APPLY "CLOSE":U TO THIS-PROCEDURE.
  RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME f-pg-dig
&Scoped-define SELF-NAME f-pg-imp
&Scoped-define SELF-NAME f-pg-par
&Scoped-define BROWSE-NAME br-centros
&Scoped-define FRAME-NAME f-ctrab
&Scoped-define SELF-NAME br-centros
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL br-centros w-relat
ON MOUSE-SELECT-DBLCLICK OF br-centros IN FRAME f-ctrab
DO:
  fi_cod-ctrab_ini:SCREEN-VALUE IN FRAME f-pg-sel = string(br-centros:GET-BROWSE-COLUMN(1):SCREEN-VALUE IN FRAME f-ctrab).
  FRAME f-ctrab:VISIBLE = FALSE.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define BROWSE-NAME br-digita
&Scoped-define FRAME-NAME f-pg-dig
&Scoped-define SELF-NAME br-digita
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL br-digita w-relat
ON DEL OF br-digita IN FRAME f-pg-dig
DO:
   apply 'choose':U to bt-retirar in frame f-pg-dig.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL br-digita w-relat
ON END-ERROR OF br-digita IN FRAME f-pg-dig
ANYWHERE 
DO:
    if  br-digita:new-row in frame f-pg-dig then do:
        if  avail tt-digita then
            delete tt-digita.
        if  br-digita:delete-current-row() in frame f-pg-dig then. 
    end.                                                               
    else do:
        get current br-digita.
        display tt-digita.CENTRO     
                tt-digita.DESC-CENTRO 
                tt-digita.GRUPO      
                tt-digita.DESC-GRUPO 
            
            
            
            with browse br-digita. 
    end.
    return no-apply.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL br-digita w-relat
ON ENTER OF br-digita IN FRAME f-pg-dig
ANYWHERE
DO:
  apply 'tab':U to self.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL br-digita w-relat
ON INS OF br-digita IN FRAME f-pg-dig
DO:
   apply 'choose':U to bt-inserir in frame f-pg-dig.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL br-digita w-relat
ON OFF-END OF br-digita IN FRAME f-pg-dig
DO:
   apply 'entry':U to bt-inserir in frame f-pg-dig.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL br-digita w-relat
ON OFF-HOME OF br-digita IN FRAME f-pg-dig
DO:
  apply 'entry':U to bt-recuperar in frame f-pg-dig.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL br-digita w-relat
ON ROW-ENTRY OF br-digita IN FRAME f-pg-dig
DO:
   /*:T trigger para inicializar campos da temp table de digitaá∆o */
   if  br-digita:new-row in frame f-pg-dig then do:
       assign tt-digita.CENTRO:screen-value in browse br-digita = ''
              tt-digita.GRUPO :screen-value in browse br-digita = ''.



   end.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL br-digita w-relat
ON ROW-LEAVE OF br-digita IN FRAME f-pg-dig
DO:
    /*:T ê aqui que a gravaá∆o da linha da temp-table Ç efetivada.
       PorÇm as validaá‰es dos registros devem ser feitas na procedure pi-executar,
       no local indicado pelo coment†rio */
    
    if br-digita:NEW-ROW in frame f-pg-dig then 
    do transaction on error undo, return no-apply:
        create tt-digita.
        assign input browse br-digita tt-digita.CENTRO     
               input browse br-digita tt-digita.DESC-CENTRO
               input browse br-digita tt-digita.GRUPO      
               input browse br-digita tt-digita.DESC-GRUPO. 
    
        br-digita:CREATE-RESULT-LIST-ENTRY() in frame f-pg-dig.
    end.
    else do transaction on error undo, return no-apply:
        assign input browse br-digita tt-digita.CENTRO      
               input browse br-digita tt-digita.DESC-CENTRO 
               input browse br-digita tt-digita.GRUPO       
               input browse br-digita tt-digita.DESC-GRUPO. 
            
            
            
            .
    end.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-relat
&Scoped-define SELF-NAME bt-ajuda
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-ajuda w-relat
ON CHOOSE OF bt-ajuda IN FRAME f-relat /* Ajuda */
DO:
   {include/ajuda.i}
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-dig
&Scoped-define SELF-NAME bt-alterar
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-alterar w-relat
ON CHOOSE OF bt-alterar IN FRAME f-pg-dig /* Alterar */
DO:
   apply 'entry':U to tt-digita.centro in browse br-digita. 
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-imp
&Scoped-define SELF-NAME bt-arquivo
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-arquivo w-relat
ON CHOOSE OF bt-arquivo IN FRAME f-pg-imp
DO:
    {include/i-rparq.i}
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-relat
&Scoped-define SELF-NAME bt-cancelar
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-cancelar w-relat
ON CHOOSE OF bt-cancelar IN FRAME f-relat /* Fechar */
DO:
   apply "close":U to this-procedure.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-imp
&Scoped-define SELF-NAME bt-config-impr
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-config-impr w-relat
ON CHOOSE OF bt-config-impr IN FRAME f-pg-imp
DO:
   {include/i-rpimp.i}
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-relat
&Scoped-define SELF-NAME bt-executar
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-executar w-relat
ON CHOOSE OF bt-executar IN FRAME f-relat /* Executar */
DO:
   do  on error undo, return no-apply:
       run pi-executar.
   end.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-dig
&Scoped-define SELF-NAME bt-inserir
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-inserir w-relat
ON CHOOSE OF bt-inserir IN FRAME f-pg-dig /* Inserir */
DO:
    assign bt-alterar:SENSITIVE in frame f-pg-dig = yes
           bt-retirar:SENSITIVE in frame f-pg-dig = yes
           bt-salvar:SENSITIVE in frame f-pg-dig  = yes.
    
    if num-results("br-digita":U) > 0 then
        br-digita:INSERT-ROW("after":U) in frame f-pg-dig.
    else do transaction:
        create tt-digita.
        
        open query br-digita for each tt-digita.
        
        apply "entry":U to tt-digita.centro in browse br-digita. 
    end.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-imp
&Scoped-define SELF-NAME bt-modelo-rtf
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-modelo-rtf w-relat
ON CHOOSE OF bt-modelo-rtf IN FRAME f-pg-imp
DO:
    def var c-arq-conv  as char no-undo.
    def var l-ok as logical no-undo.

    assign c-modelo-rtf = replace(input frame {&frame-name} c-modelo-rtf, "/", "~\").
    SYSTEM-DIALOG GET-FILE c-arq-conv
       FILTERS "*.rtf" "*.rtf",
               "*.*" "*.*"
       DEFAULT-EXTENSION "rtf"
       INITIAL-DIR "modelos" 
       MUST-EXIST
       USE-FILENAME
       UPDATE l-ok.
    if  l-ok = yes then
        assign c-modelo-rtf:screen-value in frame {&frame-name}  = replace(c-arq-conv, "~\", "/"). 

END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-dig
&Scoped-define SELF-NAME bt-recuperar
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-recuperar w-relat
ON CHOOSE OF bt-recuperar IN FRAME f-pg-dig /* Recuperar */
DO:
    {include/i-rprcd.i}
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME bt-retirar
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-retirar w-relat
ON CHOOSE OF bt-retirar IN FRAME f-pg-dig /* Retirar */
DO:
    if  br-digita:num-selected-rows > 0 then do on error undo, return no-apply:
        get current br-digita.
        delete tt-digita.
        if  br-digita:delete-current-row() in frame f-pg-dig then.
    end.
    
    if num-results("br-digita":U) = 0 then
        assign bt-alterar:SENSITIVE in frame f-pg-dig = no
               bt-retirar:SENSITIVE in frame f-pg-dig = no
               bt-salvar:SENSITIVE in frame f-pg-dig  = no.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME bt-salvar
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL bt-salvar w-relat
ON CHOOSE OF bt-salvar IN FRAME f-pg-dig /* Salvar */
DO:
   {include/i-rpsvd.i}
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-sel
&Scoped-define SELF-NAME btn-plano-acao
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL btn-plano-acao w-relat
ON CHOOSE OF btn-plano-acao IN FRAME f-pg-sel /* Plano de Aá∆o */
DO:
  RUN "_brandl/plano-acao.r".
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME ch-1
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL ch-1 w-relat
ON VALUE-CHANGED OF ch-1 IN FRAME f-pg-sel /* Por per°odo di†rio */
DO:
  IF SELF:CHECKED THEN
      ASSIGN ch-2:CHECKED IN FRAME f-pg-sel = NO.
      ASSIGN tg-2:CHECKED IN FRAME f-pg-par = NO.
      ASSIGN tg-1:CHECKED IN FRAME f-pg-par = YES.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME ch-2
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL ch-2 w-relat
ON VALUE-CHANGED OF ch-2 IN FRAME f-pg-sel /* Por per°odo mensal */
DO:
    IF SELF:CHECKED THEN
      ASSIGN ch-1:CHECKED IN FRAME f-pg-sel = NO.
      ASSIGN tg-1:CHECKED IN FRAME f-pg-par = NO.
      ASSIGN tg-2:CHECKED IN FRAME f-pg-par = YES.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-par
&Scoped-define SELF-NAME CtrlFrame
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL CtrlFrame w-relat OCX.Tick
PROCEDURE CtrlFrame.PSTimer.Tick .
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  None required for OCX.
  Notes:       
------------------------------------------------------------------------------*/
IF organiza = 1 THEN DO:
    RUN organiza-layout.
    ASSIGN organiza = 2.
END.

ELSE
    chCtrlFrame:pSTimer:ENABLED = NO.


END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-sel
&Scoped-define SELF-NAME fi_cod-ctrab_ini
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL fi_cod-ctrab_ini w-relat
ON MOUSE-SELECT-DBLCLICK OF fi_cod-ctrab_ini IN FRAME f-pg-sel /* Centro de Trabalho */
DO:
  IF FRAME f-ctrab:VISIBLE = TRUE THEN
      FRAME f-ctrab:VISIBLE = FALSE.
  ELSE
      FRAME f-ctrab:VISIBLE = TRUE.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-relat
&Scoped-define SELF-NAME im-pg-dig
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL im-pg-dig w-relat
ON MOUSE-SELECT-CLICK OF im-pg-dig IN FRAME f-relat
DO:
    run pi-troca-pagina.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME im-pg-imp
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL im-pg-imp w-relat
ON MOUSE-SELECT-CLICK OF im-pg-imp IN FRAME f-relat
DO:
    run pi-troca-pagina.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME im-pg-par
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL im-pg-par w-relat
ON MOUSE-SELECT-CLICK OF im-pg-par IN FRAME f-relat
DO:
    run pi-troca-pagina.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME im-pg-sel
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL im-pg-sel w-relat
ON MOUSE-SELECT-CLICK OF im-pg-sel IN FRAME f-relat
DO:
    run pi-troca-pagina.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-imp
&Scoped-define SELF-NAME l-habilitaRtf
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL l-habilitaRtf w-relat
ON VALUE-CHANGED OF l-habilitaRtf IN FRAME f-pg-imp /* RTF */
DO:
    &IF "{&RTF}":U = "YES":U &THEN
    RUN pi-habilitaRtf.
    &endif
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME rs-destino
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL rs-destino w-relat
ON VALUE-CHANGED OF rs-destino IN FRAME f-pg-imp
DO:
/*Alterado 15/02/2005 - tech1007 - Evento alterado para correto funcionamento dos novos widgets
  utilizados para a funcionalidade de RTF*/
do  with frame f-pg-imp:
    case self:screen-value:
        when "1" then do:
            assign c-arquivo:sensitive    = no
                   bt-arquivo:visible     = no
                   bt-config-impr:visible = YES
                   /*Alterado 17/02/2005 - tech1007 - Realizado teste de preprocessador para
                     verificar se o RTF est† ativo*/
                   &IF "{&RTF}":U = "YES":U &THEN
                   l-habilitaRtf:sensitive  = NO
                   l-habilitaRtf:SCREEN-VALUE IN FRAME f-pg-imp = "No"
                   l-habilitaRtf = NO
                   &endif
                   /*Fim alteracao 17/02/2005*/
                   .
        end.
        when "2" then do:
            assign c-arquivo:sensitive     = yes
                   bt-arquivo:visible      = yes
                   bt-config-impr:visible  = NO
                   /*Alterado 17/02/2005 - tech1007 - Realizado teste de preprocessador para
                     verificar se o RTF est† ativo*/
                   &IF "{&RTF}":U = "YES":U &THEN
                   l-habilitaRtf:sensitive  = YES
                   &endif
                   /*Fim alteracao 17/02/2005*/
                   .
        end.
        when "3" then do:
            assign c-arquivo:sensitive     = no
                   bt-arquivo:visible      = no
                   bt-config-impr:visible  = no
                   /*Alterado 17/02/2005 - tech1007 - Realizado teste de preprocessador para
                     verificar se o RTF est† ativo*/
                   &IF "{&RTF}":U = "YES":U &THEN
                   l-habilitaRtf:sensitive  = YES
                   &endif
                   /*Fim alteracao 17/02/2005*/
                   .
            /*Alterado 15/02/2005 - tech1007 - Teste para funcionar corretamente no WebEnabler*/
            &IF "{&RTF}":U = "YES":U &THEN
            IF VALID-HANDLE(hWenController) THEN DO:
                ASSIGN l-habilitaRtf:sensitive  = NO
                       l-habilitaRtf:SCREEN-VALUE IN FRAME f-pg-imp = "No"
                       l-habilitaRtf = NO.
            END.
            &endif
            /*Fim alteracao 15/02/2005*/
        end.
    end case.
end.
&IF "{&RTF}":U = "YES":U &THEN
RUN pi-habilitaRtf.
&endif
/*Fim alteracao 15/02/2005*/

RUN organiza-layout.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME rs-execucao
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL rs-execucao w-relat
ON VALUE-CHANGED OF rs-execucao IN FRAME f-pg-imp
DO:
   {include/i-rprse.i}
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-par
&Scoped-define SELF-NAME tg-1
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL tg-1 w-relat
ON VALUE-CHANGED OF tg-1 IN FRAME f-pg-par /* Por per°odo di†rio */
DO:
  IF SELF:CHECKED THEN
      ASSIGN tg-2:CHECKED IN FRAME f-pg-par      = NO.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME tg-2
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL tg-2 w-relat
ON VALUE-CHANGED OF tg-2 IN FRAME f-pg-par /* Por per°odo mensal */
DO:
   IF SELF:CHECKED THEN
      ASSIGN tg-1:CHECKED IN FRAME f-pg-par      = NO.
      MESSAGE "BOX" VIEW-AS ALERT-BOX.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define FRAME-NAME f-pg-dig
&Scoped-define BROWSE-NAME br-centros
&UNDEFINE SELF-NAME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK w-relat 


/* ***************************  Main Block  *************************** */
fi_cod-ctrab_ini:LOAD-MOUSE-POINTER("image/lupa.cur":U) IN FRAME f-pg-sel.


/* Set CURRENT-WINDOW: this will parent dialog-boxes and frames.        */
ASSIGN CURRENT-WINDOW                = {&WINDOW-NAME} 
       THIS-PROCEDURE:CURRENT-WINDOW = {&WINDOW-NAME}.


//{utp/ut9000.i "XX9999" "9.99.99.999"}

/*:T inicializaá‰es do template de relat¢rio */
{include/i-rpini.i}

/* The CLOSE event can be used from inside or outside the procedure to  */
/* terminate it.                                                        */
ON CLOSE OF THIS-PROCEDURE 
   RUN disable_UI.

{include/i-rplbl.i}

/* Best default for GUI applications is...                              */
//PAUSE 0 BEFORE-HIDE.

/* on f5 of tt-digita.CENTRO  in browse br-digita or mouse-select-dblclick of tt-digita.CENTRO in browse br-digita do:                         */
/*     if num-results("br-digita") <> 0 THEN DO:                                                                                               */
/*        {include/zoomvar.i &prog-zoom=inzoom/z01in513.w                                                                                      */
/*                           &campo=tt-digita.CENTRO                                                                                           */
/*                           &campozoom=cod-ctrab                                                                                              */
/*                           &browse=br-digita}.                                                                                               */
/*                                                                                                                                             */
/*     END.                                                                                                                                    */
/* end.                                                                                                                                        */
/*                                                                                                                                             */
/* ON 'LEAVE':U OF tt-digita.CENTRO DO:                                                                                                        */
/*      IF AVAIL tt-digita AND tt-digita.CENTRO:SCREEN-VALUE IN BROWSE br-digita <> ''                                                         */
/*             AND tt-digita.CENTRO:SCREEN-VALUE IN BROWSE br-digita <> tt-digita.CENTRO  THEN DO:                                             */
/*             FOR FIRST ctrab FIELDS (cod-ctrab des-ctrab gm-codigo )   USE-INDEX id                                                          */
/*                  WHERE ctrab.cod-ctrab = tt-digita.CENTRO:SCREEN-VALUE IN BROWSE br-digita NO-LOCK:                                         */
/*                                                                                                                                             */
/*             END.                                                                                                                            */
/*             IF NOT AVAIL ctrab  THEN DO:                                                                                                    */
/*                                                                                                                                             */
/*                                                                                                                                             */
/*                     RUN utp/ut-msgs.p(INPUT "SHOW",                                                                                         */
/*                                       INPUT 17006,                                                                                          */
/*                                       INPUT "Centro de trabalho n∆o cadastrado! ~~ Centro de trabalho Informado , n∆o esta cadastrado!":U). */
/*                                                                                                                                             */
/*                     ASSIGN tt-digita.CENTRO:SCREEN-VALUE IN BROWSE br-digita = tt-digita.CENTRO.                                            */
/*                                                                                                                                             */
/*                     RETURN NO-APPLY.                                                                                                        */
/*                                                                                                                                             */
/*                                                                                                                                             */
/*             END.                                                                                                                            */
/*             ELSE DO:                                                                                                                        */
/*                   ASSIGN tt-digita.CENTRO:SCREEN-VALUE IN BROWSE br-digita      = ctrab.cod-ctrab                                           */
/*                          tt-digita.CENTRO                                       = ctrab.cod-ctrab                                           */
/*                          tt-digita.DESC-centro:SCREEN-VALUE IN BROWSE br-digita = ctrab.des-ctrab                                           */
/*                          tt-digita.DESC-centro                                  = ctrab.des-ctrab.                                          */
/*             END.                                                                                                                            */
/*      END.                                                                                                                                   */
/*                                                                                                                                             */
/* END.                                                                                                                                        */
/* on f5 of tt-digita.GRUPO  in browse br-digita or mouse-select-dblclick of tt-digita.GRUPO in browse br-digita do:                           */
/*     if num-results("br-digita") <> 0 THEN DO:                                                                                               */
/*        {include/zoomvar.i &prog-zoom=inzoom/z01in144.w                                                                                      */
/*                           &campo=tt-digita.GRUPO                                                                                            */
/*                           &campozoom=gm-codigo                                                                                              */
/*                           &browse=br-digita}.                                                                                               */
/*                                                                                                                                             */
/*     END.                                                                                                                                    */
/* end.                                                                                                                                        */
/*                                                                                                                                             */
/* ON 'LEAVE':U OF tt-digita.GRUPO DO:                                                                                                         */
/*      IF AVAIL tt-digita AND tt-digita.GRUPO:SCREEN-VALUE IN BROWSE br-digita <> ''                                                          */
/*             AND tt-digita.GRUPO:SCREEN-VALUE IN BROWSE br-digita <> tt-digita.GRUPO  THEN DO:                                               */
/*             FOR FIRST grup-maquina FIELDS (descricao gm-codigo )                                                                            */
/*                  WHERE  grup-maquina.gm-codigo = tt-digita.GRUPO:SCREEN-VALUE IN BROWSE br-digita NO-LOCK:                                  */
/*                                                                                                                                             */
/*             END.                                                                                                                            */
/*             IF NOT AVAIL grup-maquina  THEN DO:                                                                                             */
/*                                                                                                                                             */
/*                                                                                                                                             */
/*                     RUN utp/ut-msgs.p(INPUT "SHOW",                                                                                         */
/*                                       INPUT 17006,                                                                                          */
/*                                       INPUT "Grupo de m†quinas n∆o cadastrado! ~~ Grupo de m†quinas Informado , n∆o esta cadastrado!":U).   */
/*                                                                                                                                             */
/*                     ASSIGN tt-digita.GRUPO:SCREEN-VALUE IN BROWSE br-digita = tt-digita.GRUPO.                                              */
/*                                                                                                                                             */
/*                     RETURN NO-APPLY.                                                                                                        */
/*                                                                                                                                             */
/*                                                                                                                                             */
/*             END.                                                                                                                            */
/*             ELSE DO:                                                                                                                        */
/*                   ASSIGN tt-digita.GRUPO:SCREEN-VALUE IN BROWSE br-digita      = grup-maquina.gm-codigo                                     */
/*                          tt-digita.GRUPO                                       = grup-maquina.gm-codigo                                     */
/*                          tt-digita.DESC-GRUPO:SCREEN-VALUE IN BROWSE br-digita = grup-maquina.descricao                                     */
/*                          tt-digita.DESC-GRUPO                                  = grup-maquina.descricao.                                    */
/*             END.                                                                                                                            */
/*      END.                                                                                                                                   */
/*                                                                                                                                             */
/* END.                                                                                                                                        */

if tt-digita.GRUPO:load-mouse-pointer("image/lupa.cur")   then .
if tt-digita.CENTRO:load-mouse-pointer("image/lupa.cur")   then .

/* Now enable the interface and wait for the exit condition.            */
/* (NOTE: handle ERROR and END-KEY so cleanup code will always fire.    */
MAIN-BLOCK:
DO  ON ERROR   UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK
    ON END-KEY UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK:

    RUN enable_UI.
    RUN organiza-layout.
      RUN displayPerfil.
      ASSIGN
             tg-1   :CHECKED IN FRAME f-pg-par = TRUE.


    {include/i-rpmbl.i}.
  
    IF  NOT THIS-PROCEDURE:PERSISTENT THEN
        WAIT-FOR CLOSE OF THIS-PROCEDURE.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


/* **********************  Internal Procedures  *********************** */

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE adm-create-objects w-relat  _ADM-CREATE-OBJECTS
PROCEDURE adm-create-objects :
/*------------------------------------------------------------------------------
  Purpose:     Create handles for all SmartObjects used in this procedure.
               After SmartObjects are initialized, then SmartLinks are added.
  Parameters:  <none>
------------------------------------------------------------------------------*/

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE adm-row-available w-relat  _ADM-ROW-AVAILABLE
PROCEDURE adm-row-available :
/*------------------------------------------------------------------------------
  Purpose:     Dispatched to this procedure when the Record-
               Source has a new row available.  This procedure
               tries to get the new row (or foriegn keys) from
               the Record-Source and process it.
  Parameters:  <none>
------------------------------------------------------------------------------*/

  /* Define variables needed by this internal procedure.             */
  {src/adm/template/row-head.i}

  /* Process the newly available records (i.e. display fields,
     open queries, and/or pass records on to any RECORD-TARGETS).    */
  {src/adm/template/row-end.i}

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE control_load w-relat  _CONTROL-LOAD
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

OCXFile = SEARCH( "sfcp0010.wrx":U ).
IF OCXFile = ? THEN
  OCXFile = SEARCH(SUBSTRING(THIS-PROCEDURE:FILE-NAME, 1,
                     R-INDEX(THIS-PROCEDURE:FILE-NAME, ".":U), "CHARACTER":U) + "wrx":U).

IF OCXFile <> ? THEN
DO:
  ASSIGN
    chCtrlFrame = CtrlFrame:COM-HANDLE
    UIB_S = chCtrlFrame:LoadControls( OCXFile, "CtrlFrame":U)
  .
  RUN initialize-controls IN THIS-PROCEDURE NO-ERROR.
END.
ELSE MESSAGE "sfcp0010.wrx":U SKIP(1)
             "The binary control file could not be found. The controls cannot be loaded."
             VIEW-AS ALERT-BOX TITLE "Controls Not Loaded".

&ENDIF

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE disable_UI w-relat  _DEFAULT-DISABLE
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
  IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(w-relat)
  THEN DELETE WIDGET w-relat.
  IF THIS-PROCEDURE:PERSISTENT THEN DELETE PROCEDURE THIS-PROCEDURE.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE displayPerfil w-relat 
PROCEDURE displayPerfil :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
/* DEFINE VARIABLE c-info AS CHARACTER   NO-UNDO.                  */
/* DEFINE VARIABLE i-info AS INTEGER  INITIAL 0   NO-UNDO.         */
/* DEFINE VARIABLE i-cont  AS INTEGER     NO-UNDO.                 */
/* DEFINE VARIABLE i-num AS INTEGER     NO-UNDO.                   */
/*                                                                 */
/* ASSIGN i-cont = 1                                               */
/*        i-num  = c-perfil:NUM-ITEMS IN FRAME f-pg-par.           */
/*     REPEAT WHILE i-cont <= i-num:                               */
/*         IF c-perfil:DELETE(1) IN FRAME f-pg-par THEN.           */
/*         i-cont = i-cont + 1.                                    */
/*     END.                                                        */
/*                                                                 */
/*                                                                 */
/*                                                                 */
/*                                                                 */
/*                                                                 */
/* ASSIGN c-perfil = ''.                                           */
/*        c-perfil:SCREEN-VALUE IN FRAME f-pg-par = ''.            */
/* FOR EACH CST_PERFIL                                             */
/*       WHERE CST_PERFIL.usuario = c-seg-usuario NO-LOCK          */
/*       BREAK BY CST_PERFIL.perfil:                               */
/*                                                                 */
/*     ASSIGN i-info = i-info + 1                                  */
/*            c-info = string(i-info) + " : " + CST_PERFIL.perfil. */
/*                                                                 */
/*     c-perfil:ADD-LAST(c-info)  in frame f-pg-par NO-ERROR.      */
/* END.                                                            */
/*                                                                 */
/* RUN organiza-layout.                                            */
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE enable_UI w-relat  _DEFAULT-ENABLE
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
  ENABLE im-pg-dig im-pg-par im-pg-sel im-pg-imp bt-executar bt-cancelar 
         bt-ajuda 
      WITH FRAME f-relat IN WINDOW w-relat.
  {&OPEN-BROWSERS-IN-QUERY-f-relat}
  DISPLAY fi_dt_ini fi_dt_fim fi_cod-ctrab_ini fi_estab_ini ch-1 ch-2 
          fi_gm-codigo_ini fi_gm-codigo_fim fi_cod-ctrab_fim FI_ITE-INI 
          FI_ITE-fim fi_estab_fim fi-datetime 
      WITH FRAME f-pg-sel IN WINDOW w-relat.
  ENABLE fi_dt_ini fi_dt_fim fi_cod-ctrab_ini ch-1 btn-plano-acao 
         fi_cod-ctrab_fim IMAGE-1 IMAGE-2 IMAGE-3 IMAGE-4 IMAGE-5 IMAGE-6 
         IMAGE-8 IMAGE-9 IMAGE-11 IMAGE-12 RECT-14 RECT-15 
      WITH FRAME f-pg-sel IN WINDOW w-relat.
  {&OPEN-BROWSERS-IN-QUERY-f-pg-sel}
  VIEW FRAME FRAME-C IN WINDOW w-relat.
  {&OPEN-BROWSERS-IN-QUERY-FRAME-C}
  DISPLAY rs-destino c-arquivo l-habilitaRtf c-modelo-rtf rs-execucao text-rtf 
          text-modelo-rtf 
      WITH FRAME f-pg-imp IN WINDOW w-relat.
  ENABLE rect-rtf RECT-7 RECT-9 rs-destino bt-arquivo bt-config-impr c-arquivo 
         l-habilitaRtf bt-modelo-rtf rs-execucao text-rtf text-modelo-rtf 
      WITH FRAME f-pg-imp IN WINDOW w-relat.
  {&OPEN-BROWSERS-IN-QUERY-f-pg-imp}
  ENABLE br-digita bt-inserir bt-recuperar 
      WITH FRAME f-pg-dig IN WINDOW w-relat.
  {&OPEN-BROWSERS-IN-QUERY-f-pg-dig}
  ENABLE br-centros 
      WITH FRAME f-ctrab IN WINDOW w-relat.
  {&OPEN-BROWSERS-IN-QUERY-f-ctrab}
  DISPLAY tg-1 tg-2 
      WITH FRAME f-pg-par IN WINDOW w-relat.
  ENABLE tg-1 tg-2 
      WITH FRAME f-pg-par IN WINDOW w-relat.
  {&OPEN-BROWSERS-IN-QUERY-f-pg-par}
  VIEW w-relat.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE local-exit w-relat 
PROCEDURE local-exit :
/* -----------------------------------------------------------
  Purpose:  Starts an "exit" by APPLYing CLOSE event, which starts "destroy".
  Parameters:  <none>
  Notes:    If activated, should APPLY CLOSE, *not* dispatch adm-exit.   
-------------------------------------------------------------*/
   
   APPLY "CLOSE":U TO THIS-PROCEDURE.

   RETURN.
       
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE organiza-layout w-relat 
PROCEDURE organiza-layout :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
fi_gm-codigo_ini:VISIBLE IN FRAME f-pg-sel = FALSE.
FI_ITE-INI:VISIBLE IN FRAME f-pg-sel = FALSE.
fi_gm-codigo_fim:VISIBLE IN FRAME f-pg-sel = FALSE.
FI_ITE-fim:VISIBLE IN FRAME f-pg-sel = FALSE.
fi_estab_fim:VISIBLE IN FRAME f-pg-sel = FALSE.
fi_cod-ctrab_fim:VISIBLE IN FRAME f-pg-sel = FALSE.
IMAGE-8:VISIBLE IN FRAME f-pg-sel = FALSE.
IMAGE-9:VISIBLE IN FRAME f-pg-sel = FALSE.
IMAGE-11:VISIBLE IN FRAME f-pg-sel = FALSE.
IMAGE-12:VISIBLE IN FRAME f-pg-sel = FALSE.
IMAGE-4:VISIBLE IN FRAME f-pg-sel = FALSE.
IMAGE-3:VISIBLE IN FRAME f-pg-sel = FALSE.
IMAGE-5:VISIBLE IN FRAME f-pg-sel = FALSE.
IMAGE-6:VISIBLE IN FRAME f-pg-sel = FALSE.
FRAME f-ctrab:VISIBLE = FALSE.
fi-datetime:SCREEN-VALUE IN FRAME f-pg-sel = "USUµRIO: " + v_cod_usuar_corren + "    -    DIA: " + string(DATE(TODAY)).

rt-folder:VISIBLE IN FRAME f-relat = FALSE.
im-pg-dig:VISIBLE IN FRAME f-relat = FALSE.
im-pg-par:VISIBLE IN FRAME f-relat = FALSE.
im-pg-imp:VISIBLE IN FRAME f-relat = FALSE.
im-pg-sel:VISIBLE IN FRAME f-relat = FALSE.
FRAME f-pg-dig:VISIBLE = FALSE.
FRAME f-pg-par:VISIBLE = FALSE.
FRAME f-pg-imp:VISIBLE = FALSE.
FRAME f-pg-sel:VISIBLE = FALSE.
FRAME FRAME-C:VISIBLE = FALSE.
FRAME f-relat:VISIBLE = FALSE.
FRAME f-relat:VISIBLE = TRUE.
FRAME FRAME-C:VISIBLE = TRUE.
FRAME f-pg-sel:VISIBLE = TRUE.

/* ASSIGN dt-ini = DATE(MONTH(TODAY),01,YEAR(TODAY)).                      */
/*   IF MONTH(TODAY) = 12 THEN                                             */
/*     ASSIGN dt-fim =  DATE(MONTH(TODAY) + 1 ,01,YEAR(TODAY) + 1).        */
/*   ELSE                                                                  */
/*       ASSIGN dt-fim =  DATE(MONTH(TODAY) + 1 ,01,YEAR(TODAY)).          */
/*                                                                         */
/*   ASSIGN dt-fim = dt-fim - 1.                                           */
/*       ASSIGN fi_dt_ini:SCREEN-VALUE IN FRAME f-pg-sel = STRING(dt-ini)  */
/*              fi_dt_fim:SCREEN-VALUE IN FRAME f-pg-sel = STRING(dt-fim). */

ASSIGN fi_dt_ini:SCREEN-VALUE IN FRAME f-pg-sel = STRING(data-ini)
       fi_dt_fim:SCREEN-VALUE IN FRAME f-pg-sel = STRING(data-fim)
       fi_cod-ctrab_ini:SCREEN-VALUE IN FRAME f-pg-sel = STRING(centro).

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-del-perfil w-relat 
PROCEDURE pi-del-perfil :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
/*                                                                                   */
/* DEF INPUT PARAM P-PERFIL AS CHAR NO-UNDO.                                         */
/* DEFINE VARIABLE C-PERFIL LIKE CST_PERFIL.perfil   NO-UNDO.                        */
/*                                                                                   */
/*  ASSIGN C-PERFIL = P-PERFIL                                                       */
/*         C-PERFIL = TRIM(C-PERFIL).                                                */
/* FIND FIRST CST_PERFIL                                                             */
/*       WHERE  CST_PERFIL.perfil  = C-PERFIL                                        */
/*         AND  CST_PERFIL.usuario = c-seg-usuario  EXCLUSIVE-LOCK NO-ERROR.         */
/* IF AVAIL CST_PERFIL THEN DO:                                                      */
/*    DELETE CST_PERFIL.                                                             */
/*                                                                                   */
/*                                                                                   */
/*                                                                                   */
/* END.                                                                              */
/* RUN displayPerfil.                                                                */
/*                                                                                   */
/*   ASSIGN                                                                          */
/*     fi_dt_ini       :SCREEN-VALUE IN FRAME f-pg-sel = STRING(dt-ini)              */
/*     fi_dt_fim       :SCREEN-VALUE IN FRAME f-pg-sel = STRING(dt-fim)              */
/*     fi_estab_ini    :SCREEN-VALUE IN FRAME f-pg-sel = STRING('' )                 */
/*     fi_estab_fim    :SCREEN-VALUE IN FRAME f-pg-sel = STRING( 'ZZZZZ')            */
/*     fi_cod-ctrab_ini:SCREEN-VALUE IN FRAME f-pg-sel = STRING('')                  */
/*     fi_cod-ctrab_fim:SCREEN-VALUE IN FRAME f-pg-sel = STRING('ZZZZZZZZZZZZZZZZ' ) */
/*     fi_gm-codigo_ini:SCREEN-VALUE IN FRAME f-pg-sel = STRING(''  )                */
/*     fi_gm-codigo_fim:SCREEN-VALUE IN FRAME f-pg-sel = STRING('ZZZZZZZZZZ' )       */
/*     fi_perfil       :SCREEN-VALUE IN FRAME f-pg-par = ''                          */
/*     tg-1            :CHECKED IN FRAME f-pg-par      = NO                          */
/*     tg-2            :CHECKED IN FRAME f-pg-par      = NO.                         */
/*                                                                                   */

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-executar w-relat 
PROCEDURE pi-executar :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
fi_cod-ctrab_fim:SCREEN-VALUE IN FRAME f-pg-sel = fi_cod-ctrab_ini:SCREEN-VALUE IN FRAME f-pg-sel.
fi_estab_fim:SCREEN-VALUE IN FRAME f-pg-sel = fi_estab_ini:SCREEN-VALUE IN FRAME f-pg-sel.

define var r-tt-digita as rowid no-undo.

do on error undo, return error on stop  undo, return error:
    {include/i-rpexa.i}
    /*14/02/2005 - tech1007 - Alterada condicao para n∆o considerar mai o RTF como destino*/
    if input frame f-pg-imp rs-destino = 2 and
       input frame f-pg-imp rs-execucao = 1 then do:
        run utp/ut-vlarq.p (input input frame f-pg-imp c-arquivo).
        
        if return-value = "NOK":U then do:
            run utp/ut-msgs.p (input "show":U, input 73, input "").
            
            apply "MOUSE-SELECT-CLICK":U to im-pg-imp in frame f-relat.
            apply "ENTRY":U to c-arquivo in frame f-pg-imp.
            return error.
        end.
    end.

    /*14/02/2005 - tech1007 - Teste efetuado para nao permitir modelo em branco*/
    &IF "{&RTF}":U = "YES":U &THEN
    IF ( INPUT FRAME f-pg-imp c-modelo-rtf = "" AND
         INPUT FRAME f-pg-imp l-habilitaRtf = "Yes" ) OR
       ( SEARCH(INPUT FRAME f-pg-imp c-modelo-rtf) = ? AND
         input frame f-pg-imp rs-execucao = 1 AND
         INPUT FRAME f-pg-imp l-habilitaRtf = "Yes" )
         THEN DO:
        run utp/ut-msgs.p (input "show":U, input 73, input "").        
        apply "MOUSE-SELECT-CLICK":U to im-pg-imp in frame f-relat.
        /*30/12/2004 - tech1007 - Evento removido pois causa problemas no WebEnabler*/
        /*apply "CHOOSE":U to bt-modelo-rtf in frame f-pg-imp.*/
        return error.
    END.
    &endif
    /*Fim teste Modelo*/
    
    /*:T Coloque aqui as validaá‰es da p†gina de Digitaá∆o, lembrando que elas devem
       apresentar uma mensagem de erro cadastrada, posicionar nesta p†gina e colocar
       o focus no campo com problemas */
    /*browse br-digita:SET-REPOSITIONED-ROW (browse br-digita:DOWN, "ALWAYS":U).*/
    
    for each tt-digita no-lock:
        assign r-tt-digita = rowid(tt-digita).
        
        /*:T Validaá∆o de duplicidade de registro na temp-table tt-digita */
        find first b-tt-digita where b-tt-digita.centro = tt-digita.centro and 
                                     rowid(b-tt-digita) <> rowid(tt-digita) no-lock no-error.
        if avail b-tt-digita then do:
            apply "MOUSE-SELECT-CLICK":U to im-pg-dig in frame f-relat.
            reposition br-digita to rowid rowid(b-tt-digita).
            
            run utp/ut-msgs.p (input "show":U, input 108, input "").
            apply "ENTRY":U to tt-digita.centro in browse br-digita.
            
            return error.
        end.
        
        /*:T As demais validaá‰es devem ser feitas aqui */
        if tt-digita.centro = '' then do:
            assign browse br-digita:CURRENT-COLUMN = tt-digita.centro:HANDLE in browse br-digita.
            
            apply "MOUSE-SELECT-CLICK":U to im-pg-dig in frame f-relat.
            reposition br-digita to rowid r-tt-digita.
            
            run utp/ut-msgs.p (input "show":U, input 99999, input "").
            apply "ENTRY":U to tt-digita.centro in browse br-digita.
            
            return error.
        end.
        
    end.
    
    
    /*:T Coloque aqui as validaá‰es das outras p†ginas, lembrando que elas devem 
       apresentar uma mensagem de erro cadastrada, posicionar na p†gina com 
       problemas e colocar o focus no campo com problemas */
    
    
    
    /*:T Aqui s∆o gravados os campos da temp-table que ser† passada como parÉmetro
       para o programa RP.P */
    
    create tt-param.
    assign tt-param.usuario            = c-seg-usuario
           tt-param.destino            = input frame f-pg-imp rs-destino
           tt-param.data-exec          = today
           tt-param.hora-exec          = time
           tt-param.dt_dt_ini          =  INPUT FRAME f-pg-sel fi_dt_ini          
           tt-param.dt_dt_fim          =  INPUT FRAME f-pg-sel fi_dt_fim          
           tt-param.c_estab_ini        =  INPUT FRAME f-pg-sel fi_estab_ini       
           tt-param.c_estab_fim        =  INPUT FRAME f-pg-sel fi_estab_fim       
           tt-param.c_cod-ctrab_ini    =  INPUT FRAME f-pg-sel fi_cod-ctrab_ini   
           tt-param.c_cod-ctrab_fim    =  INPUT FRAME f-pg-sel fi_cod-ctrab_fim   
           tt-param.c_gm-codigo_ini    =  INPUT FRAME f-pg-sel fi_gm-codigo_ini   
           tt-param.c_gm-codigo_fim    =  INPUT FRAME f-pg-sel fi_gm-codigo_fim  
           tt-param.c_item_ini         =  INPUT FRAME f-pg-sel FI_ITE-INI
           tt-param.c_item_fim         =  INPUT FRAME f-pg-sel FI_ITE-fim
           tt-param.l-diario           =  tg-1:CHECKED IN FRAME f-pg-par          
           tt-param.l-mensal           =  tg-2:CHECKED IN FRAME f-pg-par     .    
           &IF "{&RTF}":U = "YES":U &THEN
           tt-param.modelo-rtf      = INPUT FRAME f-pg-imp c-modelo-rtf
           /*Alterado 14/02/2005 - tech1007 - Armazena a informaá∆o se o RTF est† habilitado ou n∆o*/
           tt-param.l-habilitaRtf     = INPUT FRAME f-pg-imp l-habilitaRtf
           /*Fim alteracao 14/02/2005*/ 
           &endif
           .
    
    /*Alterado 14/02/2005 - tech1007 - Alterado o teste para verificar se a opá∆o de RTF est† selecionada*/
    if tt-param.destino = 1 
    then assign tt-param.arquivo = "".
    else if  tt-param.destino = 2
         then assign tt-param.arquivo = input frame f-pg-imp c-arquivo.
         else assign tt-param.arquivo = session:temp-directory + c-programa-mg97 + ".tmp":U.
    /*Fim alteracao 14/02/2005*/

    /*:T Coloque aqui a/l¢gica de gravaá∆o dos demais campos que devem ser passados
       como parÉmetros para o programa RP.P, atravÇs da temp-table tt-param */
    
    
    
    /*:T Executar do programa RP.P que ir† criar o relat¢rio */
    {include/i-rpexb.i}
    
    SESSION:SET-WAIT-STATE("general":U).
    
    {include/i-rprun.i SFC/sfcp0010rp.p}
    
    {include/i-rpexc.i}
    
    SESSION:SET-WAIT-STATE("":U).
    
    {include/i-rptrm.i}

end.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-perfil w-relat 
PROCEDURE pi-perfil :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
/*                                                                                                          */
/* DEF INPUT PARAM P-PERFIL AS CHAR NO-UNDO.                                                                */
/* DEFINE VARIABLE C-PERFIL LIKE CST_PERFIL.perfil   NO-UNDO.                                               */
/* EMPTY temp-table tt-digita.                                                                              */
/*  ASSIGN C-PERFIL = P-PERFIL                                                                              */
/*         C-PERFIL = TRIM(C-PERFIL).                                                                       */
/* FIND FIRST CST_PERFIL                                                                                    */
/*       WHERE  CST_PERFIL.perfil  = C-PERFIL                                                               */
/*         AND  CST_PERFIL.usuario = c-seg-usuario  NO-LOCK NO-ERROR.                                       */
/* IF AVAIL CST_PERFIL THEN DO:                                                                             */
/*    ASSIGN                                                                                                */
/*     fi_dt_ini       :SCREEN-VALUE IN FRAME f-pg-sel = STRING(CST_PERFIL.dt-ini    )                      */
/*     fi_dt_fim       :SCREEN-VALUE IN FRAME f-pg-sel = STRING(CST_PERFIL.dt-final )                       */
/*     fi_estab_ini    :SCREEN-VALUE IN FRAME f-pg-sel = STRING(CST_PERFIL.cod-estabel-ini )                */
/*     fi_estab_fim    :SCREEN-VALUE IN FRAME f-pg-sel = STRING( CST_PERFIL.cod-estabel-fim)                */
/*     fi_cod-ctrab_ini:SCREEN-VALUE IN FRAME f-pg-sel = STRING(CST_PERFIL.cod-ctrab-ini )                  */
/*     fi_cod-ctrab_fim:SCREEN-VALUE IN FRAME f-pg-sel = STRING(CST_PERFIL.cod-ctrab-fim )                  */
/*     fi_gm-codigo_ini:SCREEN-VALUE IN FRAME f-pg-sel = STRING(CST_PERFIL.gm-codigo-ini  )                 */
/*     fi_gm-codigo_fim:SCREEN-VALUE IN FRAME f-pg-sel = STRING(CST_PERFIL.gm-codigo-fim )                  */
/*     tg-1            :CHECKED IN FRAME f-pg-par      = CST_PERFIL.log-diario                              */
/*     tg-2            :CHECKED IN FRAME f-pg-par      = CST_PERFIL.log-mensal.                             */
/*                                                                                                          */
/*   FOR EACH  cst_perfil_digita                                                                            */
/*       WHERE  cst_perfil_digita.perfil     = C-PERFIL                                                     */
/*       AND     cst_perfil_digita.usuario   = c-seg-usuario  NO-LOCK:                                      */
/*                                                                                                          */
/*                                                                                                          */
/*       CREATE  tt-digita.                                                                                 */
/*       ASSIGN tt-digita.CENTRO  = cst_perfil_digita.cod-ctrab                                             */
/*              tt-digita.GRUPO   = cst_perfil_digita.gm-codigo.                                            */
/*                                                                                                          */
/*      FOR FIRST ctrab FIELDS (cod-ctrab des-ctrab gm-codigo )   USE-INDEX id                              */
/*                  WHERE ctrab.cod-ctrab = tt-digita.CENTRO NO-LOCK:                                       */
/*           ASSIGN                                                                                         */
/*                                                                                                          */
/*                  tt-digita.DESC-centro                                  = ctrab.des-ctrab.               */
/*                                                                                                          */
/*       END.                                                                                               */
/*         FOR FIRST grup-maquina FIELDS (descricao gm-codigo )                                             */
/*                  WHERE  grup-maquina.gm-codigo = tt-digita.GRUPO NO-LOCK:                                */
/*                                                                                                          */
/*                                                                                                          */
/*                   ASSIGN                                                                                 */
/*                          tt-digita.DESC-GRUPO                                  = grup-maquina.descricao. */
/*             END.                                                                                         */
/*                                                                                                          */
/*                                                                                                          */
/*   END.                                                                                                   */
/*   {&OPEN-QUERY-br-digita}                                                                                */
/*                                                                                                          */
/* END.                                                                                                     */
/*                                                                                                          */







                                                       
 
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-save-perfil w-relat 
PROCEDURE pi-save-perfil :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
/* IF INPUT FRAME f-pg-par fi_perfil  = '' THEN DO:                               */
/*      RUN utp/ut-msgs.p ( INPUT "SHOW":U,                                       */
/*                          INPUT 17006,                                          */
/*                          INPUT "Perfil deve ser informato" ).                  */
/*         RETURN "NOK":U.                                                        */
/*                                                                                */
/* END.                                                                           */
/*                                                                                */
/*                                                                                */
/* FIND FIRST CST_PERFIL                                                          */
/*       WHERE  CST_PERFIL.perfil  = INPUT FRAME f-pg-par fi_perfil               */
/*       AND    CST_PERFIL.usuario = c-seg-usuario EXCLUSIVE-LOCK NO-ERROR.       */
/* IF NOT AVAIL CST_PERFIL THEN DO:                                               */
/*    CREATE CST_PERFIL.                                                          */
/*    ASSIGN CST_PERFIL.perfil  = INPUT FRAME f-pg-par fi_perfil                  */
/*           CST_PERFIL.usuario = c-seg-usuario.                                  */
/* END.                                                                           */
/* ASSIGN  CST_PERFIL.dt-ini           =    INPUT FRAME f-pg-sel fi_dt_ini        */
/*         CST_PERFIL.dt-final         =    INPUT FRAME f-pg-sel fi_dt_fim        */
/*         CST_PERFIL.cod-estabel-ini  =    INPUT FRAME f-pg-sel fi_estab_ini     */
/*         CST_PERFIL.cod-estabel-fim  =    INPUT FRAME f-pg-sel fi_estab_fim     */
/*         CST_PERFIL.cod-ctrab-ini    =    INPUT FRAME f-pg-sel fi_cod-ctrab_ini */
/*         CST_PERFIL.cod-ctrab-fim    =    INPUT FRAME f-pg-sel fi_cod-ctrab_fim */
/*         CST_PERFIL.gm-codigo-ini    =    INPUT FRAME f-pg-sel fi_gm-codigo_ini */
/*         CST_PERFIL.gm-codigo-fim    =    INPUT FRAME f-pg-sel fi_gm-codigo_fim */
/*         CST_PERFIL.log-diario       =    tg-1:CHECKED IN FRAME f-pg-par        */
/*         CST_PERFIL.log-mensal       =    tg-2:CHECKED IN FRAME f-pg-par     .  */
/*                                                                                */
/*                                                                                */
/*  FOR EACH  tt-digita:                                                          */
/*       CREATE cst_perfil_digita.                                                */
/*       ASSIGN   cst_perfil_digita.perfil     = INPUT FRAME f-pg-par fi_perfil   */
/*                cst_perfil_digita.usuario    = c-seg-usuario                    */
/*                cst_perfil_digita.cod-ctrab  =  tt-digita.CENTRO                */
/*                cst_perfil_digita.gm-codigo  =  tt-digita.GRUPO  .              */
/*                                                                                */
/*                                                                                */
/*                                                                                */
/*                                                                                */
/*   END.                                                                         */
/*                                                                                */
/*                                                                                */
/*                                                                                */
/*                                                                                */
/*                                                                                */
/*                                                                                */
/*                                                                                */
/* RUN displayPerfil.                                                             */
/* ASSIGN fi_perfil:SCREEN-VALUE IN FRAME f-pg-par = ''.                          */

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-troca-pagina w-relat 
PROCEDURE pi-troca-pagina :
/*:T------------------------------------------------------------------------------
  Purpose: Gerencia a Troca de P†gina (folder)   
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/

//{include/i-rptrp.i}
    

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE send-records w-relat  _ADM-SEND-RECORDS
PROCEDURE send-records :
/*------------------------------------------------------------------------------
  Purpose:     Send record ROWID's for all tables used by
               this file.
  Parameters:  see template/snd-head.i
------------------------------------------------------------------------------*/

  /* Define variables needed by this internal procedure.               */
  {src/adm/template/snd-head.i}

  /* For each requested table, put it's ROWID in the output list.      */
  {src/adm/template/snd-list.i "tt-digita"}
  {src/adm/template/snd-list.i "ctrab"}

  /* Deal with any unexpected table requests before closing.           */
  {src/adm/template/snd-end.i}

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE state-changed w-relat 
PROCEDURE state-changed :
/* -----------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
-------------------------------------------------------------*/
  //DEFINE INPUT PARAMETER p-issuer-hdl AS HANDLE NO-UNDO.
  //DEFINE INPUT PARAMETER p-state AS CHARACTER NO-UNDO.
  
  //run pi-trata-state (p-issuer-hdl, p-state).
  
  
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

