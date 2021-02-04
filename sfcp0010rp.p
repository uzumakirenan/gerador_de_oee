&ANALYZE-SUSPEND _VERSION-NUMBER AB_v10r12
&ANALYZE-RESUME
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS Procedure 
/*------------------------------------------------------------------------
File        :
Purpose     :  

Syntax      :

Description :

Author(s)   :
Created     :
Notes       :                              
----------------------------------------------------------------------*/
/*          This .W file was created with the Progress AppBuilder.      */
/*----------------------------------------------------------------------*/
{include/i-prgvrs.i   SFCP0010RP  2.00.00.001}  /*** 010001 ***/
/* ***************************  Definitions  ************************** */

/*****************************************************************************
**
**       PROGRAMA:
**
**       DATA....: 16/07/2019
**
**       AUTOR...:
**
**       OBJETIVO: Gera?∆o de Estatiticas de OEE
**
*****************************************************************************/
DEF VAR c-prg-obj AS c.
DEF VAR c-prg-vrs AS c.



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

DEF TEMP-TABLE ttTurno
    FIELD cod-model-turno LIKE  turno-semana.cod-model-turno
    FIELD dia       AS int
    FIELD num-turno LIKE  turno-semana.num-turno
    FIELD qtd-tempo-util-dia-1 as decimal format ">9.99"
    FIELD qtd-tempo-util-dia-2 as decimal format ">9.99"
    FIELD qtd-tempo-util-dia-3 as decimal format ">9.99"
    FIELD qtd-tempo-util-dia-4 as decimal format ">9.99"
    FIELD qtd-tempo-util-dia-5 as decimal format ">9.99"
    FIELD qtd-tempo-util-dia-6 as decimal format ">9.99"
    FIELD qtd-tempo-util-dia-7 as decimal format ">9.99"
     index ch-codigo     cod-model-turno dia.
 
{sfc/sfapi002.i} /*** Temp-tables **/

DEF TEMP-TABLE ttDados
   field cod-ctrab         like ctrab.cod-ctrab
   field DESC-CENTRO       as character format "x(30)"
   field gm-codigo         like ctrab.gm-codigo
   FIELD it-codigo         LIKE ITEM.it-codigo
   field data-inicio       like area-produc-ctrab.dat-inic-valid
   field data-fim          like area-produc-ctrab.dat-fim-valid
   field tempo-disponivel  as decimal format ">>>9.99"
   field tempo-parada      as decimal format ">>>9.99"
   field tempo-parada-nef  as decimal format ">>>9.99"
   field tempo-utilizado   as decimal format ">>>9.99"
   field tempo-padrao      as decimal format ">>>9.99"
   field tempo-extra       as decimal format ">>>9.99"
   field tempo-par-ext     as decimal format ">>>9.99"
   field tempo-par-nef-ext as decimal format ">>>9.99"
   field carga-disponivel  as decimal format ">>>>,>>9.9999"
   field carga-utilizada   as decimal format ">>>>,>>9.9999"
   field fator-refugo      as decimal format ">>9.99" 
   field fator-retrab      as decimal format ">>9.99" 
   field eficiencia-media  as decimal format ">>9.99"
   field qtd-segs-inic     as int format ">>>>9" init ?
   field qtd-segs-fim      as int format ">>>>9" init ?
   field de-classif        as decimal format ">>9.99"
   field fator-utiliz      as decimal format ">>9.99"
   field cod-area          like area-produc-ctrab.cod-area-produc
   field qtd-padrao        as dec  format ">>,>>>,>>9.9999" 
   field qtd-aprov         like rep-oper-ctrab.qtd-operac-aprov
   field qtd-refgada       like rep-oper-ctrab.qtd-operac-refgda
   field qtd-retrab        like rep-oper-ctrab.qtd-operac-retrab
   FIELD val-eficien-ctrab as decimal format ">>>9.99"
   FIELD oee               as decimal format ">>9.99"
   FIELD c-horas           AS char
   index ch-codigo         cod-ctrab gm-codigo data-inicio data-fim.

DEF TEMP-TABLE ttDadosComp
   field cod-ctrab         like ctrab.cod-ctrab
   field DESC-CENTRO       as character format "x(30)"
   field gm-codigo         like ctrab.gm-codigo
   FIELD it-codigo         LIKE ITEM.it-codigo
   field data-inicio       like area-produc-ctrab.dat-inic-valid
   field data-fim          like area-produc-ctrab.dat-fim-valid
   field tempo-disponivel  as decimal format ">>>9.99"
   field tempo-parada      as decimal format ">>>9.99"
   field tempo-parada-nef  as decimal format ">>>9.99"
   field tempo-utilizado   as decimal format ">>>9.99"
   field tempo-padrao      as decimal format ">>>9.99"
   field tempo-extra       as decimal format ">>>9.99"
   field tempo-par-ext     as decimal format ">>>9.99"
   field tempo-par-nef-ext as decimal format ">>>9.99"
   field carga-disponivel  as decimal format ">>>>,>>9.9999"
   field carga-utilizada   as decimal format ">>>>,>>9.9999"
   field fator-refugo      as decimal format ">>9.99" 
   field fator-retrab      as decimal format ">>9.99" 
   field eficiencia-media  as decimal format ">>9.99"
   field qtd-segs-inic     as int format ">>>>9" init ?
   field qtd-segs-fim      as int format ">>>>9" init ?
   field de-classif        as decimal format ">>9.99"
   field fator-utiliz      as decimal format ">>9.99"
   field cod-area          like area-produc-ctrab.cod-area-produc
   field qtd-padrao        as dec  format ">>,>>>,>>9.9999" 
   field qtd-aprov         like rep-oper-ctrab.qtd-operac-aprov
   field qtd-refgada       like rep-oper-ctrab.qtd-operac-refgda
   field qtd-retrab        like rep-oper-ctrab.qtd-operac-retrab
   FIELD val-eficien-ctrab as decimal format ">>>9.99"
   FIELD oee               as decimal format ">>9.99"
   FIELD c-horas           AS char
   index ch-codigo         cod-ctrab gm-codigo data-inicio data-fim.
   
DEF TEMP-TABLE ttDadosMensal
   field cod-ctrab         like ctrab.cod-ctrab
   field DESC-CENTRO       as character format "x(30)"
   field gm-codigo         like ctrab.gm-codigo
   FIELD it-codigo         LIKE ITEM.it-codigo
   field data-inicio       like area-produc-ctrab.dat-inic-valid
   field data-fim          like area-produc-ctrab.dat-fim-valid
   field tempo-disponivel  as decimal format ">>>9.99"
   field tempo-parada      as decimal format ">>>9.99"
   field tempo-parada-nef  as decimal format ">>>9.99"
   field tempo-utilizado   as decimal format ">>>9.99"
   field tempo-padrao      as decimal format ">>>9.99"
   field tempo-extra       as decimal format ">>>9.99"
   field tempo-par-ext     as decimal format ">>>9.99"
   field tempo-par-nef-ext as decimal format ">>>9.99"
   field carga-disponivel  as decimal format ">>>>,>>9.9999"
   field carga-utilizada   as decimal format ">>>>,>>9.9999"
   field fator-refugo      as decimal format ">>9.99" 
   field fator-retrab      as decimal format ">>9.99" 
   field eficiencia-media  as decimal format ">>9.99"
   field qtd-segs-inic     as int format ">>>>9" init ?
   field qtd-segs-fim      as int format ">>>>9" init ?
   field de-classif        as decimal format ">>9.99"
   field fator-utiliz      as decimal format ">>9.99"
   field cod-area          like area-produc-ctrab.cod-area-produc
   field qtd-padrao        as dec  format ">>,>>>,>>9.9999" 
   field qtd-aprov         like rep-oper-ctrab.qtd-operac-aprov
   field qtd-refgada       like rep-oper-ctrab.qtd-operac-refgda
   field qtd-retrab        like rep-oper-ctrab.qtd-operac-retrab
   FIELD val-eficien-ctrab as decimal format ">>>9.99"
   FIELD oee               as decimal format ">>9.99"
   FIELD c-horas           AS CHAR
   FIELD c-mes             AS CHAR
   FIELD i-ano             AS int
   index ch-codigo         cod-ctrab c-mes.  
   
   
  DEF TEMP-TABLE ttTotalMensal
   field de-classif        as decimal format ">>9.99"
   field fator-utiliz      as decimal format ">>9.99"
   FIELD val-eficien-ctrab as decimal format ">>>9.99"
   FIELD oee               as decimal format ">>9.99"
   field eficiencia-media  as decimal format ">>9.99"
   FIELD i-mes             AS int
   FIELD i-ano             AS int
   index ch-codigo         i-ano i-mes.   
   
 DEF TEMP-TABLE ttPeriodo
        FIELD i-mes AS INT
        FIELD i-ano AS INT
        index ch-codigo         i-ano i-mes.  

DEF TEMP-TABLE ttTotalCentro
    field cod-ctrab         like ctrab.cod-ctrab
    FIELD c-horas           AS char
    field tempo-utilizado   as decimal format ">>>9.99"
    field tempo-parada      as decimal format ">>>9.99"
    field tempo-setup       as decimal format ">>>9.99"
    FIELD qtd-capac-ctrab   LIKE ctrab.qtd-capac-ctrab
    index ch-codigo         cod-ctrab c-horas.
 
DEF TEMP-TABLE ttCentro
   field cod-ctrab         like ctrab.cod-ctrab
   field DESC-CENTRO       as character format "x(30)"
   field gm-codigo         like ctrab.gm-codigo
   field DESC-GRUPO        as character format "x(30)"
   FIELD qtd-capac-ctrab   LIKE ctrab.qtd-capac-ctrab
   INDEX idx1 cod-ctrab.
   
 DEF TEMP-TABLE ttParadas
     field cod-ctrab         like ctrab.cod-ctrab
     field data-inicio       like rep-parada-ctrab.dat-reporte
     FIELD cod-parada        LIKE rep-parada-ctrab.cod-parada
     FIELD motivo            LIKE motiv-parada.des-parada
     FIELD tempo            as decimal format ">>>9.99"
     INDEX idx cod-ctrab data-inicio cod-parada.
     
     
DEF TEMP-TABLE ttParadasAux
     field cod-ctrab         like ctrab.cod-ctrab
     field data-inicio       like rep-parada-ctrab.dat-reporte
     FIELD cod-parada        LIKE rep-parada-ctrab.cod-parada
     FIELD motivo            LIKE motiv-parada.des-parada
     FIELD tempo            as decimal format ">>>9.99"
     INDEX idx cod-ctrab data-inicio cod-parada.
     
  DEF TEMP-TABLE ttListParadas
     field cod-ctrab         like ctrab.cod-ctrab
     field data-inicio       like area-produc-ctrab.dat-inic-valid
     field gm-codigo         like ctrab.gm-codigo
     FIELD cod-parada        LIKE rep-parada-ctrab.cod-parada
     FIELD it-codigo         LIKE ITEM.it-codigo
     FIELD motivo            LIKE motiv-parada.des-parada
     FIELD qtd               AS INT
     FIELD desc-oper         AS CHAR
     FIELD deparada          as decimal format ">>>9.99"
     FIELD tempo             as decimal format ">>>9.99"
     INDEX idx cod-ctrab data-inicio it-codigo cod-parada
     INDEX idx1 cod-ctrab data-inicio.

  
  DEF TEMP-TABLE ttParadasMes
     field cod-ctrab         like ctrab.cod-ctrab
     field gm-codigo         like ctrab.gm-codigo
     FIELD i-ano             AS INT
     FIELD i-mes             AS INT
     FIELD cod-parada        LIKE rep-parada-ctrab.cod-parada
     FIELD motivo            LIKE motiv-parada.des-parada
     FIELD qtd               AS INT
     FIELD tempo             as decimal format ">>>9.99"
     FIELD val-eficien-ctrab as decimal format ">>>9.99"
     INDEX idx cod-ctrab i-ano i-mes  cod-parada.
     
 DEF TEMP-TABLE ttParadasTotalMes
     field cod-ctrab         like ctrab.cod-ctrab
     FIELD i-ano             AS INT
     FIELD i-mes             AS INT
     FIELD qtd               AS INT
     FIELD tempo             as decimal format ">>>9.99"
     FIELD val-eficien-ctrab as decimal format ">>>9.99"
     INDEX idx cod-ctrab i-ano i-mes.
      
     
   
   
 DEF TEMP-TABLE ttItem
     field cod-ctrab         like ctrab.cod-ctrab
     FIELD it-codigo         LIKE ITEM.it-codigo
     INDEX idx1 cod-ctrab it-codigo.


 DEF TEMP-TABLE ttDias
     field cod-ctrab         like ctrab.cod-ctrab
     FIELD dia               AS DATE
     INDEX idx cod-ctrab  dia
     INDEX idx2 cod-ctrab.

 def temp-table tt-operac-ctrab-rp_aux NO-UNDO
    field op-codigo            like operacao.op-codigo
    field cod-ctrab            like ctrab.cod-ctrab
    field cod-roteiro          like rep-oper.cod-roteiro
    field cod-ferr-prod        like ferr-prod.cod-ferr-prod
    field it-codigo            like item.it-codigo
    field op-altern            like op-altern.op-altern
    field qtd-aprov            like rep-oper-ctrab.qtd-operac-aprov
    field qtd-refgada          like rep-oper-ctrab.qtd-operac-refgda
    field qtd-retrab           like rep-oper-ctrab.qtd-operac-retrab
    field perc-refugo          as decimal format ">>9.99"
    field cap-utilizada        as decimal format ">>>>,>>9.9999"
    field cap-especif          as decimal format ">>>>,>>9.9999"
    field qtd-tempo-real       as decimal format ">>>9.99"
    field qtd-tempo-padrao     as decimal format ">>>9.99"
    field qtd-tempo-mod        as decimal format ">>>9.99"
    field qtd-padrao           as dec  format ">>,>>>,>>9.9999"  
    field qtd-tempo-pad-mod    as decimal format ">>>9.99"
    field qtd-tempo-extra      as decimal format ">>>9.99"
    field qtd-tempo-par-rep    as decimal format ">>>9.99"  /*** Tempo de parada no reporte ***/
    field fator-item-un-al     as decimal format ">>>>9.9<<<<"
    FIELD cod-refer            LIKE ITEM.cod-refer
    index ch-indice            cod-ctrab cod-ferr-prod  it-codigo cod-refer cod-roteiro  op-codigo
    index operacao             it-codigo cod-refer cod-roteiro    op-codigo cod-ctrab
    INDEX idx cod-ctrab.


/* Transfer Definitions */
def temp-table tt-raw-digita
field raw-digita      as raw.
/****************************************************************************/
/** Defini?∆o de Par?metros **/
def input parameter raw-param as raw no-undo.
def input parameter table for tt-raw-digita.
/****************************************************************************/
/** Transfer?ncia de par?metros para temp-table padr∆o **/
create tt-param.
raw-transfer raw-param to tt-param.

FOR EACH tt-raw-digita:
    CREATE tt-digita.
    RAW-TRANSFER tt-raw-digita.raw-digita TO tt-digita.
END.


/****************************************************************************/
/** Defini?∆o de Vari‡veis **/

{include/i-rpvar.i}
/****************************************************************************/
/***** DEFINI?«O DE TEMP-TABLES/BUFFER ******************/
/********** DEFINI?ÂES DE VARIAVEIS GLOBAIS **************/
define variable h-acomp        as handle     no-undo.
/***** Excel *****************/
DEFINE VARIABLE chExcel             AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE chExcelApplication  AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE ch-Arquivo          AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE ch-planilha         AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE chWorkbook          AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE chWorksheet         AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE chSelection         AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE chPasta             AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE chPlanilha          AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE chArquivo           AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE chPlanilhaMod       AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE cFile               AS CHARACTER   NO-UNDO.
DEFINE VARIABLE i-linha             AS INTEGER     NO-UNDO.
DEFINE VARIABLE i-linha-ini         AS INTEGER     NO-UNDO.
DEFINE VARIABLE i-linha-fim         AS INTEGER     NO-UNDO.
DEFINE VARIABLE i-linha-aux         AS INTEGER     NO-UNDO.
DEFINE VARIABLE i-coluna            AS INTEGER     NO-UNDO.
DEFINE VARIABLE C-RANGE             AS CHARACTER   NO-UNDO.
DEFINE VARIABLE chImagem            AS HANDLE      NO-UNDO.
DEFINE VARIABLE i-planilha          AS INT         NO-UNDO.
DEFINE VARIABLE ch-celula           AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE c-modelo            AS CHARACTER   NO-UNDO.
DEFINE VARIABLE C-RANGE-GRAF        AS CHARACTER   NO-UNDO.
DEFINE VARIABLE C-FAIXA-GRAF        AS CHARACTER   NO-UNDO.
DEFINE VARIABLE chChart             AS COM-HANDLE  NO-UNDO. 
DEFINE VARIABLE chWorksheetRange    AS COM-HANDLE  NO-UNDO.
DEFINE VARIABLE chRange             AS COM-HANDLE NO-UNDO.
DEFINE VARIABLE i-cont-reg          AS INT         NO-UNDO.
DEFINE VARIABLE iRow                AS INTEGER    NO-UNDO.
DEFINE VARIABLE iContSeries         AS INTEGER    NO-UNDO.
DEFINE VARIABLE C-FOTO              AS CHARACTER   NO-UNDO.
/************ API SF*************/
def var h-sfapi002        as handle                                    no-undo.
def var h-boin535a        as handle                                    no-undo.
def var de-tempo-turno    as dec  format ">9.99"                       no-undo.  
def var de-fator-parada   as dec  format ">>9.99"   label "      "     no-undo.
def var de-fator-utiliz   as dec  format ">>9.99"   label "       "    no-undo.
def var de-perc-refugo    as dec  format ">>9.99"   label "      "     no-undo.
def var de-par-total      as dec  format ">,>>9.99" label "          " no-undo.
def var de-tempo-dispon   as dec  format ">,>>9.99"                    no-undo.
def var de-par-normal     as dec  format ">,>>9.99"                    no-undo.
def var de-util-normal    like tt-ctrab-rp.tempo-utilizado             no-undo.
def var de-ret-ref        like rep-oper-ctrab.qtd-operac-retrab        no-undo.
def var de-reportada      like rep-oper-ctrab.qtd-operac-aprov         no-undo.
def var c-total           as char format "x(6)"       no-undo.
def var de-tot-tp-dispon  as dec  format ">>>,>>9.99" no-undo.
def var de-tot-aprov      as dec  format ">>>>,>>9.99" no-undo.
def var de-tot-ref        as dec  format ">>,>>9.99" no-undo.
def var de-tot-rep        as dec  format ">>>>,>>9.99" no-undo.
def var de-tot-perc-ref   as dec  format ">>>,>>9.99" no-undo.
def var de-tot-real       as dec  format ">>>,>>9.99" no-undo.
def var de-tot-pad        as dec  format ">>>,>>9.99" no-undo.
def var de-tot-tp-extra   as dec  format ">>>,>>9.99"   no-undo.
def var de-tot-par-normal as dec  format ">>>,>>9.99"   no-undo.
def var de-tot-par-total  as dec  format ">>>,>>9.99"   no-undo.
def var de-tot-util-norm  as dec  format ">>>,>>9.99" no-undo.
def var de-tot-tp-utiliz  as dec  format ">>>,>>9.99" no-undo.
def var de-tot-perc-util  as dec  format ">>9.99"     no-undo.
def var de-tot-perc-para  as dec  format ">>9.99"     no-undo.
def var v-val-refer-inic  as dec  format ">>>>>>>>,>>>>9":U no-undo.
def var v-val-refer-fim   as dec  format ">>>>>>>>,>>>>9":U no-undo.

/* Variaveis Calculo de paradas por eficiencia */
def var data-ini as date. 
def var data-fim as date.  
def var data-ini-paradas as date. 
def var data-fim-paradas as date. 
def var v-ctrab as char.
def var de_parada_geral       like rep-parada-ctrab.qtd-tempo-parada.
def var de_parada_total_geral like rep-parada-ctrab.qtd-tempo-parada.
def var de_parada_alt_efic    like rep-parada-ctrab.qtd-tempo-parada.
def var de_parada_n_alt_efic  like rep-parada-ctrab.qtd-tempo-parada. 
def var i_cont as integer format "->>>9" .

DEF VAR v-teorica         AS INT NO-UNDO.
DEF VAR v-real            AS INT NO-UNDO.
DEF VAR v-dia-semana      AS INT NO-UNDO.
def var c-hora-ini        as CHAR  no-undo.
DEF VAR cHora             AS CHAR NO-UNDO.
DEF VAR cMesIni           AS CHAR NO-UNDO.
DEF VAR cMesFim           AS CHAR NO-UNDO.
/**** BO E API****/
def var  h-boin469b      AS HANDLE      NO-UNDO.
def var h-sfapi001       AS HANDLE      NO-UNDO.
/*****************/

def var de-setup-tempo    as dec                           no-undo.

def var de-qtd-parada      as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var de-qtd-parada-nao  as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var de-qtd-parada-set  as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.    
def var de-qtd-par-tex     as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var de-qtd-par-nao-tex as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var de-qtd-par-set-tex as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var v-carga-teorica    as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var deutiliz           as decimal format ">>>9.99"  NO-UNDO.
def var desoma             as decimal format ">>>9.99" no-undo.
def var deSTemp            as decimal format ">>>9.99" no-undo.
def var c-cod-calend       as char no-undo.
DEF VAR c-cod-refer        AS CHAR NO-UNDO.
def var v-tempo-padrao     as decimal.
def var v-tempo-homem      as decimal. 
def var v-capac            as decimal.
def var v-total1           as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var v-total2           as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var v-total3           as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var v-total4           as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var v-total5           as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
def var v-total6           like rep-oper-ctrab.qtd-operac-aprov no-undo.
def var v-total7           as decimal format ">>9.99" no-undo.
def var v-total8           as decimal format ">>9.99" no-undo.
def var v-total9           as decimal format ">>9.99" no-undo.
def var v-total10          as decimal format ">>>9.99"  no-undo.
def var v-total11          as decimal format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.
/**** BUFFER ******/
DEF BUFFER b-gm FOR grup-maquina.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 



/* ********************  Preprocessor Definitions  ******************** */

&Scoped-define PROCEDURE-TYPE Procedure
&Scoped-define DB-AWARE no



/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME



/* *********************** Procedure Settings ************************ */

&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: Procedure
   Allow: 
   Frames: 0
   Add Fields to: Neither
   Other Settings: CODE-ONLY COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS

/* *************************  Create Window  ************************** */

&ANALYZE-SUSPEND _CREATE-WINDOW
/* DESIGN Window definition (used by the UIB) 
  CREATE WINDOW Procedure ASSIGN
         HEIGHT             = 20.04
         WIDTH              = 60.
/* END WINDOW DEFINITION */
                                                                        */
&ANALYZE-RESUME

 


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK Procedure 


/* ***************************  Main Block  *************************** */

/** Parametriza padr‰es de cabe?alho e rodap? a serem exibidos **/
run piInicial in this-procedure.

/** Imprime cabe?alho e abre o output para arquivo **/
{include/i-rpcab.i}
{include/i-rpout.i}  /** Abertura do output do programa **/

/** Procedure para inicializar c‡lculos e impress∆o **/
run piPrincipal in this-procedure.

/** Fecha o output para arquivo **/
{include/i-rpclo.i}

return "OK":U.
/*--- Fim do Programa ---*/

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


/* **********************  Internal Procedures  *********************** */

&IF DEFINED(EXCLUDE-buscaParadas) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE buscaParadas Procedure 
PROCEDURE buscaParadas :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/

    def  Input param p-num-action      as int  format ">>>>,>>9" no-undo.
    def  Input param p-cod-ctrab       as char format "x(16)"    no-undo.
    def  Input param p-cod-area        as char format "x(8)"     no-undo.
    def  Input param p-dat-inic-period as date format "99/99/9999"  no-undo.
    def  input param p-qtd-segs-inic   as int  format ">>>>9"       no-undo.
    def  Input param p-dat-fim-period  as date format "99/99/9999"  no-undo.
    def  input param p-qtd-segs-fim    as int  format ">>>>9"       no-undo.
    def  Input param p-log-parada-nef  as log                       no-undo.    /*** Paradas que n∆o alteram efici?ncia ***/
    def  input param p-log-parada-set  as log                       no-undo.    /*** Paradas de Setup                   ***/ 
    def var v-dat-fim-period                 as date            no-undo. 
    def var v-dat-inic-period                as date            no-undo. 
    def var v-dat-relogio                    as date            no-undo. 
    def var v-num-turno                      as integer         no-undo. 
    def var v-qtd-segs-relogio               as decimal         no-undo. 
    def var v-qtd-parada-nef                 as decimal         no-undo. /*** Paradas n∆o afetam efici?ncia ***/
    def var v-qtd-parada-set                 as decimal         no-undo. /*** Paradas de setup              ***/
    def var v-qtd-par-tex                    as decimal         no-undo.
    def var v-qtd-par-nef-tex                as decimal         no-undo.
    def var v-qtd-par-set-tex                as decimal         no-undo.
    def var v-val-refer-fim-parada           as decimal         no-undo. 
    def var v-val-refer-inic-parada          as decimal         no-undo. 
    def var v-val-refer-relogio              as decimal         no-undo. 
    def var v-val-tempo-calcul               as decimal         no-undo.
    def var v-val-tempo-calcul-tex           as decimal         no-undo.
    def var v-num-action                     as integer         no-undo.
    def var c-gm-codigo                      as char            no-undo.
    
    def var v-dat-inic                       as date            no-undo. 
    def var v-qtd-segs-inic                  as integer         no-undo. 
    def var v-dat-fim                        as date            no-undo. 
    def var v-qtd-segs-fim                   as integer         no-undo. 
    def VAR v-qtd-parada                    as dec  format "->>,>>>,>>>,>>9.99999999" decimals 8 no-undo.

    
    /*** Avalia tempo parado para o centro de trabalho ***/
    if  p-num-action >= 3 then do:
    
        find first ctrab
             where ctrab.cod-ctrab = p-cod-ctrab no-lock no-error.
        
        run pi-converte-data-segs-valor (Input p-dat-inic-period,
                                         Input p-qtd-segs-inic,
                                         output v-val-refer-inic-parada).

        run pi-converte-data-segs-valor (Input p-dat-fim-period,
                                         Input p-qtd-segs-fim,
                                         output v-val-refer-fim-parada).

        assign v-val-refer-relogio = v-val-refer-inic-parada
               v-dat-relogio       = p-dat-inic-period
               v-qtd-segs-relogio  = p-qtd-segs-inic.

        /*** O loop vai pesquisando as paradas ordenadas pela data de finaliza?∆o e vai paralelamente atualizando
             um rel¢gio com a data e hora final da parada para evitar contabilizar overlap entre duas paradas
             consecutivas ***/

        reporte-parada:
        for each rep-parada-ctrab
            where rep-parada-ctrab.cod-ctrab            = ctrab.cod-ctrab
              and rep-parada-ctrab.val-refer-fim-parada > v-val-refer-inic-parada no-lock:

            find first grup-maquina
                 where grup-maquina.gm-codigo = ctrab.gm-codigo no-lock no-error. 
         
           /*** Acha tipo de parada para saber se efici?ncia ≤ afetada ou n∆o ***/     
            find motiv-parada where
                 motiv-parada.cod-parada = rep-parada-ctrab.cod-parada no-lock.
                 
           /*** Quando o in°cio da parada ultrapase os limites, sair do loop ***/
            if  rep-parada-ctrab.val-refer-inic-parada >= v-val-refer-fim-parada then
                leave reporte-parada.

            assign v-val-tempo-calcul     = 0
                   v-val-tempo-calcul-tex = 0
                   v-qtd-parada           = 0.

           /*** Se a parada tem data de in°cio menor do que a posi?∆o do rel¢gio, trunca in°cio no rel¢gio ***/
            if  v-val-refer-relogio > rep-parada-ctrab.val-refer-inic-parada and
                v-val-refer-relogio < rep-parada-ctrab.val-refer-fim-parada then do:
                if  v-val-refer-fim-parada > rep-parada-ctrab.val-refer-fim-parada then do:
                    assign v-val-tempo-calcul = ?
                           v-dat-inic      = v-dat-relogio
                           v-qtd-segs-inic = v-qtd-segs-relogio
                           v-dat-fim       = rep-parada-ctrab.dat-fim-parada 
                           v-qtd-segs-fim  = rep-parada-ctrab.qtd-segs-fim.

                    assign c-cod-calend = ?
                           c-gm-codigo  = ctrab.gm-codigo.
                    run CalcularTemposCtrab in h-boin469b (input-output p-cod-area,
                                                           input-output c-gm-codigo,
                                                           input        ctrab.cod-ctrab,
                                                           input-output c-cod-calend,
                                                           Input-output v-dat-fim,
                                                           Input-output v-qtd-segs-fim,
                                                           Input-output v-dat-relogio,
                                                           Input-output v-qtd-segs-relogio,
                                                           Input-output v-val-tempo-calcul).
                                         
                    if v-val-tempo-calcul = 0 then
                        assign v-val-tempo-calcul-tex = (v-dat-fim        - v-dat-relogio     ) * 24 +
                                                        (v-qtd-segs-fim   - v-qtd-segs-relogio) / 3600.
                end.
                else do:
                    assign v-val-tempo-calcul = ?
                           v-dat-inic      = v-dat-relogio
                           v-qtd-segs-inic = v-qtd-segs-relogio
                           v-dat-fim       = p-dat-fim-period 
                           v-qtd-segs-fim  = p-qtd-segs-fim.
            
                    assign c-cod-calend = ?
                           c-gm-codigo  = ctrab.gm-codigo.
                    run CalcularTemposCtrab in h-boin469b (input-output p-cod-area,
                                                           input-output c-gm-codigo,
                                                           input        ctrab.cod-ctrab,
                                                           input-output c-cod-calend,
                                                           Input-output v-dat-fim,
                                                           Input-output v-qtd-segs-fim,
                                                           Input-output v-dat-relogio,
                                                           Input-output v-qtd-segs-relogio,
                                                           Input-output v-val-tempo-calcul).

                  
                end.
            end.
           /*
             message "/*** A parada inicia dentro do intervalo de avaliaÁ„o ***/" skip
             "v-val-refer-relogio :" v-val-refer-relogio skip
             "val-refer-inic-parada :" rep-parada-ctrab.val-refer-inic-parada skip
             "val-refer-fim-parada  :" rep-parada-ctrab.val-refer-fim-parada 
             
             view-as alert-box.
             */
             
             
            /*** A parada inicia dentro do intervalo de avaliaÁ„o ***/
            if  v-val-refer-relogio <= rep-parada-ctrab.val-refer-inic-parada and
                v-val-refer-relogio  < rep-parada-ctrab.val-refer-fim-parada then do:

                /*** A parada inicia e finaliza dentro do intervalo de avalia?∆o ***/

                if  v-val-refer-fim-parada > rep-parada-ctrab.val-refer-fim-parada then do:
                    assign v-val-tempo-calcul     = rep-parada-ctrab.qtd-tempo-parada
                           v-val-tempo-calcul-tex = rep-parada-ctrab.qtd-tempo-ext
                           v-dat-inic             = rep-parada-ctrab.dat-inic-parada
                           v-qtd-segs-inic        = rep-parada-ctrab.qtd-segs-inic
                           v-dat-fim              = rep-parada-ctrab.dat-fim-parada 
                           v-qtd-segs-fim         = rep-parada-ctrab.qtd-segs-fim.
                end.       
                else do:
                    assign v-val-tempo-calcul = ?
                           v-dat-inic         = rep-parada-ctrab.dat-inic-parada
                           v-qtd-segs-inic    = rep-parada-ctrab.qtd-segs-inic
                           v-dat-fim          = p-dat-fim-period 
                           v-qtd-segs-fim     = p-qtd-segs-fim.

                    assign c-cod-calend = ?
                           c-gm-codigo  = ctrab.gm-codigo.
                    run CalcularTemposCtrab in h-boin469b (input-output p-cod-area,
                                                           input-output c-gm-codigo,
                                                           input        ctrab.cod-ctrab,
                                                           input-output c-cod-calend,
                                                           Input-output v-dat-fim,
                                                           Input-output v-qtd-segs-fim,
                                                           Input-output v-dat-inic,
                                                           Input-output v-qtd-segs-inic,
                                                           Input-output v-val-tempo-calcul).

                    if v-val-tempo-calcul = 0 then
                        assign v-val-tempo-calcul-tex = (p-dat-fim-period - v-dat-inic     ) * 24 +
                                                        (p-qtd-segs-fim   - v-qtd-segs-inic) / 3600.
                end.
            end.
           
           /*** Soma tempo das paradas que alteram efici?ncia ***/
            if motiv-parada.log-alter-eficien = yes then do:
                assign v-qtd-parada       = v-qtd-parada  + v-val-tempo-calcul.
/*                 message "log-alter-eficien: " log-alter-eficien skip     */
/*                         "v-qtd-parada: " v-qtd-parada view-as alert-box. */
            END.
           
           

           /*** Soma tempo das paradas que s∆o setup ***/
            if motiv-parada.parada-setup = yes  and  motiv-parada.log-alter-eficien = no  then do:
                assign v-qtd-parada   = v-qtd-parada  + v-val-tempo-calcul.
                message "log-alter-eficien: " log-alter-eficien skip 
                        "v-qtd-parada: " v-qtd-parada view-as alert-box.
            END.


             FIND FIRST ttParadas
                  WHERE ttParadas.cod-ctrab   = ctrab.cod-ctrab
                    AND ttParadas.data-inicio = rep-parada-ctrab.dat-inic-parada
                    AND ttParadas.cod-parada  = rep-parada-ctrab.cod-parada
                     NO-ERROR.
            IF NOT AVAIL  ttParadas THEN DO:
                CREATE  ttParadas.
                ASSIGN ttParadas.cod-ctrab   = ctrab.cod-ctrab
                       ttParadas.data-inicio = rep-parada-ctrab.dat-inic-parada
                       ttParadas.cod-parada  = rep-parada-ctrab.cod-parada
                       ttParadas.motivo      = motiv-parada.des-parada.
            END.
            ASSIGN ttParadas.tempo = ttParadas.tempo + v-qtd-parada.
            
            FIND FIRST ttListParadas
                 WHERE ttListParadas.cod-ctrab   = ctrab.cod-ctrab
                   AND ttListParadas.data-inicio = rep-parada-ctrab.dat-inic-parada
                   AND ttListParadas.cod-parada  = rep-parada-ctrab.cod-parada
                   AND ttListParadas.it-codigo   = split-operac.it-codigo
                   NO-ERROR.

            IF NOT AVAIL ttListParadas THEN DO:
                CREATE ttListParadas.
                ASSIGN 
                ttListParadas.cod-ctrab   = ctrab.cod-ctrab
                ttListParadas.data-inicio = rep-parada-ctrab.dat-inic-parada
                ttListParadas.cod-parada  = rep-parada-ctrab.cod-parada
                ttListParadas.it-codigo   = split-operac.it-codigo
                ttListParadas.gm-codigo   = ctrab.gm-codigo
                ttListParadas.motivo      = motiv-parada.des-parada
                ttListParadas.desc-oper   = oper-ord.descricao.
           END.
           ASSIGN ttListParadas.deparada = ttListParadas.deparada + v-qtd-parada
                  ttListParadas.qtd      = ttListParadas.qtd + 1.
            
                    
            /*** Atualiza vari‡vel rel¢gio ***/
            assign v-dat-relogio      = rep-parada-ctrab.dat-fim-parada
                   v-qtd-segs-relogio = rep-parada-ctrab.qtd-segs-fim.

            run pi-converte-data-segs-valor (Input rep-parada-ctrab.dat-fim-parada,
                                             Input rep-parada-ctrab.qtd-segs-fim,
                                             output v-val-refer-relogio).
            
            
          
        end. /*** blk-reporte-parada ***/
    end.

   
  
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-createTTDados) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE createTTDados Procedure 
PROCEDURE createTTDados :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
    DEFINE VARIABLE De-tempo-ini LIKE rep-oper-ctrab.val-refer-inic-rep NO-UNDO.
    DEFINE VARIABLE de-tempo-fim LIKE rep-oper-ctrab.val-refer-fim-rep NO-UNDO.
    
   ASSIGN de-tempo-ini = Dec(string(YEAR(tt-param.dt_dt_ini),"9999") + string(MONTH(tt-param.dt_dt_ini),"99") + string(DAY(tt-param.dt_dt_ini),"99") + "00000") 
          de-tempo-fim = Dec(string(YEAR(tt-param.dt_dt_fim),"9999") + string(MONTH(tt-param.dt_dt_fim),"99") + string(DAY(tt-param.dt_dt_fim),"99") + "99999"). 
          
        /*  c_estab_ini    
          c_estab_fim    
          
       */   
       EMPTY TEMP-TABLE tt-ctrab-rp.
       EMPTY TEMP-TABLE ttDados.
       EMPTY TEMP-TABLE ttCentro.
       EMPTY TEMP-TABLE ttItem.
       EMPTY TEMP-TABLE ttTotalCentro.
       EMPTY TEMP-TABLE ttDias.
       EMPTY TEMP-TABLE tt-operac-ctrab-rp_aux.
       EMPTY TEMP-TABLE ttParadas.
       EMPTY TEMP-TABLE ttListParadas.

       EMPTY TEMP-TABLE ttTurno.
       
       RUN criaTurno.

       run sfc/sfapi001.p  persistent set h-sfapi001.
       run iniciarBOs in h-sfapi001.
       run inbo/boin469b.p persistent set h-boin469b.
       /*run sfc/sfapi002.p persistent set h-sfapi002.*/
       run inbo/boin535a.p persistent set h-boin535a.
      /* RUN recebeReferencia IN h-sfapi002 (INPUT "*").*/
      
          
        IF  NOT CAN-FIND(FIRST tt-digita) THEN DO:   
          
           FOR EACH ctrab
               WHERE ctrab.gm-codigo >= tt-param.c_gm-codigo_ini
                 AND ctrab.gm-codigo <= tt-param.c_gm-codigo_fim
                 AND ctrab.cod-ctrab >= tt-param.c_cod-ctrab_ini
                 AND ctrab.cod-ctrab <= tt-param.c_cod-ctrab_fim NO-LOCK
                  BREAK BY ctrab.gm-codigo :                      
                 
                FIND FIRST area-produc-ctrab
                            WHERE area-produc-ctrab.cod-ctrab       =  ctrab.cod-ctrab
                              and area-produc-ctrab.dat-inic-valid  <= TODAY
                              and area-produc-ctrab.dat-fim-valid   >  TODAY NO-LOCK NO-ERROR.
                IF NOT AVAIL  area-produc-ctrab THEN NEXT.
                IF VALID-HANDLE(h-acomp) THEN               
                      RUN pi-acompanhar in h-acomp (input "Ctrab." + ctrab.cod-ctrab + " / ":U + ctrab.gm-codigo).
                find grup-maquina where
                       grup-maquina.gm-codigo = ctrab.gm-codigo no-lock no-error.


                RUN criaTTCentro (INPUT  ctrab.cod-ctrab).


                FOR EACH rep-oper-ctrab EXCLUSIVE-LOCK WHERE 
                           rep-oper-ctrab.cod-ctrab          =  ctrab.cod-ctrab AND
                           rep-oper-ctrab.val-refer-inic-rep >= de-tempo-ini and 
                           rep-oper-ctrab.val-refer-inic-rep <= de-tempo-fim  USE-INDEX ctrab-i: 
                           IF rep-oper-ctrab.val-refer-fim-rep > de-tempo-fim  THEN NEXT.
                           
                             find first tt-ctrab-rp
                                    where tt-ctrab-rp.cod-area         = area-produc-ctrab.cod-area-produ
                                     and  tt-ctrab-rp.gm-codigo        = ctrab.gm-codigo
                                     and  tt-ctrab-rp.cod-ctrab        = ctrab.cod-ctrab no-error.
                               if not avail tt-ctrab-rp then do:
                                  
                                     create tt-ctrab-rp.
                                     assign tt-ctrab-rp.cod-area         = area-produc-ctrab.cod-area-produ
                                            tt-ctrab-rp.gm-codigo        = ctrab.gm-codigo
                                            tt-ctrab-rp.cod-ctrab        = ctrab.cod-ctrab
                                            tt-ctrab-rp.cod-ctrab        = ctrab.cod-ctrab
                                            tt-ctrab-rp.descricao        = ctrab.des-ctrab
                                            tt-ctrab-rp.data-inicio      = max(tt-param.dt_dt_ini, area-produc-ctrab.dat-inic-valid)
                                            tt-ctrab-rp.data-fim         = min(tt-param.dt_dt_fim,area-produc-ctrab.dat-fim-valid)
                                            tt-ctrab-rp.qtd-capac-ctrab  = ctrab.qtd-capac-ctrab.
                      end.

                END. /*** for each ctrab****/
                   
            END. /*** for each tt-digita****/
       END.
       ELSE DO:
          FOR EACH tt-digita BREAK BY tt-digita.grupo 
                                   BY tt-digita.centro:
              FOR EACH ctrab
                   WHERE ctrab.gm-codigo = tt-digita.grupo
                     AND ctrab.cod-ctrab = tt-digita.centro NO-LOCK
                      BREAK BY ctrab.gm-codigo :                      
                     
                  FIND FIRST area-produc-ctrab
                            WHERE area-produc-ctrab.cod-ctrab       =  ctrab.cod-ctrab
                              and area-produc-ctrab.dat-inic-valid  <= TODAY
                              and area-produc-ctrab.dat-fim-valid   >  TODAY NO-LOCK NO-ERROR.
                IF NOT AVAIL  area-produc-ctrab THEN NEXT.
                IF VALID-HANDLE(h-acomp) THEN               
                    RUN pi-acompanhar in h-acomp (input "Ctrab." + ctrab.cod-ctrab + " / ":U + ctrab.gm-codigo).
                find FIRST grup-maquina where
                     grup-maquina.gm-codigo = ctrab.gm-codigo no-lock no-error.

                RUN criaTTCentro (INPUT  ctrab.cod-ctrab).
                FOR EACH rep-oper-ctrab EXCLUSIVE-LOCK WHERE 
                               rep-oper-ctrab.cod-ctrab          =  ctrab.cod-ctrab AND
                               rep-oper-ctrab.val-refer-inic-rep >= de-tempo-ini and 
                               rep-oper-ctrab.val-refer-inic-rep <= de-tempo-fim  USE-INDEX ctrab-i: 
                        IF rep-oper-ctrab.val-refer-fim-rep > de-tempo-fim  THEN NEXT.
                        
                        find first tt-ctrab-rp
                           where tt-ctrab-rp.cod-area         = area-produc-ctrab.cod-area-produ
                            and  tt-ctrab-rp.gm-codigo        = ctrab.gm-codigo
                            and  tt-ctrab-rp.cod-ctrab        = ctrab.cod-ctrab no-error.
                      if not avail tt-ctrab-rp then do:
                         
                            create tt-ctrab-rp.
                            assign tt-ctrab-rp.cod-area         = area-produc-ctrab.cod-area-produ
                                   tt-ctrab-rp.gm-codigo        = ctrab.gm-codigo
                                   tt-ctrab-rp.cod-ctrab        = ctrab.cod-ctrab
                                   tt-ctrab-rp.cod-ctrab        = ctrab.cod-ctrab
                                   tt-ctrab-rp.descricao        = ctrab.des-ctrab
                                   tt-ctrab-rp.data-inicio      = max(tt-param.dt_dt_ini, area-produc-ctrab.dat-inic-valid)
                                   tt-ctrab-rp.data-fim         = min(tt-param.dt_dt_fim,area-produc-ctrab.dat-fim-valid)
                                   tt-ctrab-rp.qtd-capac-ctrab  = ctrab.qtd-capac-ctrab.
                      end.
                           
               END.
                       
             END. /*** for each ctrab****/
          END. /*** for each tt-digita****/
       END.
       run pi-converte-data-segs-valor (input tt-param.dt_dt_ini,
                                        input 0,
                                        output v-val-refer-inic).

       run pi-converte-data-segs-valor (input tt-param.dt_dt_fim,
                                        input 86400,
                                        output v-val-refer-fim).
       FOR EACH tt-ctrab-rp use-index ch-codigo :

           IF VALID-HANDLE(h-acomp) THEN               
              RUN pi-acompanhar in h-acomp (input "Cal. Tempos:" + tt-ctrab-rp.cod-ctrab + " / ":U + tt-ctrab-rp.gm-codigo).
              

            if  tt-ctrab-rp.qtd-segs-inic = ? 
                or  tt-ctrab-rp.qtd-segs-fim  = ? then 
            run inicializarIntervaloTurno in h-boin469b (input tt-ctrab-rp.cod-area,
                                                         input tt-ctrab-rp.gm-codigo,
                                                         input tt-ctrab-rp.cod-ctrab,
                                                         input tt-ctrab-rp.data-inicio,
                                                         input tt-ctrab-rp.data-fim,
                                                         output tt-ctrab-rp.data-inicio,
                                                         output tt-ctrab-rp.qtd-segs-inic,
                                                         output tt-ctrab-rp.data-fim,
                                                         output tt-ctrab-rp.qtd-segs-fim).
                                                         
             assign de-tempo-dispon = ?
                    c-cod-calend    = ?.

        /*** Tempo dispon°vel para o centro de trabalho ***/
        run CalcularTemposCtrab in h-boin469b (input-output tt-ctrab-rp.cod-area,
                                               input-output tt-ctrab-rp.gm-codigo,
                                               input        tt-ctrab-rp.cod-ctrab,
                                               input-output c-cod-calend,
                                               Input-output tt-ctrab-rp.data-fim,
                                               Input-output tt-ctrab-rp.qtd-segs-fim,
                                               Input-output tt-ctrab-rp.data-inicio,
                                               Input-output tt-ctrab-rp.qtd-segs-inic,
                                               Input-output de-tempo-dispon).
        
        /*** O tempo de parada de-qtd-parada ? o tempo de parada em tempo normal como ? o default 
             de c‡lculo da pi-ctrab-tempo-paradas  ***/
        
             
        assign tt-ctrab-rp.tempo-dispon       = de-tempo-dispon.                                             
      END.

      for each tt-ctrab-rp use-index ch-codigo :
      
      
        find ctrab  where 
              ctrab.cod-ctrab = tt-ctrab-rp.cod-ctrab no-lock no-error.
      
           find FIRST b-gm where 
                 b-gm.gm-codigo = tt-ctrab-rp.gm-codigo no-lock no-error.
          
          IF VALID-HANDLE(h-acomp) THEN  
             run pi-acompanhar in h-acomp (INPUT "Gerando dados" + tt-ctrab-rp.cod-ctrab).
             
           

           
                                             
                                             
                                           
         ASSIGN de-setup-tempo = 0.   
        for each  rep-oper-ctrab USE-INDEX ctrab-i
            where rep-oper-ctrab.cod-ctrab           = tt-ctrab-rp.cod-ctrab
              and rep-oper-ctrab.val-refer-inic-rep >= v-val-refer-inic
              and rep-oper-ctrab.val-refer-inic-rep <  v-val-refer-fim  no-lock:

            IF VALID-HANDLE(h-acomp) THEN  
               run pi-acompanhar in h-acomp (INPUT "Lendos ap produ?∆o" + STRING(rep-oper-ctrab.nr-ord-produ)).
            
            /*** Pendencia: Truncar reportes que iniciam fora do intervalo e
                 terminam fora do intervalo, ou reportes que iniciam no intervalo
                 e terminam fora do intervalo ***/
             ASSIGN v-carga-teorica = 0
                    de-qtd-parada   = 0.

             /*** Tempo de paradas para o centro de trabalho ***/
        run pi-ctrab-tempo-paradas in h-sfapi001 (input 3,
                                                  input rep-oper-ctrab.cod-ctrab,
                                                  input tt-ctrab-rp.cod-area,
                                                  input rep-oper-ctrab.dat-inic-reporte,
                                                  input rep-oper-ctrab.qtd-segs-inic,
                                                  input rep-oper-ctrab.dat-fim-reporte,
                                                  input rep-oper-ctrab.qtd-segs-fim,
                                                  input no,
                                                  input no,
                                                  output de-qtd-parada,
                                                  output de-qtd-parada-nao,
                                                  output de-qtd-parada-set,
                                                  output de-qtd-par-tex,
                                                  output de-qtd-par-nao-tex,
                                                  output de-qtd-par-set-tex).
              
            find first split-operac USE-INDEX id
                 where split-operac.nr-ord-produ     = rep-oper-ctrab.nr-ord-produ
                   and split-operac.num-operac-sfc   = rep-oper-ctrab.num-operac-sfc
                   and split-operac.num-split-operac = rep-oper-ctrab.num-split-oper 
                   AND split-operac.it-codigo        >= tt-param.c_item_ini
                   AND split-operac.it-codigo        <= tt-param.c_item_fim no-lock no-error.
                   find first oper-ord
                 where oper-ord.nr-ord-produ = split-operac.nr-ord-produ
                   and oper-ord.it-codigo    = split-operac.it-codigo
                   and oper-ord.cod-roteiro  = split-operac.cod-roteiro
                   and oper-ord.op-codigo    = split-operac.op-codigo no-lock no-error.
                   
              IF NOT AVAIL split-operac THEN NEXT.
              
             IF de-qtd-parada <> 0 THEN DO:
             
                RUN buscaParadas(  input 3,
                                   input rep-oper-ctrab.cod-ctrab,
                                   input tt-ctrab-rp.cod-area,
                                   input rep-oper-ctrab.dat-inic-reporte,
                                   input rep-oper-ctrab.qtd-segs-inic,
                                   input rep-oper-ctrab.dat-fim-reporte,
                                   input rep-oper-ctrab.qtd-segs-fim,
                                   input no,
                                   input NO).
             
             
             END.

            /*** Falta verificar se grupo de m‡quina bate com grupo de m‡quina
                 da oper-ord, caso contrario a procedure a executar ser‡ outra ***/
            
            assign v-tempo-padrao     = ?.

            if avail oper-ord and oper-ord.num-id-operacao <> 0 then do:
                find operacao where 
                     operacao.num-id-operacao = oper-ord.num-id-operacao no-lock no-error.
                if  avail operacao and operacao.ind-tempo-operac = 0 then do:
                        find grup-maquina where 
                             grup-maquina.gm-codigo = operacao.gm-codigo no-lock no-error.
                        find current operacao exclusive-lock.
                        assign operacao.qtd-carga-batch  = grup-maquina.qtd-carga-batch 
                               operacao.ind-tempo-operac = grup-maquina.ind-tempo-operac.
                        
                    
                    /*** Grupo de m‡quina do reporte difere do grupo de m‡quina da opera?∆o padr∆o ou 
                         opera?∆o da ordem ? opera?∆o alternativa ***/
                    if tt-ctrab-rp.gm-codigo <> operacao.gm-codigo or 
                       oper-ord.op-altern    <> 0 then do:
                        find first op-altern where
                                   op-altern.num-id-operacao = operacao.num-id-operacao and
                                   op-altern.gm-codigo       = tt-ctrab-rp.gm-codigo no-lock no-error.
                        if available op-altern then do:
                            if op-altern.ind-tempo-operac = 0 then do:
                                find grup-maquina where 
                                     grup-maquina.gm-codigo = op-altern.gm-codigo no-lock no-error.
                                find current op-altern exclusive-lock.
                                assign op-altern.qtd-carga-batch  = grup-maquina.qtd-carga-batch 
                                       op-altern.ind-tempo-operac = grup-maquina.ind-tempo-operac.
    
                            end.
                            run CalcularTempoPadraoA in h-boin535a (buffer op-altern,
                                                                    input tt-ctrab-rp.cod-ctrab,
                                                                    input split-operac.cod-ferr-prod, 
                                                                    input (rep-oper-ctrab.qtd-operac-aprov + rep-oper-ctrab.qtd-operac-refgda),
                                                                    output v-tempo-padrao,
                                                                    output v-tempo-homem,
                                                                    output v-capac).
                        end.
                    end. /* avail oper-ord */
                    /*** C‡lculo dos padr‰es em fun?∆o do tempo padr∆o da opera?∆o engenharia ***/
                    if v-tempo-padrao = ? then
                        run CalcularTempoPadraoE in h-boin535a (buffer operacao,
                                                                input tt-ctrab-rp.cod-ctrab,
                                                                input split-operac.cod-ferr-prod, 
                                                                input (rep-oper-ctrab.qtd-operac-aprov + rep-oper-ctrab.qtd-operac-refgda),
                                                                output v-tempo-padrao,
                                                                output v-tempo-homem,
                                                                output v-capac).                                                               
                
                    end . /* avail operacao */           
                END.
                /*** Caso de tempo padr∆o ainda n∆o houver sido calculado ***/
                if v-tempo-padrao = ? then do:
                
                   if oper-ord.ind-tempo-operac = 0 then do:
                          find grup-maquina where 
                                grup-maquina.gm-codigo = oper-ord.gm-codigo no-lock no-error.
                           find current oper-ord exclusive-lock.
                           assign oper-ord.qtd-carga-batch  = grup-maquina.qtd-carga-batch 
                                  oper-ord.ind-tempo-operac = grup-maquina.ind-tempo-operac.
                   end.
    
                       run CalcularTempoPadraoO in h-boin535a (buffer oper-ord,
                                                               input tt-ctrab-rp.cod-ctrab,
                                                               input split-operac.cod-ferr-prod, 
                                                               input (rep-oper-ctrab.qtd-operac-aprov + rep-oper-ctrab.qtd-operac-refgda),
                                                               output v-tempo-padrao,
                                                               output v-tempo-homem,
                                                               output v-capac).
                                                     
                                                               
                          
                   assign v-tempo-padrao = v-tempo-padrao / 3600
                          v-tempo-homem  = v-tempo-homem  / 3600.
                   
                   
                end.
            
                /*** Quando a unidade de capacidade ? branco a capacidade da m‡quina ? a quantidade reportada ***/ 
                           
                if b-gm.cod-unid-capac = "" then 
                   assign v-capac = rep-oper-ctrab.qtd-operac-aprov + rep-oper-ctrab.qtd-operac-refgda. 
                   
                assign   v-carga-teorica                      =    if rep-oper-ctrab.qtd-tempo-reporte > 0 then 
                                                                   (v-capac * v-tempo-padrao /
                                                                                rep-oper-ctrab.qtd-tempo-reporte)
                                                              else 0.


              ASSIGN de-tempo-turno = 0
                     v-dia-semana   = weekday(rep-oper-ctrab.dat-inic-reporte).
               FIND FIRST ttTurno
                  WHERE ttTurno.cod-model-turno = 'Turno 1'
                   AND  ttTurno.dia       = v-dia-semana NO-ERROR.
                   
               IF AVAIL ttTurno THEN DO:
               
              
               CASE ttTurno.dia:
                  WHEN 1 THEN ASSIGN de-tempo-turno = ttTurno.qtd-tempo-util-dia-1.   
                  WHEN 2 THEN ASSIGN de-tempo-turno = ttTurno.qtd-tempo-util-dia-2.   
                  WHEN 3 THEN ASSIGN de-tempo-turno = ttTurno.qtd-tempo-util-dia-3.   
                  WHEN 4 THEN ASSIGN de-tempo-turno = ttTurno.qtd-tempo-util-dia-4.   
                  WHEN 5 THEN ASSIGN de-tempo-turno = ttTurno.qtd-tempo-util-dia-5.   
                  WHEN 6 THEN ASSIGN de-tempo-turno = ttTurno.qtd-tempo-util-dia-6.   
                  WHEN 7 THEN ASSIGN de-tempo-turno = ttTurno.qtd-tempo-util-dia-7. 
                END case.
                
                 
               END.
               
               /**** Calculo capacidades teorica e real*******/
               
               ASSIGN v-teorica = 0
                      v-real    = 0
                      v-teorica = rep-oper-ctrab.qtd-tempo-reporte * (( 60 / operacao.tempo-maquin) * 100)

                      v-real    = rep-oper-ctrab.qtd-operac-refgda + rep-oper-ctrab.qtd-operac-aprov. //rep-oper-ctrab.qtd-tempo-reporte * (( 60 / operacao.tempo-maquin) * 100).

              // asas 191213

             assign data-ini  = rep-oper-ctrab.dat-inic-reporte 
                    data-fim  = rep-oper-ctrab.dat-fim-reporte
                    v-ctrab   = tt-ctrab-rp.cod-ctrab.

             IF data-ini-paradas = ? THEN
                assign data-ini-paradas  = rep-oper-ctrab.dat-inic-reporte. 

                       
             ASSIGN data-fim-paradas = rep-oper-ctrab.dat-fim-reporte.
           

              
            run PI_CALC_PARADAS.
            
            //asas renan

              FIND FIRST ttDadosComp
                   WHERE ttDadosComp.cod-ctrab   = tt-ctrab-rp.cod-ctrab
                     AND ttDadosComp.data-inicio = rep-oper-ctrab.dat-inic-reporte 
                     AND ttDadosComp.it-codigo   = split-operac.it-codigo NO-ERROR.
                     

              IF NOT AVAIL ttDadosComp THEN DO:
                            
                      CREATE ttDadosComp.
                      ASSIGN
                      ttDadosComp.cod-ctrab         = tt-ctrab-rp.cod-ctrab
                      ttDadosComp.data-inicio       = rep-oper-ctrab.dat-inic-reporte
                      ttDadosComp.it-codigo         = split-operac.it-codigo.
              END.

              FIND FIRST ttDados
                   WHERE ttDados.cod-ctrab   = tt-ctrab-rp.cod-ctrab
                     AND ttDados.data-inicio = rep-oper-ctrab.dat-inic-reporte NO-ERROR.
                     

              IF NOT AVAIL ttDados THEN DO:
                            
                      CREATE ttDados.
                      ASSIGN
                      ttDados.cod-ctrab         = ttDadosComp.cod-ctrab //tt-ctrab-rp.cod-ctrab
                      ttDados.data-inicio       = ttDadosComp.data-inicio //rep-oper-ctrab.dat-inic-reporte
                      ttDados.tempo-disponivel  = (de-tempo-turno - de_parada_n_alt_efic)
                      ttDados.tempo-utilizado   = rep-oper-ctrab.qtd-tempo-reporte
                      ttDados.carga-disponivel  = (de-tempo-turno - de_parada_n_alt_efic) * (( 60 / operacao.tempo-maquin) * 100) //v-teorica
                      ttDados.carga-utilizada   = v-real 
                      ttDados.qtd-aprov         = rep-oper-ctrab.qtd-operac-aprov
                      ttDados.qtd-refgada       = rep-oper-ctrab.qtd-operac-refgda + rep-oper-ctrab.qtd-operac-aprov
                      ttDados.fator-utiliz      = ttDados.tempo-utilizado / ttDados.tempo-disponivel   
                      ttDados.val-eficien-ctrab = ctrab.val-eficien-ctrab
                      ttDados.oee               = ttDados.fator-utiliz * ttDados.eficiencia-media * ttDados.de-classif NO-ERROR.

                      FIND FIRST CST_Item
                           WHERE CST_Item.it-codigo = ttDadosComp.it-codigo NO-ERROR.
        
                           IF AVAIL CST_Item THEN DO:
                              IF CST_Item.LOG_Composto = TRUE THEN DO:
                                    ASSIGN ttDados.tempo-utilizado = (rep-oper-ctrab.qtd-tempo-reporte / 2).
                              END.
                           END.
                              
                                  

              END.
              
              ELSE DO:

                   FIND FIRST CST_Item
                       WHERE CST_Item.it-codigo = ttDadosComp.it-codigo NO-ERROR.
        
                       IF AVAIL CST_Item THEN DO:
                          IF CST_Item.LOG_Composto = TRUE THEN DO:
                              
                              ASSIGN
                              ttDados.carga-disponivel        =  (de-tempo-turno - de_parada_n_alt_efic) * (( 60 / operacao.tempo-maquin) * 100) //ttDados.carga-disponivel + v-teorica
                              ttDados.carga-utilizada         =  ttDados.carga-utilizada  +  v-real
                              ttDados.tempo-utilizado         =  ttDados.tempo-utilizado + (rep-oper-ctrab.qtd-tempo-reporte / 2)
                              ttDados.qtd-aprov               =  ttDados.qtd-aprov        + rep-oper-ctrab.qtd-operac-aprov
                              ttDados.qtd-refgada             =  ttDados.qtd-refgada      + rep-oper-ctrab.qtd-operac-refgda + rep-oper-ctrab.qtd-operac-aprov
                              ttDados.fator-utiliz            =  ttDados.tempo-utilizado / ttDados.tempo-disponivel
                              ttDados.oee                     =  ttDados.fator-utiliz * ttDados.eficiencia-media * ttDados.de-classif.

                          END.
                          ELSE DO:
                              
                              ASSIGN
                              ttDados.carga-disponivel        =  (de-tempo-turno - de_parada_n_alt_efic) * (( 60 / operacao.tempo-maquin) * 100) //ttDados.carga-disponivel + v-teorica
                              ttDados.carga-utilizada         =  ttDados.carga-utilizada  +  v-real
                              ttDados.tempo-utilizado         =  ttDados.tempo-utilizado + rep-oper-ctrab.qtd-tempo-reporte
                              ttDados.qtd-aprov               =  ttDados.qtd-aprov        + rep-oper-ctrab.qtd-operac-aprov
                              ttDados.qtd-refgada             =  ttDados.qtd-refgada      + rep-oper-ctrab.qtd-operac-refgda + rep-oper-ctrab.qtd-operac-aprov
                              ttDados.fator-utiliz            =  ttDados.tempo-utilizado / ttDados.tempo-disponivel
                              ttDados.oee                     =  ttDados.fator-utiliz * ttDados.eficiencia-media * ttDados.de-classif.
                          END.

                       END.

                       ELSE DO:
                           
                           ASSIGN
                           ttDados.carga-disponivel        =  (de-tempo-turno - de_parada_n_alt_efic) * (( 60 / operacao.tempo-maquin) * 100) //ttDados.carga-disponivel + v-teorica
                           ttDados.carga-utilizada         =  ttDados.carga-utilizada  +  v-real
                           ttDados.tempo-utilizado         =  ttDados.tempo-utilizado + rep-oper-ctrab.qtd-tempo-reporte
                           ttDados.qtd-aprov               =  ttDados.qtd-aprov        + rep-oper-ctrab.qtd-operac-aprov
                           ttDados.qtd-refgada             =  ttDados.qtd-refgada      + rep-oper-ctrab.qtd-operac-refgda + rep-oper-ctrab.qtd-operac-aprov
                           ttDados.fator-utiliz            =  ttDados.tempo-utilizado / ttDados.tempo-disponivel
                           ttDados.oee                     =  ttDados.fator-utiliz * ttDados.eficiencia-media * ttDados.de-classif.
                       END.

                          
              END.

              

    /*
             message "FASE 1:" 
                     rep-oper-ctrab.dat-inic-reporte   skip
                     rep-oper-ctrab.qtd-tempo-reporte   skip
                     ttDados.cod-ctrab    skip
                     ttDados.tempo-utilizado  skip
                     
                      de_parada_geral   skip      
           de_parada_total_geral   skip
           de_parada_alt_efic      skip 
           de_parada_n_alt_efic   skip
                     
                     
                     view-as alert-box.  
                     */

            //ESTAVA AQUI
              FIND FIRST ttItem USE-INDEX idx
                WHERE ttItem.cod-ctrab  = tt-ctrab-rp.cod-ctrab
                  AND  ttItem.it-codigo = split-operac.it-codigo NO-ERROR.
               IF NOT AVAIL ttItem  THEN DO:
                   CREATE ttItem.
                   ASSIGN 
                   ttItem.cod-ctrab = tt-ctrab-rp.cod-ctrab
                   ttItem.it-codigo = split-operac.it-codigo.
               END.


            FIND FIRST ttDias USE-INDEX idx
                 WHERE  ttDias.cod-ctrab = rep-oper-ctrab.cod-ctrab
                  AND   ttDias.dia       = rep-oper-ctrab.dat-inic-reporte no-error.
            IF NOT AVAIL ttDias THEN DO:
                CREATE ttDias.
                ASSIGN 
                ttDias.cod-ctrab = rep-oper-ctrab.cod-ctrab        
                ttDias.dia       = rep-oper-ctrab.dat-inic-reporte .

            END.


             run pi-sec-to-formatted-time (input rep-oper-ctrab.qtd-segs-inic ,
                                           output c-hora-ini).
              ASSIGN cHora = SUBSTRING(c-hora-ini,1,2).

           
              IF VALID-HANDLE(h-acomp) THEN  
                 run pi-acompanhar in h-acomp (INPUT "Ctrab/Hora" + tt-ctrab-rp.cod-ctrab + ' - ' + cHora).
                   FIND FIRST ttTotalCentro USE-INDEX ch-codigo
                         WHERE ttTotalCentro.cod-ctrab = tt-ctrab-rp.cod-ctrab
                           AND ttTotalCentro.c-horas   = cHora NO-ERROR.
                  IF NOT AVAIL ttTotalCentro THEN DO:
                      CREATE ttTotalCentro.
                      ASSIGN  ttTotalCentro.cod-ctrab       = tt-ctrab-rp.cod-ctrab 
                              ttTotalCentro.c-horas         = cHora 
                              ttTotalCentro.tempo-utilizado = rep-oper-ctrab.qtd-tempo-reporte
                              ttTotalCentro.tempo-parada    = de-qtd-parada
                              ttTotalCentro.tempo-setup     = ( (split-operac.qtd-segs-fim-setup  -  split-operac.qtd-segs-inic-setup) / 3600)
                              ttTotalCentro.qtd-capac-ctrab = tt-ctrab-rp.qtd-capac-ctrab.
        
                  END.
                  ELSE DO:
        
                      ASSIGN  ttTotalCentro.tempo-utilizado = ttTotalCentro.tempo-utilizado +  rep-oper-ctrab.qtd-tempo-reporte
                              ttTotalCentro.tempo-parada    = ttTotalCentro.tempo-parada    + de-qtd-parada
                              ttTotalCentro.tempo-setup     = ttTotalCentro.tempo-setup + ( (split-operac.qtd-segs-fim-setup  -  split-operac.qtd-segs-inic-setup) / 3600)
                              ttTotalCentro.qtd-capac-ctrab = tt-ctrab-rp.qtd-capac-ctrab.
        
                  END.

        END.




          FIND FIRST  ttCentro USE-INDEX idx
               WHERE  ttCentro.cod-ctrab  = tt-ctrab-rp.cod-ctrab NO-ERROR.
          IF NOT AVAIL ttCentro THEN DO:
             CREATE ttCentro.
             ASSIGN
             ttCentro.cod-ctrab    = tt-ctrab-rp.cod-ctrab
             ttCentro.DESC-CENTRO  = tt-ctrab-rp.descricao
             ttCentro.gm-codigo    = tt-ctrab-rp.gm-codigo.
             ttCentro.qtd-capac-ctrab =  tt-ctrab-rp.qtd-capac-ctrab.
             find FIRST grup-maquina where 
                  grup-maquina.gm-codigo = tt-ctrab-rp.gm-codigo no-lock no-error.
             ASSIGN ttCentro.DESC-GRUPO  =  grup-maquina.descricao .
          END.
        
          
      END.


    run finalizarBOs in h-sfapi001.
    delete procedure h-sfapi001.
   /* delete procedure h-sfapi002.   */
    delete procedure h-boin469b.
    delete procedure h-boin535a.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-criaTTCentro) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE criaTTCentro Procedure 
PROCEDURE criaTTCentro :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEF INPUT PARAM p-centro like ctrab.cod-ctrab NO-UNDO.

DEFINE VARIABLE i-cont AS INTEGER     NO-UNDO.
DEFINE VARIABLE c-hora AS CHARACTER   NO-UNDO.

  DO i-cont = 1 TO 22:
      ASSIGN c-hora = string(i-cont).
      IF LENGTH( c-hora) < 2 THEN
          ASSIGN  c-hora = '0' + string(i-cont).


      FIND FIRST ttTotalCentro USE-INDEX ch-codigo
              WHERE ttTotalCentro.cod-ctrab = p-centro
                AND ttTotalCentro.c-horas   =  c-hora NO-ERROR.
       IF NOT AVAIL ttTotalCentro THEN DO:
           CREATE ttTotalCentro.
           ASSIGN  ttTotalCentro.cod-ctrab       = p-centro
                   ttTotalCentro.c-horas         =  c-hora. 
       END.

  END.

  
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-dadosGrafico) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE dadosGrafico Procedure 
PROCEDURE dadosGrafico :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
        
        //TABELA QUE ALIMENTA OS GRAFICOS OEE
    
        DEF VAR deValor as decimal format ">>>9.99"  NO-UNDO.
        

        //DADOS GRAFICO OEE
        C-RANGE   = "B" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = ttDados.data-inicio.  
        C-RANGE   = "C" +  string(i-linha-aux).
        deValor   = DEC(ttDados.fator-utiliz * ttDados.eficiencia-media * ttDados.de-classif).
        chPlanilha:Range( C-RANGE):NumberFormat = "##0,00".
        chPlanilha:Range( C-RANGE):value              = deValor * 100.

        //DADOS GRAFICO PERFORMANCE
        C-RANGE = "E" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = ttDados.data-inicio.  
        C-RANGE = "F" +  string(i-linha-aux).
        deValor  = DEC(ttDados.eficiencia-media).
        chPlanilha:Range( C-RANGE):NumberFormat = "##0,00".
        chPlanilha:Range( C-RANGE):value              =  deValor * 100.
        
        //DADOS GRAFICO DISPONIBILIDADE
        C-RANGE = "H" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = ttDados.data-inicio.  
        C-RANGE = "I" +  string(i-linha-aux).
        deValor  = DEC(ttDados.fator-utiliz).
        chPlanilha:Range( C-RANGE):NumberFormat = "##0,00".
        chPlanilha:Range( C-RANGE):value              = deValor * 100.

        //DADOS GRAFICO QUALIDADE
        C-RANGE = "K" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = ttDados.data-inicio.  
        C-RANGE = "L" +  string(i-linha-aux).
        deValor  = DEC(ttDados.de-classif).
        chPlanilha:Range( C-RANGE):NumberFormat = "##0,00".
        chPlanilha:Range( C-RANGE):value              =  deValor * 100.
      
        iRow = iRow + 1.

        
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-planoAcao) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE planoAcao Procedure 
PROCEDURE planoAcao :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
DEF VAR celula AS CHAR.
DEF VAR cont-linha AS INT INIT 78.

FOR EACH CST_plano_acao NO-LOCK
    WHERE CST_plano_acao.cod-ctrab = tt-param.c_cod-ctrab_ini
      AND CST_plano_acao.data >= tt-param.dt_dt_ini
      AND CST_plano_acao.data <= tt-param.dt_dt_fim.

    celula   = "B" +  string(cont-linha).
    chPlanilha:Range(celula):value              = CST_plano_acao.data.

    celula   = "C" +  string(cont-linha).
    chPlanilha:Range(celula):value              = CST_plano_acao.equipamento.

    celula   = "D" +  string(cont-linha).
    chPlanilha:Range(celula):value              = CST_plano_acao.eficiencia.

    celula   = "E" +  string(cont-linha).
    chPlanilha:Range(celula):value              = CST_plano_acao.causa.

    celula   = "J" +  string(cont-linha).
    chPlanilha:Range(celula):value              = CST_plano_acao.acao.

    celula   = "R" +  string(cont-linha).
    chPlanilha:Range(celula):value              = CST_plano_acao.responsavel.

    celula   = "T" +  string(cont-linha).
    chPlanilha:Range(celula):value              = CST_plano_acao.prazo.

    ASSIGN cont-linha = cont-linha + 1.
END.


  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-dadosMensal) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE dadosMensal Procedure 
PROCEDURE dadosMensal :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
        DEFINE VARIABLE MesPrint AS CHAR NO-UNDO.
        RUN pi-retorna-mes (INPUT INT(ttDadosMensal.c-mes), OUTPUT MesPrint).
        C-RANGE = "A" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = ttDadosMensal.cod-ctrab.
        RUN formataCelula(C-RANGE).
        C-RANGE = "B"  + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  MesPrint .
        RUN formataCelula(C-RANGE).
        C-RANGE = "C" + STRING(i-linha).
        chPlanilha:Range( C-RANGE):value              = ttDadosMensal.tempo-disponivel.  
        RUN formataCelula(C-RANGE).
        C-RANGE = "D" + STRING(i-linha).
        chPlanilha:Range( C-RANGE):value              = ttDadosMensal.tempo-utilizado.
        RUN formataCelula(C-RANGE).
        C-RANGE = "E" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = ttDadosMensal.carga-disponivel.
        RUN formataCelula(C-RANGE).
        C-RANGE = "F" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  ttDadosMensal.carga-utilizada.
        RUN formataCelula(C-RANGE).
        C-RANGE = "G" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = ttDadosMensal.qtd-aprov.
        RUN formataCelula(C-RANGE).
        C-RANGE = "H" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  ttDadosMensal.qtd-refgada .
        RUN formataCelula(C-RANGE).
        C-RANGE = "I" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = "%" + STRING(ttDadosMensal.fator-utiliz,">>>9.99") . 
        RUN formataCelula(C-RANGE).
        C-RANGE = "J" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = "%" + STRING(ttDadosMensal.eficiencia-media,">>>9.99") . 
        RUN formataCelula(C-RANGE).
        C-RANGE = "K" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  "%" + STRING(ttDadosMensal.de-classif,">>>9.99").
        RUN formataCelula(C-RANGE).
        C-RANGE = "L" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  "%" + STRING(ttDadosMensal.oee,">>>9.99").
        RUN formataCelula(C-RANGE).
        C-RANGE = "M" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  "%" + STRING(ttDadosMensal.val-eficien-ctrab,">>>9.99").
        RUN formataCelula(C-RANGE).
        
        
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-dadosParadaAltEfic) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE dadosParada Procedure 

//asas

//GRAFICO PARADAS QUE ALTERAM EFICIENCIA
PROCEDURE dadosParadaAltEfic :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
  DEF VAR motivo-parada-alt-efic AS CHAR.
  DEF VAR motivo-parada-n-alt-efic AS CHAR.
  DEF VAR i-linha-alt-efic AS INT.
  DEF VAR i-linha-n-alt-efic AS INT.
    
    assign de_parada_alt_efic     = 0 
           de_parada_n_alt_efic   = 0
           i-linha-alt-efic = 16
           i-linha-n-alt-efic = 16.

    FIND first ctrab
         WHERE ctrab.gm-codigo = ttCentro.gm-codigo
           AND ctrab.cod-ctrab = ttCentro.cod-ctrab NO-LOCK NO-ERROR.


    FOR EACH rep-parada-ctrab no-lock
       where  (rep-parada-ctrab.dat-inic-parada >= data-ini-paradas AND rep-parada-ctrab.dat-fim-parada <= data-fim-paradas)
         AND  rep-parada-ctrab.cod-ctrab = v-ctrab,
                  
         FIRST split-operac WHERE split-operac.cod-ctrab = rep-parada-ctrab.cod-ctrab,
         FIRST oper-ord WHERE oper-ord.it-codigo = split-operac.it-codigo,
         FIRST motiv-parada where motiv-parada.cod-parada = rep-parada-ctrab.cod-parada
         break by rep-parada-ctrab.cod-parada:
               
               
         IF motiv-parada.log-alter-eficien = YES THEN DO:
            
             if first-of(rep-parada-ctrab.cod-parada) then  
                assign motivo-parada-alt-efic = motiv-parada.des-parada.

 
             assign de_parada_alt_efic = de_parada_alt_efic + rep-parada-ctrab.qtd-tempo-parada.
                    
    
             if last-of(rep-parada-ctrab.cod-parada) then DO:
    
                    C-RANGE = "F" + STRING(i-linha-alt-efic).
                    chPlanilha:Range (C-RANGE):VALUE = motivo-parada-alt-efic. 
                    RUN formataCelula(C-RANGE). 
                    C-RANGE = "G" + STRING(i-linha-alt-efic).
                    chPlanilha:Range (C-RANGE):VALUE = de_parada_alt_efic. 
                    RUN formataCelula(C-RANGE).
            
                    ASSIGN i-linha-alt-efic  = i-linha-alt-efic + 1.
    
                    assign  de_parada_alt_efic = 0
                            motivo-parada-alt-efic = "".
    
             END.
         END.
  
    END. //  FOR EACH rep-parada-ctrab no-lock


    //GRAFICO PARADA QUE ALTERA EFICIENCIA      
    
     ASSIGN C-RANGE-GRAF = 'F15' + ":" + "G" + STRING(i-linha-alt-efic).
     chRange = chPlanilha:Range(C-RANGE-GRAF).
     chPlanilha:ChartObjects:Add(15,195,930,650):Activate.
     chExcelApplication:ActiveChart:ChartWizard(chRange,  3, 7, 2, 1, 1, TRUE, "Paradas (Altera EficiÍncia)", "", "").
     chExcelApplication:ActiveChart:ChartType = 5.
/*      chExcelApplication:ActiveChart:ChartArea:Select.                         */
/*      chExcelApplication:ActiveChart:Axes(2,1):Select.                         */
/*      chExcelApplication:ActiveChart:Axes(2,1):DisplayUnit = -2.               */
/*      chExcelApplication:ActiveChart:Axes(2,1):TickLabels:NumberFormat = "0%". */
/*      chExcelApplication:ActiveChart:Axes(2,1):HasMajorGridlines = True.       */
/*      chExcelApplication:ActiveChart:Axes(2,1):MinimumScale = 0.               */
/*      chExcelApplication:ActiveChart:Axes(2,1):MaximumScale = 100.             */
/*      chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.                  */
/*      chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 1.                  */
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:Select.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = '#0.#0"%"'.
     chExcelApplication:ActiveChart:SeriesCollection(1):Name = "Percentual Ineficiencia".

  
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-dadosParadaAltEfic) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE dadosParada Procedure 

//GRAFICO PARADAS QUE NAO ALTERAM EFICIENCIA
PROCEDURE dadosParadaNaoAltEfic :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
  DEF VAR motivo-parada-n-alt-efic AS CHAR.
  DEF VAR i-linha-n-alt-efic AS INT.
    
    assign de_parada_n_alt_efic   = 0
           i-linha-n-alt-efic = 16.

    FIND first ctrab
         WHERE ctrab.gm-codigo = ttCentro.gm-codigo
           AND ctrab.cod-ctrab = ttCentro.cod-ctrab NO-LOCK NO-ERROR.


    FOR EACH rep-parada-ctrab no-lock
       where  (rep-parada-ctrab.dat-inic-parada >= data-ini-paradas AND rep-parada-ctrab.dat-fim-parada <= data-fim-paradas)
         AND  rep-parada-ctrab.cod-ctrab = v-ctrab,
                  
         FIRST split-operac WHERE split-operac.cod-ctrab = rep-parada-ctrab.cod-ctrab,
         FIRST oper-ord WHERE oper-ord.it-codigo = split-operac.it-codigo,
         FIRST motiv-parada where motiv-parada.cod-parada = rep-parada-ctrab.cod-parada
         break by rep-parada-ctrab.cod-parada:

             IF motiv-parada.log-alter-eficien = NO THEN DO:
            
                 if first-of(rep-parada-ctrab.cod-parada) then  
                    assign motivo-parada-n-alt-efic        = motiv-parada.des-parada.
    
     
                 assign de_parada_n_alt_efic = de_parada_n_alt_efic + rep-parada-ctrab.qtd-tempo-parada.
                    
    
                 if last-of(rep-parada-ctrab.cod-parada) then DO:
        
                        C-RANGE = "H" + STRING(i-linha-n-alt-efic).
                        chPlanilha:Range (C-RANGE):VALUE = motivo-parada-n-alt-efic. 
                        RUN formataCelula(C-RANGE). 
                        C-RANGE = "I" + STRING(i-linha-n-alt-efic).
                        chPlanilha:Range (C-RANGE):VALUE = de_parada_n_alt_efic. 
                        RUN formataCelula(C-RANGE).
                
                        ASSIGN i-linha-n-alt-efic  = i-linha-n-alt-efic + 1.
        
                        assign  de_parada_n_alt_efic = 0
                                motivo-parada-n-alt-efic = "".
        
                 END.
             END.
  
    END. //  FOR EACH rep-parada-ctrab no-lock


    //GRAFICO PARADA QUE NAO ALTERA EFICIENCIA      
    
     ASSIGN C-RANGE-GRAF = 'H15' + ":" + "I" + STRING(i-linha-n-alt-efic).
     chRange = chPlanilha:Range(C-RANGE-GRAF).
     //chPlanilha:ChartObjects:Add(15,195,930,650):Activate.
     chPlanilha:ChartObjects:Add(15,195,930,650):Activate.
     chExcelApplication:ActiveChart:ChartWizard(chRange,  3, 7, 2, 1, 1, TRUE, "Paradas (N∆o Altera Eficiància)", "", "").
     chExcelApplication:ActiveChart:ChartType = 5.
/*      chExcelApplication:ActiveChart:ChartArea:Select.                         */
/*      chExcelApplication:ActiveChart:Axes(2,1):Select.                         */
/*      chExcelApplication:ActiveChart:Axes(2,1):DisplayUnit = -2.               */
/*      chExcelApplication:ActiveChart:Axes(2,1):TickLabels:NumberFormat = "0%". */
/*      chExcelApplication:ActiveChart:Axes(2,1):HasMajorGridlines = True.       */
/*      chExcelApplication:ActiveChart:Axes(2,1):MinimumScale = 0.               */
/*      chExcelApplication:ActiveChart:Axes(2,1):MaximumScale = 100.             */
/*      chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.                  */
/*      chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 1.                  */
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:Select.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = '#0.#0"%"'.
     chExcelApplication:ActiveChart:SeriesCollection(1):Name = "Percentual Ineficiencia".

  
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-dadosRelatorio) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE dadosRelatorio Procedure 
PROCEDURE dadosRelatorio :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
     
    //asas
      
        C-RANGE = "B" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = ttDados.data-inicio.
        RUN formataCelula(C-RANGE).
        C-RANGE = "C" + STRING(i-linha).
        chPlanilha:Range( C-RANGE):value              = ttDados.tempo-disponivel.  
        RUN formataCelula(C-RANGE).
        C-RANGE = "D" + STRING(i-linha).
        chPlanilha:Range( C-RANGE):value              = ttDados.tempo-utilizado.
        RUN formataCelula(C-RANGE).
        C-RANGE = "E" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = ttDados.carga-disponivel.
        RUN formataCelula(C-RANGE).
        C-RANGE = "F" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  ttDados.carga-utilizada.
        RUN formataCelula(C-RANGE).
        C-RANGE = "G" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = ttDados.qtd-aprov.
        RUN formataCelula(C-RANGE).
        C-RANGE = "H" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  ttDados.qtd-refgada .
        RUN formataCelula(C-RANGE).
        C-RANGE = "I" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  DEC(ttDados.fator-utiliz). //DISPONIBILIDADE POR APONTAMENTO
        RUN formataCelula(C-RANGE).
        C-RANGE = "J" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  DEC(ttDados.eficiencia-media). //PERFORMANCE POR APONTAMENTO
        RUN formataCelula(C-RANGE).
        C-RANGE = "K" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  DEC(ttDados.de-classif). //QUALIDADE POR APONTAMENTO
        RUN formataCelula(C-RANGE).
        C-RANGE = "L" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  DEC(ttDados.fator-utiliz * ttDados.eficiencia-media * ttDados.de-classif).  //PERCENTUAL OEE POR APONTAMENTO
        RUN formataCelula(C-RANGE).


        
        
        
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-displayDados) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE displayDados Procedure 
PROCEDURE displayDados :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
------------------------------------------------------------------------------*/
DEFINE VARIABLE I-CONT AS INTEGER     NO-UNDO.
  
RUN createTTDados.

CREATE "Excel.Application"  chExcelApplication.
ASSIGN chExcelApplication:SheetsInNewWorkbook       = 1
       chPasta                                      = chExcelApplication:Workbooks:Add() /* Cria uma Pasta */
       chExcelApplication:APPLICATION:DisplayAlerts = FALSE.
       chExcelApplication:ActiveWindow:DisplayGridlines = False.
       chExcelApplication:VISIBLE                   = FALSE. 
       
 IF tt-param.l-diario = true THEN
    RUN geraExcell NO-ERROR.
IF tt-param.l-mensal = true THEN
    RUN geraExcellMensal NO-ERROR.
    
ASSIGN chExcelApplication:APPLICATION:DisplayAlerts = FALSE.
       chExcelApplication:VISIBLE                   = TRUE.

if valid-handle(chWorkSheet) then 
   RELEASE OBJECT chWorkSheet.  

if valid-handle(chPasta) then 
   RELEASE OBJECT chPasta. 

IF valid-handle(chWorkbook) then
   RELEASE OBJECT chWorkbook.

if valid-handle(chPlanilha) then
   RELEASE OBJECT chPlanilha.

if valid-handle(chExcelApplication) then
   RELEASE OBJECT chExcelApplication.
 
IF VALID-HANDLE(ch-celula) THEN
    release object ch-celula no-error.
    
 IF VALID-HANDLE(chChart) THEN
        release object      chChart no-error.
     IF VALID-HANDLE(chRange) THEN
        release object      chRange no-error.
    
IF VALID-HANDLE(chWorksheetRange) THEN
   release OBJECT chWorksheetRange no-error.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-displayParametros) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE displayParametros Procedure 
PROCEDURE displayParametros :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
  
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-formataCelula) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE formataCelula Procedure 
PROCEDURE formataCelula :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEF INPUT PARAM p-range AS CHAR NO-UNDO.

        chPlanilha:Range(p-range):interior:colorindex = 2.
        chPlanilha:Range(p-range):font:colorindex     = 1.
        chPlanilha:Range(p-range):Borders(7):Weight   = 2. /* Esquerda */
        chPlanilha:Range(p-range):Borders(8):Weight   = 2. /* Superior */
        chPlanilha:Range(p-range):Borders(9):Weight   = 2. /* Inferior */
        chPlanilha:Range(p-range):Borders(11):Weight  = 2. /* Interna Vertical */
        chPlanilha:Range(p-range):Borders(12):Weight  = 2. /* Interna Horizontal */
        chPlanilha:Range(p-range):Borders(10):Weight  = 2. /* Direita */

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-geraCabecalho) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE geraCabecalho Procedure 
PROCEDURE geraCabecalho :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
  DEF INPUT PARAM p-docto AS INT NO-UNDO.
  DEF VAR c-name AS CHAR.
  
                     
    FIND first ctrab
         WHERE ctrab.gm-codigo = ttCentro.gm-codigo
           AND ctrab.cod-ctrab = ttCentro.cod-ctrab NO-LOCK NO-ERROR.
  CASE p-docto:
     WHEN 1 THEN DO:
         if  search('Layout\Productivity.xlt') <> ?  AND search('Layout\Productivity.xlt') <> "" then do:
              file-info:file-name = 'Layout\Productivity.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = STRING(ttCentro.cod-ctrab) + " - Prod".
     END.
     WHEN 2 THEN DO:
      if  search('Layout\Relatorio.xlt') <> ?  AND search('Layout\Relatorio.xlt') <> "" then do:
              file-info:file-name = 'Layout\Relatorio.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = STRING(ttCentro.cod-ctrab) + "Rel".
     END.
      WHEN 3 THEN DO:
      if  search('Layout\GraficoOEE.xlt') <> ?  AND search('Layout\GraficoOEE.xlt') <> "" then do:
              file-info:file-name = 'Layout\GraficoOEE.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = STRING(ttCentro.cod-ctrab) + " - OEE".
     END.
      WHEN 4 THEN DO:
      if  search('Layout\GraficoPerf.xlt') <> ?  AND search('Layout\GraficoPerf.xlt') <> "" then do:
              file-info:file-name = 'Layout\GraficoPerf.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = STRING(ttCentro.cod-ctrab) + " - Perf".
     END.
      WHEN 5 THEN DO:
      if  search('Layout\GraficoDisp.xlt') <> ?  AND search('Layout\GraficoDisp.xlt') <> "" then do:
              file-info:file-name = 'Layout\GraficoDisp.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = STRING(ttCentro.cod-ctrab) + " - Disp".
     END.
      WHEN 6 THEN DO:
      if  search('Layout\GraficoQual.xlt') <> ?  AND search('Layout\GraficoQual.xlt') <> "" then do:
              file-info:file-name = 'Layout\GraficoQual.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = STRING(ttCentro.cod-ctrab) + " - Qual".
     END.
     WHEN 7 THEN DO:
      if  search('Layout\GraficoParadas.xlt') <> ?  AND search('Layout\GraficoParadas.xlt') <> "" then do:
              file-info:file-name = 'Layout\GraficoParadas.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = STRING(ttCentro.cod-ctrab) + " - Par Alt Efic".
     END.
     WHEN 8 THEN DO:
      if  search('Layout\RelatorioParadas.xlt') <> ?  AND search('Layout\RelatorioParadas.xlt') <> "" then do:
              file-info:file-name = 'Layout\RelatorioParadas.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = STRING(ttCentro.cod-ctrab) + " - Rel.Parada".
     END.
     WHEN 9 THEN DO:
      if  search('Layout\RelatorioMes.xlt') <> ?  AND search('Layout\RelatorioMes.xlt') <> "" then do:
              file-info:file-name = 'Layout\RelatorioMes.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = " - Relatorio Mensal - ".
     END.
      WHEN 10 THEN DO:
      if  search('Layout\GraficoMes.xlt') <> ?  AND search('Layout\GraficoMes.xlt') <> "" then do:
              file-info:file-name = 'Layout\GraficoMes.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = " - Grafico Mensal - ".
     END.
      WHEN 11 THEN DO:
      if  search('Layout\ParadasMes.xlt') <> ?  AND search('Layout\ParadasMes.xlt') <> "" then do:
              file-info:file-name = 'Layout\ParadasMes.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = " - Graf. Pradas Mensal - ".
     END.
      WHEN 12 THEN DO:
      if  search('Layout\GraficoParadas.xlt') <> ?  AND search('Layout\GraficoParadas.xlt') <> "" then do:
              file-info:file-name = 'Layout\GraficoParadas.xlt'.
            if file-info:pathname <> ? then do:
               ASSIGN  c-modelo = file-info:pathname.
             END.
        END.
        c-name = STRING(ttCentro.cod-ctrab) + " - Par n∆o Alt Efic".
     END.
  END CASE.
  
   OUTPUT TO "CLIPBOARD" NO-CONVERT.
   PUT "".
   OUTPUT CLOSE.
   
  ASSIGN i-planilha = i-planilha  + 1
         i-linha    = 6.

   chPasta:Sheets:ADD(,chPasta:Sheets:Item(i-planilha - 1)) no-error.
   IF VALID-HANDLE(chPlanilha) THEN
        release object chPlanilha no-error.

   ASSIGN chPlanilha      = chPasta:Sheets:Item(i-planilha) no-error.
          chPlanilha:Name = c-name .


   assign cFile         = search(c-modelo)
          chArquivo     = chExcelApplication:WorkBooks:Open(cFile).

   IF i-planilha <> 4 THEN DO:
        /*  chArquivo:Windows(1):DisplayGridLines = false.*/
        chPlanilhaMod = chArquivo:Sheets:Item(1) no-error.
        chSelection   = chPlanilhaMod:Rows("1:106") no-error. //"1:106"
   END.

   ELSE IF i-planilha = 4 THEN DO:
        chPlanilhaMod = chArquivo:Sheets:Item(1) no-error.
        chSelection   = chPlanilhaMod:Rows("1:45") no-error. //"1:106"
   END.
        


   chSelection:COPY().
   chPlanilha:PASTE().

        //AJUSTANDO LAYOUT DE IMPRESSAO

        //Modo Paisagem
        chPlanilha:PageSetup:ORIENTATION = 2.

        //Diminuir Margens para 0,5 cm
        chPlanilha:PageSetup:LeftMargin = 15.
        chPlanilha:PageSetup:RightMargin = 15.
        chPlanilha:PageSetup:TopMargin = 15.
        chPlanilha:PageSetup:BottomMargin = 15.
        chPlanilha:PageSetup:HeaderMargin = 0.
        chPlanilha:PageSetup:FooterMargin = 0.

        //Ajustar Colunas em uma pagina
        chPlanilha:PageSetup:Zoom = FALSE.
        chPlanilha:PageSetup:FitToPagesWide = 1.
        chPlanilha:PageSetup:FitToPagesTall = FALSE.

   OUTPUT TO "CLIPBOARD" NO-CONVERT.
   PUT "".
   OUTPUT CLOSE.

   chArquivo:Close(YES).

   release object chPlanilhaMod no-error.
   release object chArquivo  no-error.
   RELEASE OBJECT chSelection NO-ERROR.
   
   chExcelApplication:ActiveWindow:DisplayGridlines = False.
  CASE p-docto:
     when 1 THEN DO:

     
      ASSIGN C-RANGE = "G2" .
      
      chPlanilha:Range(C-RANGE):value = ' PERIODO:  ' + STRING(tt-param.dt_dt_ini,'99/99/9999') + '  -  ' + STRING(tt-param.dt_dt_fim,'99/99/9999').
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):HorizontalAlignment = -4108.
      chPlanilha:Range(C-RANGE):font:colorindex     = 1.
      chPlanilha:Range(C-RANGE):FONT:BOLD  = TRUE.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      
      ASSIGN C-RANGE = "G3" .
     
      chPlanilha:Range(C-RANGE):value = '    CENTRO:  ' + STRING(ttCentro.cod-ctrab) + '  -  ' + STRING(ttCentro.gm-codigo) + STRING(ttCentro.DESC-GRUPO).
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):HorizontalAlignment = -4108.
      chPlanilha:Range(C-RANGE):font:colorindex     = 1.
      chPlanilha:Range(C-RANGE):FONT:BOLD  = TRUE.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
    

     END.
      WHEN 9 
        OR WHEN 10
        OR WHEN  11 THEN DO:
     
       FIND FIRST  ttPeriodo USE-INDEX ch-codigo NO-ERROR.
         IF  AVAIL ttPeriodo THEN DO:
         
            RUN pi-retorna-mes (INPUT ttPeriodo.i-mes, OUTPUT cMesIni).
       
             cMesIni  = cMesIni + "/" + STRING(ttPeriodo.i-ano).
            
         END.
         
         FIND last  ttPeriodo USE-INDEX ch-codigo NO-ERROR.
         IF  AVAIL ttPeriodo THEN DO:
             RUN pi-retorna-mes (INPUT ttPeriodo.i-mes, OUTPUT cMesFim).
             
             cMesFim  = cMesFim + "/" + STRING(ttPeriodo.i-ano).
            
         END.
      
      
      
      
      ASSIGN C-RANGE = "C3" .
      chPlanilha:Range(C-RANGE):MERGE.
      chPlanilha:Range(C-RANGE):value = ' PERIODO:  ' + cMesIni  + '  -  ' + cMesFim .
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):HorizontalAlignment = -4108.
      chPlanilha:Range(C-RANGE):font:colorindex     = 1.
      chPlanilha:Range(C-RANGE):FONT:BOLD  = TRUE.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
     
     
       
  END.
  WHEN 8888 THEN DO:
  
                           
  ASSIGN C-RANGE = "C3" .
      
      chPlanilha:Range(C-RANGE):value = 'PERIODO:  ' + STRING(tt-param.dt_dt_ini,'99/99/9999') + '  -  ' + STRING(tt-param.dt_dt_fim,'99/99/9999').
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):HorizontalAlignment = -4108.
      chPlanilha:Range(C-RANGE):font:colorindex     = 1.
      chPlanilha:Range(C-RANGE):FONT:BOLD  = TRUE.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      
         
           ASSIGN C-RANGE = "C4" .
      
      chPlanilha:Range(C-RANGE):value = 'CENTRO:  ' + STRING(ttCentro.cod-ctrab) + '  -  ' + STRING(ttCentro.gm-codigo) +  "    " +  STRING(ttCentro.DESC-GRUPO).
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):HorizontalAlignment = -4108.
      chPlanilha:Range(C-RANGE):font:colorindex     = 1.
      chPlanilha:Range(C-RANGE):FONT:BOLD  = TRUE.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
     
      
      
      ASSIGN C-RANGE = "C9" .
     
      chPlanilha:Range(C-RANGE):value = STRING(ctrab.val-eficien-ctrab,">>>9.99") + "%" .
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):HorizontalAlignment = -4108.
      chPlanilha:Range(C-RANGE):font:colorindex     = 1.
      chPlanilha:Range(C-RANGE):FONT:BOLD  = TRUE.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      
      
      
      
      
      
      
      

  END.
  OTHERWISE DO:
      
 
      ASSIGN C-RANGE = "C3" .
      chPlanilha:Range(C-RANGE):value               = 'PERIODO:  ' + STRING(tt-param.dt_dt_ini,'99/99/9999') + '  -  ' + STRING(tt-param.dt_dt_fim,'99/99/9999').
      chPlanilha:Range(C-RANGE):font:colorindex     = 1.
      chPlanilha:Range(C-RANGE):FONT:BOLD           = TRUE.
      
      ASSIGN C-RANGE = "C4" .
      chPlanilha:Range(C-RANGE):value               = 'CENTRO:  ' + STRING(ttCentro.cod-ctrab) + '  -  ' + STRING(ttCentro.gm-codigo) + " - " +   STRING(ttCentro.DESC-GRUPO).
      chPlanilha:Range(C-RANGE):font:colorindex     = 1.
      chPlanilha:Range(C-RANGE):FONT:BOLD           = TRUE.
     
      ASSIGN C-RANGE = "C9" .
      chPlanilha:Range(C-RANGE):VALUE               = STRING(ctrab.val-eficien-ctrab,">>>9.99") + "%" .
      chPlanilha:Range(C-RANGE):font:colorindex     = 1.
      chPlanilha:Range(C-RANGE):FONT:BOLD           = TRUE.
    
   
       
       
   
  END.    
END CASE.    
    

   
  
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-geraExcell) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE geraExcell Procedure 
PROCEDURE geraExcell :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
  DEFINE VARIABLE decapcaidade LIKE ctrab.qtd-capac-ctrab     NO-UNDO.
  DEFINE VARIABLE degeral      LIKE ctrab.qtd-capac-ctrab     NO-UNDO.
  DEFINE VARIABLE deaplicado   LIKE ctrab.qtd-capac-ctrab     NO-UNDO.
  DEFINE VARIABLE deValor      as decimal format ">>9.99"     NO-UNDO.
  DEFINE VARIABLE c-ctrab      AS char NO-UNDO.
  DEFINE VARIABLE iDias        AS INTEGER     NO-UNDO.
  ASSIGN c-ctrab    = ''
         i-planilha = 0.

  EMPTY TEMP-TABLE ttParadasAux.

  
   FOR EACH ttParadas
       WHERE ttParadas.cod-ctrab   = ttCentro.cod-ctrab
       BREAK BY ttParadas.data-inicio:
       FIND FIRST ttParadasAux
           WHERE ttParadasAux.cod-ctrab  = ttParadas.cod-ctrab
             AND ttParadasAux.cod-parada = ttParadas.cod-parada NO-ERROR.
        IF NOT AVAIL ttParadasAux THEN DO:
            CREATE ttParadasAux.
            ASSIGN 
            ttParadasAux.cod-ctrab   = ttParadas.cod-ctrab
            ttParadasAux.cod-parada  = ttParadas.cod-parada
            ttParadasAux.motivo      = ttParadas.motivo
            ttParadasAux.data-inicio = ttParadas.data-inicio.
            
        END.
     END.
  

  FOR EACH ttCentro USE-INDEX idx1:

      IF VALID-HANDLE(h-acomp) THEN  
         run pi-acompanhar in h-acomp (INPUT "Gerando Planilha " + ttCentro.cod-ctrab).
    IF  c-ctrab <> ttCentro.cod-ctrab THEN DO:
        RUN geraCabecalho (1).
        ASSIGN c-ctrab = ttCentro.cod-ctrab.
    END.
    ASSIGN iDias = 0.
    FOR EACH ttDias USE-INDEX idx2
        WHERE ttDias.cod-ctra = ttCentro.cod-ctrab:
        ASSIGN iDias = iDias + 1.
    END.
    /** Capacidade para o periodo 21 horas X capacidade cadastadas
    ** no SF0104 (por hora) *  numero de dias em que ocorre report de produ?∆o
    ****/
    
    ASSIGN decapcaidade = (21 * iDias ) 
           deaplicado   = 0
           deutiliz     = 0
           desoma       = 0
           deSTemp      = 0
           degeral      = 0.
   
    FOR EACH  ttTotalCentro
        WHERE ttTotalCentro.cod-ctrab = ttCentro.cod-ctrab
        BREAK BY ttTotalCentro.c-horas :
        RUN pi-busca-linha (INPUT ttTotalCentro.c-horas,
                            OUTPUT i-linha).
          IF i-linha = 0 THEN
          DO:
                  NEXT.
          END.
         ASSIGN 
               chPlanilha:Range("C" + STRING(i-linha)):value        = ttTotalCentro.tempo-utilizado.
               chPlanilha:Range("C" + STRING(i-linha)):numberformat =  "##0,00".
               
        ASSIGN deaplicado = deaplicado + ttTotalCentro.tempo-utilizado
               deValor    = deValor    + ttTotalCentro.tempo-parada
               deSTemp    = deSTemp    + ttTotalCentro.tempo-setup
               deutiliz   = deutiliz + (  ttTotalCentro.tempo-utilizado + ttTotalCentro.tempo-parada + ttTotalCentro.tempo-setup ).
    END.
   
            
        ASSIGN deaplicado = (deaplicado * 100) / deutiliz
               deValor    = (deValor    * 100) / deutiliz
               deSTemp    = (deSTemp    * 100) / deutiliz.
    chPlanilha:Range("H7"):value  = deValor .
    chPlanilha:Range("H7"):numberformat =  "##0,00".
    chPlanilha:Range("H9"):value = deSTemp.
    chPlanilha:Range("H9"):numberformat =  "##0,00".
    chPlanilha:Range("H13"):value = deaplicado.
    chPlanilha:Range("H13"):numberformat =  "##0,00".
    ASSIGN degeral =  deaplicado + deValor + deSTemp.
    chPlanilha:Range("J15"):value = decapcaidade.
    chPlanilha:Range("K15"):value = degeral. 
    
    ASSIGN i-linha = 14.
    
    IF VALID-HANDLE(ch-celula) THEN
        release object ch-celula no-error.
    
    
    
    FOR EACH ttItem
       WHERE ttItem.cod-ctrab = ttCentro.cod-ctrab:
        IF i-linha >= 27 THEN DO:
           ch-celula = chPlanilha:Cells(ch-celula:ROW + 1,ch-celula:COLUMN).
           chPlanilha:Cells(ch-celula:ROW - 1,ch-celula:COLUMN):COPY(ch-celula).
           ch-celula:INSERT(-4121). /*xlDown*/
        END.


        chPlanilha:Range("L" + STRING(i-linha)):value = STRING(ttItem.it-codigo).
        chPlanilha:Range("L" + STRING(i-linha)):EntireColumn:AutoFit.
        ASSIGN i-linha = i-linha + 1.
        ch-celula = chPlanilha:Cells(i-linha,12).


    END.
    ASSIGN i-linha = 14.
     FOR EACH ttParadasAux
       WHERE ttParadasAux.cod-ctrab   = ttCentro.cod-ctrab
       BREAK BY ttParadasAux.data-inicio:
        chPlanilha:Range("M" + STRING(i-linha)):value = STRING(ttParadasAux.motivo).
        chPlanilha:Range("M" + STRING(i-linha)):EntireColumn:AutoFit.
        ASSIGN i-linha = i-linha + 1.
     END.
     if  search('Layout\logo_mini.png') <> ?  AND search('Layout\logo_mini.png') <> "" then do:
       file-info:file-name = 'Layout\logo_mini.png'.
       if file-info:pathname <> ? then do:
            ASSIGN  C-FOTO = file-info:pathname.
        END.
         if  C-FOTO <> ?  and
          C-FOTO <> "" then do:
          C-RANGE = "A1".
          chImagem =  chPlanilha:Shapes:AddPicture(string(C-FOTO),FALSE,TRUE,chPlanilha:Range(C-RANGE):LEFT,chPlanilha:Range(C-RANGE):TOP,TRUE,TRUE).
         END.
     END.
    chExcelApplication:COLUMNS("A:A"):ColumnWidth = 0. //modificado
    chExcelApplication:COLUMNS("B:B"):ColumnWidth = 5.00.
    chExcelApplication:COLUMNS("C:C"):ColumnWidth = 5.00.
    chExcelApplication:COLUMNS("D:D"):ColumnWidth = 5.00.
    chExcelApplication:COLUMNS("E:E"):ColumnWidth = 10.00.
    chExcelApplication:COLUMNS("F:F"):ColumnWidth = 10.00.
    chExcelApplication:COLUMNS("G:G"):ColumnWidth = 10.00.
    chExcelApplication:COLUMNS("H:H"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("I:I"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("J:J"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("K:K"):ColumnWidth = 14.00.
    chExcelApplication:COLUMNS("L:L"):ColumnWidth = 12.00.
  
    RUN geraCabecalho (2).    
    ASSIGN i-linha    = 13
           v-total1   = 0  
           v-total2   = 0 
           v-total3   = 0 
           v-total4   = 0 
           v-total5   = 0 
           v-total6   = 0 
           v-total7   = 0 
           v-total8   = 0
           v-total9   = 0
           i-cont-reg = 0
           v-total10  = 0
           v-total11  = 0.
           
    DEF VAR calc-oee AS DECIMAL.
      
    FOR EACH ttDados
        WHERE ttDados.cod-ctrab = ttCentro.cod-ctrab
        BREAK BY ttDados.data-inicio:
        

        ASSIGN  ttDados.eficiencia-media   = ttDados.carga-utilizada / ttDados.carga-disponivel
                ttDados.oee                = ttDados.fator-utiliz * ttDados.eficiencia-media * ttDados.de-classif
                ttDados.de-classif         = ttDados.qtd-aprov / ttDados.qtd-refgada.
                

        //RENAN
                
        RUN dadosRelatorio NO-ERROR.   
        ASSIGN
        calc-oee   = DEC((ttDados.tempo-utilizado / ttDados.tempo-disponivel) * (ttDados.carga-utilizada / ttDados.carga-disponivel) * (ttDados.qtd-aprov / ttDados.qtd-refgada))
        i-cont-reg = i-cont-reg + 1
        v-total1   = v-total1  + ttDados.tempo-disponivel
        v-total2   = v-total2  + ttDados.tempo-utilizado
        v-total3   = v-total3  + ttDados.carga-disponivel
        v-total4   = v-total4  + ttDados.carga-utilizada
        v-total5   = v-total5  + ttDados.qtd-aprov
        v-total6   = v-total6  + ttDados.qtd-refgada
        v-total7   = v-total7  + ttDados.fator-utiliz
        v-total8   = v-total8  + ttDados.eficiencia-media 
        v-total9   = v-total9  + ttDados.de-classif
        v-total10  = v-total10 + calc-oee
        i-linha    = i-linha   + 1.

        ASSIGN calc-oee = 0.
    END.

    ASSIGN 
        v-total1   = v-total1  / i-cont-reg
        v-total2   = v-total2  / i-cont-reg
        v-total3   = v-total3  / i-cont-reg
        v-total4   = v-total4  / i-cont-reg
        v-total5   = v-total5  / i-cont-reg
        v-total6   = v-total6  / i-cont-reg
        v-total7   = v-total7  / i-cont-reg
        v-total8   = v-total8  / i-cont-reg 
        v-total9   = v-total9  / i-cont-reg
        v-total10  = v-total10 / i-cont-reg.
    
    chExcelApplication:COLUMNS("A:A"):ColumnWidth = 0.
    chExcelApplication:COLUMNS("B:B"):ColumnWidth = 26.00.
    chExcelApplication:COLUMNS("C:C"):ColumnWidth = 18.00.
    chExcelApplication:COLUMNS("D:D"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("E:E"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("F:F"):ColumnWidth = 14.00.
    chExcelApplication:COLUMNS("G:G"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("H:H"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("I:I"):ColumnWidth = 15.00.
    chExcelApplication:COLUMNS("J:J"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("K:K"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("L:L"):ColumnWidth = 14.00.
    RUN geraRodape. 
    if  search('Layout\logo_mini.png') <> ?  AND search('Layout\logo_mini.png') <> "" then do:
       file-info:file-name = 'Layout\llogo_mini.png'.
       if file-info:pathname <> ? then do:
            ASSIGN  C-FOTO = file-info:pathname.
        END.
         if  C-FOTO <> ?  and
          C-FOTO <> "" then do:
          C-RANGE = "B2".
          chImagem =  chPlanilha:Shapes:AddPicture(string(C-FOTO),FALSE,TRUE,chPlanilha:Range(C-RANGE):LEFT,chPlanilha:Range(C-RANGE):TOP,TRUE,TRUE).
         END.
     END.
    
    RUN geraCabecalho (3). 
    RUN putValores.
    
    if  search('Layout\logo_mini.png') <> ?  AND search('Layout\logo_mini.png') <> "" then do:
         file-info:file-name = 'Layout\logo_mini.png'.
         if file-info:pathname <> ? then do:
            ASSIGN  C-FOTO = file-info:pathname.
          END.
          
      if  C-FOTO <> ?  AND C-FOTO <> "" then do:
          C-RANGE = "B2".
          chImagem =  chPlanilha:Shapes:AddPicture(string(C-FOTO),FALSE,TRUE,chPlanilha:Range(C-RANGE):LEFT,chPlanilha:Range(C-RANGE):TOP,TRUE,TRUE).
      END.
    END.
    chExcelApplication:COLUMNS("A:A"):ColumnWidth = 0.
    chExcelApplication:COLUMNS("B:B"):ColumnWidth = 26.00.
    chExcelApplication:COLUMNS("C:C"):ColumnWidth = 18.00.
    chExcelApplication:COLUMNS("D:D"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("E:E"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("F:F"):ColumnWidth = 14.00.
    chExcelApplication:COLUMNS("G:G"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("H:H"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("I:I"):ColumnWidth = 15.00.
    chExcelApplication:COLUMNS("J:J"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("K:K"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("L:L"):ColumnWidth = 14.00.
    RUN geraCabecalho (8).
    RUN relatorioParadas.
   if  search('Layout\logo_mini.png') <> ?  AND search('Layout\logo_mini.png') <> "" then do:
       file-info:file-name = 'Layout\logo_mini.png'.
       if file-info:pathname <> ? then do:
          ASSIGN  C-FOTO = file-info:pathname.
        END.
       if  C-FOTO <> ?  and
          C-FOTO <> "" then do:
          C-RANGE = "B2".
          chImagem =  chPlanilha:Shapes:AddPicture(string(C-FOTO),FALSE,TRUE,chPlanilha:Range(C-RANGE):LEFT,chPlanilha:Range(C-RANGE):TOP,TRUE,TRUE).
       END.
   END.
    chExcelApplication:COLUMNS("A:A"):ColumnWidth = 0.
    chExcelApplication:COLUMNS("B:B"):ColumnWidth = 26.00.
    chExcelApplication:COLUMNS("C:C"):ColumnWidth = 18.00.
    chExcelApplication:COLUMNS("D:D"):ColumnWidth = 20.00.
    chExcelApplication:COLUMNS("E:E"):ColumnWidth = 20.00.
    chExcelApplication:COLUMNS("F:F"):ColumnWidth = 30.00.
    chExcelApplication:COLUMNS("G:G"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("H:H"):ColumnWidth = 20.00.
    chExcelApplication:COLUMNS("I:I"):ColumnWidth = 20.00.
    chExcelApplication:COLUMNS("J:J"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("K:K"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("L:L"):ColumnWidth = 14.00.

    RUN geraCabecalho (7).
    RUN dadosParadaAltEfic.
    chExcelApplication:COLUMNS("A:A"):ColumnWidth = 0.
    chExcelApplication:COLUMNS("B:B"):ColumnWidth = 26.00.
    chExcelApplication:COLUMNS("C:C"):ColumnWidth = 18.00.
    chExcelApplication:COLUMNS("D:D"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("E:E"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("F:F"):ColumnWidth = 14.00.
    chExcelApplication:COLUMNS("G:G"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("H:H"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("I:I"):ColumnWidth = 15.00.
    chExcelApplication:COLUMNS("J:J"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("K:K"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("L:L"):ColumnWidth = 14.00.
   if  search('Layout\logo_mini.png') <> ?  AND search('Layout\logo_mini.png') <> "" then do:
       file-info:file-name = 'Layout\logo_mini.png'.
       if file-info:pathname <> ? then do:
          ASSIGN  C-FOTO = file-info:pathname.
        END.
       if  C-FOTO <> ?  and
          C-FOTO <> "" then do:
          C-RANGE = "B2".
          chImagem =  chPlanilha:Shapes:AddPicture(string(C-FOTO),FALSE,TRUE,chPlanilha:Range(C-RANGE):LEFT,chPlanilha:Range(C-RANGE):TOP,TRUE,TRUE).
       END.
   END.  
   
   RUN geraCabecalho (12).
    RUN dadosParadaNaoAltEfic.
    chExcelApplication:COLUMNS("A:A"):ColumnWidth = 0.
    chExcelApplication:COLUMNS("B:B"):ColumnWidth = 26.00.
    chExcelApplication:COLUMNS("C:C"):ColumnWidth = 18.00.
    chExcelApplication:COLUMNS("D:D"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("E:E"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("F:F"):ColumnWidth = 14.00.
    chExcelApplication:COLUMNS("G:G"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("H:H"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("I:I"):ColumnWidth = 15.00.
    chExcelApplication:COLUMNS("J:J"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("K:K"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("L:L"):ColumnWidth = 14.00.
   if  search('Layout\logo_mini.png') <> ?  AND search('Layout\logo_mini.png') <> "" then do:
       file-info:file-name = 'Layout\logo_mini.png'.
       if file-info:pathname <> ? then do:
          ASSIGN  C-FOTO = file-info:pathname.
        END.
       if  C-FOTO <> ?  and
          C-FOTO <> "" then do:
          C-RANGE = "B2".
          chImagem =  chPlanilha:Shapes:AddPicture(string(C-FOTO),FALSE,TRUE,chPlanilha:Range(C-RANGE):LEFT,chPlanilha:Range(C-RANGE):TOP,TRUE,TRUE).
       END.
   END.  

         
      
 END.

   
  
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-geraExcellMensal) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE geraExcellMensal Procedure 
PROCEDURE geraExcellMensal :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE VARIABLE cMes AS CHAR NO-UNDO.

EMPTY TEMP-TABLE ttDadosMensal.
EMPTY TEMP-TABLE ttPeriodo.
EMPTY TEMP-TABLE ttTotalMensal.
EMPTY TEMP-TABLE ttParadasMes.
EMPTY TEMP-TABLE ttParadasTotalMes.

cMesIni = ''.
i-cont-reg  = 0.
FOR EACH ttCentro USE-INDEX idx1:
      IF VALID-HANDLE(h-acomp) THEN  
         run pi-acompanhar in h-acomp (INPUT "Gerando Dados Mensal: " + ttCentro.cod-ctrab).
   FOR EACH ttDados
        WHERE ttDados.cod-ctrab = ttCentro.cod-ctrab
        BREAK BY ttDados.data-inicio:
        

        ASSIGN  ttDados.eficiencia-media   = ttDados.carga-utilizada / ttDados.carga-disponivel
                ttDados.oee                = ttDados.fator-utiliz * ttDados.eficiencia-media * ttDados.de-classif //( (ttDados.fator-utiliz + ttDados.eficiencia-media + ttDados.de-classif  ) * 100 ) / ttDados.val-eficien-ctrab 
                ttDados.de-classif         = ( ttDados.qtd-aprov * 100 ) / ttDados.qtd-refgada.
                
         ASSIGN cMes = string(MONTH(ttDados.data-inicio)).    
         
       FIND FIRST ttDadosMensal
            WHERE ttDadosMensal.cod-ctrab = ttDados.cod-ctrab
              AND ttDadosMensal.c-mes     = cMes NO-ERROR.
       IF NOT AVAIL ttDadosMensal THEN DO:
           CREATE ttDadosMensal.
           ASSIGN ttDadosMensal.cod-ctrab = ttDados.cod-ctrab
                  ttDadosMensal.c-mes     = cMes.
                   
                  
       END.
       ASSIGN 
       ttDadosMensal.tempo-disponivel    =  ttDadosMensal.tempo-disponivel   + ttDados.tempo-disponivel
       ttDadosMensal.tempo-utilizado     =  ttDadosMensal.tempo-utilizado    + ttDados.tempo-utilizado 
       ttDadosMensal.carga-utilizada     =  ttDadosMensal.carga-utilizada    + ttDados.carga-utilizada 
       ttDadosMensal.carga-disponivel    =  ttDadosMensal.carga-disponivel   + ttDados.carga-disponivel
       ttDadosMensal.qtd-aprov           =  ttDadosMensal.qtd-aprov          + ttDados.qtd-aprov       
       ttDadosMensal.qtd-refgada         =  ttDadosMensal.qtd-refgada        + ttDados.qtd-refgada     
       ttDadosMensal.fator-utiliz        =  ttDadosMensal.fator-utiliz       + ttDados.fator-utiliz
       ttDadosMensal.eficiencia-media    =  ttDadosMensal.eficiencia-media   + ttDados.eficiencia-media
       ttDadosMensal.oee                 =  ttDadosMensal.oee                + ttDados.oee              
       ttDadosMensal.de-classif          =  ttDadosMensal.de-classif         + ttDados.de-classif
       ttDadosMensal.val-eficien-ctrab   =  ttDados.val-eficien-ctrab.
       
       
       FIND FIRST  ttPeriodo
        WHERE ttPeriodo.i-ano  = YEAR(ttDados.data-inicio)
          and ttPeriodo.i-mes =  MONTH(ttDados.data-inicio)  NO-ERROR.
       IF NOT AVAIL ttPeriodo THEN DO:
           CREATE ttPeriodo.
           ASSIGN 
           ttPeriodo.i-ano  = YEAR(ttDados.data-inicio)
           ttPeriodo.i-mes =  MONTH(ttDados.data-inicio).
           
       END.
       
       FIND FIRST ttTotalMensal
            WHERE ttTotalMensal.i-ano  = YEAR(ttDados.data-inicio)
              AND ttTotalMensal.i-mes =  MONTH(ttDados.data-inicio) NO-ERROR.
       IF NOT AVAIL ttTotalMensal THEN DO:
           CREATE ttTotalMensal.
           ASSIGN ttTotalMensal.i-ano  = YEAR(ttDados.data-inicio)
                  ttTotalMensal.i-mes =  MONTH(ttDados.data-inicio).
                  i-cont-reg = i-cont-reg + 1.
       END.
       ASSIGN 
       ttTotalMensal.fator-utiliz        =  ttDadosMensal.fator-utiliz       + ttDados.fator-utiliz
       ttTotalMensal.eficiencia-media    =  ttDadosMensal.eficiencia-media   + ttDados.eficiencia-media
       ttTotalMensal.oee                 =  ttDadosMensal.oee                + ttDados.oee              
       ttTotalMensal.de-classif          =  ttDadosMensal.de-classif         + ttDados.de-classif
       ttTotalMensal.val-eficien-ctrab   =  ttDados.val-eficien-ctrab.
       
            
   END.
  

END.
 ASSIGN i-linha    = 8
           v-total1   = 0  
           v-total2   = 0 
           v-total3   = 0 
           v-total4   = 0 
           v-total5   = 0 
           v-total6   = 0 
           v-total7   = 0 
           v-total8   = 0
           v-total9   = 0
           i-cont-reg = 0
           v-total10  = 0
           v-total11  = 0.
RUN geraCabecalho(9).
ASSIGN i-linha    = 8.
FOR EACH ttDadosMensal:
    IF VALID-HANDLE(h-acomp) THEN  
      run pi-acompanhar in h-acomp (INPUT "Planilha Mensal: " + ttDadosMensal.cod-ctrab + " - " + ttDadosMensal.c-mes).
      RUN dadosMensal.
        ASSIGN
        i-cont-reg = i-cont-reg + 1
        v-total1   = v-total1  + ttDadosMensal.tempo-disponivel
        v-total2   = v-total2  + ttDadosMensal.tempo-utilizado
        v-total3   = v-total3  + ttDadosMensal.carga-disponivel
        v-total4   = v-total4  + ttDadosMensal.carga-utilizada
        v-total5   = v-total5  + ttDadosMensal.qtd-aprov
        v-total6   = v-total6  + ttDadosMensal.qtd-refgada 
        v-total7   = v-total7  + ttDadosMensal.fator-utiliz
        v-total8   = v-total8  + ttDadosMensal.eficiencia-media 
        v-total9   = v-total9  + ttDadosMensal.de-classif
        v-total10  = v-total10 + ttDadosMensal.oee
        i-linha    = i-linha   + 1.

      

END.
    chExcelApplication:COLUMNS("A:A"):ColumnWidth = 17.00.
    chExcelApplication:COLUMNS("B:B"):ColumnWidth = 17.00.
    chExcelApplication:COLUMNS("C:C"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("D:D"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("E:E"):ColumnWidth = 16.00.
    chExcelApplication:COLUMNS("F:F"):ColumnWidth = 10.00.
    chExcelApplication:COLUMNS("G:G"):ColumnWidth = 10.00.
    chExcelApplication:COLUMNS("H:H"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("I:I"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("J:J"):ColumnWidth = 12.00.
    chExcelApplication:COLUMNS("K:K"):ColumnWidth = 14.00.
    chExcelApplication:COLUMNS("L:L"):ColumnWidth = 12.00.
RUN rodapeMensal.
 if  search('Layout\LOGO.png') <> ?  AND search('Layout\LOGO.png') <> "" then do:
              file-info:file-name = 'Layout\LOGO.png'.
            if file-info:pathname <> ? then do:
               ASSIGN  C-FOTO = file-info:pathname.
             END.
             
        END.

     if  C-FOTO <> ?  and
              C-FOTO <> "" then do:

             C-RANGE = "B2".

             chImagem =  chPlanilha:Shapes:AddPicture(string(C-FOTO),FALSE,TRUE,chPlanilha:Range(C-RANGE):LEFT,chPlanilha:Range(C-RANGE):TOP,TRUE,TRUE).

             IF VALID-HANDLE (chImagem) THEN DO:
                /*chImagem:LockAspectRatio = TRUE.*/
    
                chImagem:HEIGHT = 30.
                chImagem:WIDTH   = 140.
            END.
           
      END.
RUN geraCabecalho(10).
RUN graficoMensal.
 if  search('Layout\LOGO.png') <> ?  AND search('Layout\LOGO.png') <> "" then do:
              file-info:file-name = 'Layout\LOGO.png'.
            if file-info:pathname <> ? then do:
               ASSIGN  C-FOTO = file-info:pathname.
             END.
             
        END.

     if  C-FOTO <> ?  and
              C-FOTO <> "" then do:

           

               C-RANGE = "B2".

             chImagem =  chPlanilha:Shapes:AddPicture(string(C-FOTO),FALSE,TRUE,chPlanilha:Range(C-RANGE):LEFT,chPlanilha:Range(C-RANGE):TOP,TRUE,TRUE).

             IF VALID-HANDLE (chImagem) THEN DO:
                /*chImagem:LockAspectRatio = TRUE.*/
    
                chImagem:HEIGHT = 30.
                chImagem:WIDTH   = 140.
            END.
           
      END.
RUN geraCabecalho(11).
RUN paradasMensal.
 if  search('Layout\LOGO.png') <> ?  AND search('Layout\LOGO.png') <> "" then do:
              file-info:file-name = 'Layout\LOGO.png'.
            if file-info:pathname <> ? then do:
               ASSIGN  C-FOTO = file-info:pathname.
             END.
             
        END.

     if  C-FOTO <> ?  and
              C-FOTO <> "" then do:


               C-RANGE = "B2".

             chImagem =  chPlanilha:Shapes:AddPicture(string(C-FOTO),FALSE,TRUE,chPlanilha:Range(C-RANGE):LEFT,chPlanilha:Range(C-RANGE):TOP,TRUE,TRUE).

             IF VALID-HANDLE (chImagem) THEN DO:
                /*chImagem:LockAspectRatio = TRUE.*/
    
                chImagem:HEIGHT = 30.
                chImagem:WIDTH   = 140.
            END.
           
      END.


END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-geraGrafico) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE geraGrafico Procedure 
PROCEDURE geraGrafico :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/

    DEF INPUT PARAMETER chExcel        AS COM-HANDLE NO-UNDO.
    DEF INPUT PARAMETER chPasta        AS COM-HANDLE NO-UNDO.
    DEF INPUT PARAMETER chPlanilha     AS COM-HANDLE NO-UNDO.
    DEF INPUT PARAMETER pNomePasta     AS CHAR       NO-UNDO.
    DEF INPUT PARAMETER pRange         AS CHAR       NO-UNDO.

    DEF VAR chGrafico AS COM-HANDLE            NO-UNDO.
    DEF VAR i-num-max AS INTEGER    INIT 2     NO-UNDO.
    
    chGrafico = chPasta:Charts:ADD().

    {utp/ut-liter.i Performance_de_ProduÁ„o *}.

    chGrafico:ChartType = 65. /*Grafico de Linha*/
    chGrafico:SetSourceData(chPlanilha:Range(pRange), 1).
    chGrafico:HasTitle = True.
    chGrafico:ChartTitle:Characters:Text =  RETURN-VALUE + " - ":U +  pNomePasta .
    chGrafico:HasLegend = True.
    chGrafico:Legend:Select.
    chGrafico:PlotArea:Select.
    chGrafico:Legend:Select.

    chGrafico:SeriesCollection(1):ChartType = 51. /*Barras*/
    chGrafico:Location(2,pNomePasta).
    chExcelApplication:ActiveSheet:Shapes:ITEM(1):LEFT   = 125.
    chExcelApplication:ActiveSheet:Shapes:ITEM(1):TOP    = 10.
    chExcelApplication:ActiveSheet:Shapes:ITEM(1):WIDTH  = (i-coluna) * 50.
    chExcelApplication:ActiveSheet:Shapes:ITEM(1):HEIGHT = 305.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-geraRelatorio) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE geraRelatorio Procedure 
PROCEDURE geraRelatorio :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
   

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-geraRodape) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE geraRodape Procedure 
PROCEDURE geraRodape :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
        
        C-RANGE = "B" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value              = "MÇdia".  
        chPlanilha:Range(C-RANGE):interior:colorindex = 15.
        chPlanilha:Range(C-RANGE):font:colorindex     = 2.
        chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
  
        C-RANGE = "C" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = v-total1.  
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "D" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):VALUE               = v-total2.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "E" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):VALUE              = v-total3.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "F" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  v-total4.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "G" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):VALUE               = v-total5.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "H" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =   v-total6.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */

       
        //RENAN

        C-RANGE = "I" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value                = v-total7. //DEC(v-total2 / v-total1). //DISPONIBILIDADE MENSAL
        
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "J" + STRING(i-linha).
        
        chPlanilha:Range(C-RANGE):value                = v-total8. //DEC(v-total4 / v-total3).  //PERFORMANCE MENSAL
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "K" + STRING(i-linha).
        
        chPlanilha:Range(C-RANGE):value               =  v-total9. //DEC(v-total5 / v-total6). //QUALIDADE MENSAL
        chPlanilha:Range(C-RANGE):interior:colorindex = 2.
        chPlanilha:Range(C-RANGE):font:colorindex     = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
         C-RANGE = "L" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value                = v-total10. //DEC((v-total2 / v-total1) * (v-total4 / v-total3) * (v-total5 / v-total6)). //PERCENTUAL OEE MENSAL
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        
        
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-graficoMensal) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE graficoMensal Procedure 
PROCEDURE graficoMensal :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE VARIABLE deValor  as decimal format ">>>9.99"  NO-UNDO.
DEFINE VARIABLE MesPrint AS CHAR NO-UNDO.
DEFINE VARIABLE lFirst   AS LOG INIT NO NO-UNDO.
ASSIGN i-linha     = 65
       i-linha-fim = 0
       i-linha-aux = 11
       i-linha-ini = 11
       iRow       = 0.
       
FOR EACH ttTotalMensal
    USE-INDEX ch-codigo:
      RUN pi-retorna-mes (INPUT ttTotalMensal.i-mes, OUTPUT MesPrint).
      C-RANGE = "A" +  string(i-linha-aux).
      chPlanilha:Range( C-RANGE):value              = MesPrint.  
      C-RANGE = "B" +  string(i-linha-aux).
      deValor  = ROUND(ttTotalMensal.fator-utiliz,2).
      chPlanilha:Range( C-RANGE):value              = deValor.
      C-RANGE = "C" +  string(i-linha-aux).
      deValor  = ROUND(ttTotalMensal.eficiencia-media,2).
      chPlanilha:Range( C-RANGE):value              = deValor.
      C-RANGE = "D" +  string(i-linha-aux).
      deValor  = ROUND(ttTotalMensal.de-classif,2).
      chPlanilha:Range( C-RANGE):value              =  deValor.
      C-RANGE = "E" +  string(i-linha-aux).
      deValor  = ROUND(ttTotalMensal.oee,2).
      chPlanilha:Range( C-RANGE):value              = deValor.
     /* IF  lFirst = no THEN DO:*/
        C-RANGE = "F" +  string(i-linha-aux).
        deValor  = ROUND(ttTotalMensal.val-eficien-ctrab,2).
        chPlanilha:Range( C-RANGE):value              = deValor.
        lFirst = TRUE.
   /*   END.*/
      iRow = iRow + 1.
      IF  i-cont-reg = 1 THEN DO:
        i-linha-aux = i-linha-aux + 1.
        RUN pi-retorna-mes (INPUT ttTotalMensal.i-mes, OUTPUT MesPrint).
        C-RANGE = "A" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = MesPrint.  
        C-RANGE = "B" +  string(i-linha-aux).
        deValor  = ROUND(ttTotalMensal.fator-utiliz,2).
        chPlanilha:Range( C-RANGE):value              = deValor.
        C-RANGE = "C" +  string(i-linha-aux).
        deValor  = ROUND(ttTotalMensal.eficiencia-media,2).
        chPlanilha:Range( C-RANGE):value              = deValor.
        C-RANGE = "D" +  string(i-linha-aux).
        deValor  = ROUND(ttTotalMensal.de-classif,2).
        chPlanilha:Range( C-RANGE):value              =  deValor.
        C-RANGE = "E" +  string(i-linha-aux).
        deValor  = ROUND(ttTotalMensal.oee,2).
        chPlanilha:Range( C-RANGE):value              = deValor.
         iRow = iRow + 1.
      END.
      C-RANGE = "A" +  string(i-linha).
      chPlanilha:Range( C-RANGE):value              = MesPrint.  
      C-RANGE = "B" +  string(i-linha).
      chPlanilha:Range( C-RANGE):value              = "%" + STRING(ttTotalMensal.fator-utiliz,">>>9.99").
      C-RANGE = "C" +  string(i-linha).
      chPlanilha:Range( C-RANGE):value              = "%" + STRING(ttTotalMensal.eficiencia-media,">>>9.99").
      C-RANGE = "D" +  string(i-linha).
      chPlanilha:Range( C-RANGE):value              =  "%" + STRING(ttTotalMensal.de-classif,">>>9.99").
      C-RANGE = "E" +  string(i-linha).
      chPlanilha:Range( C-RANGE):value              = "%" + STRING(ttTotalMensal.oee,">>>9.99").
      C-RANGE = "F" +  string(i-linha).
      chPlanilha:Range( C-RANGE):value              = "%" + STRING(ttTotalMensal.val-eficien-ctrab,">>>9.99").
    
      
      
      
      ASSIGN i-linha-aux  = i-linha-aux  + 1
             i-linha      = i-linha      + 1.
    
END.

ASSIGN i-linha-fim  = i-linha-aux - 1
       C-FAIXA-GRAF = "A11:" + "F" + STRING(i-linha-fim).
   

     IF VALID-HANDLE(chChart) THEN
        release object      chChart no-error.
     IF VALID-HANDLE(chRange) THEN
        release object      chRange no-error.   
        
    chExcelApplication:COLUMNS("B:B"):ColumnWidth = 35.00.
    chExcelApplication:COLUMNS("C:C"):ColumnWidth = 80.00.     
          
     chRange = chPlanilha:Range(C-FAIXA-GRAF).
     chPlanilha:ChartObjects:Add(50,280,980,600):Activate.
     chExcelApplication:ActiveChart:ChartWizard(chRange,  3, 7, 2, 1, 1, TRUE, "PERF. PRODUCAO MENSAL", "EFICIENCIA", "Capacidade").
     chExcelApplication:ActiveChart:ChartArea:Select.
     chExcelApplication:ActiveChart:Axes(2,1):Select.
     chExcelApplication:ActiveChart:Axes(2,1):DisplayUnit = -2.
     chExcelApplication:ActiveChart:Axes(2,1):TickLabels:NumberFormat = "0%".
     chExcelApplication:ActiveChart:Axes(2,1):HasMajorGridlines = True.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 5.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnitIsAuto = True.
     chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 5.
     chExcelApplication:ActiveChart:Axes(2,1):MaximumScale = 120.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnitIsAuto = True.
     chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:Select.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = "0.00%".
     chExcelApplication:ActiveChart:SeriesCollection(2):DataLabels:Select.
     chExcelApplication:ActiveChart:SeriesCollection(2):DataLabels:NumberFormat = "0.00%".








END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-paradasMensal) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE paradasMensal Procedure 
PROCEDURE paradasMensal :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE VARIABLE deValor  as decimal format ">>>9.99"  NO-UNDO.
DEFINE VARIABLE MesPrint AS CHAR NO-UNDO.
DEFINE VARIABLE lFirst   AS LOG INIT NO NO-UNDO.
i-cont-reg  = 0.
iRow        = 0.
FOR EACH ttListParadas
    BREAK BY ttListParadas.cod-ctrab
          BY ttListParadas.data-inicio:
          
          FIND first ctrab
                 WHERE ctrab.gm-codigo = ttListParadas.gm-codigo
                   AND ctrab.cod-ctrab = ttListParadas.cod-ctrab NO-LOCK NO-ERROR.
          
      FIND FIRST ttParadasMes
            WHERE ttParadasMes.cod-ctrab  = ttListParadas.cod-ctrab
              AND ttParadasMes.i-ano      = YEAR(ttListParadas.data-inicio) 
              AND ttParadasMes.i-mes      = MONTH(ttListParadas.data-inicio)
              AND ttParadasMes.cod-parada = ttListParadas.cod-parada                     NO-ERROR.
       IF NOT AVAIL ttParadasMes THEN DO:
            
           CREATE ttParadasMes.
           ASSIGN ttParadasMes.cod-ctrab         = ttListParadas.cod-ctrab
                  ttParadasMes.i-ano             = YEAR(ttListParadas.data-inicio) 
                  ttParadasMes.i-mes             = MONTH(ttListParadas.data-inicio)
                  ttParadasMes.cod-parada        = ttListParadas.cod-parada
                  ttParadasMes.gm-codigo         = ttListParadas.gm-codigo
                  ttParadasMes.motivo            = ttListParadas.motivo 
                  ttParadasMes.val-eficien-ctrab = ctrab.val-eficien-ctrab.
                
       END.
       ASSIGN 
       ttParadasMes.tempo               = ttParadasMes.tempo + ttListParadas.tempo
       ttParadasMes.qtd                 = ttParadasMes.qtd   + 1.
       
        FIND FIRST ttParadasTotalMes
            WHERE ttParadasTotalMes.i-ano      = YEAR(ttListParadas.data-inicio) 
              AND ttParadasTotalMes.i-mes      = MONTH(ttListParadas.data-inicio)  NO-ERROR.
       IF NOT AVAIL ttParadasTotalMes THEN DO:
          CREATE ttParadasTotalMes.
          ASSIGN
          ttParadasTotalMes.i-ano      = YEAR(ttListParadas.data-inicio) 
          ttParadasTotalMes.i-mes      = MONTH(ttListParadas.data-inicio)
          ttParadasTotalMes.val-eficien-ctrab = ctrab.val-eficien-ctrab
          i-cont-reg             = i-cont-reg + 1.
          
       END.
       ASSIGN 
       ttParadasTotalMes.tempo               = ttParadasTotalMes.tempo + ttListParadas.tempo
       ttParadasTotalMes.qtd                 = ttParadasTotalMes.qtd   + 1.
 END.
 ASSIGN i-linha    = 68
       i-linha-fim = 0
       i-linha-aux = 11
       i-linha-ini = 11.
 FOR EACH ttParadasTotalMes
   BREAK BY ttParadasTotalMes.i-ano
         BY ttParadasTotalMes.i-mes:
      RUN pi-retorna-mes (INPUT ttParadasTotalMes.i-mes, OUTPUT MesPrint).
      C-RANGE = "A" +  string(i-linha-aux).
      chPlanilha:Range( C-RANGE):value              = MesPrint.  
      C-RANGE = "B" +  string(i-linha-aux).
      deValor  = ROUND(ttParadasTotalMes.tempo,2).
      chPlanilha:Range( C-RANGE):value              = deValor.
      iRow = iRow + 1.
 /*     IF  lFirst = no THEN DO:*/
        C-RANGE = "C" +  string(i-linha-aux).
        deValor  = ROUND(ttParadasTotalMes.val-eficien-ctrab,2).
        chPlanilha:Range( C-RANGE):value              = deValor.
        lFirst = TRUE.
 /*     END.*/
      IF  i-cont-reg = 1  THEN DO:
          iRow = iRow + 1.
          i-linha-aux = i-linha-aux + 1.
          RUN pi-retorna-mes (INPUT ttParadasTotalMes.i-mes, OUTPUT MesPrint).
          C-RANGE = "A" +  string(i-linha-aux).
          chPlanilha:Range( C-RANGE):value              = MesPrint.  
          C-RANGE = "B" +  string(i-linha-aux).
          deValor  = ROUND(ttParadasTotalMes.tempo,2).
          chPlanilha:Range( C-RANGE):value              = deValor.
          
      END.
      
       C-RANGE = "A" +  string(i-linha).
      chPlanilha:Range( C-RANGE):value              = MesPrint.  
      C-RANGE = "B" +  string(i-linha).
      chPlanilha:Range( C-RANGE):value              = "%" + STRING(ttParadasTotalMes.tempo,">>>9.99").
      C-RANGE = "C" +  string(i-linha).
      chPlanilha:Range( C-RANGE):value              = "%" + STRING(ttParadasTotalMes.val-eficien-ctrab,">>>9.99").
      
      ASSIGN i-linha-aux  = i-linha-aux  + 1
             i-linha      = i-linha      + 1.  
       
 END.
 
 ASSIGN i-linha-fim  = i-linha-aux - 1
       C-FAIXA-GRAF = "A11:" + "C" + STRING(i-linha-fim).
   

     IF VALID-HANDLE(chChart) THEN
        release object      chChart no-error.
     IF VALID-HANDLE(chRange) THEN
        release object      chRange no-error.   
        
         
          
     chRange = chPlanilha:Range(C-FAIXA-GRAF).
     chPlanilha:ChartObjects:Add(0,165,515,800):Activate.
     chExcelApplication:ActiveChart:ChartWizard(chRange,  3, 7, 2, 1, 1, TRUE, "PARADAS MENSAL", "Paradas", "Capacidade").
     chExcelApplication:ActiveChart:ChartArea:Select.
     chExcelApplication:ActiveChart:Axes(2,1):Select.
     chExcelApplication:ActiveChart:Axes(2,1):DisplayUnit = -2.
     chExcelApplication:ActiveChart:Axes(2,1):TickLabels:NumberFormat = "0%".
     chExcelApplication:ActiveChart:Axes(2,1):HasMajorGridlines = True.
      chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 5.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnitIsAuto = True.
     chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 5.
    chExcelApplication:ActiveChart:Axes(2,1):MaximumScale = 120.
    chExcelApplication:ActiveChart:Axes(2,1):MinorUnitIsAuto = True.
    chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.
    chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:Select.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = "0%".
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = "0,00%".
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = "0,00%".
     chExcelApplication:ActiveChart:SeriesCollection(1):Name = "PARADAS MENSAL".
     
 i-linha      = i-linha      + 2.
  ASSIGN C-RANGE = "A" +  string(i-linha) .
  
      chPlanilha:Range(C-RANGE):MERGE.
      chPlanilha:Range(C-RANGE):value =  'DETALHAMENTO'.
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):font:colorindex     = 2.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      i-linha      = i-linha      + 1.
      chPlanilha:Range(C-RANGE):MERGE.
      chPlanilha:Range(C-RANGE):value =  'MES'.
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):font:colorindex     = 2.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      chPlanilha:Range(C-RANGE):MERGE.
      C-RANGE = "B" +  string(i-linha) .
       chPlanilha:Range(C-RANGE):MERGE.
      chPlanilha:Range(C-RANGE):value =  'CTOS. TRABALHO'.
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):font:colorindex     = 2.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      chPlanilha:Range(C-RANGE):MERGE.
      C-RANGE = "C" +  string(i-linha) .
      chPlanilha:Range(C-RANGE):MERGE.
      chPlanilha:Range(C-RANGE):value =  'MOTIVO'.
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):font:colorindex     = 2.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      chPlanilha:Range(C-RANGE):MERGE.
       C-RANGE = "D" +  string(i-linha) .
      chPlanilha:Range(C-RANGE):MERGE.
      chPlanilha:Range(C-RANGE):value =  'TEMPO PARADA'.
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):font:colorindex     = 2.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      chPlanilha:Range(C-RANGE):MERGE.
       C-RANGE = "E" +  string(i-linha) .
      chPlanilha:Range(C-RANGE):MERGE.
      chPlanilha:Range(C-RANGE):value =  'QTDE. PARADA'.
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):font:colorindex     = 2.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      chPlanilha:Range(C-RANGE):MERGE.
       C-RANGE = "C" +  string(i-linha) .
      chPlanilha:Range(C-RANGE):MERGE.
      chPlanilha:Range(C-RANGE):value =  'EFICIENCIA'.
      chPlanilha:Range(C-RANGE):interior:colorindex = 15.
      chPlanilha:Range(C-RANGE):font:colorindex     = 2.
      chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
      chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
      chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
      chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
      chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
      chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
      chPlanilha:Range(C-RANGE):MERGE.
       i-linha      = i-linha      + 1.
       
/*   chPlanilha:Range("A1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("B1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("C1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("D1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("E1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("F1"):EntireColumn:AutoFit. */
       
 FOR EACH ttParadasMes
     BREAK BY ttParadasMes.cod-ctrab
           BY ttParadasMes.i-ano
           BY ttParadasMes.i-MES
           BY ttParadasMes.cod-parada:
           
        RUN pi-retorna-mes (INPUT  ttParadasMes.i-MES, OUTPUT MesPrint).   
        C-RANGE = "A" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = MesPrint.
        RUN formataCelula(C-RANGE).
        C-RANGE = "B" + STRING(i-linha).
        chPlanilha:Range( C-RANGE):value              = ttParadasMes.cod-ctrab.  
        RUN formataCelula(C-RANGE).
        C-RANGE = "C" + STRING(i-linha).
        chPlanilha:Range( C-RANGE):value              = ttParadasMes.MOTIVO.
        RUN formataCelula(C-RANGE).
        C-RANGE = "D" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = "%" + STRING(ttDados.fator-utiliz,">>>9.99").
        RUN formataCelula(C-RANGE).
        C-RANGE = "E" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  ttParadasMes.qtd  .
        RUN formataCelula(C-RANGE).
        C-RANGE = "F" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = "%" + STRING(ttParadasMes.val-eficien-ctrab,">>>9.99").
        i-linha = i-linha + 1.
 END.
/*   chPlanilha:Range("A1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("B1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("C1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("D1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("E1"):EntireColumn:AutoFit. */
/*   chPlanilha:Range("F1"):EntireColumn:AutoFit. */
 

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-pi-busca-linha) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-busca-linha Procedure 
PROCEDURE pi-busca-linha :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    def  input param p-hora               as char  no-undo.
    def output param p-linha              as int no-undo.

    CASE p-hora:
       WHEN '04' THEN ASSIGN p-linha = 27.
       WHEN '05' THEN ASSIGN p-linha = 26.
       WHEN '06' THEN ASSIGN p-linha = 25.
       WHEN '07' THEN ASSIGN p-linha = 24.
       WHEN '08' THEN ASSIGN p-linha = 23.
       WHEN '09' THEN ASSIGN p-linha = 22.
       WHEN '10' THEN ASSIGN p-linha = 21.
       WHEN '11' THEN ASSIGN p-linha = 20.
       WHEN '12' THEN ASSIGN p-linha = 19.
       WHEN '13' THEN ASSIGN p-linha = 18.
       WHEN '14' THEN ASSIGN p-linha = 17.
       WHEN '15' THEN ASSIGN p-linha = 16.
       WHEN '16' THEN ASSIGN p-linha = 15.
       WHEN '17' THEN ASSIGN p-linha = 14.
       WHEN '18' THEN ASSIGN p-linha = 13.
       WHEN '19' THEN ASSIGN p-linha = 12.
       WHEN '20' THEN ASSIGN p-linha = 11.
       WHEN '21' THEN ASSIGN p-linha = 10.
       WHEN '22' THEN ASSIGN p-linha = 9.
       WHEN '23' THEN ASSIGN p-linha = 8.
       WHEN '24' THEN ASSIGN p-linha = 7.
    END CASE.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-pi-converte-data-segs-valor) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-converte-data-segs-valor Procedure 
PROCEDURE pi-converte-data-segs-valor :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    def Input param p-dat-refer-1           as date format "99/99/9999"     no-undo.
    def Input param p-qtd-segs-refer-1      as dec  format "->>>>,>>9.9999" decimals 4 NO-UNDO.
    def output param p-val-refer-perf-capac as dec  format "9999999999999"  decimals 0 no-undo.

    assign p-val-refer-perf-capac = decimal(string(year(p-dat-refer-1),"9999") +
                                            string(month(p-dat-refer-1),"99")  +
                                            string(day(p-dat-refer-1),"99")    +
                                            string(p-qtd-segs-refer-1,"99999")).
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-pi-retorna-mes) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-retorna-mes Procedure 
PROCEDURE pi-retorna-mes :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER p-mes     AS INT NO-UNDO.
DEFINE OUTPUT PARAMETER p-retorno AS CHAR FORMAT "15" NO-UNDO.

CASE p-mes:
    WHEN 1 THEN DO:
        ASSIGN p-retorno = "Janeiro" .
    END.
    WHEN 2 THEN DO:
        ASSIGN p-retorno = "Fevereiro" .
    END.
    WHEN 3 THEN DO:
        ASSIGN p-retorno = "Mar?o" .
    END.
    WHEN 4 THEN DO:
        ASSIGN p-retorno = "Abril" .
    END.
    WHEN 5 THEN DO:
        ASSIGN p-retorno = "Maio" .
    END.
    WHEN 6 THEN DO:
        ASSIGN p-retorno = "Junho" .
    END.
    WHEN 7 THEN DO:
        ASSIGN p-retorno = "Julho" .
    END.
    WHEN 8 THEN DO:
        ASSIGN p-retorno = "Agosto" .
    END.
    WHEN 9 THEN DO:
        ASSIGN p-retorno = "Setembro" .
    END.
    WHEN 10 THEN DO:
        ASSIGN p-retorno = "Outubo" .
    END.
    WHEN 11 THEN DO:
        ASSIGN p-retorno = "Novembro" .
    END.
    WHEN 12 THEN DO:
        ASSIGN p-retorno = "Dezembro" .
    END.
END CASE.
        
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-pi-sec-to-formatted-time) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-sec-to-formatted-time Procedure 
PROCEDURE pi-sec-to-formatted-time :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    def  input param p-num-seconds        as integer   format ">>>,>>9"  no-undo.
    def output param p-hra-formatted-time as character format "99:99:99" no-undo.

    assign p-hra-formatted-time = replace(string(p-num-seconds,"hh:mm:ss":U),":","").

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-piInicial) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE piInicial Procedure 
PROCEDURE piInicial :
/*------------------------------------------------------------------------------
  Purpose:     piInicial
  Parameters:  <none>
  Notes:       Define os valores que ser∆o mostrados o cabe?alho e rodap?
  ------------------------------------------------------------------------------*/
  assign c-programa = "SFC/SFCP0010RP"
         c-versao   = "2.00"
         c-revisao  = "001".
  /** Define o destino do arquivo a ser gerado **/
  
  
  
  
  /** Busca os par?metros criados na interface gr‡fica **/
  find first tt-param no-lock no-error.
  
  /** Busca empresa padr∆o **/
  for first param-global fields(grupo) no-lock:
    assign c-empresa = param-global.grupo.
  end.
  
  
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-piPrincipal) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE piPrincipal Procedure 
PROCEDURE piPrincipal :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
  IF c-dir-spool-servid-exec = ''   THEN DO:
    IF  NOT VALID-HANDLE (h-acomp) THEN
    RUN utp/ut-acomp.p PERSISTENT SET h-acomp.
    
    IF  VALID-HANDLE (h-acomp) THEN  DO:
      RUN pi-inicializar IN h-acomp (INPUT 'GeraÁ„o dados OEE').
      RUN pi-acompanhar  IN h-acomp (INPUT 'Calculando...').
    END.
    
  END.
  
  
  /*********************************************************************
  Neste espa?o deve-se colocar a l¢gica para busca as informa?‰es e
  tamb?m para fazer o display, pode-se inclui a chamada de procedures
  ou fazer a l¢gica neste espa?o.
  *********************************************************************/
  RUN displayDados.
  /*********************************************************************
  Fim do espa?o para l¢gica de c‡lculo e display das informa?‰es
  *********************************************************************/
  /** Mostra par?metros selecionados **/
  run displayParametros in this-procedure.
  
  
  IF VALID-HANDLE (h-acomp) THEN
  RUN pi-finalizar IN h-acomp.
  
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-putValores) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE putValores Procedure 
PROCEDURE putValores :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEF VAR media-oee AS DECIMAL.
DEF VAR media-performance AS DECIMAL.
DEF VAR media-qualidade AS DECIMAL.
DEF VAR media-disponibilidade AS DECIMAL.
DEF VAR contador AS INT.

//RENAN
    
ASSIGN //i-linha      = 108 //69
       i-linha-ini  = 15
       i-linha-fim  = 0
       i-linha-aux  = 15
       iRow         = 0
       i-cont-reg   = 1
       contador = 0.

        C-RANGE   = "B" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = ''.  
        C-RANGE   = "C" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = 0.
        C-RANGE = "E" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = ''.  
        C-RANGE = "F" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              =  0.
        C-RANGE = "H" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = ''.  
        C-RANGE = "I" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = 0.
        C-RANGE = "K" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              = ''.  
        C-RANGE = "L" +  string(i-linha-aux).
        chPlanilha:Range( C-RANGE):value              =  0.

          
 ASSIGN 
       
       i-linha-aux  = i-linha-aux  + 1.       
        
FOR EACH ttDados
    WHERE ttDados.cod-ctrab = ttCentro.cod-ctrab
    BREAK BY ttDados.data-inicio
          BY ttDados.cod-ctrab:
    
    //RUN dadosRelatorio NO-ERROR.
    RUN dadosGrafico.
    RUN planoAcao.

    //i-linha   = i-linha   + 1.
    i-linha-aux =  i-linha-aux + 1.
    contador = contador + 1.

    ASSIGN media-oee = media-oee + DEC(ttDados.fator-utiliz * ttDados.eficiencia-media * ttDados.de-classif)
           media-performance = media-performance + DEC(ttDados.eficiencia-media)
           media-qualidade = media-qualidade + DEC(ttDados.de-classif)
        .  media-disponibilidade = media-disponibilidade + DEC(ttDados.fator-utiliz).


END.

    ASSIGN media-oee = (media-oee / contador)
           media-performance = (media-performance / contador)
           media-qualidade = (media-qualidade / contador)
           media-disponibilidade = (media-disponibilidade / contador).

    
    //INSERE MEDIA OEE NO GRAFICO
    C-RANGE   = "B" +  string(i-linha-aux).
    chPlanilha:Range( C-RANGE):VALUE = "MÇdia". 
    C-RANGE   = "C" +  string(i-linha-aux).
    chPlanilha:Range( C-RANGE):VALUE = media-oee * 100. 

    //INSERE MEDIA PERFORMANCE NO GRAFICO
    C-RANGE   = "E" +  string(i-linha-aux).
    chPlanilha:Range( C-RANGE):VALUE = "MÇdia". 
    C-RANGE   = "F" +  string(i-linha-aux).
    chPlanilha:Range( C-RANGE):VALUE = media-performance * 100. 

    //INSERE MEDIA QUALIDADE NO GRAFICO
    C-RANGE   = "K" +  string(i-linha-aux).
    chPlanilha:Range( C-RANGE):VALUE = "MÇdia". 
    C-RANGE   = "L" +  string(i-linha-aux).
    chPlanilha:Range( C-RANGE):VALUE = media-qualidade * 100. 

    //INSERE MEDIA DISPONIBILIDADE NO GRAFICO
    C-RANGE   = "H" +  string(i-linha-aux).
    chPlanilha:Range( C-RANGE):VALUE = "MÇdia". 
    C-RANGE   = "I" +  string(i-linha-aux).
    chPlanilha:Range( C-RANGE):VALUE = media-disponibilidade * 100. 


ASSIGN i-linha-aux = i-linha-aux + 1
       i-linha-fim  = i-linha-aux - 1
       C-RANGE-GRAF = "B" + STRING(i-linha-ini) + ":" + "D" + STRING(i-linha-fim).

     
   

     IF VALID-HANDLE(chChart) THEN
        release object      chChart no-error.
     IF VALID-HANDLE(chRange) THEN
        release object      chRange no-error.   

     //GRAFICO OEE MODIFICADO
        
     ASSIGN C-FAIXA-GRAF =   "B15:" +   "C" + STRING(i-linha-fim).
     chRange = chPlanilha:Range(C-FAIXA-GRAF).
     chPlanilha:ChartObjects:Add(15,170,420,400):Activate.
     chExcelApplication:ActiveChart:ChartWizard(chRange, 3, 7, 2, 1, 1, TRUE, "OEE", "", "").
     chExcelApplication:ActiveChart:ChartArea:Select.
     chExcelApplication:ActiveChart:Axes(2,1):Select.
     chExcelApplication:ActiveChart:Axes(2,1):DisplayUnit = -2.
     chExcelApplication:ActiveChart:Axes(2,1):TickLabels:NumberFormat = "0%".
     chExcelApplication:ActiveChart:Axes(2,1):HasMajorGridlines = True.
     chExcelApplication:ActiveChart:Axes(2,1):MinimumScale = 0.
     chExcelApplication:ActiveChart:Axes(2,1):MaximumScale = 100.
     chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 1.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:Select.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = "#0.#0%".
     chExcelApplication:ActiveChart:SeriesCollection(1):Name = "Oee".
             
     IF VALID-HANDLE(chChart) THEN
        release object      chChart no-error.
     IF VALID-HANDLE(chRange) THEN
        release object      chRange no-error.   

     //GRAFICO PERFORMANCE MODIFICADO
        
     ASSIGN C-FAIXA-GRAF =   "E15:" +   "F" + STRING(i-linha-fim).
     chRange = chPlanilha:Range(C-FAIXA-GRAF).
     chPlanilha:ChartObjects:Add(440,170,570,400):Activate.
     chExcelApplication:ActiveChart:ChartWizard(chRange, 3, 7, 2, 1, 1, TRUE, "PERFORMANCE", "", "").
     /* 3, 6, 1, 1, 1 */
     chExcelApplication:ActiveChart:ChartArea:Select.
     chExcelApplication:ActiveChart:Axes(2,1):Select.
     chExcelApplication:ActiveChart:Axes(2,1):DisplayUnit = -2.
     chExcelApplication:ActiveChart:Axes(2,1):TickLabels:NumberFormat = "0%".
     chExcelApplication:ActiveChart:Axes(2,1):HasMajorGridlines = True.
     chExcelApplication:ActiveChart:Axes(2,1):MinimumScale = 0.
     chExcelApplication:ActiveChart:Axes(2,1):MaximumScale = 100.
     chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 1.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:Select.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = "#0.#0%".
     chExcelApplication:ActiveChart:SeriesCollection(1):Name = "Performance".
          
     IF VALID-HANDLE(chChart) THEN
        release object      chChart no-error.
     IF VALID-HANDLE(chRange) THEN
        release object      chRange no-error.   

     //GRAFICO DISPONIBILIDADE
        
     ASSIGN C-FAIXA-GRAF =   "H15:" +   "I" + STRING(i-linha-fim).
     chRange = chPlanilha:Range(C-FAIXA-GRAF).
     chPlanilha:ChartObjects:Add(15,580,420,400):Activate.
     chExcelApplication:ActiveChart:ChartWizard(chRange,  3, 7, 2, 1, 1, TRUE, "DISPONIBILIDADE", "", "").
     chExcelApplication:ActiveChart:ChartArea:Select.
     chExcelApplication:ActiveChart:Axes(2,1):Select.
     chExcelApplication:ActiveChart:Axes(2,1):DisplayUnit = -2.
     chExcelApplication:ActiveChart:Axes(2,1):TickLabels:NumberFormat = "0%".
     chExcelApplication:ActiveChart:Axes(2,1):HasMajorGridlines = FALSE.
     chExcelApplication:ActiveChart:Axes(2,1):MinimumScale = 0.
     chExcelApplication:ActiveChart:Axes(2,1):MaximumScale = 100.
     chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 1.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:Select.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = "#0.#0%".
     chExcelApplication:ActiveChart:SeriesCollection(1):Name = "Disponibilidade".
     
    
     
         
     IF VALID-HANDLE(chChart) THEN
        release object      chChart no-error.
     IF VALID-HANDLE(chRange) THEN
        release object      chRange no-error.   

     //GRAFICO QUALIDADE
        
     ASSIGN C-FAIXA-GRAF =   "K15:" +   "L" + STRING(i-linha-fim).
     chRange = chPlanilha:Range(C-FAIXA-GRAF).
     chPlanilha:ChartObjects:Add(440,580,570,400):Activate.
     chExcelApplication:ActiveChart:ChartWizard(chRange,  3, 7, 2, 1, 1, TRUE, "QUALIDADE", "", "").
     chExcelApplication:ActiveChart:ChartArea:Select.
     chExcelApplication:ActiveChart:Axes(2,1):Select.
     chExcelApplication:ActiveChart:Axes(2,1):DisplayUnit = -2.
     chExcelApplication:ActiveChart:Axes(2,1):TickLabels:NumberFormat = "0%".
     chExcelApplication:ActiveChart:Axes(2,1):HasMajorGridlines = True.
     chExcelApplication:ActiveChart:Axes(2,1):MinimumScale = 0.
     chExcelApplication:ActiveChart:Axes(2,1):MaximumScale = 100.
     chExcelApplication:ActiveChart:Axes(2,1):MajorUnit = 5.
     chExcelApplication:ActiveChart:Axes(2,1):MinorUnit = 1.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:Select.
     chExcelApplication:ActiveChart:SeriesCollection(1):DataLabels:NumberFormat = "#0.#0%".
     chExcelApplication:ActiveChart:SeriesCollection(1):Name = "Qualidade".
         
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-relatorioParadas) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE relatorioParadas Procedure 


//asas lista de paradas


PROCEDURE relatorioParadas :
/*------------------------------------------------------------------------------
  Purpose:
  Parameters:  <none>
  Notes:
  ------------------------------------------------------------------------------*/
  DEF VAR data-parada AS DATE.
  DEF VAR codigo-operacao AS CHAR.
  DEF VAR descricao-operacao AS CHAR.
  DEF VAR codigo-parada AS CHAR.
  DEF VAR motivo-parada AS CHAR.

    
  assign   i_cont                 = 0
           de_parada_geral        = 0
           de_parada_total_geral  = 0
           de_parada_alt_efic     = 0 
           de_parada_n_alt_efic   = 0
           i-linha = 11.


    FIND FIRST ctrab
         WHERE ctrab.gm-codigo = ttCentro.gm-codigo
           AND ctrab.cod-ctrab = ttCentro.cod-ctrab NO-LOCK NO-ERROR.

    FOR EACH rep-parada-ctrab no-lock
       where  (rep-parada-ctrab.dat-inic-parada >= data-ini-paradas AND rep-parada-ctrab.dat-fim-parada <= data-fim-paradas)
         AND  rep-parada-ctrab.cod-ctrab = v-ctrab,

         FIRST motiv-parada where motiv-parada.cod-parada = rep-parada-ctrab.cod-parada
         break by rep-parada-ctrab.dat-inic-parada
               by rep-parada-ctrab.cod-parada:
   

         if first-of(dat-inic-parada) then DO: 
                assign de_parada_alt_efic   = 0
                       de_parada_n_alt_efic = 0.

                
         END.

         //Verifica o item que estava produzindo no momento da parada que altera a eficiencia
         FOR FIRST rep-oper-ctrab 
             WHERE rep-oper-ctrab.cod-ctrab = rep-parada-ctrab.cod-ctrab
               AND rep-oper-ctrab.dat-inic-reporte = rep-parada-ctrab.dat-inic-parada
               AND rep-oper-ctrab.qtd-segs-inic-reporte <= rep-parada-ctrab.qtd-segs-inic
               AND rep-oper-ctrab.qtd-segs-fim-reporte >= rep-parada-ctrab.qtd-segs-fim,

               FIRST split-operac WHERE split-operac.nr-ord-prod = rep-oper-ctrab.nr-ord-prod.
                     
              ASSIGN codigo-operacao = split-operac.it-codigo
                     descricao-operacao = oper-ord.descricao.
               
         END.



         IF FIRST-OF(rep-parada-ctrab.cod-parada) THEN DO:

                        ASSIGN data-parada = rep-parada-ctrab.dat-inic-parada
                               codigo-parada = rep-parada-ctrab.cod-parada
                               motivo-parada = motiv-parada.des-parada.
                               
         END.
 
         assign i_cont = i_cont + 1 
                de_parada_geral        = de_parada_geral        + rep-parada-ctrab.qtd-tempo-parada
                de_parada_total_geral  = de_parada_total_geral  + rep-parada-ctrab.qtd-tempo-parada
                de_parada_alt_efic     = if motiv-parada.log-alter-eficien = yes then (de_parada_alt_efic   + rep-parada-ctrab.qtd-tempo-parada) else (de_parada_alt_efic)   
                de_parada_n_alt_efic   = if motiv-parada.log-alter-eficien = no  then (de_parada_n_alt_efic + rep-parada-ctrab.qtd-tempo-parada) else (de_parada_n_alt_efic).
        
        
 
         if last-of(rep-parada-ctrab.cod-parada) then do:
            C-RANGE = "B" + STRING(i-linha).
            chPlanilha:Range( C-RANGE):value              = data-parada.
            chPlanilha:Range(C-RANGE):interior:colorindex = 2.
            chPlanilha:Range(C-RANGE):font:colorindex     = 1.
            chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
            chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
            chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
            chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
            chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
            chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
            C-RANGE = "C" + STRING(i-linha).
            chPlanilha:Range( C-RANGE):value              = codigo-operacao.
            chPlanilha:Range(C-RANGE):interior:colorindex = 2.
            chPlanilha:Range(C-RANGE):font:colorindex     = 1.
            chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
            chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
            chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
            chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
            chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
            chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
            C-RANGE = "D" + STRING(i-linha).
            chPlanilha:Range( C-RANGE):value              = descricao-operacao.
            chPlanilha:Range(C-RANGE):interior:colorindex = 2.
            chPlanilha:Range(C-RANGE):font:colorindex     = 1.
            chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
            chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
            chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
            chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
            chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
            chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
            C-RANGE = "E" + STRING(i-linha).
            chPlanilha:Range( C-RANGE):value              = codigo-parada.
            chPlanilha:Range(C-RANGE):interior:colorindex = 2.
            chPlanilha:Range(C-RANGE):font:colorindex     = 1.
            chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
            chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
            chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
            chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
            chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
            chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
            C-RANGE = "F" + STRING(i-linha).
            chPlanilha:Range( C-RANGE):value              = motivo-parada.
            chPlanilha:Range(C-RANGE):interior:colorindex = 2.
            chPlanilha:Range(C-RANGE):font:colorindex     = 1.
            chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
            chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
            chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
            chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
            chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
            chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
            C-RANGE = "G" + STRING(i-linha).
            chPlanilha:Range( C-RANGE):value              = i_cont.
            chPlanilha:Range(C-RANGE):interior:colorindex = 2.
            chPlanilha:Range(C-RANGE):font:colorindex     = 1.
            chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
            chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
            chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
            chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
            chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
            chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
            C-RANGE = "H" + STRING(i-linha).
           
            chPlanilha:Range( C-RANGE):value              = DEC(de_parada_geral). //STRING(rep-parada-ctrab.qtd-tempo-parada,">>>9.99") .
            chPlanilha:Range(C-RANGE):interior:colorindex = 2.
            chPlanilha:Range(C-RANGE):font:colorindex     = 1.
            chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
            chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
            chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
            chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
            chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
            chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
            C-RANGE = "I" + STRING(i-linha).
            chPlanilha:Range( C-RANGE):value              = "N/A". //,">>>9.99") + "%".
            chPlanilha:Range(C-RANGE):interior:colorindex = 2.
            chPlanilha:Range(C-RANGE):font:colorindex     = 1.
            chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
            chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
            chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
            chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
            chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
            chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
            
            ASSIGN i-linha  = i-linha + 1.
             
            assign  i_cont          = 0
                    de_parada_geral = 0
                    codigo-operacao = ""
                    descricao-operacao = ""
                    codigo-parada = ""
                    motivo-parada = "".
         
         END.
  
    END. //  FOR EACH rep-parada-ctrab no-lock

        C-RANGE = "G" + STRING(i-linha).
        chPlanilha:Range( C-RANGE):value              = "Tempo Total:" . 
        chPlanilha:Range(C-RANGE):interior:colorindex = 4.
        chPlanilha:Range(C-RANGE):font:colorindex     = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
  
        C-RANGE = "H" + STRING(i-linha).
         chPlanilha:Range( C-RANGE):value              = STRING(de_parada_total_geral,">>>9.99") .
        chPlanilha:Range(C-RANGE):interior:colorindex = 4.
        chPlanilha:Range(C-RANGE):font:colorindex     = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
        ASSIGN i-linha  = 0.
        
  
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-rodapeMensal) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE rodapeMensal Procedure 
PROCEDURE rodapeMensal :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
        C-RANGE = "A" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value              = "Media".  
        chPlanilha:Range(C-RANGE):interior:colorindex = 15.
        chPlanilha:Range(C-RANGE):font:colorindex     = 2.
        chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
  
         C-RANGE = "C" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               = v-total1.  
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "D" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):VALUE               = v-total2.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "E" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):VALUE              = v-total3.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "F" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =  v-total4.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "G" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):VALUE               = v-total5.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "H" + STRING(i-linha).
        chPlanilha:Range(C-RANGE):value               =   v-total6.
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "I" + STRING(i-linha).
        chPlanilha:Range( C-RANGE):value               =  "%" + STRING(v-total7,">>>9.99").
        
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        C-RANGE = "J" + STRING(i-linha).
        
        chPlanilha:Range(C-RANGE):value        =  "%" + STRING(v-total8,">>>9.99").
        chPlanilha:Range(C-RANGE):interior:colorindex = 2.
        chPlanilha:Range(C-RANGE):font:colorindex     = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
        C-RANGE = "K" + STRING(i-linha).
        
        chPlanilha:Range(C-RANGE):value               =  "%" + STRING(v-total9,">>>9.99") .
        chPlanilha:Range(C-RANGE):interior:colorindex = 2.
        chPlanilha:Range(C-RANGE):font:colorindex     = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight   = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight   = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight   = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight  = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight  = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight  = 2. /* Direita */
         C-RANGE = "L" + STRING(i-linha).
         
        chPlanilha:Range(C-RANGE):value               =  "%" + STRING(v-total10,">>>9.99").
        chPlanilha:Range(C-RANGE):interior:colorindex  = 2.
        chPlanilha:Range(C-RANGE):font:colorindex      = 1.
        chPlanilha:Range(C-RANGE):Borders(7):Weight    = 2. /* Esquerda */
        chPlanilha:Range(C-RANGE):Borders(8):Weight    = 2. /* Superior */
        chPlanilha:Range(C-RANGE):Borders(9):Weight    = 2. /* Inferior */
        chPlanilha:Range(C-RANGE):Borders(11):Weight   = 2. /* Interna Vertical */
        chPlanilha:Range(C-RANGE):Borders(12):Weight   = 2. /* Interna Horizontal */
        chPlanilha:Range(C-RANGE):Borders(10):Weight   = 2. /* Direita */
        
        
  return "OK":U.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-totalizaValores) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE totalizaValores Procedure 
PROCEDURE totalizaValores :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF
&IF DEFINED(EXCLUDE-criaTurno) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE criaTurno Procedure 
PROCEDURE criaTurno :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
              

 FOR EACH turno-semana  NO-LOCK 
    WHERE turno-semana.cod-model-turno = 'Turno 1':
           FIND FIRST ttTurno
              WHERE ttTurno.cod-model-turno =  turno-semana.cod-model-turno
               AND  ttTurno.dia       = 1 NO-ERROR.
          IF NOT AVAIL ttTurno THEN DO:
             CREATE ttTurno.
             ASSIGN ttTurno.cod-model-turno =  turno-semana.cod-model-turno
                    ttTurno.dia                  = 1
                    ttTurno.qtd-tempo-util-dia-1 = turno-semana.qtd-tempo-util-dia[1].
          END.
          ELSE do:
             ASSIGN ttTurno.qtd-tempo-util-dia-1 = ttTurno.qtd-tempo-util-dia-1 + turno-semana.qtd-tempo-util-dia[1].
                 
          END.
          
          FIND FIRST ttTurno
              WHERE ttTurno.cod-model-turno =  turno-semana.cod-model-turno
               AND  ttTurno.dia       = 2 NO-ERROR.
          IF NOT AVAIL ttTurno THEN DO:
             CREATE ttTurno.
             ASSIGN ttTurno.cod-model-turno =  turno-semana.cod-model-turno
                    ttTurno.dia                  = 2
                    ttTurno.qtd-tempo-util-dia-2 = turno-semana.qtd-tempo-util-dia[2].
          END.
          ELSE do:
             ASSIGN ttTurno.qtd-tempo-util-dia-2 = ttTurno.qtd-tempo-util-dia-2 + turno-semana.qtd-tempo-util-dia[2].
          END.
          FIND FIRST ttTurno
              WHERE ttTurno.cod-model-turno =  turno-semana.cod-model-turno 
               AND  ttTurno.dia       = 3 NO-ERROR.
          IF NOT AVAIL ttTurno THEN DO:
             CREATE ttTurno.
             ASSIGN ttTurno.cod-model-turno =  turno-semana.cod-model-turno
                    ttTurno.dia                  = 3
                    ttTurno.qtd-tempo-util-dia-3 = turno-semana.qtd-tempo-util-dia[3].
          END.
          ELSE do:
             ASSIGN ttTurno.qtd-tempo-util-dia-3 = ttTurno.qtd-tempo-util-dia-3 + turno-semana.qtd-tempo-util-dia[3].
          END.
          FIND FIRST ttTurno
              WHERE ttTurno.cod-model-turno =  turno-semana.cod-model-turno 
               AND  ttTurno.dia       = 4 NO-ERROR.
          IF NOT AVAIL ttTurno THEN DO:
             CREATE ttTurno.
             ASSIGN ttTurno.cod-model-turno =  turno-semana.cod-model-turno
                    ttTurno.dia                  = 4
                    ttTurno.qtd-tempo-util-dia-4 = turno-semana.qtd-tempo-util-dia[4].
          END.
          ELSE do:
             ASSIGN ttTurno.qtd-tempo-util-dia-4 = ttTurno.qtd-tempo-util-dia-4 + turno-semana.qtd-tempo-util-dia[4].
          END.
          FIND FIRST ttTurno
              WHERE ttTurno.cod-model-turno =  turno-semana.cod-model-turno 
               AND  ttTurno.dia       = 5 NO-ERROR.
          IF NOT AVAIL ttTurno THEN DO:
             CREATE ttTurno.
             ASSIGN ttTurno.cod-model-turno =  turno-semana.cod-model-turno
                    ttTurno.dia                  = 5
                    ttTurno.qtd-tempo-util-dia-5 = turno-semana.qtd-tempo-util-dia[5].
          END.
          ELSE do:
             ASSIGN ttTurno.qtd-tempo-util-dia-5 = ttTurno.qtd-tempo-util-dia-5 + turno-semana.qtd-tempo-util-dia[5].
          END.
          FIND FIRST ttTurno
              WHERE ttTurno.cod-model-turno =  turno-semana.cod-model-turno 
               AND  ttTurno.dia       = 6 NO-ERROR.
          IF NOT AVAIL ttTurno THEN DO:
             CREATE ttTurno.
             ASSIGN ttTurno.cod-model-turno =  turno-semana.cod-model-turno
                    ttTurno.dia                  = 6
                    ttTurno.qtd-tempo-util-dia-6 = turno-semana.qtd-tempo-util-dia[6].
          END.
          ELSE do:
             ASSIGN ttTurno.qtd-tempo-util-dia-6 = ttTurno.qtd-tempo-util-dia-6 + turno-semana.qtd-tempo-util-dia[6].
          END.
          FIND FIRST ttTurno
              WHERE ttTurno.cod-model-turno =  turno-semana.cod-model-turno 
               AND  ttTurno.dia       = 7 NO-ERROR.
          IF NOT AVAIL ttTurno THEN DO:
             CREATE ttTurno.
             ASSIGN ttTurno.cod-model-turno =  turno-semana.cod-model-turno
                    ttTurno.dia                  = 7
                    ttTurno.qtd-tempo-util-dia-7 = turno-semana.qtd-tempo-util-dia[7].
          END.
          ELSE do:
             ASSIGN ttTurno.qtd-tempo-util-dia-7 = ttTurno.qtd-tempo-util-dia-7 + turno-semana.qtd-tempo-util-dia[7].
          END.
          
           
       END.
       
    
       
END PROCEDURE.


PROCEDURE PI_CALC_PARADAS:
 

    assign i_cont                 = 0
           de_parada_geral        = 0
           de_parada_total_geral  = 0
           de_parada_alt_efic     = 0 
           de_parada_n_alt_efic   = 0.


    FOR EACH rep-parada-ctrab no-lock
       where  (rep-parada-ctrab.dat-inic-parada >= data-ini AND rep-parada-ctrab.dat-fim-parada <= data-fim)
         AND  rep-parada-ctrab.cod-ctrab = v-ctrab
           ,
                  
         FIRST motiv-parada where motiv-parada.cod-parada = rep-parada-ctrab.cod-parada
         break by rep-parada-ctrab.dat-inic-parada
               by rep-parada-ctrab.cod-parada:
               
               
         if first-of(dat-inic-parada) then  
                assign de_parada_alt_efic   = 0
                       de_parada_n_alt_efic = 0.
 
         assign i_cont = i_cont + 1 
                de_parada_geral        = de_parada_geral        + rep-parada-ctrab.qtd-tempo-parada
                de_parada_total_geral  = de_parada_total_geral  + rep-parada-ctrab.qtd-tempo-parada
                de_parada_alt_efic     = if motiv-parada.log-alter-eficien = yes then (de_parada_alt_efic   + rep-parada-ctrab.qtd-tempo-parada) else (de_parada_alt_efic)   
                de_parada_n_alt_efic   = if motiv-parada.log-alter-eficien = no  then (de_parada_n_alt_efic + rep-parada-ctrab.qtd-tempo-parada) else (de_parada_n_alt_efic)  .
 
         if last-of(rep-parada-ctrab.cod-parada)  
            then assign i_cont          = 0
                        de_parada_geral = 0.
  
    END. //  FOR EACH rep-parada-ctrab no-lock
 

END. //PROCEDURE PI_CALC_PARADAS

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

