DEFINE NEW GLOBAL SHARED VAR data-ini AS DATE NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR data-fim AS DATE NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR centro AS CHAR NO-UNDO.
DEF VAR data-atu AS DATE.
DEF VAR cont-temp-prod AS DECIMAL.
DEF VAR cont-pc-boas AS DECIMAL.
DEF VAR cont-pc-ruins AS DECIMAL.
DEF VAR cont-turno AS DECIMAL.
DEF VAR cont-tempo-par-alt AS DECIMAL.
DEF VAR cont-tempo-par-nao-alt AS DECIMAL.
DEF VAR cont-dias-mes AS INT.
DEF VAR cont-prod-teorica AS DECIMAL.

DEFINE NEW GLOBAL SHARED TEMP-TABLE tt-oee NO-UNDO
    FIELD data AS DATE LABEL "Data Apontamento"
    FIELD turno AS DECIMAL LABEL "Tempo Programado"
    FIELD cod-item LIKE split-operac.it-codigo LABEL "Item" 
    FIELD temp-prod AS DECIMAL LABEL "Tempo Produzindo"
    FIELD prod-teorica AS DECIMAL LABEL "Produ‡Æo Teorica"
    FIELD prod-real AS INT LABEL "Produ‡Æo Real"
    FIELD pc-boas AS INT LABEL "Pe‡as Boas"
    FIELD pc-ruins AS INT LABEL "Pe‡as Ruins"
    FIELD pc-boas-ruins AS INT LABEL "Pe‡as Boas + Ruins"
    FIELD temp-par-alt AS DECIMAL
    FIELD temp-par-nao-alt AS DECIMAL
    FIELD oee-disp AS DECIMAL LABEL "Disponibilidade"
    FIELD oee-perf AS DECIMAL LABEL "Performance"
    FIELD oee-quali AS DECIMAL LABEL "Qualidade"
    FIELD oee-percent AS DECIMAL LABEL "Percentual OEE"
    FIELD mes AS INT.

DEFINE NEW GLOBAL SHARED TEMP-TABLE tt-parada NO-UNDO
    FIELD cod-ctrab LIKE rep-parada-ctrab.cod-ctrab LABEL "teste"
    FIELD cod-parada AS CHAR
    FIELD des-parada LIKE motiv-parada.des-parada
    FIELD log-alter-eficien LIKE motiv-parada.log-alter-eficien
    FIELD dat-inic-parada LIKE rep-parada-ctrab.dat-inic-parada
    FIELD qtd-segs-inic LIKE rep-parada-ctrab.qtd-segs-inic
    FIELD dat-fim-parada LIKE rep-parada-ctrab.dat-fim-parada
    FIELD qtd-segs-fim LIKE rep-parada-ctrab.qtd-segs-fim
    FIELD it-codigo LIKE split-operac.it-codigo
    FIELD inicio AS DATETIME
    FIELD fim AS DATETIME
    FIELD tempo AS DECIMAL.

ASSIGN cont-temp-prod = 0.
/*        data-ini = DATE("07/01/2021") */
/*        data-fim = DATE("07/01/2021") */
/*        centro = "4120101SP09".       */

EMPTY TEMP-TABLE tt-oee.

//CAPTURA DE PARADAS

        FOR EACH rep-parada-ctrab NO-LOCK
           WHERE rep-parada-ctrab.cod-ctrab = centro
             AND rep-parada-ctrab.dat-inic-parada >= data-ini
             AND rep-parada-ctrab.dat-fim-parada <= data-fim,

             FIRST motiv-parada WHERE motiv-parada.cod-parada = rep-parada-ctrab.cod-parada.

             CREATE tt-parada.
             ASSIGN tt-parada.cod-ctrab = string(rep-parada-ctrab.cod-ctrab)
                    tt-parada.cod-parada = rep-parada-ctrab.cod-parada
                    tt-parada.des-parada = motiv-parada.des-parada
                    tt-parada.log-alter-eficien = motiv-parada.log-alter-eficien
                    tt-parada.inicio = DATETIME(rep-parada-ctrab.dat-inic-parada, INT(rep-parada-ctrab.qtd-segs-inic * 1000))
                    tt-parada.fim = DATETIME(rep-parada-ctrab.dat-fim-parada, INT(rep-parada-ctrab.qtd-segs-fim * 1000))
                    tt-parada.dat-inic-parada = rep-parada-ctrab.dat-inic-parada
                    tt-parada.qtd-segs-inic = rep-parada-ctrab.qtd-segs-inic
                    tt-parada.dat-fim-parada = rep-parada-ctrab.dat-fim-parada
                    tt-parada.qtd-segs-fim = rep-parada-ctrab.qtd-segs-fim
                    tt-parada.tempo = rep-parada-ctrab.qtd-tempo-parada + rep-parada-ctrab.qtd-tempo-ext.

             FOR FIRST rep-oper-ctrab 
                 WHERE rep-oper-ctrab.cod-ctrab = rep-parada-ctrab.cod-ctrab
                   AND rep-oper-ctrab.dat-inic-reporte = rep-parada-ctrab.dat-inic-parada
                   AND rep-oper-ctrab.qtd-segs-inic-reporte <= rep-parada-ctrab.qtd-segs-inic
                   AND rep-oper-ctrab.qtd-segs-fim-reporte >= rep-parada-ctrab.qtd-segs-fim,
    
                 FIRST split-operac WHERE split-operac.nr-ord-prod = rep-oper-ctrab.nr-ord-prod.
                 
                 ASSIGN tt-parada.it-codigo = split-operac.it-codigo.

             END.
        END.
        
/* FOR EACH tt-parada no-lock.                                                                                                      */
/*     DISP tt-parada.dat-inic-parada tt-parada.cod-parada tt-parada.des-parada tt-parada.tempo tt-parada.it-codigo WITH WIDTH 400. */
/* END.                                                                                                                             */

//FINAL PARADAS **************************************************

//CALCULA O DESCONTO DO TEMPO PROGRAMADO
FOR EACH rep-parada-ctrab
    WHERE rep-parada-ctrab.cod-ctrab = centro
    AND rep-parada-ctrab.dat-inic-parada >= data-ini
    AND rep-parada-ctrab.dat-fim-parada <= data-fim,

    EACH motiv-parada WHERE motiv-parada.cod-parada = rep-parada-ctrab.cod-parada NO-LOCK BREAK BY rep-parada-ctrab.dat-inic-parada.

    IF FIRST-OF(rep-parada-ctrab.dat-inic-parada) THEN DO:
        ASSIGN data-atu = rep-parada-ctrab.dat-inic-parada.

        FIND FIRST tt-oee
             WHERE tt-oee.data = data-atu NO-ERROR.

        IF NOT AVAIL tt-oee THEN DO:
            CREATE tt-oee.

            ASSIGN tt-oee.data = data-atu.
        END.
        
    END.
    
    //Paradas que alteram eficiencia
    IF motiv-parada.log-alter-eficien = TRUE THEN DO:
        ASSIGN cont-tempo-par-alt = cont-tempo-par-alt + qtd-tempo-parada + qtd-tempo-ext.
    END.

    //Paradas que NÇO alteram eficiencia
    IF motiv-parada.log-alter-eficien = FALSE THEN DO:
        ASSIGN cont-tempo-par-nao-alt = cont-tempo-par-nao-alt + qtd-tempo-parada + qtd-tempo-ext.
    END.

    IF LAST-OF(rep-parada-ctrab.dat-inic-parada) THEN DO:

        FIND FIRST tt-oee
             WHERE tt-oee.data = data-atu NO-ERROR.

        IF AVAIL tt-oee THEN DO:
            
            ASSIGN tt-oee.temp-par-alt = cont-tempo-par-alt
                   tt-oee.temp-par-nao-alt = cont-tempo-par-nao-alt.
        END.

        IF NOT AVAIL tt-oee THEN DO:
            MESSAGE "ERRO! Algo deu errado 1." VIEW-AS ALERT-BOX.
        END.
        
        ASSIGN cont-tempo-par-alt = 0
               cont-tempo-par-nao-alt = 0.
               
    END.

END.

//CALCULA O DESCONTO DO TEMPO PROGRAMADO **************************************************


//CAPTURANDO HORAS EQUIVALENTE OS TURNOS DE UM DIA ***************

FOR EACH turno-semana.
    ASSIGN cont-turno = cont-turno + qtd-tempo-util-sem.
END.

ASSIGN cont-turno = cont-turno / 5.

//FINAL TURNOS **********************************************/


//BUSCA DE APONTAMENTOS DATA, TEMPO PRODUZINDO, PRODU€ÇO REAL, PE€AS *************

ASSIGN cont-temp-prod = 0
       cont-pc-boas = 0
       cont-pc-ruins = 0
       cont-dias-mes = 0
       cont-prod-teorica = 0.

FOR EACH rep-oper-ctrab
    WHERE rep-oper-ctrab.cod-ctrab = centro
      AND rep-oper-ctrab.dat-inic-report >= data-ini
      AND rep-oper-ctrab.dat-fim-report <= data-fim,

    FIRST split-operac WHERE split-operac.nr-ord-produ = rep-oper-ctrab.nr-ord-produ,
    FIRST rep-oper WHERE rep-oper.nr-ord-produ = split-operac.nr-ord-produ AND rep-oper.nr-reporte = rep-oper-ctrab.nr-reporte,
    FIRST oper-ord WHERE oper-ord.nr-ord-produ = rep-oper.nr-ord-produ AND oper-ord.op-codigo = rep-oper.op-codigo  NO-LOCK BREAK BY rep-oper-ctrab.dat-inic-report.

    IF FIRST-OF(rep-oper-ctrab.dat-inic-report) THEN DO:
        ASSIGN data-atu = rep-oper-ctrab.dat-inic-report.

        FIND FIRST tt-oee
             WHERE tt-oee.data = data-atu NO-ERROR.

        IF AVAIL tt-oee THEN DO:
            ASSIGN tt-oee.turno = cont-turno - tt-oee.temp-par-nao-alt.
        END.

        IF NOT AVAIL tt-oee THEN DO:
            CREATE tt-oee.

            ASSIGN tt-oee.data = data-atu
                   tt-oee.turno = cont-turno - tt-oee.temp-par-nao-alt.
        END.
    END.

    //VERIFICA SE O ITEM Ô COMPOSTO
    FIND FIRST CST_Item
         WHERE CST_Item.it-codigo = split-operac.it-codigo NO-ERROR.
    
         
         IF AVAIL CST_Item THEN DO:

            //SE O ITEM  COMPOSTO FA€A ISSO:
            IF CST_Item.LOG_Composto = TRUE THEN DO:
                
                  //MESSAGE "COMPOSTO" VIEW-AS ALERT-BOX.
                  ASSIGN cont-dias-mes = cont-dias-mes + 1
                         cont-temp-prod = cont-temp-prod + (rep-oper-ctrab.qtd-tempo-reporte / 2)
                         cont-pc-boas = cont-pc-boas + rep-oper-ctrab.qtd-operac-aprov
                         cont-pc-ruins = cont-pc-ruins + rep-oper-ctrab.qtd-operac-refgda
                         cont-prod-teorica = cont-prod-teorica + (ROUND((((oper-ord.nr-unidades / oper-ord.tempo-maquin) * 60) * (rep-oper-ctrab.qtd-tempo-reporte / 2)),0)).
            END.

            //SE N¶O  COMPOSTO FA€A ISSO:
            ELSE DO:

                
                //MESSAGE "NÇO COMPOSTO" VIEW-AS ALERT-BOX.
                ASSIGN cont-dias-mes = cont-dias-mes + 1
                    cont-temp-prod = cont-temp-prod + rep-oper-ctrab.qtd-tempo-reporte
                    cont-pc-boas = cont-pc-boas + rep-oper-ctrab.qtd-operac-aprov
                    cont-pc-ruins = cont-pc-ruins + rep-oper-ctrab.qtd-operac-refgda
                    cont-prod-teorica = cont-prod-teorica + (ROUND((((oper-ord.nr-unidades / oper-ord.tempo-maquin) * 60) * rep-oper-ctrab.qtd-tempo-reporte),0)).

            END.

         END.
        
/*           MESSAGE   "ordem: " + string(rep-oper-ctrab.nr-ord-produ) SKIP                                                                                         */
/*                     "item: " + STRING(oper-ord.it-codigo) SKIP                                                                                                   */
/*                     "nr-unidade: " + STRING(oper-ord.nr-unidades) SKIP                                                                                           */
/*                     "tempo-maquin: " + string(oper-ord.tempo-maquin) SKIP                                                                                        */
/*                                                                                                                                                                  */
/*                     "Ciclo: " + STRING((oper-ord.tempo-maquin / oper-ord.nr-unidades) * 100) SKIP                                                                */
/*                     "=========" SKIP                                                                                                                             */
/*                     "reporte: " + string(rep-oper-ctrab.qtd-tempo-reporte) SKIP                                                                                  */
/*                     "PE€AS: " + STRING((ROUND((((oper-ord.nr-unidades / oper-ord.tempo-maquin) * 60) * rep-oper-ctrab.qtd-tempo-reporte),0))) VIEW-AS ALERT-BOX. */
         
         IF NOT AVAIL CST_Item THEN DO:

             ASSIGN cont-dias-mes = cont-dias-mes + 1
                    cont-temp-prod = cont-temp-prod + rep-oper-ctrab.qtd-tempo-reporte
                    cont-pc-boas = cont-pc-boas + rep-oper-ctrab.qtd-operac-aprov
                    cont-pc-ruins = cont-pc-ruins + rep-oper-ctrab.qtd-operac-refgda
                    cont-prod-teorica = cont-prod-teorica + (ROUND((((oper-ord.nr-unidades / oper-ord.tempo-maquin) * 60) * rep-oper-ctrab.qtd-tempo-reporte),0)).

         END.
    
    IF LAST-OF(rep-oper-ctrab.dat-inic-report) THEN DO:

        FIND FIRST tt-oee
             WHERE tt-oee.data = data-atu NO-ERROR.

        IF AVAIL tt-oee THEN DO:
            
            ASSIGN tt-oee.temp-prod = cont-temp-prod
                   tt-oee.prod-real = cont-pc-boas + cont-pc-ruins
                   tt-oee.pc-boas = cont-pc-boas
                   tt-oee.pc-ruins = cont-pc-ruins
                   tt-oee.pc-boas-ruins = cont-pc-boas + cont-pc-ruins
                   tt-oee.cod-item = split-operac.it-codigo
                   tt-oee.prod-teorica = cont-prod-teorica // cont-dias-mes
                   tt-oee.oee-disp = ROUND(((tt-oee.temp-prod / tt-oee.turno) * 100),0)
                   tt-oee.oee-perf = ROUND(((tt-oee.prod-real / tt-oee.prod-teorica) * 100),0)
                   tt-oee.oee-quali = ROUND(((tt-oee.pc-boas / tt-oee.pc-boas-ruins) * 100),0)
                   tt-oee.oee-percent = ROUND(((tt-oee.oee-disp * tt-oee.oee-perf * tt-oee.oee-quali) / 10000),0).
        END.

        IF NOT AVAIL tt-oee THEN DO:
            MESSAGE "ERRO! Algo deu errado 2." VIEW-AS ALERT-BOX.
        END.
        
        ASSIGN cont-temp-prod = 0
               cont-pc-boas = 0
               cont-pc-ruins = 0
               cont-dias-mes = 0
               cont-prod-teorica = 0.

    END.
    
END.

//FINAL BUSCA APONTAMENTOS***********************************************************

/* FOR EACH tt-oee.                                                          */
/*     DISP tt-oee.data tt-oee.prod-real tt-oee.prod-teorica WITH WIDTH 400. */
/* END.                                                                      */
