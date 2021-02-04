DEF VAR data-ini AS DATE.
DEF VAR data-fim AS DATE.
DEF VAR data-atu AS DATE.
DEF VAR cont-temp-prod AS DECIMAL.
DEF VAR cont-pc-boas AS DECIMAL.
DEF VAR cont-pc-ruins AS DECIMAL.
DEF VAR cont-turno AS DECIMAL.
DEF VAR cont-tempo-par-alt AS DECIMAL.
DEF VAR cont-tempo-par-nao-alt AS DECIMAL.
DEF VAR cont-dias-mes AS INT.
DEF VAR cont-prod-teorica AS DECIMAL.

DEFINE NEW GLOBAL SHARED VAR centro AS CHAR NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR oee-ano AS INT NO-UNDO.

DEFINE TEMP-TABLE tt-oee NO-UNDO
    FIELD data AS DATE LABEL "Data Apontamento"
    FIELD turno AS DECIMAL LABEL "Tempo Programado"
    FIELD cod-item LIKE split-operac.it-codigo LABEL "Item" 
    FIELD temp-prod AS DECIMAL LABEL "Tempo Produzindo"
    FIELD prod-teorica AS DECIMAL LABEL "Produ��o Teorica"
    FIELD prod-real AS INT LABEL "Produ��o Real"
    FIELD pc-boas AS INT LABEL "Pe�as Boas"
    FIELD pc-ruins AS INT LABEL "Pe�as Ruins"
    FIELD pc-boas-ruins AS INT LABEL "Pe�as Boas + Ruins"
    FIELD temp-par-alt AS DECIMAL
    FIELD temp-par-nao-alt AS DECIMAL
    FIELD oee-disp AS DECIMAL LABEL "Disponibilidade"
    FIELD oee-perf AS DECIMAL LABEL "Performance"
    FIELD oee-quali AS DECIMAL LABEL "Qualidade"
    FIELD oee-percent AS DECIMAL LABEL "Percentual OEE"
    FIELD mes AS INT.


DEFINE NEW GLOBAL SHARED TEMP-TABLE tt-oee-men NO-UNDO
    FIELD mes-nome AS CHAR FORMAT "x(10)" LABEL "Mes"
    FIELD turno AS DECIMAL LABEL "Tempo Programado"
    FIELD cod-item LIKE split-operac.it-codigo LABEL "Item" 
    FIELD temp-prod AS DECIMAL LABEL "Tempo Produzindo"
    FIELD prod-teorica AS DECIMAL LABEL "Produ��o Teorica"
    FIELD prod-real AS INT LABEL "Produ��o Real"
    FIELD pc-boas AS INT LABEL "Pe�as Boas"
    FIELD pc-ruins AS INT LABEL "Pe�as Ruins"
    FIELD pc-boas-ruins AS INT LABEL "Pe�as Boas + Ruins"
    FIELD temp-par-alt AS DECIMAL
    FIELD temp-par-nao-alt AS DECIMAL
    FIELD oee-disp AS DECIMAL LABEL "Disponibilidade"
    FIELD oee-perf AS DECIMAL LABEL "Performance"
    FIELD oee-quali AS DECIMAL LABEL "Qualidade"
    FIELD oee-percent AS DECIMAL LABEL "Percentual OEE"
    FIELD mes AS INT.

DEFINE NEW GLOBAL SHARED TEMP-TABLE tt-parada-men NO-UNDO
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
    FIELD tempo AS DECIMAL
    FIELD mes AS INT.



ASSIGN cont-temp-prod = 0.
       //oee-ano = 2020
       //centro = "4138001PH07".

ASSIGN data-ini = DATE(STRING("01/01/" + string(oee-ano) ))
       data-fim = DATE(STRING("31/12/" + string(oee-ano) )).

EMPTY TEMP-TABLE tt-oee.
EMPTY TEMP-TABLE tt-oee-men.
EMPTY TEMP-TABLE tt-parada-men.

//CAPTURA DE PARADAS

        FOR EACH rep-parada-ctrab NO-LOCK
           WHERE rep-parada-ctrab.cod-ctrab = centro
             AND rep-parada-ctrab.dat-inic-parada >= data-ini
             AND rep-parada-ctrab.dat-fim-parada <= data-fim,

             FIRST motiv-parada WHERE motiv-parada.cod-parada = rep-parada-ctrab.cod-parada.

             CREATE tt-parada-men.
             ASSIGN tt-parada-men.cod-ctrab = string(rep-parada-ctrab.cod-ctrab)
                    tt-parada-men.cod-parada = rep-parada-ctrab.cod-parada
                    tt-parada-men.des-parada = motiv-parada.des-parada
                    tt-parada-men.log-alter-eficien = motiv-parada.log-alter-eficien
                    tt-parada-men.inicio = DATETIME(rep-parada-ctrab.dat-inic-parada, INT(rep-parada-ctrab.qtd-segs-inic * 1000))
                    tt-parada-men.fim = DATETIME(rep-parada-ctrab.dat-fim-parada, INT(rep-parada-ctrab.qtd-segs-fim * 1000))
                    tt-parada-men.dat-inic-parada = rep-parada-ctrab.dat-inic-parada
                    tt-parada-men.qtd-segs-inic = rep-parada-ctrab.qtd-segs-inic
                    tt-parada-men.dat-fim-parada = rep-parada-ctrab.dat-fim-parada
                    tt-parada-men.qtd-segs-fim = rep-parada-ctrab.qtd-segs-fim
                    tt-parada-men.tempo = rep-parada-ctrab.qtd-tempo-parada + rep-parada-ctrab.qtd-tempo-ext
                    tt-parada-men.mes = int(MONTH(tt-parada-men.dat-inic-parada)).

             FOR FIRST rep-oper-ctrab 
                 WHERE rep-oper-ctrab.cod-ctrab = rep-parada-ctrab.cod-ctrab
                   AND rep-oper-ctrab.dat-inic-reporte = rep-parada-ctrab.dat-inic-parada
                   AND rep-oper-ctrab.qtd-segs-inic-reporte <= rep-parada-ctrab.qtd-segs-inic
                   AND rep-oper-ctrab.qtd-segs-fim-reporte >= rep-parada-ctrab.qtd-segs-fim,
    
                 FIRST split-operac WHERE split-operac.nr-ord-prod = rep-oper-ctrab.nr-ord-prod.
                 
                 ASSIGN tt-parada-men.it-codigo = split-operac.it-codigo.

             END.
        END.

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

    //Paradas que N�O alteram eficiencia
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
            MESSAGE "ERRO! Algo deu errado." VIEW-AS ALERT-BOX.
        END.

        ASSIGN cont-tempo-par-alt = 0
               cont-tempo-par-nao-alt = 0.

    END.

END.
//FINAL CALCULA O DESCONTO DO TEMPO PROGRAMADO **************************************************



//CAPTURANDO HORAS EQUIVALENTE OS TURNOS DE UM DIA ***************

FOR EACH turno-semana.
    ASSIGN cont-turno = cont-turno + qtd-tempo-util-sem.
END.

ASSIGN cont-turno = cont-turno / 5.

//FINAL TURNOS **********************************************


//BUSCA DE APONTAMENTOS DATA, TEMPO PRODUZINDO, PRODU��O REAL, PE�AS *************

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
    FIRST oper-ord WHERE oper-ord.nr-ord-produ = rep-oper-ctrab.nr-ord-produ NO-LOCK BREAK BY rep-oper-ctrab.dat-inic-report.
    //FIRST operacao WHERE operacao.it-codigo = split-operac.it-codigo NO-LOCK BREAK BY rep-oper-ctrab.dat-inic-report.


    IF FIRST-OF(rep-oper-ctrab.dat-inic-report) THEN DO:
        ASSIGN data-atu = rep-oper-ctrab.dat-inic-report.

        FIND FIRST tt-oee
             WHERE tt-oee.data = data-atu NO-ERROR.

        IF AVAIL tt-oee THEN DO:
            ASSIGN tt-oee.turno = cont-turno - tt-oee.temp-par-nao-alt
                   tt-oee.mes = INT(SUBSTR(STRING(data-atu), 4,2)).
        END.

        IF NOT AVAIL tt-oee THEN DO:
            CREATE tt-oee.

            ASSIGN tt-oee.data = data-atu
                   tt-oee.turno = cont-turno - tt-oee.temp-par-nao-alt
                   tt-oee.mes = INT(SUBSTR(STRING(data-atu), 4,2)).

        END.


    END.

    //VERIFICA SE O ITEM � COMPOSTO
    FIND FIRST CST_Item
         WHERE CST_Item.it-codigo = split-operac.it-codigo NO-ERROR.

         
         IF AVAIL CST_Item THEN DO:

            //SE O ITEM � COMPOSTO FA�A ISSO:
            IF CST_Item.LOG_Composto = TRUE THEN DO:
                  //MESSAGE "COMPOSTO" VIEW-AS ALERT-BOX.
                  ASSIGN cont-dias-mes = cont-dias-mes + 1
                         cont-temp-prod = cont-temp-prod + (rep-oper-ctrab.qtd-tempo-reporte / 2)
                         cont-pc-boas = cont-pc-boas + rep-oper-ctrab.qtd-operac-aprov
                         cont-pc-ruins = cont-pc-ruins + rep-oper-ctrab.qtd-operac-refgda
                         cont-prod-teorica = cont-prod-teorica + (ROUND((((oper-ord.nr-unidades / oper-ord.tempo-maquin) * 60) * (rep-oper-ctrab.qtd-tempo-reporte / 2)),0)).
            END.


            //SE N�O � COMPOSTO FA�A ISSO:
            ELSE DO:
                //MESSAGE "N�O � COMPOSTO" VIEW-AS ALERT-BOX.
                ASSIGN cont-dias-mes = cont-dias-mes + 1
                       cont-temp-prod = cont-temp-prod + rep-oper-ctrab.qtd-tempo-reporte
                       cont-pc-boas = cont-pc-boas + rep-oper-ctrab.qtd-operac-aprov
                       cont-pc-ruins = cont-pc-ruins + rep-oper-ctrab.qtd-operac-refgda
                       cont-prod-teorica = cont-prod-teorica + (ROUND((((oper-ord.nr-unidades / oper-ord.tempo-maquin) * 60) * rep-oper-ctrab.qtd-tempo-reporte),0)).
            END.
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
                   tt-oee.prod-teorica = cont-prod-teorica / cont-dias-mes.
        END.

        IF NOT AVAIL tt-oee THEN DO:
            MESSAGE "ERRO! Algo deu errado." VIEW-AS ALERT-BOX.
        END.

        //DISP data-atu LABEL "Data Apontamento" tempo-total LABEL "Tempo Produzindo".
        ASSIGN cont-temp-prod = 0
               cont-pc-boas = 0
               cont-pc-ruins = 0
               cont-dias-mes = 0
               cont-prod-teorica = 0.

    END.

END.

//FINAL BUSCA APONTAMENTOS***********************************************************

ASSIGN cont-temp-prod = 0
       cont-pc-boas = 0
       cont-pc-ruins = 0
       cont-turno = 0
       cont-dias-mes = 0
       cont-prod-teorica = 0.

FOR EACH tt-oee BREAK BY tt-oee.mes.

    ASSIGN cont-temp-prod = cont-temp-prod + tt-oee.temp-prod
           cont-pc-boas = cont-pc-boas + tt-oee.pc-boas
           cont-pc-ruins = cont-pc-ruins + tt-oee.pc-ruins
           cont-turno = cont-turno + tt-oee.turno
           cont-dias-mes = cont-dias-mes + 1
           cont-prod-teorica = cont-prod-teorica + tt-oee.prod-teorica.

    IF LAST-OF(tt-oee.mes) THEN DO:

        FIND FIRST tt-oee-men
             WHERE tt-oee-men.mes = tt-oee.mes NO-ERROR.

        IF NOT AVAIL tt-oee-men THEN DO:
            CREATE tt-oee-men.

            ASSIGN tt-oee-men.turno = cont-turno / cont-dias-mes
                   tt-oee-men.temp-prod = cont-temp-prod / cont-dias-mes
                   tt-oee-men.prod-real = (cont-pc-boas + cont-pc-ruins)  / cont-dias-mes
                   tt-oee-men.pc-boas = cont-pc-boas / cont-dias-mes
                   tt-oee-men.pc-ruins = cont-pc-ruins / cont-dias-mes
                   tt-oee-men.pc-boas-ruins = (cont-pc-boas + cont-pc-ruins) / cont-dias-mes
                   tt-oee-men.cod-item = split-operac.it-codigo
                   tt-oee-men.prod-teorica = cont-prod-teorica / cont-dias-mes
                   tt-oee-men.mes = tt-oee.mes
                   tt-oee-men.oee-disp = ROUND(((tt-oee-men.temp-prod / tt-oee-men.turno) * 100),0)
                   tt-oee-men.oee-perf = ROUND(((tt-oee-men.prod-real / tt-oee-men.prod-teorica) * 100),0)
                   tt-oee-men.oee-quali = ROUND(((tt-oee-men.pc-boas / tt-oee-men.pc-boas-ruins) * 100),0)
                   tt-oee-men.oee-percent = ROUND(((tt-oee-men.oee-disp * tt-oee-men.oee-perf * tt-oee-men.oee-quali) / 10000),0).

            IF tt-oee-men.mes = 1 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Janeiro".
            END.

            IF tt-oee-men.mes = 2 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Fevereiro".
            END.

            IF tt-oee-men.mes = 3 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Mar�o".
            END.

            IF tt-oee-men.mes = 4 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Abril".
            END.

            IF tt-oee-men.mes = 5 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Maio".
            END.

            IF tt-oee-men.mes = 6 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Junho".
            END.

            IF tt-oee-men.mes = 7 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Julho".
            END.

            IF tt-oee-men.mes = 8 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Agosto".
            END.

            IF tt-oee-men.mes = 9 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Setembro".
            END.

            IF tt-oee-men.mes = 10 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Outubro".
            END.

            IF tt-oee-men.mes = 11 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Novembro".
            END.

            IF tt-oee-men.mes = 12 THEN DO:
                ASSIGN tt-oee-men.mes-nome = "Dezembro".
            END.

        END.

        ASSIGN cont-temp-prod = 0
               cont-pc-boas = 0
               cont-pc-ruins = 0
               cont-turno = 0
               cont-dias-mes = 0
               cont-prod-teorica = 0.

    END.

END.

/* FOR EACH tt-oee WHERE tt-oee.mes = 1.                                                                                                                      */
/*     DISP tt-oee.mes tt-oee.cod-item tt-oee.turno tt-oee.temp-prod tt-oee.prod-teorica tt-oee.prod-real tt-oee.pc-boas tt-oee.pc-boas-ruins WITH WIDTH 400. */
/* END.                                                                                                                                                       */

/* FOR EACH tt-oee-men WHERE tt-oee-men.mes > 0.                                                                                                     */
/*     DISP tt-oee-men.mes-nome tt-oee-men.turno tt-oee-men.oee-disp tt-oee-men.oee-perf tt-oee-men.oee-quali tt-oee-men.oee-percent WITH WIDTH 400. */
/* END.                                                                                                                                              */