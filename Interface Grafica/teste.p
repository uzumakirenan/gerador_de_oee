/* FOR EACH rep-oper-ctrab                                                         */
/*     WHERE rep-oper-ctrab.cod-ctrab = "4120101SP09"                              */
/*       AND rep-oper-ctrab.dat-inic-report >= DATE("07/01/2021")                  */
/*       AND rep-oper-ctrab.dat-fim-report <= DATE("07/01/2021"),                  */
/*                                                                                 */
/*     FIRST rep-oper WHERE rep-oper.nr-ord-produ = rep-oper-ctrab.nr-ord-produ,   */
/*     FIRST split-operac WHERE split-operac.nr-ord-produ = rep-oper.nr-ord-produ. */
/*                                                                                 */
/*     MESSAGE rep-oper-ctrab.num-seq-rep VIEW-AS ALERT-BOX.                       */
/*                                                                                 */
/*     DISP rep-oper-ctrab.dat-inic-report split-operac.op-codigo.                 */
/*                                                                                 */
/* END.                                                                            */
    

/* FOR EACH rep-oper WHERE rep-oper.nr-ord-produ = 9439.  */
/*     MESSAGE string(rep-oper.char-1) VIEW-AS ALERT-BOX. */
/* END.                                                   */
            
                
FOR EACH rep-oper WHERE rep-oper.nr-ord-produ = 9439,
    
    FIRST rep-oper-ctrab 
        WHERE rep-oper-ctrab.nr-ord-produ = rep-oper.nr-ord-produ
        AND rep-oper-ctrab.dat-inic-report >= DATE("07/01/2021")
        AND rep-oper-ctrab.dat-fim-report <= DATE("07/01/2021").


    DISP rep-oper-ctrab.dat-inic-reporte rep-oper-ctrab.qtd-tempo-reporte rep-oper.op-codigo rep-oper.quantidade.


END.
