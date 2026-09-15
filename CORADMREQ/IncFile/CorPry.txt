<%

GPRYNOM = "NO"
GPRYSTS = 0
GPRYCOD = "999"
GPRYFSP = 0
GPRYFEP = 0
GPRYFSR = 0
GPRYFER = 0

XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
XFCURR  = XFSTART
XFHOY   = XFSTART
XFREAL  = XFSTART
XTACUM  = 1
XTACUMR = 1
XTFTSK  = 0

XMFMIN = 0
XMFMAX = 0
MFMIN  = 0
MFMAX  = 0

DIM ATAREA(400,15)
DIM AETAPA(100,2)

DIM WTAREA(400,15)
DIM WETAPA(100,2)

ITART = 0
ITATI = 0
TTAR  = 0

IETPT = 0
IETPI = 0 
TETP  = 0

CDESP = 0
CHASP = 0
CDESR = 0
CHASR = 0
CMAX  = 0   

FUNCTION GETPRYSTS(YPRY)



    StrSql = "Select * From PROYECTO where CODIGO = '" & YPRY & "'  "
    set rsXP = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsXP.eof then
       GPRYNOM = rsXP(2)
       GPRYSTS = 0
       GPRYCOD = YPRY
       GPRYFSP = rsXP(3)
       GPRYFEP = rsXP(5)
       GPRYFSR = rsXP(4)
       GPRYFER = rsXP(6)
       GETPRYSTS = (rsXP(11) > 1)
    end if

END FUNCTION

FUNCTION GETFSTART(YPRY,YTSK)

  YXFSTART = 0

  StrSql = "Select * From PROYTAREASLST WHERE CODIGO = '" & YPRY & "' and TAREA = '" & YTSK & "' AND ESTADO = 0 order by orden "
  set rsXp2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do while not rsXp2.eof
     YXFSTARTD = 0
     if rsxp2(5) <> YTSK then
        StrSql = "Select * From PROYTAREAS WHERE TAREA = '" & RSXp2(5) & "' "
        set rsxp3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT rsxp3.EOF THEN
           IF  rsxp3(10) = 0 THEN
               YXFSTARTD = GETFEND(rsxp3(3),rsxp3(1))
           ELSE
               YXFSTARTD = rsxp3(10)
           END IF 

        END IF
        call ofv.cerrarconsulta(rsxp3)

        IF YXFSTARTD > YXFSTART THEN YXFSTART = YXFSTARTD

    END IF


  rsxp2.MOVENEXT
  LOOP
  
  call ofv.cerrarconsulta(rsxp2)

  GETFSTART = YXFSTART

END FUNCTION

FUNCTION GETFEND(XTSK,XPRY)

    YXFEND = 0
    YXSALE = "N"

    StrSql = "Select * From PROYecto where codigo = '" & XPRY & "'  "
    set rsXP = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsXP.eof then
       if rsXP(4) > 0 then
          YXFSTART = rsXP(4)
       else
          YXFSTART = rsXP(3)
       end if
 
    end if
    YXFCURR = YXFSTART

  StrSql = "Select * From PROYetapa where codigo = '" & XPRY & "' order by orden "
  set rsXP = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsXP.eof then
     do while not rsXP.eof


  StrSql = "Select * From PROYtareas where codigo = '" & XPRY & "' and etapa = '" & rsXP(2) & "' order by orden "
  set rsXP1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rsXP1.eof then
     do while not rsXP1.eof



    IF RSXP1(15) > 0 THEN
       YXAGRD = rsXP1(15)
       if rsXP1(16) = 1 then
          YXAGRD = CINT(YXAGRD / 8)
       elseif rsXP1(16) = 3 then
          YXAGRD = YXAGRD * 7
       else

       end if 
    ELSE
       YXAGRD = 1
    END IF 
          
    YTSKFS = RSXP1(8)
    YTSKFE = RSXP1(10)

    IF RSXP1(7) > 0 THEN

       YXTSKSTART = GETFSTART(RSXP1(3))
       IF YXTSKSTART > 0 THEN
          YXFCURR = YXTSKSTART 
          YTSKFS = YXFCURR
       END IF
    END IF 


    IF YTSKFS > 0  THEN
       if YTSKFS = 11111 then
          YXFCURR = YXFSTART
       end if 

    END IF 

    YXFCURR = YXFCURR + YXAGRD

    IF RSXP1(3) = XTSK THEN
       YXFEND = YXFCURR
       YXSALE = "S"
       EXIT DO
    END IF

    rsXP1.movenext
    loop
 
  end if

    IF YXSALE = "S" THEN
       EXIT DO
    END IF

    rsXP.movenext
    loop
 
  end if

  GETFEND = YXFEND


END FUNCTION


FUNCTION GETPRYCALC(YPRY)

' ''''''' CABECERA DEL PROYECTO

    StrSql = "Select * From PROYecto where codigo = '" & YPRY & "'  "
    set rsXP0 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsXP0.eof then
       if rsXP0(4) > 0 then
          XFSTART = rsXP0(4)
          XFCURR  = XFSTART
          XFREAL  = XFSTART
       else
          XFSTART = rsXP0(3)
          XFCURR  = XFSTART
          XFREAL  = XFSTART
       end if

       XPRYNOM = rsXP0(2)
 
    end if

    call ofv.cerrarconsulta(rsxp0)

      XFCURR = XFSTART

' ''''''' ETAPAS Y TAREAS

    StrSql = "Select MAX(ORDEN) From PROYETAPA where CODIGO = '" & YPRY & "' AND ORDEN < 999 "
    set rsXP0 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if rsXP0.eof then
       XORDEN = 10
    ELSE
       XORDEN = rsXP0(0)
       IF XORDEN <> "" THEN
          XORDEN = XORDEN + 10
       ELSE
          XORDEN = 10
       END IF
    END IF

    call ofv.cerrarconsulta(rsxp0)


  StrSql = "Select * From PROYETAPA where CODIGO = '" & YPRY & "' order by ORDEN "
  set rsxp0 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsxp0.eof then
     do while not rsxp0.eof


    SINPUT = rsxp0(3)
    IF SINPUT = 999 THEN SINPUT = XORDEN
    XETPORD = SINPUT


    sInput = rsxp0(4)

    IETPT            = IETPT + 1
    AETAPA (IETPT,1) = rsxp0(4)
    AETAPA (IETPT,2) = XETPORD
    TETP             = IETPT 



    StrSql = "Select * From PROYTAREAS where CODIGO = '" & YPRY & "' and ETAPA = '" & rsxp0(2) & "' order by ORDEN "
 '' response.write strsql
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not  rs1.eof then
       do while not rs1.eof

          sinput1 = sImagenE0
          sInputn0 = rs1(5)


    IF RS1(13) <> "NO" THEN
       SINPUT4 = SIMAGEN4S
       XCOLOR = "#DDDDDD"
    ELSE
       SINPUT4 = "<B>" & SIMAGEN4N & "</B>"
       XCOLOR = "#AA0000"
    END IF 

    IF RS1(15) > 0 THEN
       XAGRD = rs1(15)
       if rs1(16) = 1 then
          XAGRD = CINT(XAGRD / 8)
       elseif rs1(16) = 3 then
          XAGRD = XAGRD * 7
       else

       end if 

    ELSE

       XAGRD = 0
    END IF 


    TSKst = RS1(18) & ""

    if TSKst = "" then TSKst = "0"


    TSKFS = RS1(8)
    TSKFE = RS1(10)
    TSKFSr = RS1(9)
    TSKFEr = RS1(11)

    IF RS1(7) > 0 THEN

       XTSKSTART = GETFSTART(RS1(3))
       IF XTSKSTART > 0 THEN
          XFREAL = XTSKSTART 
          TSKFSr = XFREAL
       END IF
    END IF 

     if TSKFS = 0 or TSKFS = 11111 then
        if TSKFSr > 0 then 
           TSKFS  = TSKFSr
           TSKFE  = TSKFS + XAGRD
           TSKFEr = TSKFE

        else
           TSKFS = XFREAL
           TSKFSr = XFREAL
           TSKFE  = TSKFS + XAGRD
           TSKFEr = TSKFE

        end if
    end if         

       FSPLAN = TSKFS
       FEPLAN = TSKFE



   if TSKst = "0" then

        FSREAL = XFREAL
        XFREAL = XFREAL + XAGRD
        FEREAL = XFREAL

        XAGRR = XAGRD


   elseif TSKst = "2" then



         FSREAL = TSKFSR
         FEREAL = TSKFER
         XFREAL = TSKFER

         XAGRR = FEREAL - FSREAL


   else


          XFREAL = TSKFSR

          PADV = 0 


          StrSql = "Select * From PROYTAREASADV where codigo = '" & session("NCliente") & "' AND TSK = " & RS1(0) & " ORDER BY id desc "
        ''  response.write strsql & "<br>"
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then


             PADV = RSX(5)

             IF XFHOY < TSKFE THEN

                XFREAL = XFREAL + XAGRD
             
                XAGRR = XAGRD


             ELSEIF XFHOY = TSKFE AND PADV > 99  THEN

                XFREAL = XFREAL + XAGRD

                XAGRR = XAGRD

             ELSE

                XAGRR =  (XFHOY - XFREAL) +  XAGRD - CINT((XAGRD * PADV / 100)) 

                XFREAL = XFHOY + XAGRD - CINT((XAGRD * PADV / 100)) 


             END IF

          else

            IF XFHOY < TSKFE THEN

                XFREAL = XFREAL + XAGRD
             
                XAGRR = XAGRD



             ELSE



                XAGRR =  (XFHOY - XFREAL) +  XAGRD 

                XFREAL = XFHOY + XAGRD 



             END IF



          END IF




        FSREAL = TSKFSR
        FEREAL = XFREAL


   end if


    IF MFMAX < FEREAL THEN MFMAX = FEREAL

    mXTACUM  = FSPLAN - MFMIN + 1 + XAGRD
    mXTACUMR = FSPLAN - MFMIN + 1 + XAGRR

    XTACUM   = FSPLAN - MFMIN + 1
    XTACUMR  = FSREAL - MFMIN + 1


    CDESP    = XTACUM
    XTACUM   = XTACUM + XAGRD
    CHASP    = XTACUM       

    CDESR    = XTACUMR
    XTACUMR  = XTACUMR + XAGRR
    CHASR    = XTACUMR



    if RS1(13) = "NO" then
       sniv = "SIN AFECTAR"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & rs1(13) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
    END IF



    ITART             = ITART + 1
    ATAREA (ITART,0)  = RS1(0)
    ATAREA (ITART,1)  = IETPT
    ATAREA (ITART,2)  = RS1(5)
    ATAREA (ITART,3)  = RS1(4)
    ATAREA (ITART,4)  = CDESP
    ATAREA (ITART,5)  = CDESR
    ATAREA (ITART,6)  = CHASP
    ATAREA (ITART,7)  = CHASR 
    ATAREA (ITART,8)  = TSKst
    ATAREA (ITART,9)  = sniv
    ATAREA (ITART,10) = PADV
    ATAREA (ITART,11) = FSREAL
    ATAREA (ITART,12) = FEREAL
    ATAREA (ITART,13) = FSPLAN
    ATAREA (ITART,14) = FEPLAN
    ATAREA (ITART,15) = XAGRD
    TTAR              = ITART 





    rs1.movenext
    loop
    call ofv.cerrarconsulta(rs1)
 
  end if



    rsxp0.movenext
    loop
    call ofv.cerrarconsulta(rsxp0)
 
  end if

    CMAX = MFMAX - MFMIN + 1

    TSEM = CMAX / 7

    IF (CINT(TSEM) * 7) <> CMAX THEN
       TSEM = CINT(TSEM) + 1
       CMAX = TSEM * 7
    ELSE
       TSEM = CINT(TSEM)
    END IF  


END FUNCTION



%>