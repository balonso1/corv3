<%@LCID = 11274%> 
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


     EMP = SESSION("VEMP")

     VSESS = SESSION.SESSIONID


     set cn = ofv.conectar(ofv.strconn2)



     strsql = "Insert into SYSDEFDSKTOP (OBJUN,DEFSET,DSCSET,USRGRP,TIPO) "
     strsql = STRSQL & " values (" & EMP & ",'NO','" & VSESS & "','NO','NO')"
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)



     Strsql = "Select OBJID From SYSDEFDSKTOP WHERE DSCSET = '" & VSESS & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        VKEY = RS(0) 
        strsql = "UPDATE SYSDEFDSKTOP SET DSCSET = 'NUEVO ESCRITORIO' WHERE DSCSET = '" & VSESS & "'"
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)
     END IF


     Strsql = "Select * From SYSDEFDSKTOPDET WHERE CODSET = 0 ORDER BY OBJID "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  Do While (Not rs.EOF)

     VCOD = RS(2)

     VPRTY = SESSION(VCOD)    

     IF VPRTY <> "" THEN
        VPRTY = UCASE(VPRTY)
     END IF

     IF VPRTY = "" THEN

        VFORMAT = RS(6)

        SELECT CASE VFORMAT
               CASE "COLOR" 
                     VPRTY = "#FFFFFF"
               CASE "COLORFT" 
                     VPRTY = "#000080"
               CASE "SIZE" 
                     VPRTY = "8"
               CASE "FONT" 
                     VPRTY = "VERDANA"
               CASE "IMAGE" 
                     VPRTY = " "
               CASE "BOLD" 
                     VPRTY = "0"
               CASE "ITAL" 
                     VPRTY = "0"
               CASE "SUBR" 
                     VPRTY = "0"

        END SELECT 



     END IF

     strsql = "Insert into SYSDEFDSKTOPDET (CODSET,PRTCOD,PRTY1) "
     strsql = STRSQL & " values (" & VKEY & ",'" & VCOD & "','" & VPRTY & "')"
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    rs.Movenext
  Loop




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  Response.redirect "defescridet.asp"

%>