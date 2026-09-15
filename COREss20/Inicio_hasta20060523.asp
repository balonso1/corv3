<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorInit.inc" -->
<%

  set oIni = server.createobject("CorSecurityTCDC.Inicio")

  Set oIni.Request = Request
  Set oIni.Server = Server
  Set oIni.Response = Response
  Set oIni.Session = Session
  Set oIni.Ofv = Ofv
      oIni.Modal = Session("Modal")

  session("noved") = "NO"

  Opcion = Request.querystring("Opcion")

  if Opcion = "ISET" then
     Opcion = "I"
     Session("EXPLOG") = "N"
     Session("SETLOG") = "S"
  end if


  if Opcion = "I" then
  '   Conset(session("CONSET"))
  '   ofv.wstyles = CorStyler()    

     IF NOT (Session("Usuario") & "") <> "" THEN

        xddat1 =  request.querystring("ddat")
        xddat = REPLACE(cstr(xddat1),"___",chr(9))
        xddatV = split(xddat,chr(9),-1,1)  

        Session("Usuario")     = REPLACE(REPLACE(xddatV(0),"_"," "),"*","")
        Session("UsuarNT")     = REPLACE(REPLACE(xddatV(1),"_"," "),"*","")
        Session("DOMINIO")     = REPLACE(REPLACE(xddatV(2),"_"," "),"*","")
        SESSION("LOGIDUSRDET") = REPLACE(REPLACE(xddatV(3),"_"," "),"*","")

     END IF

     set cn2 = ofv.conectar(ofv.strconn2)

     XXXX = GETUSRLOGUPD("LOGOFF")

     call ofv.cerrarconn(cn2)


    sHTML = "<html><script language=javascript > function reinicia() { document.parentWindow.top.navigate('inicio.asp?opcion=IX'); } </script>"
    sHTML = shtml & "<body bgcolor=#f7eedd onload='reinicia();' ><p align=center ><font face=verdana size=2 color=#000080><b><br><br>Cor E-Solution Suite:<br>Logoff Request</font></body></html>"
    Response.write sHTML 
  elseif Opcion = "IX" then
    xcinit = corinit() 
   
    if  Session("Modal") = "MSNTT" and  Session("EXPLOG") = "S" then
        sHTML = DIRECTLOG()
    else
        sHTML = oIni.Inicio("")
    end if

    Response.write sHTML
  elseif Opcion = "V" then
    if Session("APLBGD") = "" then 
       Conset(session("CONSET"))
       ofv.wstyles = CorStyler()
    end if
    session("TINI") = "I"
    sHTML = oIni.Valida
    if len(shtml) > 2 then
       Response.write sHTML
    else
       If sHTML = "1" or sHTML = "2"  or sHTML = "3"  or sHTML = "6"  or sHTML = "7"  or sHTML = "8"  or sHTML = "9"  or sHTML = "11"  Then
          sHTML = oIni.Inicio(sHTML)
          Response.write sHTML
       ElseIf sHTML = "0" Then
          if Session("Usuario") = UCase(Request.Form("Usuarnt")) then
             session("TINI") = "Z"
          end if

'''''''' Log ingreso '''''
           set cn2 = ofv.conectar(ofv.strconn2)

           XXXX = GENUSRLOG()
           XXXX = GETUSRLOGUPD("LOGON")

           call ofv.cerrarconn(cn2)
'''''''''''''

          Response.redirect "Corinicio.asp?Opcion=IS"
       elseIf sHTML = "4" or sHTML = "5" Then

'''''''' Log ingreso '''''
           set cn2 = ofv.conectar(ofv.strconn2)

           XXXX = GENUSRLOG()
           XXXX = GETUSRLOGUPD("LOGON")

           call ofv.cerrarconn(cn2)
'''''''''''''
          Response.redirect "Corinicio.asp?Opcion=CP"
       Else
          Response.write sHTML
       End If
    end if 
  elseif Opcion = "IS" then
    xcinit = UsrInit()

    sHTML = oIni.iniciosesion
   ' sHTML = replace(sHTML,"<html>","<html><head><bgsound src=" & chr(34) & session("sound") &  chr(34) & "></head> ")
    Response.write sHTML
  elseif Opcion = "CP" then
    sHTML = oIni.CambiaPass("")
    Response.write sHTML
  elseif Opcion = "CPOK" then
    sHTML = oIni.CambiaPassOK
    if len(shtml) > 1 then
       Response.write sHTML
    else
       If sHTML = "0" Then
          Response.redirect "inicio.asp?Opcion=IS"
       Else
          sHTML = oIni.CambiaPass(sHTML)
          Response.write sHTML
       End If
    end if 
  end if

FUNCTION DIRECTLOG()

      vu = OFV.ValidarUsuario()
      If vu = "OK" Then

         fondo = " TOPMARGIN=1 LEFTMARGIN=1 ONLOAD='Inicio.submit();'   "

         UserNt = UCase(Request.ServerVariables("LOGON_USER"))
         UserNt = Mid(UserNt, InStrRev(UserNt, "\") + 1, Len(UserNt))

         DOMAINNT = UCase(Request.ServerVariables("LOGON_USER"))
         DOMAINNT = Left(DOMAINNT, InStr(DOMAINNT, "\")) & "\"
         DOMAINNT = Replace(DOMAINNT, "\", "")

         XHTML = OFV.MenuHeader("#FFFFFF", fondo) & Chr(13) & "<center><BR><BR><BR><BR><BR><BR><BR><BR>"

         XHTML = XHTML & "<font face=verdana size=4 color=#6699cc ><b>COR E-SOLUTION SUITE</B></FONT><BR><BR><BR>"


         sInput = "<BIG>VERIFICANDO ACCESO PARA<BR><BR>USUARIO:&nbsp;" & UserNt & "<BR><BR>DOMINIO:&nbsp;" & DOMAINNT

         Celda = OFV.gencelda("", "vt VALIGN=MIDDLE ", "center", "80%", "50%", "", "", sInput, "si")
         Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

         sInput1 = OFV.GenerarInput("usuarNt", UserNt, "HIDDEN", "", "80", "at", "no")
         sInput2 = OFV.GenerarInput("DOMAINNt", DOMAINNT, "HIDDEN", "", "80", "at", "no")
         sInput3 = OFV.GenerarInput("usuar", "", "HIDDEN", "", "62", "at", "no")
         sInput4 = OFV.GenerarInput("clave", "", "HIDDEN", "", "62", "at", "no")

         Celda = OFV.gencelda("", "ut", "center", "80%", "20", "", "", sInput1 &  sInput2 &  sInput3 &  sInput4, "si")
         Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

         Formulario = OFV.GenForm("Inicio", "", "Inicio.asp?Opcion=V", "", "post", "", Fila, "si")

         XHTML = XHTML & OFV.GenTabla("", "", "", "border=2 BGCOLOR=#BBDDFF ", "", "0", "", "0", "0", "", "", Formulario, "si")
         XHTML = XHTML & "</center></BODY></html>"

       end if

         DIRECTLOG = XHTML


END FUNCTION

%>
