<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorInit.inc" -->
<!-- #INCLUDE FILE="IncFile/CorSDK.inc" -->
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

   if session("WEBACC")="S" and Opcion <> "IXWV" then
      Opcion="IXW"
   end if


  if Opcion = "ISET" then
     Opcion = "I"
     Session("EXPLOG") = "N"
     Session("SETLOG") = "S"
  end if


  if Opcion = "I" then
     Conset(session("CONSET"))
     ofv.wstyles = CorStyler()    

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

     sHTML = LogFinal()
     Response.write sHTML 

  elseif Opcion = "IX" then
    xcinit = corinit() 

    SESSION("CLIEMP") = 0

    SESSION("USREXPIR") = "D"
   
    if  Session("Modal") = "MSNTT" and  Session("EXPLOG") = "S" then
        sHTML = DIRECTLOG()
    else
      ''  sHTML = oIni.Inicio("")
        sHTML = Inicio("")
    end if
  
    
    
    Response.write sHTML
    
   elseif Opcion = "IXW" then
      xcinit = corinit() 
  
      SESSION("CLIEMP") = 0
  
      SESSION("USREXPIR") = "D"
     
      
       sHTML = webvalida()
     
      
      
      Response.write sHTML
      
    elseif Opcion = "IXWV" then
       
        
         vu = webvalidabt()
       
        
        If vu = "OK" Then 
          response.redirect "inicio.asp?opcion=VD"
        else
         response.redirect "inicio.asp?opcion=IXW"
        end if
        
        Response.write sHTML
      
    
  elseif Opcion = "IXS" then

         Session("EXPLOG") = "N"
         Session("DELLOG") = "N"
         Session("ESPLOG") = "S"

    xcinit = corinit() 

    SESSION("CLIEMP") = 0
   
      ''  sHTML = oIni.Inicio("")
        sHTML = Inicio("")


    Response.write sHTML
  elseif Opcion = "V" or Opcion = "VD" then
    if Session("APLBGD") = "" then 
       Conset(session("CONSET"))
       ofv.wstyles = CorStyler()
    end if
    session("TINI") = "I"


    if  Opcion = "VD" then 
        SHTML = "0"
    else
        sHTML = oIni.Valida
    end if

    if len(shtml) > 2 then
       Response.write sHTML
    else
       If sHTML = "1" or sHTML = "2"  or sHTML = "3"  or sHTML = "6"  or sHTML = "7"  or sHTML = "8"  or sHTML = "9"  or sHTML = "11"  Then
         ''' sHTML = oIni.Inicio(sHTML)
          sHTML = Inicio(sHTML)
          Response.write sHTML
       ElseIf sHTML = "0" Then
        '''  response.buffer=true
          if Session("Usuario") = UCase(Request.Form("Usuarnt")) then
             session("TINI") = "Z"
          end if



''''''''''''''
  set XFCN = ofv.conectar(ofv.strconn0)


  XFStrSql = "Select sessexp From Grupos where id = '" & Session("SegGrp") & "' "

''''  response.write XFStrSql
  set XFrs = ofv.crearconsultaEx(XFStrSql,XFCN,1,parametros)

  if  not XFrs.eof then
      SESSION("USREXPIR") = "D"

       if XFrs(0) <> "D" then SESSION("USREXPIR") = cint("0" & XFrs(0)) * 60
  else
      SESSION("USREXPIR") = "D"
  end if
  call ofv.cerrarconsulta(XFrs)
  call ofv.cerrarconn(XFCN)
  
 ''''''''''' call ChkGrp()


'''''''' Log ingreso '''''
           set cn2 = ofv.conectar(ofv.strconn2)

           XXXX = GENUSRLOG()
           XXXX = GETUSRLOGUPD("LOGON")

           call ofv.cerrarconn(cn2)
'''''''''''''

'''''''''''''' CONFIGURA EMPRESA DEL USUARIO
      IF SESSION("CLIEMP") = 0 THEN
           SESSION("SETEMP") = "N"
           set cn = ofv.conectar(ofv.strconn2)

           StrSql = "Select SITE From USUARIOS WHERE ID = '" & SESSION("USUARIO") & "' "
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              SESSION("CLIEMP") = clng(RS(0))
           ELSE
              SESSION("CLIEMP") = 1
           END IF

           StrSql = "Select * From Clienteplan WHERE ID = " & SESSION("CLIEMP")
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              Session("Cliente") = rs(1)
              Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
              if (rs(11) & "" ) <> "" then Session("CliLogoP") = "../COREss20/Empresa/" & rs(11)
              Session("USRSET")  = RS(10)
           END IF

           StrSql = "Select * From SYSDEFDSKTOP WHERE OBJID = 0" & Session("USRSET")
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              IF rs(2) <> "SI" THEN
                 XCONFUSRSET = CONUSRSET()
              else
                Conset(session("CONSET"))
              END IF
           END IF

           call ofv.cerrarconn(cn)
      END IF
''''''''''''''

           'Response.redirect "http://server-41/coresv3/coress20/Corinicio.asp?Opcion=IS"
         ' Response.redirect "Corinicio.asp?Opcion=IS"
          
          xshtml = "<html><body ><form name='ejec' method='post' action='Corinicio.asp?Opcion=IS' style='visibility:hidden;' ><input  type='submit' value='n'   ></form>"
          xshtml = xshtml & "<script language='javascript'>document.ejec.submit();</script></body></html>"
          
          Response.write xshtml
          
          
          
       elseIf sHTML = "4" or sHTML = "5" Then

'''''''' Log ingreso '''''
           set cn2 = ofv.conectar(ofv.strconn2)

           XXXX = GENUSRLOG()
           XXXX = GETUSRLOGUPD("LOGON")

           call ofv.cerrarconn(cn2)
'''''''''''''

'''''''''''''' CONFIGURA EMPRESA DEL USUARIO
      IF SESSION("CLIEMP") = 0 THEN
           SESSION("SETEMP") = "N"
           set cn = ofv.conectar(ofv.strconn2)

           StrSql = "Select SITE From USUARIOS WHERE ID = '" & SESSION("USUARIO") & "' "
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              SESSION("CLIEMP") = clng(RS(0))
           ELSE
              SESSION("CLIEMP") = 1
           END IF

           StrSql = "Select * From Clienteplan WHERE ID = " & SESSION("CLIEMP")
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              Session("Cliente") = rs(1)
              Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
              if (rs(11) & "" ) <> "" then Session("CliLogoP") = "../COREss20/Empresa/" & rs(11)
              Session("USRSET")  = RS(10)
           END IF

           StrSql = "Select * From SYSDEFDSKTOP WHERE OBJID = 0" & Session("USRSET")
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              IF rs(2) <> "SI" THEN
                 XCONFUSRSET = CONUSRSET()
              else
                Conset(session("CONSET"))
              END IF
           END IF

           call ofv.cerrarconn(cn)
      END IF
''''''''''''''
          
          Response.redirect "Corinicio.asp?Opcion=CP"
       Else
          Response.write sHTML
       End If
    end if 
  elseif Opcion = "IS" then

    xcinit = UsrInit()

   ''' sHTML = oIni.iniciosesion
    sHTML = iniciosesion

'''''''''''''' CONFIGURA EMPRESA DEL USUARIO
    '  IF SESSION("CLIEMP") = 0 THEN
           SESSION("SETEMP") = "N"
           set cn = ofv.conectar(ofv.strconn2)

           StrSql = "Select SITE From USUARIOS WHERE ID = '" & SESSION("USUARIO") & "' "
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              SESSION("CLIEMP") = clng(RS(0))
           ELSE
              SESSION("CLIEMP") = 1
           END IF

           StrSql = "Select * From Clienteplan WHERE ID = " & SESSION("CLIEMP")
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              Session("Cliente") = rs(1)
              Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
              if (rs(11) & "" ) <> "" then Session("CliLogoP") = "../COREss20/Empresa/" & rs(11)
              Session("USRSET")  = RS(10)
           END IF

           StrSql = "Select * From SYSDEFDSKTOP WHERE OBJID = 0" & Session("USRSET")
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              IF rs(2) <> "SI" THEN
                 XCONFUSRSET = CONUSRSET()
              else
                Conset(session("CONSET"))
              END IF
           END IF


    '  END IF
''''''''''''''
     '' 
     '''''''' CONFIGURACION GLOBAL MULTIEMPRESA
     ''
     if SESSION("EMPMUN") = "S" THEN
           StrSql = "Select * From Clienteplan WHERE ID = 1"
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              SESSION("GLBLOG") = RS(8)
              SESSION("GLBMUN") = clng(RS(9))
           ELSE
              SESSION("GLBLOG") = 1
              SESSION("GLBMUN") = 1
           END IF

     ELSE
        SESSION("GLBLOG") = 1
        SESSION("GLBMUN") = 1
     END IF   
           call ofv.cerrarconn(cn)

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

    if  Session("DELLOG") = "S" then

        vu = OFV.ValidarUsuario()

        vu = "OK"

        usuario=UserLogon()
        call getInfoUser(Usuario)

        Session("Usuario")   = session("SGU_userid")
        Session("UsrOpcion") = session("SGU_nombre")
        Session("SkrDefT")   = "0"
        Session("SkrDefE")   = "0"
        Session("SkrDefI")   = "0"
        Session("UsrPmin")   = "0"
        Session("UsrPmax")   = "0"
        Session("UsrHist")   = "0"
        Session("SegUsuario") = session("SGU_nombre")
        Session("SegNivel")  = cdbl("8")
        Session("SegGrp")    = "DEF"
        Session("SegPerfil") = Session("Usuario")
        Session("UsuarNT")   = Session("Usuario")

        Session("UsrSector") = session("SGU_idcostcenter")
        Session("UsrSecDsc") = session("SGU_descrisector")
        Session("UsrPuesto") = session("SGU_codigopuesto")
        Session("UsrPstDsc") = session("SGU_descripuesto")
        Session("UsrMail")   = session("SGU_mail")

        call getGroupsUser(Usuario)

        Session("SegGrpOpt")    = "S"

        Session("SegGrpLst")    = session("SGU_grupos")

        call ChkGrp()

        XFURL = "Inicio.asp?Opcion=VD"

    else
        vu    = OFV.ValidarUsuario()
        XFURL = "Inicio.asp?Opcion=V"
    end if


      If vu = "OK" Then 

         fondo = " TOPMARGIN=1 LEFTMARGIN=1 ONLOAD='document.INIT.submit();'   "

         UserNt = UCase(Request.ServerVariables("LOGON_USER"))
         UserNt = Mid(UserNt, InStrRev(UserNt, "\") + 1, Len(UserNt))

         DOMAINNT = UCase(Request.ServerVariables("LOGON_USER"))
         DOMAINNT = Left(DOMAINNT, InStr(DOMAINNT, "\")) & "\"
         DOMAINNT = Replace(DOMAINNT, "\", "")

         XHTML = OFV.MenuHeader("", fondo) & Chr(13) & "<center><BR><BR><BR><BR><BR>"

         XHTML = XHTML & "<font face=verdana size=5 color=#6699cc ><b>COR E-SOLUTION SUITE</B></FONT><BR><BR><BR>"


         sInput = "<BIG>VERIFICANDO ACCESO PARA</BIG>"

         Celda = OFV.gencelda("", "ttt VALIGN=MIDDLE ", "center", "80%", "50%", "", "", sInput, "si")
         Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

         sInput1 = OFV.GenerarInput("usuarNt", UserNt, "HIDDEN", "", "80", "at", "no")
         sInput2 = OFV.GenerarInput("DOMAINNt", DOMAINNT, "HIDDEN", "", "80", "at", "no")
         sInput3 = OFV.GenerarInput("usuar", "", "HIDDEN", "", "62", "at", "no")
         sInput4 = OFV.GenerarInput("clave", "", "HIDDEN", "", "62", "at", "no")

         sInput = "<BR><BR>USUARIO:&nbsp;" & UserNt & "<BR><BR>DOMINIO:&nbsp;" & DOMAINNT & "<BR><BR>"

         Celda = OFV.gencelda("", "uts", "center", "80%", "20", "", "", sInput & sInput1 &  sInput2 &  sInput3 &  sInput4, "si")
         Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

         Formulario = OFV.GenForm("INIT", "", XFURL, "", "post", "", Fila, "si")

         XHTML = XHTML & OFV.GenTabla("", "AAA STYLE='BORDER-STYLE:solid;border-width:2;border-color:#000080;' ", "", "", "", "0", "", "0", "0", "", "", Formulario, "si")
         XHTML = XHTML & "</center></BODY></html>"

       end if

         DIRECTLOG = XHTML


END FUNCTION


FUNCTION WebValida() 

        vu    = OFV.ValidarUsuario()
        XFURL = "Inicio.asp?Opcion=IXWV"
  

         fondo = " TOPMARGIN=1 LEFTMARGIN=1    "

       

         
          set cn = ofv.conectar(ofv.strconn2)
	     StrSql = "Select * From Clienteplan WHERE ID =1 " 
	        Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	        IF NOT RS.EOF THEN
	 	  Session("Cliente") = rs(1)
	 	  Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
	 	  if (rs(11) & "" ) <> "" then Session("CliLogoP") = "../COREss20/Empresa/" & rs(11)
	 	  Session("USRSET")  = RS(10)
	        END IF
	  call ofv.cerrarconn(cn)

 
         XHTML =  "<img src=" & Session("CliLogo") &">"
         XHTML = XHTML & OFV.MenuHeader("", fondo) & Chr(13) & "<center><BR><BR><BR><BR><BR>"
         XHTML = XHTML & "<font face=verdana size=5 color=#6699cc ><b>COR E-MANAGER </B></FONT><BR>"
         XHTML = XHTML & "<font face=verdana size=2 color=#6699cc ><b>BASE DE CONOCIMIENTOS </B></FONT><BR><BR><BR>"

         
         
          FILA=""
          Celda = OFV.gencelda("", "ttt ", "center", "80%", "20", "", "", "<B>USUARIO:</B>", "si")
	  Fila = Fila & OFV.GenRow("u11", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
	     
	  sInput = OFV.GenerarInput("usuar", "", "agregar", "", "40", "dt", "no")
	  Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
          Fila = Fila & OFV.GenRow("u12", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", "<B>PASSWORD:</B>", "si")
          Fila = Fila & OFV.GenRow("u13", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")

          sInput = OFV.GenerarInput("clave", "", "password", "", "40", "dt", "no")
          Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
          Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
              
          sInput = OFV.GenerarInput("xxx", "login", "submit", "", "62", "bt", "no")
          Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	  Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
	  
	  sInput =session("mensajeWeb")
	  Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	  Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
	  
          
          Formulario =OFV.GenForm("INIT", "", XFURL, "", "post", "", Fila , "si")

         XHTML = XHTML & OFV.GenTabla("", "AAA STYLE='BORDER-STYLE:solid;border-width:2;border-color:#000080;' ", "", "", "", "0", "", "0", "0", "", "", Formulario , "si")
         XHTML = XHTML & "</center></BODY></html>"

      

         WebValida = XHTML


END FUNCTION


function ChkGrp()

'    SESSION("USREXPIR") = 50

  SESSION("USREXPIR") = "0"

  set XFCN = ofv.conectar(ofv.strconn0)

  XFStrSql = "Select id,sessexp From Grupos where "

  session("SegGrpLst") = Session("SegGrp") & chr(9) & session("SegGrpLst")


  XMFTGRP = split(session("SegGrpLst"),chr(9))
  XMFSEP  = "" 
  XFStrSql = XFStrSql & " id in ("

  for XMFI = 0 to ubound(XMFTGRP)
      XFStrSql  = XFStrSql  & XMFSEP & "'" & replace(XMFTGRP(XMFI) & "","'","''") & "'"
      XMFSEP  = "," 
  next

  XFStrSql = XFStrSql & ") "
  XFStrSql = XFStrSql & " Order By id"
''''  response.write XFStrSql
  set XFrs = ofv.crearconsultaEx(XFStrSql,XFCN,1,parametros)

  if  not XFrs.eof then
      XMFSEP               = ""
      session("SegGrpLst") = ""
      do until XFRS.EOF 

         if XFrs(1) <> "D" and SESSION("USREXPIR") <> "D" then
            if (cint("0" & XFrs(1)) * 60) > cint(SESSION("USREXPIR")) then
               SESSION("USREXPIR") = cint("0" & XFrs(1)) * 60
            end if
          else
           SESSION("USREXPIR") = "D"
          end if  

         session("SegGrpLst") = session("SegGrpLst") & XMFSEP & XFRS(0)
         XMFSEP  = chr(9)
      XFRS.MOVENEXT
      LOOP     

      IF  session("SegGrpLst") = "" THEN session("SegGrpLst") = Session("SegGrp")

  else
      session("SegGrpLst") = Session("SegGrp")
  end if
  call ofv.cerrarconsulta(XFrs)
  call ofv.cerrarconn(XFCN)
 


end function

%>
