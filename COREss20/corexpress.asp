<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorInit.inc" -->
<!-- #INCLUDE FILE="IncFile/CorSDK.inc" -->

<%

'  Session("EXPLOG") = "S"
'  Session("DELLOG") = "S"
  session("UN") = 1
'  Session("SegNivel") = "8"

' http://server-41/coresV3/coress20/corexpress.asp?opcion=IX&opt=CC&id=103-MNN-1&durl=codigo*103_tipo*MN_perfil*OBJARC

' http://server-41/coresV3/coress20/corexpress.asp?opt=CC-103-MN


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  set oIni = server.createobject("CorSecurityTCDC.Inicio")

  Set oIni.Request = Request
  Set oIni.Server = Server
  Set oIni.Response = Response
  Set oIni.Session = Session
  Set oIni.Ofv = Ofv
      oIni.Modal = Session("Modal")



''   xid     = Request.querystring("id")
   xoptin    = Request.querystring("opt") & "" 
''   xdurl   = Request.querystring("durl")
   xBus = Request.querystring("Xbus")
   
   if Xbus = "" then xBus="N"
    
   xoptopn = instr(xoptin,"-")

   xopt = left(xoptin,xoptopn-1)

   xoptin2    = right(xoptin,len(xoptin) - xoptopn)

   xoptopn2 = instr(xoptin2,"-")

   xincodobj = left(xoptin2,xoptopn2-1)

   xintipobj = right(xoptin2,len(xoptin2) - xoptopn2)


 '  xacceso = Request.querystring("durl")
 '  xacceso = replace(xacceso,"_","&")
 '  xacceso = replace(xacceso,"*","=")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  session("noved") = "NO"

  Opcion = Request.querystring("Opcion")

  if opcion = "" then opcion = "IX"




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
      SESSION("USREXPIR") = 0
   
      if  Session("Modal") = "MSNTT" and  Session("EXPLOG") = "S" then
          sHTML = DIRECTLOG()
      else
      ''  sHTML = oIni.Inicio("")
          sHTML = Inicio("")
      end if

      shtml = replace(shtml,"Inicio.asp?","Corexpress.asp?Xbus="& xbus & "&opt=" & xoptin & "&")

      Response.write sHTML
  elseif Opcion = "IXS" then

      Session("EXPLOG") = "N"
      Session("DELLOG") = "N"
      Session("ESPLOG") = "S"
      xcinit = corinit() 

      SESSION("CLIEMP") = 0
   
      ''  sHTML = oIni.Inicio("")
      sHTML = Inicio("")

      shtml = replace(shtml,"Inicio.asp?","Corexpress.asp?Xbus="& xbus & "&opt=" & xoptin & "&")


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






    if len(shtml) > 1 then
       Response.write sHTML
    else
       If sHTML = "1" or sHTML = "2"  or sHTML = "3"  or sHTML = "6"  or sHTML = "7"  or sHTML = "8"  or sHTML = "9"  or sHTML = "11"  Then

          sHTML = Inicio(sHTML)
          Response.write sHTML

       ElseIf sHTML = "0" Then


''''''''''''''''
       '       session("SegGrpLst") = session("SegGrp") & chr(9) & "CONS" & chr(9) & "RR"

''''''''''''''
              set XFCN = ofv.conectar(ofv.strconn0)
              XFStrSql = "Select sessexp From Grupos where id = '" & Session("SegGrp") & "' "
              set XFrs = ofv.crearconsultaEx(XFStrSql,XFCN,1,parametros)

              if  not XFrs.eof then
                  SESSION("USREXPIR") = "D"
                  if XFrs(0) <> "D" then SESSION("USREXPIR") = cint("0" & XFrs(0)) * 60
              else
                  SESSION("USREXPIR") = "D"
              end if
              call ofv.cerrarconsulta(XFrs)
              call ofv.cerrarconn(XFCN)
  



'''''''''''' Log ingreso '''''
             set cn2 = ofv.conectar(ofv.strconn2)

             XXXX = GENUSRLOG()
             XXXX = GETUSRLOGUPD("LOGON")

             call ofv.cerrarconn(cn2)
'''''''''''''

'''''''''''' CONFIGURA EMPRESA DEL USUARIO
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
             xcinit = UsrInit()


     '' 
     '''''''' CONFIGURACION GLOBAL MULTIEMPRESA
     ''
           set cn = ofv.conectar(ofv.strconn2)

     if SESSION("EMPMUN") = "S" THEN
           StrSql = "Select * From Clienteplan WHERE ID = 1"
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              SESSION("GLBLOG") = RS(8)
              SESSION("GLBMUN") = RS(9)
           ELSE
              SESSION("GLBLOG") = 1
              SESSION("GLBMUN") = 1
           END IF

     ELSE
        SESSION("GLBLOG") = 1
        SESSION("GLBMUN") = 1
     END IF   
                
           ' response.write request.querystring  & "2222222222222"
            if xBus="S" then
              xincodobj=BuscarId(xincodobj)
            end if
             call ofv.cerrarconn(cn)
             xacceso = "codigo*" & xincodobj & "_tipo*" & xintipobj & "_perfil*OBJARC"
             xacceso = replace(xacceso,"_","&")
             xacceso = replace(xacceso,"*","=")


             if xopt = "CC" then

                session("ModMain") = "CL" 
                
                 'response.write "../Coremanager/Consultar.asp?ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
                 Response.redirect "../Coremanager/Consultar.asp?ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
               
             else

                session("ModMain") = "ED" 

                if xopt = "EV" then Session("Form2") = "TEMAS"
                if xopt = "ED" then Session("Form2") = "DESARROLLO"
                if xopt = "EA" then Session("Form2") = "PARA APROBAR"

                editen(xoptin2 & "N-1")
                Response.redirect "../Coremanager/editar.asp?ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
              end if
        Else
            Response.write sHTML
        End If




    end if

 end if

'''''''''''''''''''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'''''''''''''''''

FUNCTION DIRECTLOG()

    if  Session("DELLOG") = "S" then

        vu = OFV.ValidarUsuario()

        vu = "OK"

        usuario=UserLogon()

       ' usuario="ADM"
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
        Session("SegNivel")  = 8
        Session("SegGrp")    = "MAS"
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function editen(sElemento)


  n = instr(sElemento,"-")
  nCodobjeto = left(sElemento,n-1)
  n2 = instr(n+1,sElemento,"-")
  sTipobjeto = mid(sElemento,n+1,n2-(n+1))
  sEstado = mid(sElemento,n2+1)
  sEsqSup = trim(sTipobjeto)
  sEsqSup = mid(sEsqSup,len(sEsqSup))
  sTipobjeto = left(sTipobjeto,len(sTipobjeto)-1)

  ofv.ObtenerAtributos Session("Form2"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  scodigo = nCodobjeto
  stipo = sTipobjeto

  srestric = "N"


  if sEsqSup <> "Y" then
     redim preserve sPath(0)
     i = 1
     do

       strsql = "select * from csbarr where codojasoc = " & sCodigo
       strsql = strsql & " and tipobjasoc = '" & sTipo & "'"
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      if not rs.eof then
        redim preserve sPath(i)
        sCodigo = rs(1)
        sTipo = rs(2)
        if not isnull(isnull(rs(0))) then
          sPath(i) = rs(0)
        else
          i = i + 1
          exit do
        end if
        i = i + 1
      else
        exit do
      end if
      call ofv.cerrarconsulta(rs)
    loop

  end if

    strsql = "SELECT * FROM PRODSERV where codprodserv = " & sCodigo
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then
       if rs("segaceso") <> 0 or rs("segactual") <> 0 then srestric = "S"
    end if
    call ofv.cerrarconsulta(rs)

    if srestric = "S" then
            StrSql = "Select * From subtemaskr "
            StrSql = StrSql & "Where codprodserv = " & sCodigo & " and tiposeg = 'T' "
            StrSql = StrSql & " and (grpusr = 'USR-" & Ucase(session("usuario")) & "' "

            if  Session("DELLOG") = "S" then

                XMFTGRP = split(session("SegGrpLst"),chr(9))
                xcount  = ubound(XMFTGRP) + 1
                XMFSEP  = "" 
                StrSql = StrSql & " or grpusr in ("

                for XMFI = 0 to ubound(XMFTGRP)
                    StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                    XMFSEP  = "," 
                next

                StrSql = StrSql & ") "

            else
                StrSql = StrSql & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
            end if 

            StrSql = StrSql & ")"

       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs.eof then
          srestric = "N"
       else
          srestric = "S"
       end if
       call ofv.cerrarconsulta(rs)
    end if



  if srestric = "S" then
     bAgregar = 0
     bModificar = 0
     bSuprimir = 0
  end if

  session("Nomodificar") = "N"



  if isnumeric(sEstado) then
     if Session("Form2") <> "TEMAS" And sEstado = 3 then
        bSuprimir = False
        session("Nomodificar") = "S"
     end if
     if Session("Form2") <> "DESARROLLO" And sEstado = 1 then
        bSuprimir = False
        session("Nomodificar") = "S"
     end if
     if Session("Form2") <> "PARA APROBAR" And sEstado = 2 then
        bSuprimir = False
        session("Nomodificar") = "S"
     end if
  end if

  if Session("Form2") = "TEMAS" or Session("Form2") = "ELIMINADAS" then
     bAgregar = False
  end if

  session("EsqSup") = sEsqSup
  Session("Segct") = "S"
  Session("Segmo") = "S"
  Session("Segco") = "S"
  Session("Segce") = "S"

  if sEsqSup = "Y" then
     xtipo = "SUBTEMAS"
     strsql = "SELECT DESCRI FROM PRODSERV where codprodserv = " & NCODOBJETO
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then sNombredOC = rs(0)


     Strsql = "SELECT SUBTEMA.* FROM PRODSERV, SUBTEMA where "
     Strsql = Strsql & "PRODSERV.marcaps = SUBTEMA.tipobjeto and "
     Strsql = Strsql & " prodserv.marcaps = '" & sTipobjeto & "'"
  else
     xtipo = "ESQUEMAS"
     strsql = "SELECT DESCRI FROM OBJETOS where CODOBJETO = " & NCODOBJETO
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then sNombredOC = rs(0)

     Strsql = "Select * From Tipobjeto Where tipobjeto='" & sTipobjeto & "'"
  end if

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  x = 0
  redim Niveles(1)
  if not rs.eof then
     sNombre  = rs(1)
     sTIPOdOC = rs(1)
     sPerfil = rs(3)
     for z = 5 to 14
         if rs(z) Then
            x = x + 1
            redim preserve Niveles(x)
            Niveles(x) = right(rs(z).name,1)
         end if
     next
   end if

   call ofv.cerrarconsulta(rs)


   if sEsqSup = "N" then

      set cn = ofv.conectar(ofv.strconn2)

      Strsql = "Select * From OBJSDK Where tipobjeto='" & sTipobjeto & "' and codobjeto = "  & nCodobjeto
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rs.eof then


        if Session("Usuario") <> rs(7) then

           if rs(3) then Session("Segct") = "N"
           if rs(4) then Session("Segmo") = "N"
           if rs(5) then Session("Segco") = "N"
           if rs(6) then Session("Segce") = "N"

           call ofv.cerrarconsulta(rs)
           Strsql = "Select * From OBJSDKDET Where tipobjeto='" & sTipobjeto
           Strsql = Strsql & "' and codobjeto = "  & nCodobjeto
           Strsql = Strsql & " and ("
           Strsql = Strsql & "(grpusr = 'USR-" & Session("Usuario") & "') "

           if  Session("DELLOG") = "S" then

                XMFTGRP = split(session("SegGrpLst"),chr(9))
                xcount  = ubound(XMFTGRP) + 1
                XMFSEP  = "" 
                StrSql = StrSql & " or ( grpusr in ("

                for XMFI = 0 to ubound(XMFTGRP)
                    StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                    XMFSEP  = "," 
                next

                StrSql = StrSql & ")) "

           else
               Strsql = Strsql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')"
           end if 

           Strsql = Strsql & ") "
           set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           do until rs.eof

              select case rs(3)
                   case "1"
                        Session("Segct") = "S"
                        Session("Segmo") = "S"
                        Session("Segco") = "S"
                        Session("Segce") = "S"

                   case "2"
                        Session("Segmo") = "S"
                        Session("Segco") = "S"

                   case "3"
                        Session("Segco") = "S"

                   case "4"
                        Session("Segco") = "S"
                        Session("Segce") = "S"

              end select
              rs.movenext
          loop
      end if
    end if
    call ofv.cerrarconsulta(rs)

    if srestric = "S" then
       Session("Segct") = "N"
       Session("Segmo") = "N"
       Session("Segce") = "N"
    end if


    if Session("Segce") = "S" and session("Nomodificar") <> "S" then
       session("Nomodificarce") = "N"
    else
       session("Nomodificarce") = "S"
    end if

    if Session("Segct") <> "S" and Session("Segmo") <> "S" then
       session("Nomodificar") = "S"
    end if

    BCOPYSPC = "N"

    IF  SESSION("WMATRMEMP") = "S" AND  SESSION("WMATRMEMPACT") = "N" THEN

        bAgregar = 0
        bSuprimir = 0
        session("Nomodificar") = "S"
        session("Nomodificarce") = "S"

        IF bModificar and Session("Segct") = "S" THEN
           BCOPYSPC = "S"
        END IF

        bModificar = 0

        Session("Segct") = "N"
        Session("Segmo") = "N"
        Session("Segce") = "N"

    END IF

  end if

  call ofv.cerrarconn(cn)

end function

function BuscarId(id)
 dim strsql,rsid
  
  strsql="select objrdo from objetos where codobjeto=0" &id
  'response.write strsql
  set rsid = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 
  Orem= rsid(0)
  call ofv.cerrarconsulta(rsid)
  
   strsql="select a.objrdo,a.codobjeto, b.estado from objetos a,objprty b where a.codobjeto=0" &Orem
   strsql= strsql & " and a.codobjeto=b.codobjeto "
  ' response.write "<br>" & strsql &"<BR>"
   set rsid = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if  not rsid.eof then
    do while not rsid.eof   
      Orem=rsid(0)
      if (rsid("codobjeto")= rsid("objrdo")) or (rsid("estado") <> 4) then exit do
     strsql="select a.objrdo, b.estado ,a.codobjeto  from objetos a,objprty b where a.codobjeto=0" & Orem
     strsql=  strsql & " and a.codobjeto=b.codobjeto "
     'response.write "<br>" & strsql &"<BR>"
     call ofv.cerrarconsulta(rsid)
     set rsid = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    
    loop
    
      
     if not rsid.eof  then 
       if rsid("estado") <> 4 then Orem=rsid("codobjeto")
     
     end if
     
   end if  
   
    if Orem=0 then
     BuscarId=id
    else
     BuscarId=Orem
    end if
    
   'RESPONSE.WRITE "222" & BuscarId
end function
%>
