<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)
  set cn4 = ofv.conectar(ofv.strconn6)


  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")
 WKFPSC = request.querystring("WKFPSC")

 if WKFPSC = "" then
    WKFPSC = session("WKFPSC")
 else
    session("WKFPSC") = WKFPSC
 end if 


    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenF = "<B><FONT FACE=webDINGS COLOR=#993333 SIZE=5>q</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
    sImagenSI = "<B><FONT FACE=WINGDINGS COLOR=#0000aa SIZE=4>&#0236;</FONT></B>"
    sImagenSF = "<B><FONT FACE=WINGDINGS COLOR=#000000 SIZE=4>&#0237;</FONT></B>"
    sImagenSA = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=5>&#0252;</FONT></B>"
    sImagenSR = "<B><FONT FACE=WINGDINGS COLOR=#aa0000 SIZE=5>&#0251;</FONT></B>"
    sImagenPA = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=4>&#0231;</FONT></B>"
    sImagenPS = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=4>&#0232;</FONT></B>"


    formularioA = ""
    formularioB = ""
    formularioC = ""


  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & session("WKF") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
'     IF RS(3) THEN
'        bAgregar = FALSE
'        bModificar = FALSE
'        bSuprimir = FALSE
'     END IF


    sInput = "CIRCUITO:"
    celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

    sInput = RS(1) & " - <B>" & RS(2) & "</B>"
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    OBJTIPO = RS(6)

    SELECT CASE OBJTIPO

           CASE "REQ"
                ESTTABLA = "REQESTADOS"
                ESTCLAVE = "ID"
                ESTORDER = "   Order By DESCRI"
                SET CNX  = CN4

           CASE "COR"
                ESTTABLA = "ESTADOS"
                ESTCLAVE = "OBJID"
                ESTORDER = "  Order By ESTADO"
                SET CNX  = CN3

           CASE ELSE
                ESTTABLA = "ESTADOS"
                ESTCLAVE = "ID"
                ESTORDER = "   Order By DESCRI"
                SET CNX  = CN

    END SELECT


  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")


  END IF



    celdas = ""
    xrespo =  ""

  IF RS(5) = "M" THEN

    sInput = "RESPONSABLE:"
    celdas = ofv.GenCelda("","vt","","10%","","","",sInput,"si")

  StrSql = "Select * From WKFPASOS WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then


        if rs1(6)  then sInput = "ORIGINANTE DEL CIRCUITO"



        if rs1(9)  then sInput = "ORIGINANTE DEL CIRCUITO"



        if rs1(7) OR RS1(8) then
           if rs1(12) = "U" then sInput = "A ASIGNAR EN EL ORIGEN"
           if rs1(12) = "R" then sInput = "LISTA DE RESPONSABLES AUTOMATICA"
         END IF




  END IF



    celdas = celdas & ofv.GenCelda("","vt COLSPAN=2","","","","","","<b>" & sInput & "</b>","si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    xrespo =  filas

  ELSE

  StrSql = "Select * From WKFPSCRESP WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF RSX.EOF THEN
     GENWKFPASRESP()
     StrSql = "Select * From WKFPSCRESP WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
     set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF



  if not rsX.eof then


    sInput = "RESPONSABLE:"
    celdas = ofv.GenCelda("","vt","","10%","","","",sInput,"si")

       sInputD = "S/D"
    StrSql = "Select * From WKFRESP WHERE ID = " & RSX(3)
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then


        if rs(2) = "USR" then
           sInputT = "USUARIO: "
           strsql = "Select * From USUARIOS where ID = '" & rs(3) & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInputD = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
        end if
        if rs(2) = "GRP" then
           sInputT = "GRUPO: "
           strsql = "Select * From GRUPOS where ID = '" & rs(3) & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInputD = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
        end if
        if rs(2) = "SEC" then
           sInputT = "SECTOR: "
           strsql = "Select * From ECO where ECO = " & rs(3)
           set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
           if not rsniv.eof then sInputD = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
        end if
    ELSE
       sInputD = "SIN ASIGNAR"
    END IF

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href=WKFPSCRESP.asp?WKFPSC=" & session("WKFPSC")
        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 


    celdas = celdas & ofv.GenCelda("","vt","","","","","","<b>&nbsp;" & sInputT & sInputD & "&nbsp;" & sinputK & "</b>","si")

 ''      sinput = "<a  href=WKFPSCRESP.asp?WKFPSC=" & session("WKFPSC")
 ''      sinput = sinput  & " ><b><U>MODIFICAR</U></b></a>" 
 ''      celdas = celdas & ofv.GenCelda("","tt","center","15%","","","",sInput,"si")

       celdas = celdas & ofv.GenCelda("","ut","center","15%","","","","&nbsp;","si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    xrespo =  filas


 end if

 END IF

  call ofv.cerrarconsulta(rsX)
'  call ofv.cerrarconn(cn2)
'  call ofv.cerrarconn(cn3)


  StrSql = "Select * From WKFPASOS WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

    SFORM = "WKF00"

    sInput = "PASO:"
    celdas = ofv.GenCelda("","vt","","10%","","","",sInput,"si")

       if bModificar then
          sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sform,"70","dt STYLE='FONT-WEIGHT:BOLD' ","")
       else
          sInput = ofv.GenerarInput("",rs(4),"readonly","","70","dt STYLE='FONT-WEIGHT:BOLD' ","")
       end if

    sInput = RS(3) & ".&nbsp;" & " - " & sInput
    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

 ''   sinput = "<a class=nt href=WKFPASOS.asp"
 ''   sinput = sinput  & " >" & sImagenF & "&nbsp;<b><big>Volver&nbsp;al&nbsp;Circuito</big></b></a>" 

        sinputK = "<a class=nt style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href=WKFPASOS.asp"
        sinputK = sinputK  & "  TITLE='Volver&nbsp;al&nbsp;Circuito'  ><b>" & sImagenF & "&nbsp;<b><big>Volver&nbsp;al&nbsp;Circuito</big></b></b></a>" 


    celdas = celdas & ofv.GenCelda("","vt","CENTER","15%","","","",sInputK,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOS"
      formulario = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = ofv.GenForm("","","","","post","",filas,"si")
    end if

    XRESPO2 = formulario

'' sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si") & "<BR>"

''  sHTML = sHTML  & XRESPO

    SFORM = "WKF02"

    sInput = "ITEM:"
    celdas = ofv.GenCelda("","vt","","10%","","","",sInput,"si")


        if rs(6)  then sInput = "<b>" & sImagenPI & "</b>&nbsp;" & "PASO DE INICIO"
        if rs(9)  then sInput = "<b>" & sImagenPF & "</b>&nbsp;" & "PASO DE FIN"
        if rs(7) OR RS(8) then sInput = "<b>" & sImagenPC & "</b>&nbsp;" & "PASO DE CIRCUITO"
      ''  sInput = ofv.GenerarInput("",sInput,"readonly","","20","dt","")



    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")


  ''     sinput = "<a  href=WKFPSCPRTY.asp?WKFPSC=" & RS(2)
  ''     sinput = sinput  & " ><b><U>DATOS</U></b></a>" 
  ''     celdas = celdas & ofv.GenCelda("","tt","center","15%","","","",sInput,"si")

       celdas = celdas & ofv.GenCelda("","ut","center","15%","","","","&nbsp;","si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")




    if bModificar then
      sAccion = "asp/ActualizarWK.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOS"
      formulario = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = ofv.GenForm("","","","","post","",filas,"si")
    end if

    XRESPO3 = formulario



  sHTML = sHTML  & ofv.GenTabla("","aaa style='filter:shadow(color=#cccccc,direction=125);' ","","100%","","0","","0","0","","",XRESPO2 & XRESPO & XRESPO3,"si") & "<BR>"


  Xpanel = GenEntries()


formulario = ""


  formulario =   ofv.GenTabla("","","","98%","","0","","0","0","","",Xpanel,"si")
  celdas = ofv.GenCelda("","","center","","","","",formulario,"si")
  formulario = ofv.GenRow("","","","","","","",celdas,"si")
''  celdas = ofv.GenCelda("","","center","","","","","aaaaaaaa&nbsp;","si")
''  formulario = formulario & ofv.GenRow("","","","","","","",celdas,"si")
  sHTML = sHTML  & ofv.GenTabla("","aaa  style='filter:shadow(color=#cccccc,direction=125);' ","","100%","","0","","0","0","","",formulario,"si") & "<BR>"


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




  'Botonera
  Agregar = ""


  Buscar = ""
  

''  sHTML = sHTML  & "<HR WIDTH=100% SIZE=1 COLOR=#993333>"





''' ANTERIOR
    STAB = "&nbsp;"
  StrSql = "Select * From WKFPSCTPC WHERE PASOSIG = '" & session("WKFPSC") & "' "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS1.EOF THEN
    STAB = ""
     DO WHILE NOT RS1.EOF

        sInput = sImagenPA & "&nbsp;&nbsp;" & GENWKFPASO(RS1(2))
        
        STABC = ofv.GenCelda("","ut  ","","","","","",sInput,"si")
        STAB = STAB & ofv.GenRow("","","","","","","",STABC,"si")
     RS1.MOVENEXT
     LOOP
    STAB = ofv.GenTabla("","","","80%","","0","","0","0","","",STAB,"si")
  END IF

''' ANTERIOR
    celdas = ofv.GenCelda("","vt STYLE='BORDER-bottom-STYLE:solid;BORDER-bottom-WIDTH:2;BORDER-bottom-color:#993333;' colspan=2 ","CENTER","","","","","&nbsp;","si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")

    celdas = ofv.GenCelda("","vt bgcolor=#ffffff ","CENTER","15%","","","","<big>ANTERIORES</big>","si")
    celdas = celdas & ofv.GenCelda("","ut  ","CENTER","85%","","","",STAB,"si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")
    STAB = ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")

    celdas = ofv.GenCelda("","","CENTER","","","","",STAB,"si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")
 
    sHTML = sHTML  & "<CENTER>"
'' & ofv.GenTabla("","AAA STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:2;' ","","80%","","0","","0","0","","",filas,"si") & "<BR>"


  sHTML = sHTML  & ofv.GenTabla("","aaa style='filter:shadow(color=#cccccc,direction=125);' ","","100%","","0","","0","0","","",filas,"si") & "<BR>"

  END IF


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENWKFPSCTPC(TIPO)

    strsql = "insert into WKFPSCTPC (WKFCOD,WKFPASCOD,TIPO,CAMEST,ESTADO,ANULA) values ('" & session("WKF") & "','" & session("WKFPSC") & "','" & TIPO & "','N','0','N')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION

FUNCTION GENWKFPASO(PASO)

  GENWKFPASO = "SIN ASOCIAR"

  StrSql = "Select * From WKFPASOS WHERE WKFPASCOD = '" & PASO & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs2.eof then
     sinput = "<a CLASS=CC href=WKFPASOSPROP.asp?WKFPSC=" & RS2(2)
     sinput = sinput  & " ><b><U>" & replace(RS2(3) & ". " & RS2(4)," ","&nbsp;") & "</U></b></a>" 
     GENWKFPASO = sinput
  END IF
  call ofv.cerrarconsulta(rs2)

END FUNCTION

FUNCTION GENWKFPASRESP()

  StrSql = "Select * From WKFRESP WHERE WKFCOD = '" & session("WKF") & "' "
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then 

    strsql = "insert into WKFPSCRESP (WKFCOD,WKFPASCOD,CODRESP) values ('" & session("WKF") & "','" & session("WKFPSC") & "','" & RS3(0) & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)


  END IF

END FUNCTION



function GenEntries()  

   sc = "vt"
   formulario = "" 

   xpanI = "" 
   xpanF = "" 
   xpanA = "" 
   xpanR = "" 

IF RS(6) THEN
   xpanI = ArmaEntries("I")
end if

IF RS(9) THEN
   xpanF = ArmaEntries("F")
end if

IF NOT RS(6) AND  NOT RS(9) THEN

IF RS(7) THEN
   xpanA = ArmaEntries("A")
end if

IF RS(8) THEN
   xpanR = ArmaEntries("R")
end if

end if




   GenEntries = xpanI & xpanF & xpanA & xpanR

end function



function ArmaEntries(Opc)


    sInput = "&nbsp;"
    celdas = ofv.GenCelda("","AAa COLSPAN=4 style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#f7aa22;' ","","10%","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")
    formularioC = filas








formulario = ""
sc = "vt"

    select case Opc
           case "I"
              irs  = 6
              frm1 = "N"
              frmN = "INICIAL"
              sInput = "<big>" & sImagenSI & "&nbsp;" & frmN & "</big>"
              FStrSql = "Select * From WKFPSCTPC WHERE WKFPASCOD = '" & session("WKFPSC") & "' AND TIPO = 'INI' "
              FOPC = "INI"
           case "F"
              irs  = 9
              frm1 = "N"
              frmN = "FINAL"
              sInput = "<big>" & sImagenSF & "&nbsp;" & frmN & "</big>"
              FStrSql = "Select * From WKFPSCTPC WHERE WKFPASCOD = '" & session("WKFPSC") & "' AND TIPO = 'FIN' "
              FOPC = "FIN"
           case "A"
              irs  = 7
              frm1 = "S"
              frmN = "APRUEBA"
              sInput = "<big>" & sImagenSA & "&nbsp;" & frmN & "</big>"
              FStrSql = "Select * From WKFPSCTPC WHERE WKFPASCOD = '" & session("WKFPSC") & "' AND TIPO = 'APR' "
              FOPC = "APR"
           case "R"
              irs  = 8
              frm1 = "S"
              frmN = "RECHAZA"
              sInput = "<big>" & sImagenSR & "&nbsp;" & frmN & "</big>"
              FStrSql = "Select * From WKFPSCTPC WHERE WKFPASCOD = '" & session("WKFPSC") & "' AND TIPO = 'RCH' "
              FOPC = "RCH"
    end select





    formulario = FORMULARIO & formularioC




   if frm1 = "S" then

    SFORM = "WKF" & irs & "S"
    celdas = ofv.GenCelda("","vt rowspan=3 bgcolor=#ffffff ","center","15%","","","",sInput,"si")

        sinputK = ""
        sinputt = ""

    if bModificar  then
        sInput = "<font class=at><select name=" & RS(irs).NAME & " onchange=submit()>"
        sInput = sInput & "<option value=SI "
        if rs(irs) then sInput = sInput & "selected"
        sInput = sInput & " >SI<option value=NO "
        if NOT rs(irs) then sInput = sInput & "selected"
        sInput = sInput & " >NO</select></font>"
        sinputt = "<div id=AA" & irs & "U name=xxx style='display:none;' >&nbsp;<big>Activo&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</big>&nbsp;" &  sInput & "</div>"

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='javascript:void(0);' onclick='AA" & irs & "U.style.display=" & chr(34) & "block" & chr(34) & ";AA" & irs & "D.style.display=" & chr(34) & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

    end if


        if rs(irs)  then
           sInput = "SI"
        ELSE
           sInput = "NO"
        END IF

        sInput = "<div id=AA" & irs & "D name=aaa style='display:block;' ><span>&nbsp;<big>Activo&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</big>&nbsp;" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt


    celdas = celdas & ofv.GenCelda("","vt COLSPAN=3","","","","","",sInput,"si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/ActualizarWK.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOS"
      formulario = FORMULARIO & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = FORMULARIO & ofv.GenForm("","","","","post","",filas,"si")
    end if

   else

    celdas = ofv.GenCelda("","vt rowspan=3 bgcolor=#ffffff ","center","15%","","","",sInput,"si")
    sInput = "&nbsp;"
    celdas = celdas & ofv.GenCelda("","vt COLSPAN=3","","","","","",sInput,"si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")
    formulario = FORMULARIO & filas

   end if

    SFORM = "WKF" & irs  & "P"
     StrSql = FStrSql
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF RS1.EOF THEN
     GENWKFPSCTPC(FOPC)
     StrSql = FStrSql
     set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF



  if not rs1.eof then

      formulario = FORMULARIO & formularioA

 ''   sInput = "&nbsp;"
 ''   celdas = ofv.GenCelda("","","","10%","","","",sInput,"si")
celdas = ""


        sinputK = ""
        sinputt = ""

    if bModificar  then
        sInput = "<font class=at><select name=" & rs1(4).name & " onchange=submit()>"
        sInput = sInput & "<option value=N "
        if rs1(4) = "N" then sInput = sInput & "selected"
        sInput = sInput & " >NO<option value=M "
        if rs1(4) = "M" then sInput = sInput & "selected"
        sInput = sInput & " >MANUAL<option value=A "
        if rs1(4) = "A" then sInput = sInput & "selected"
        sInput = sInput & " >AUTOMATICO</select></font>"
        sinputt = "<div id=A" & irs & "U name=xxx style='display:none;' >&nbsp;<big>Cambio&nbsp;de&nbsp;Estado:</big>&nbsp;" &  sInput & "</div>"

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='javascript:void(0);' onclick='A" & irs & "U.style.display=" & chr(34) & "block" & chr(34) & ";A" & irs & "D.style.display=" & chr(34) & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

    end if


        if rs1(4) = "N" then sInput = "&nbsp;<big>Cambio&nbsp;de&nbsp;Estado:</big>&nbsp;" & "NO"
        if rs1(4) = "M" then sInput = "&nbsp;<big>Cambio&nbsp;de&nbsp;Estado:</big>&nbsp;" & "MANUAL"
        if rs1(4) = "A" then sInput = "&nbsp;<big>Cambio&nbsp;de&nbsp;Estado:</big>&nbsp;" & "AUTOMATICO"

        sInput = "<div id=A" & irs & "D name=aaa style='display:block;' ><span>" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt





 '''   sInput = "&nbsp;<big>Cambio&nbsp;de&nbsp;Estado:</big>&nbsp;" & sInput




    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

   if rs1(4) = "A" then 




        sinputK = ""
        sinputt = ""

    if bModificar  then
        strsql = "Select * From " & ESTTABLA & ESTORDER
        sInput = ofv.GenerarCombo(sform,rs1(5).name,rs1(5),strsql,cnX,0,1,"")
        sinputt = "<div id=AB" & irs & "U name=xxx style='display:none;' ><big>Estado:</big>&nbsp;&nbsp;" &  sInput & "</div>"

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='javascript:void(0);' onclick='AB" & irs & "U.style.display=" & chr(34) & "block" & chr(34) & ";AB" & irs & "D.style.display=" & chr(34) & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

    end if

        sInput = "&nbsp;"
        strsql = "Select * From " & ESTTABLA & " where " & ESTCLAVE & " = 0" & rs1(5)
      ''  RESPONSE.WRITE STRSQL & "<BR>"
        set rsniv = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
        if not rsniv.eof then sInput = rsniv(1)
        call ofv.cerrarconsulta(rsniv)

        sInput = "<div id=AB" & irs & "D name=aaa style='display:block;' ><span><big>Estado:</big>&nbsp;" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt




   ELSE
    sInput = "&nbsp;"
   END IF



    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

  if Opc = "x" then
    if bModificar  then
        sInput = "<font class=at><select name=" & rs1(7).name & " onchange=submit()>"
        sInput = sInput & "<option value=N "
        if rs1(7) = "N" then sInput = sInput & "selected"
        sInput = sInput & " >NO<option value=P "
        if rs1(7) = "P" then sInput = sInput & "selected"
        sInput = sInput & " >PROPIOS<option value=G "
        if rs1(7) = "G" then sInput = sInput & "selected"
        sInput = sInput & " >GENERALES</select></font>"
    else
        if rs1(7) = "N" then sInput = "NO"
        if rs1(7) = "P" then sInput = "PROPIOS"
        if rs1(7) = "G" then sInput = "GENERALES"
        sInput = ofv.GenerarInput("",sInput,"readonly","","20","dt","")
    end if

  else
 ''   sInput = "----"
    sInput = "&nbsp;"
  end if
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs1(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPSCTPC"
      formulario = FORMULARIO & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = FORMULARIO & ofv.GenForm("","","","","post","",filas,"si")
    end if

  if Opc <> "F" then

      formulario = FORMULARIO & formularioB



 ''   sInput = "&nbsp;"
 ''   celdas = ofv.GenCelda("","","","10%","","","",sInput,"si")
celdas = ""

   XPASO = RS1(6) & ""
   if XPASO <> "" then 
     sInput = GENWKFPASO(RS1(6))
   ELSE
    sInput = "SIN ASIGNAR"
   END IF

    sInput = "&nbsp;<big>Paso siguiente</big>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & sImagenPS & "&nbsp;" & sinput

    celdas = celdas & ofv.GenCelda("",sc & " COLSPAN=2 ","","","","","",sInput,"si")


 ''   sinput = "<a  href=WKFPSCSELPAS.asp?WKFPSCTPC=" & RS1(0)
 ''   sinput = sinput  & " ><b>" & sImagenI3 & "</b></a>" 

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href=WKFPSCSELPAS.asp?WKFPSCTPC=" & RS1(0)
        sinputK = sinputK  & "  TITLE='Configurar'  ><b>" & sImagenI3 & "</b></a>" 



    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInputK,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")


    formulario = FORMULARIO & ofv.GenForm("","","","","post","",filas,"si")

  else
 ''   sInput = "&nbsp;"
 ''   celdas = ofv.GenCelda("","","","10%","","","",sInput,"si")
celdas = ""
    celdas = celdas & ofv.GenCelda("",sc & " COLSPAN=3 ","","","","","","&nbsp;","si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")


    formulario = FORMULARIO & ofv.GenForm("","","","","post","",filas,"si")

  end if


  END IF



ArmaEntries = formulario

end function




%>