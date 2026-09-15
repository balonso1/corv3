<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = replace(ofv.FormHeader(session("FormN")),"<body ","<body onload='uno();' ")  & chr(13)
  shtml = shtml & "<script  language=javascript>" & chr(13)
  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function uno() {  " & chr(13)
  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - document.all.odiv.clientHeight; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)

'  sHTML = sHTML  & "<script language=javascript> "
'  sHTML = sHTML  & "function AbrirCons(Datos) { "
'  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
'  sHTML = sHTML  & " } "
'  sHTML = sHTML  & "</script>"

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")
 WKF = request.querystring("WKF")

 if WKF = "" then
    WKF = session("WKF")
 else
    session("WKF") = WKF
 end if 

 if sopcion = "SI" then
    GENWKFPASO()

 end if
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
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



  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & WKF & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
'     IF RS(3) THEN
'        bAgregar = FALSE
'        bModificar = FALSE
'        bSuprimir = FALSE
'     END IF

    IDWKF = RS(0)  

    sInput = "CIRCUITO:"
    celdas = ofv.GenCelda("","tt","","20%","","","",sInput,"si")

    sInput = RS(1) & " - <B>" & RS(2) & "</B>"
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"

    filas = ""
    celdas = ofv.GenCelda("","vt","","20%","","","","<big>" & sImagenE & "&nbsp;Modo inicio:</big>","si")

        sinputK = ""
        sinputt = ""

    if bModificar  then
        sInput = "<font class=at><select style='border-style:outset;border-width:5;' name=" & rs(4).name & " onchange=submit() >"
        sInput = sInput & "<option value=A "
        if rs(4) = "A" then sInput = sInput & "selected"
        sInput = sInput & " >AUTOMATICO<option value=S "
        if rs(4) = "S" then sInput = sInput & "selected"
        sInput = sInput & " >A SOLICITUD</select></font>"
        sinputt = "<div id=A11 name=xxx style='display:none;' >" &  sInput & "</div>"

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='javascript:void(0);' onclick='A11.style.display=" & chr(34) & "block" & chr(34) & ";A12.style.display=" & chr(34) & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

    end if


        if rs(4) = "A" then sInput = "AUTOMATICO"
        if rs(4) = "S" then sInput = "A SOLICITUD"

        sInput = "<div id=A12 name=aaa style='display:block;' ><span>" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt



    celdas = celdas & ofv.GenCelda("","vt bgcolor=#ffffff","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","vt","","20%","","","","<big>" & sImagenC & "&nbsp;Clase:</big>","si")


        sinputK = ""
        sinputt = ""

    if bModificar  then
        sInput = "<font class=at><select name=" & rs(5).name & " onchange=submit()>"
        sInput = sInput & "<option value=E "
        if rs(5) = "E" then sInput = sInput & "selected"
        sInput = sInput & " >EFECTIVO<option value=M "
        if rs(5) = "M" then sInput = sInput & "selected"
        sInput = sInput & " >MODELO</select></font>"
        sinputt = "<div id=A21 name=xxx style='display:none;' >" &  sInput & "</div>"

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='javascript:void(0);' onclick='A21.style.display=" & chr(34) & "block" & chr(34) & ";A22.style.display=" & chr(34) & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

    end if


        if rs(5) = "E" then sInput = "EFECTIVO"
        if rs(5) = "M" then sInput = "MODELO"

        sInput = "<div id=A22 name=aaa style='display:block;' ><span>" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt

    celdas = celdas & ofv.GenCelda("","vt bgcolor=#ffffff","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt","","20%","","","","<big>" & sImagenO & "&nbsp;Tipo:</big>","si")

        sinputK = ""
        sinputt = ""

    if bModificar then
      strsql = "Select * From WKFTIPOS Order By DESCRIPCION"
      sInput = ofv.GenerarCombo("Usuarios",rs(6).name,rs(6),strsql,cn,1,2,"")
      sinputt = "<div id=A31 name=xxx style='display:none;' >" &  sInput & "</div>"

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='javascript:void(0);' onclick='A31.style.display=" & chr(34) & "block" & chr(34) & ";A32.style.display=" & chr(34) & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

    end if

      sInput = "S/D"
      strsql = "Select * From WKFTIPOS where TIPO = '" & rs(6) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sInput = rsniv(2)
        call ofv.cerrarconsulta(rsniv)

        sInput = "<div id=A32 name=aaa style='display:block;' ><span>" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt




    celdas = celdas & ofv.GenCelda("","vt bgcolor=#ffffff","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(IDWKF," ","*") & "&volver=" & vuelta & "&Tabla=WKFCIRCUITO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    filas = ofv.GenRow("","","","","","","",formulario0,"si")

  sHTML = sHTML  & ofv.GenTabla("","aaa style='filter:shadow(color=#cccccc,direction=125);' ","","100%","","0","","-2","-1","","",filas,"si") & "<BR>"

''  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


  END IF



    celdas = ""



  'Encabezados
  formulario =  ofv.GenCelda("","tt colspan=7","center","","","","","PASOS","si")
'''  formulario =  formulario & ofv.Encabezados("#","CODIGO","DESCRIPCION","TIPO","&nbsp;")


  StrSql = "Select * From WKFPASOS WHERE WKFCOD = '" & WKF & "' ORDER BY ORDEN"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF RS.EOF THEN
     GENWKFPASOIF()
     StrSql = "Select * From WKFPASOS WHERE WKFCOD = '" & WKF & "' ORDER BY ORDEN"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF


  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "WKF" & rs(0)

    if rs(5) then
       sc = "CC"
    else
       sc = "TT"
    end if

       sc = "VT"

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtask.asp?opcion=SBT&TSK="
     sinput = sinput  & " TITLE='SUBIR TAREA' ><b>" & sImagen1 & "</b></a>" 


    celdas = ofv.GenCelda("","vt  bgcolor=#6699cc ","","","","","",sinput,"si")



       sInput = rs(3) & ".&nbsp;"
       celdas = celdas & ofv.GenCelda("",sc,"right","","","","",sInput,"si")

 ''      sInput = ofv.GenerarInput("",rs(2),"readonly","","50","dt","")
 ''      celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


       if bModificar then
          sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sform,"60","dt STYLE='FONT-WEIGHT:BOLD' ","")
       else
          sInput = ofv.GenerarInput("",rs(4),"readonly","","60","dt STYLE='FONT-WEIGHT:BOLD' ","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


        if rs(6)  then     sInput = "<b>" & sImagenPI & "</b>"
        if rs(9)  then     sInput = "<b>" & sImagenPF & "</b>"
        if rs(7) OR RS(8) then     sInput = "<b>" & sImagenPC & "</b>"



    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



      sInput = "<b>" & sImagenCP & "</b>"
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")








        sinputK = ""
        sinputt = ""

    if  bModificar AND RS(3) > 10 AND RS(3) < 999 then
        sInput = "<font class=at><select name=" & rs(12).name & " onchange=submit()>"
        sInput = sInput & "<option value=U "
        if rs(12) = "U" then sInput = sInput & "selected"
        sInput = sInput & " >NORMAL<option value=R "
        if rs(12) = "R" then sInput = sInput & "selected"
        sInput = sInput & " >RECURRENTE</select></font>"
        sinputt = "<div id=AX" & rs(0) & " name=xxx style='display:none;' >" &  sInput & "</div>"

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='javascript:void(0);' onclick='AX" & rs(0) & ".style.display=" & chr(34) & "block" & chr(34) & ";AY" & rs(0) & ".style.display=" & chr(34) & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

    end if


        if rs(12) = "U" then sInput = "NORMAL"
        if rs(12) = "R" then sInput = "RECURRENTE"

        sInput = "<div id=AY" & rs(0) & " name=aaa style='display:block;' ><span>" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt





    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")




''    if rs(5) then
''       sinput = "<a  href=WKFPASOSPROP.asp?WKFPSC=" & RS(2)
''       sinput = sinput  & " ><b>MODIFICAR</b></a>" 
''    else
''       sinput = "<a  href=WKFPASOSPROP.asp?WKFPSC=" & RS(2)
''       sinput = sinput  & " ><b>CONFIGURAR</b></a>" 
''    end if 

     sinput = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=WKFPASOSPROP.asp?WKFPSC=" & RS(2)
     sinput = sinput  & "  TITLE='Configuracion'  ><b>" & sImagenI3 & "</b></a>" 


       celdas = celdas & ofv.GenCelda("",sc,"center","","","","",sInput,"si")






    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOS"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop
  sHTML = sHTML  & ofv.GenTabla("","aaa style='filter:shadow(color=#cccccc,direction=125);' ","","100%","","0","","-2","-1","","",formulario,"si")
'''  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "WKFPASOS.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  


  cnbot = "<input type=submit class=bt value='Volver&nbsp;a&nbsp;Circuito' >"
  Cons = ofv.gencelda("","","center","","1%","","",cnbot,"si")
  sAccion = "WKFCIRCUITOS.asp"
  Cons = ofv.genform("cons","",sAccion,"","post","",Cons,"si")


  fill = space(50)

  fill = replace(fill," ","&nbsp;")

  fill = ofv.gencelda("","vt","right","","100%","","",fill & "<big>Cor&nbsp;<font color=#f7aa44 size=3><i>e</i></font>-Manager&nbsp;</big>&reg;&nbsp;","si")

      sInput = "<b>" & sImagenI & "&nbsp;Modificar</b>"
      celdas = ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenCP & "&nbsp;Modo&nbsp;del&nbsp;paso</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenPI & "&nbsp;Paso&nbsp;inicial</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenPF & "&nbsp;Paso&nbsp;final</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenPC & "&nbsp;Paso&nbsp;del&nbsp;circuito</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenI3 & "&nbsp;Configuracion</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")


    formu = ofv.GenRow("","aaa height=30","","","","","",celdas,"si")
    formu = ofv.GenTabla("","aaa ","","100%","","0","","-2","-1","","",formu,"si")


  sHTML = sHTML & "<div  id=odiv name=odiv style='position:absolute;top:-10px;filter:shadow(color=" & Session("APLSHW") & ",direction=125);'  >"
  sHTML = sHTML & formu & replace(ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,Cons,fill),"id=btnas width=2%","id=btnas width=100% style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ")
  sHTML = sHTML & "</div>"


  
 '' sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Cons,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENWKFPASO()

    WKFCOD = GTWKFPSC()

  StrSql = "Select MAX(ORDEN) From WKFPASOS WHERE WKFCOD = '" & session("WKF") & "' AND ORDEN < 999 "
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then 
     NORDEN = RS3(0) + 10
  ELSE

     NORDEN = 20
  END IF

    strsql = "insert into WKFPASOS (WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,APR,RCH,TIPO) values ('" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','NUEVO PASO WKF','1','1','U')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION

FUNCTION GENWKFPASOIF()

    WKFCOD = GTWKFPSC() & "I"
    NORDEN = 10


    strsql = "insert into WKFPASOS (WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,INI,TIPO) values ('" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','INICIO','1','U')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    session("WKFPSC") = WKFCOD
    GENWKFPSCTPC("INI")

    WKFCOD = GTWKFPSC() & "F"
    NORDEN = 999


    strsql = "insert into WKFPASOS (WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,FIN,TIPO) values ('" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','FINAL','1','U')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    session("WKFPSC") = WKFCOD
    GENWKFPSCTPC("FIN")

END FUNCTION

FUNCTION GENWKFPSCTPC(TIPO)

    strsql = "insert into WKFPSCTPC (WKFCOD,WKFPASCOD,TIPO,CAMEST,ESTADO,ANULA) values ('" & session("WKF") & "','" & session("WKFPSC") & "','" & TIPO & "','N','0','N')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION

%>