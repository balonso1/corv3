<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

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
    session("CKT") = ""

 if sopcion = "SI" then
    GENWKF()

 end if


    celdas = ""



  'Encabezados
  formulario =  ofv.GenCelda("","tt colspan=6","center","","","","","CIRCUITOS","si")
''  formulario =  formulario & ofv.Encabezados("DESCRIPCION","&nbsp;","&nbsp;","&nbsp;")


  StrSql = "Select * From WKFCIRCUITO ORDER BY VALIDO, DESCRIPCION "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then
     rs.move(CantRegMoverx)
     formulario = formulario & GENLISTA(RS)

  END IF




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "WKFCIRCUITOS.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  


  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENWKF()

    WKFCOD = GTWKF()

    strsql = "insert into WKFCIRCUITO (WKFCOD,DESCRIPCION,INICIO,CLASE,OBJTIPO) values ('" & WKFCOD & "','NUEVO CIRCUITO WKF','S','E','GEN')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION


FUNCTION GENLISTA(RSX)

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>L</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=4>&#064;</FONT></B>"

  posi = 0
  prow = "N" 
  Do While (Not rsX.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rsX(0)

    if prow = "S" then
       prow = "N"
       sr   = " bgcolor=#eeeeee "
    else 
       prow = "S"
       sr   = ""
    end if

       sr   = ""
       sc = "vt"

    xlink = "<A  href='&nbsp;'  "
    xlink = xlink & " onclick=" & chr(34) & "AbrirCons('ConsPRY.asp?codigo=" & rsx(1) & "'); return false;" & chr(34)
    xlink = xlink & " ><IMG  src='Imagenes/ICorAbierto.jpg'  border=0 ></A>"

    sInput = xlink

      celdas = ofv.GenCelda("","vt rowspan=2 ","","1%","","","",sInput,"si")



       if bModificar then
          sInput = ofv.GenerarInput(rsx(2).name,rsx(2),"text",sform,"80","dt","")
       else
          sInput = ofv.GenerarInput("",rsx(2),"readonly","","80","dt","")
       end if

      celdas = celdas & ofv.GenCelda("",sc & " rowspan=2 style='border-right-style:solid;border-right-width:2;border-right-color:#6699cc;' ","","30%","","","",sInput,"si")


'''''''''


        if rsx(4) = "A" then sInput = "AUTOMATICO"
        if rsx(4) = "S" then sInput = "A SOLICITUD"
        sMwk = sInput


        if rs(5) = "E" then sInput = "EFECTIVO"
        if rs(5) = "M" then sInput = "MODELO"
        sCwk = sInput

        sniv = "S/D"
        strsql = "Select * From WKFTIPOS where TIPO = '" & rsx(6) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(2)
        call ofv.cerrarconsulta(rsniv)
        sTwk = sniv




'''''''''

      sInput = "<b>" & sImagenE & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","3%","","","",sInput,"si")


      sInput = sMwk
      celdas = celdas & ofv.GenCelda("",sc & " bgCOLOR=#ffffff","","37%","","","",sInput,"si")

      sInput = "<b>" & sImagenC & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","3%","","","",sInput,"si")

      sInput = sCwk
      celdas = celdas & ofv.GenCelda("",sc & " bgCOLOR=#ffffff","","37%","","","",sInput,"si")

 
      filas = ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")



      sInput = "<b>" & sImagenO & "</b>"
      celdas = ofv.GenCelda("","vt","","3%","","","",sInput,"si")

      sInput = sTwk
      celdas = celdas & ofv.GenCelda("",sc & " bgCOLOR=#ffffff" ,"","37%","","","",sInput,"si")


     sinput = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=WKFRESP.asp?WKF=" & RSx(1)
     sinput = sinput  & "  TITLE='Responsables'  ><b>" & sImagenI1 & "</b></a>" 
     sinput1 = sinput

     sinput = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "TARGET=_blank   href=WKFver.asp?WKF=" & RSx(1)
     sinput = sinput  & "  TITLE='Visualizacion'  ><b>" & sImagenI2 & "</b></a>" 
     sinput2 = sinput

     sinput = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=WKFPASOS.asp?WKF=" & RSx(1)
     sinput = sinput  & "  TITLE='Configuracion'  ><b>" & sImagenI3 & "</b></a>" 
     sinput3 = sinput



      sInput = sinput1 & "&nbsp;&nbsp;" & sinput2 & "&nbsp;&nbsp;" & sinput3
      celdas = celdas & ofv.GenCelda("","vt colspan=2","","3%","","","",sInput,"si")





    filas = filas & ofv.GenRow("","aaa  " & sr,"","","","","",celdas,"si")

    filas = ofv.GenTabla("","aaa style='filter:shadow(color=#cccccc,direction=125);' ","","100%","","0","","-2","-1","","",filas,"si")

    celdas = ofv.GenCelda("","aaa colspan=6 style='border-style:solid;border-width:2;border-color:#dddddd;' ","","","","","",filas,"si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFCIRCUITO"
      formu = formu & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formu = formu & ofv.GenForm("","","","","post","",filas,"si")
    end if


    posi = posi + 1
    rsX.Movenext
  loop

      sInput = "<b>" & sImagenE & "&nbsp;Modo Inicio</b>"
      celdas = ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenC & "&nbsp;Clase</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenO & "&nbsp;Tipo</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenI1 & "&nbsp;Responsables</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenI2 & "&nbsp;Visualizacion</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenI3 & "&nbsp;Configuracion</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")


    formu = formu & ofv.GenRow("","aaa height=30","","","","","",celdas,"si")

  GENLISTA = FORMU

END FUNCTION


function ppp

  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "WKF" & rs(0)

    if rs(3) then
       sc = "CC"
    else
       sc = "TT"
    end if


    celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")



'       sInput = ofv.GenerarInput("",rs(1),"readonly","","50","dt","")
'       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


       if bModificar then
          sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sform,"80","dt","")
       else
          sInput = ofv.GenerarInput("",rs(2),"readonly","","80","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

'    if rs(3) then
'          sInput = ofv.GenerarInput("","SI","readonly","","3","dt","")
'    else
'          sInput = ofv.GenerarInput("","NO","readonly","","3","dt","")
'    end if 
'       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    if rs(3) then
       sinput = "<a class=tt  href=WKFPASOS.asp?WKF=" & RS(1)
       sinput = sinput  & " ><b>MODIFICAR</b></a>" 
    else
       sinput = "<a class=tt  href=WKFPASOS.asp?WKF=" & RS(1)
       sinput = sinput  & " ><b>CONFIGURAR</b></a>" 
    end if 
       celdas = celdas & ofv.GenCelda("","tt","center","","","","",sInput,"si")


'    if rs(3) then
'       sinput = "<font face=verdana size=1 color=#ffffff ><b>---</b></font>" 
'    else
'       sinput = "<a class=tt href=WKFVAL.asp?OPC=WKF&WKF=" & RS(1)
'       sinput = sinput  & " ><b>VALIDAR</b></a>" 
'    end if 
'       celdas = celdas & ofv.GenCelda("","tt","center","","","","",sInput,"si")

       sinput = "<a class=tt  href=WKFRESP.asp?WKF=" & RS(1)
       sinput = sinput  & " ><b>RESPONSABLES</b></a>" 
       celdas = celdas & ofv.GenCelda("","tt","center","","","","",sInput,"si")

       sinput = "<a class=tt  href=WKFver.asp?WKF=" & RS(1)
       sinput = sinput  & " TARGET=_blank ><b>ver</b></a>" 
       celdas = celdas & ofv.GenCelda("","tt","center","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFCIRCUITO"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop
end function



%>