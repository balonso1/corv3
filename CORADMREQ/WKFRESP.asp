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
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

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
    GENWKFRESP()

 end if

  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & WKF & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     IF RS(3) THEN
        bAgregar = FALSE
        bModificar = FALSE
        bSuprimir = FALSE
     END IF


    sInput = "CIRCUITO:"
    celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

    sInput = RS(1) & " - " & RS(2)
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


  END IF


    celdas = ""


  on error resume next
  'Encabezados
  formulario =  ofv.GenCelda("","tt colspan=4","center","","","","","RESPONSABLES","si")
  formulario =  formulario & ofv.Encabezados("TIPO","RESPONSABLE","&nbsp;")


  StrSql = "Select * From WKFresp WHERE WKFCOD = '" & WKF & "' ORDER BY tipo, id "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "WKF" & rs(0)

    if rs(3) = "" then
       sc = "tt"
    else
       sc = "cc"
    end if


    celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")


    if bModificar  then
        sInput = "<font class=at><select name=" & rs(2).name & " onchange=submit()>"
        sInput = sInput & "<option value=USR "
        if rs(2) = "USR" then sInput = sInput & "selected"
        sInput = sInput & " >USUARIO<option value=GRP "
        if rs(2) = "GRP" then sInput = sInput & "selected"
        sInput = sInput & " >GRUPO<option value=SEC "
        if rs(2) = "SEC" then sInput = sInput & "selected"
        sInput = sInput & " >SECTOR</select></font>"
    else
        if rs(2) = "USR" then sInput = "USUARIO"
        if rs(2) = "GRP" then sInput = "GRUPO"
        if rs(2) = "SEC" then sInput = "SECTOR"
        sInput = ofv.GenerarInput("",sInput,"readonly","","20","dt","")
    end if

       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    XRESPO = RS(3) & ""

    

    if XRESPO = "" then
       sInput = "--- NO ASIGNADO ---------------------"
    else
       sInput = "S/D"
        if rs(2) = "USR" then
           strsql = "Select * From USUARIOS where ID = '" & rs(3) & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
        end if
        if rs(2) = "GRP" then
           strsql = "Select * From GRUPOS where ID = '" & rs(3) & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
        end if
        if rs(2) = "SEC" then
           strsql = "Select * From ECO where ECO = " & rs(3)
           set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
        end if

    end if




       sInput = ofv.GenerarInput("",sInput,"readonly","","70","dt","")
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    if bModificar  then

       sinput = "<a  href=WKFRESPsel.asp?wkfresp=" & rs(0)
       sinput = sinput  & " ><b>SELECCIONAR</b></a>" 

    ELSE
       sinput = " "
    END IF


       celdas = celdas & ofv.GenCelda("","TT","center","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFRESP"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "WKFRESP.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  
  sinput = "<a CLASS=TT href=WKFCIRCUITOS.asp"
  sinput = sinput  & " ><b>Volver&nbsp;a&nbsp;CIRCUITO</b></a>" 
  Buscar = ofv.GenCelda("","TT","","","","","",sInput,"si")

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)

  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENWKFRESP()



    strsql = "insert into WKFRESP (WKFCOD,TIPO) values ('" & WKF & "','USR')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>