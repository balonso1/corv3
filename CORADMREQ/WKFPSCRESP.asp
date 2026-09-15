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

    sInput = RS(1) & " - " & RS(2)
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


  END IF



    celdas = ""





  StrSql = "Select * From WKFPASOS WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then



    sInput = "PASO:"
    celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

    sInput = RS(3) & ":" & RS(2) & " - " & RS(4)
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


 END IF

  StrSql = "Select * From WKFPSCRESP WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF RSX.EOF THEN
     GENWKFPASRESP()
     StrSql = "Select * From WKFPSCRESP WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
     set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF



  if not rsX.eof then


    sInput = "RESPONSABLE "
    celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

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

    celdas = celdas & ofv.GenCelda("","cc","","","","","","<b>" & sInputT & sInputD & "</b>","si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"




  StrSql = "Select * From WKFRESP WHERE WKFCOD = '" & session("WKF") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

    sInput = "SELECCIONAR"
    celdas = ofv.GenCelda("","TT","CENTER","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")
    formularioA = filas
    formularioC = ""

     DO WHILE NOT RS.EOF

    SFORM = "WKF00" & RS(0)
    sinput = "<a CLASS=CC href='&nbsp;' onclick=" & chr(34) & "submit(); return false;" & chr(34)


       sInputD = "S/D"
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


    sinput = sinput  & " ><b><U>" & sInputT & sInputD & "</U></b></a>" 
    sinput = sinput  & ofv.GenerarInput(rsX(3).name,rs(0),"hidden",sform,"50","dt","")
    celdas = ofv.GenCelda("","CC","CENTER","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")
    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rsX(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPSCRESP"
      formularioC = formularioC & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formularioC = formularioC & ofv.GenForm("","","","","post","",filas,"si")
    end if


    RS.MOVENEXT
 LOOP

  sHTML = sHTML  & "<center>" & ofv.GenTabla("","","","70%","","0","","0","0","","",formularioA & formularioC,"si")  & "</center>" & "<BR>"



END IF

  call ofv.cerrarconsulta(rs)

 ELSE



 END IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




  'Botonera
  Agregar = ""


  Buscar = ""
  
  sinput = "<a  href=WKFPASOSPROP.asp"
  sinput = sinput  & " ><b>Volver&nbsp;aL&nbsp;PASO</b></a>" 
  Buscar = ofv.GenCelda("","TT","CENTER","","","","",sInput,"si")
  Buscar = BUSCAR & ofv.GenCelda("","","","80%","","","",sInput,"si")
      filas = ofv.GenRow("","","","","","","",Buscar,"si")
  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")





  call ofv.cerrarconsulta(rsX)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)

  sHTML = sHTML  & "</body></html>"



end if

response.write sHTML

FUNCTION GENWKFPASRESP()


  StrSql = "Select * From WKFRESP WHERE WKFCOD = '" & session("WKF") & "' "
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then 

    strsql = "insert into WKFPSCRESP (WKFCOD,WKFPASCOD,CODRESP) values ('" & session("WKF") & "','" & session("WKFPSC") & "','" & RS3(0) & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)


  END IF




END FUNCTION

FUNCTION GENWKFPASO(PASO)

  GENWKFPASO = "SIN ASOCIAR"

  StrSql = "Select * From WKFPASOS WHERE WKFPASCOD = '" & PASO & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs2.eof then
     sinput = RS2(3) & ":" & RS2(2) & " - " & RS2(4)
     GENWKFPASO = sinput
  END IF
  call ofv.cerrarconsulta(rs2)

END FUNCTION



%>