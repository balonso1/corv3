<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  sHTML = ofv.FormHeader("")

  CantRegAMover = cdbl(request.querystring("CantRegAMover"))
  CantRegMover = cdbl(request.querystring("CantRegMover"))
  if CantRegMover > 0 then CantRegAMover = CantRegMover
  CantRegAMostrar = 1
  vuelta = ofv.vueltaasp("../",CantRegAMostrar,CantRegAMover)
  vuelta2 = ofv.vueltaaspEx("",CantRegAMostrar,CantRegAMover)
  OpcAgregar = request.querystring("OpcAgregar")

  session("CRAM") = CantRegAMover

  set cn = ofv.conectar(ofv.strconn0)


  ofv.ObtenerAtributos request.querystring("seg"),""
  Session("Form") = request.querystring("seg")

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  StrSql = "Select * From Grupos Order By descripcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegAMover)
  if not rs.eof then
    nID = rs(0)


    sHTML = sHTML & "<center>"


  celdas = ofv.GenCelda("","TTT colspan=3","","","","","","GRUPO: " & rs(0),"si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","tt","","","","","","Descripcion:","si")

    'Eliminar
    if bSuprimir then
      sHref = "asp/Eliminar.asp?ID=" & rs(0) & "&Descripcion=" & ofv.convertircar(rs(1)," ","_")
      sHref = sHref & "&volver=" & vuelta & "&Eliminar=NSNC&CantRegAMover=" & CantRegAMover
      sHref = sHref & "&Tabla=Grupos"
     ' response.write shref
      celdas = celdas & ofv.BotonEliminar2(rs(1),sHref,"")
    else
      celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    end if

    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","Grupos","90","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    XNOMBRE = rs(1)    


    scomK = "D" & chr(9) & "L" & chr(9) & "R"
    scomD = "DEFAULT" & chr(9) & "LOCAL" & chr(9) & "REMOTO"

    celdas = ofv.GenCelda("","tt","","","","","","Acceso:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
       sInput = ofv.gencomboFX("Grupos",rs(2).name,rs(2),scomK,scomD," ")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","90","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    scomK = "D" & chr(9) & "A" & chr(9) & "B"
    scomD = "DEFAULT" & chr(9) & "ALTA" & chr(9) & "BAJA"

    celdas = ofv.GenCelda("","tt","","","","","","Resolucion:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
       sInput = ofv.gencomboFX("Grupos",rs(3).name,rs(3),scomK,scomD," ")
    else
      sInput = ofv.GenerarInput("",rs(3),"readonly","","90","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    scomK = "D" & chr(9) & "5" & chr(9) & "10" & chr(9) & "15" & chr(9) & "20" & chr(9) & "30" & chr(9) & "40" & chr(9) & "60"  & chr(9) & "120" & chr(9) & "240" & chr(9) & "360" & chr(9) & "480"
    scomD = "DEFAULT" & chr(9) & "5 m." & chr(9) & "10 m." & chr(9) & "15 m." & chr(9) & "20 m." & chr(9) & "30 m." & chr(9) & "40 m." & chr(9) & "60 m." & chr(9) & "120 m." & chr(9) & "240 m." & chr(9) & "360 m." & chr(9) & "480 m."

    celdas = ofv.GenCelda("","tt","","","","","","Time-Out sesion:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
       sInput = ofv.gencomboFX("Grupos",rs(4).name,rs(4),scomK,scomD," ")
    else
      sInput = ofv.GenerarInput("",rs(4),"readonly","","20","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=Grupos"
      formulario0 = ofv.GenForm("Grupos","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    filas = ofv.GenRow("","","","","","","",formulario0,"si")

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"
    rs.movenext
  end if

  if OpcAgregar = "si" then
    sHTML = sHTML & "<center>"
    celdas = ofv.GenCelda("","cc colspan=3","","","","","","NUEVO GRUPO","si")
    filas  = ofv.GenRow("","","","","10","","",celdas,"si")

    sAccion = "asp/Agregar.asp?Tabla=Grupos&volver=" & vuelta
    filas = filas & ofv.GenForm("Agregar","",sAccion,"","post","","","no")

    celdas = ofv.GenCelda("","tt","","","","","","ID:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sInput = ofv.GenerarInput("ID","","agregar","","10","dt","")
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","","","","","Descripcion:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sInput = ofv.GenerarInput("Descripcion","","agregar","","100","dt","")
    sInput = sinput & ofv.GenerarInput("acceso","D","hidden","","100","dt","")
    sInput = sinput & ofv.GenerarInput("resol","D","hidden","","100","dt","")
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    celdas = celdas & ofv.GenFForm()
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")

    sInput = ofv.GenerarInput("","Cancelar","submit","","10","bt","")
    sInput = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "GruposX.asp?OpcAgregar=no"
    celdas = celdas & ofv.GenForm("Cancelar","",sAccion,"","post","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center>"

     sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")

  else
    sHTML = sHTML & "<center>"



    sHTML = sHTML & "</center><br><br>"

' Lista de Grupos
    sInput = ofv.GenerarInput("","LISTA","submit","","10","bt","")
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "CoGrupos.asp"
    form1 = ofv.GenForm("lisgrp","",sAccion,"_blank","post","",celdas,"si")
' Lista de Usuarios
    sInput = ofv.GenerarInput("","INTEGRANTES","submit","","10","bt","")
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "CoGrpUsr.asp?Grupo=" & nID
    form2 = ofv.GenForm("lisusr","",sAccion,"_blank","post","",celdas,"si")
' Permisos del grupo
    sInput = ofv.GenerarInput("","PERMISOS","submit","","10","bt","")
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "CoPermisos.asp?ID=" & nID & "&t=G"
    form3 = ofv.GenForm("pergrp","",sAccion,"_blank","post","",celdas,"si")
' Perfil del grupo
    sInput = ofv.GenerarInput("boton","PERFIL","submit","","10","bt","")
  ''  sAccion = "GrpUsrMar.asp?Opcion=" & nID & "&t=G&o=A&volver=" & vuelta2

    sAccion = "Aplicaciones.asp?Opcion=" & nID & "&t=G&o=A&volver=" & vuelta2
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    bperf = ofv.GenForm("init","",sAccion,"","post","",celdas,"si")

' EMPRESAS
    sInput = ofv.GenerarInput("boton","EMPRESAS","submit","","10","bt","")
    sAccion = "USREMPRE.asp?Opcion=" & nID & "&t=G&o=A&volver=" & vuelta2 & "&NOMBRE=" & REPLACE(XNOMBRE," ","_")
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    bperf2 = ofv.GenForm("init","",sAccion,"","post","",celdas,"si")

    celdas = ofv.GenCelda("","ttt colspan=5","","","","","","OPCIONES","si")
    filas  = ofv.GenRow("","","","","10","","",celdas,"si")


    filas = filas & ofv.GenRow("","","","","40","","",form1 & form2 & form3 & bperf & bperf2,"si")

    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","tbpx","","10%","","0","","2","0","","",filas,"si")
    sHTML = sHTML & "</center>"



    str1 = "Select descripcion From Grupos Order By descripcion "    ' orden pagina
    str2 = "Select descripcion From Grupos Order By descripcion "    ' orden alfabetico
    Buscar = ofv.Combobuscar(str1,str2,cn,cantregamostrar,CantRegaMover,ofv.vueltaasp2(""))

      Agregar = ""
    if bModificar then
      sAccion = "GruposX.asp?OpcAgregar=si"
      Agregar = ofv.BotonAgregar(sAccion)
     '' sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Buscar,"")
      sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,"")
    else
  ''    sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no","no",Buscar,"")
      sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,"")
    end if


  end if

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)





  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>