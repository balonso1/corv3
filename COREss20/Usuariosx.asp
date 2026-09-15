<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  CantRegAMover = cdbl(request.querystring("CantRegAMover"))
  CantRegMover = cdbl(request.querystring("CantRegMover"))
  if CantRegMover > 0 then CantRegAMover = CantRegMover
  CantRegAMostrar = 1
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta2 = ofv.vueltaaspEx("",CantRegAMostrar,CantRegAMover)
  OpcAgregar = request.querystring("OpcAgregar")


  nusuario = request.querystring("Nusuario")

  if nusuario <> "" then
     session("Nusuario") = nusuario
  else
     nusuario = session("Nusuario")
  end if

  set cn = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos request.querystring("seg"),""
  Session("Form") = request.querystring("seg")

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  if nusuario <> "" then
     StrSql = "Select * From Usuarios where id = '" & nusuario & "' "
  else
     StrSql = "Select * From Usuarios Order By descripcion"
  end if
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'  if not rs.eof then rs.move(CantRegAMover)
  if not rs.eof then
    nID = rs(0)

'    sHTML = sHTML & "<div style='position:absolute;top:10;left:10;visibility:visible;'>"
    sHTML = sHTML & "<center>"

  celdas = ofv.GenCelda("","ttt colspan=3","","","","","","USUARIO:&nbsp;&nbsp;" & rs(0) & "&nbsp;(" & rs(1) & ")","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")


   '' celdas = ofv.GenCelda("","tt","","","","","","ID:","si")
   '' celdas = celdas & ofv.GenCelda("","","","","","","","","si")
   '' celdas = celdas & ofv.GenCelda("","cc","","20","","","",rs(0),"si")
   '' filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    session("sletra") = rs(1)

    celdas = ofv.GenCelda("","tt","","","","","","Descripcion:","si")

    'Eliminar
    if bSuprimir and rs("ID") <> "COREMANAGER" then
     ' sHref = "asp/Eliminar.asp?ID=" & rs(0) & "&Descripcion=" & ofv.convertircar(rs(1)," ","_")
     ' sHref = sHref & "&volver=" & vuelta & "&Eliminar=NSNC&CantRegAMover=" & CantRegAMover
     ' sHref = sHref & "&Tabla=Usuarios"
      sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=usuarios&volver=" & vuelta & "&Eliminar=NSNC"
      
      celdas = celdas & ofv.BotonEliminar(rs(1),sHref)
      
      'celdas = celdas & ofv.BotonEliminar2(rs(1),sHref,"")
    else
      celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    end if




   if bModificar and rs("ID") <> "COREMANAGER" then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","Usuarios","100","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","100","dt","")
    end if

    XNOMBRE = rs(1)

    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

 if  Session("Modal") = "MSNTT" then

    celdas = ofv.GenCelda("","aa colspan=3 bgcolor=#f7aa22 ","","","2","","","","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     simagen1 = "<font face='webdings' size=2 color=#f7aa22>4</font>"

'Usuario Nt

    celdas = ofv.GenCelda("","tt  ","","","","","","Usuario de Red:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

   if bModificar  then
      sInput = ofv.GenerarInput(rs(26).name,rs(26),"text","Usuarios","100","dt","")
    else
      sInput = ofv.GenerarInput("",rs(26),"readonly","","100","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","aa ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

'DOMINIO

   celdas = ofv.GenCelda("","tt ","","","","","","DOMINIO:","si")
       celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

   if bModificar and rs("ID") <> "COREMANAGER" then
      strsql = "Select * From dominios Order by descripcion"
      sInput = ofv.GenerarCombo("Usuarios",rs("dominio").name,rs("dominio"),strsql,cn,0,1,"")

    else
      strsql = "Select * From dominios where id = " & "0" & rs("dominio")
	   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	   if not rsniv.eof then sniv = rsniv(1)
	   call ofv.cerrarconsulta(rsniv)
	   sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","aa  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","aa colspan=3 bgcolor=#f7aa22 ","","","2","","","","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


 end if


    celdas = ofv.GenCelda("","tt","","","","","","Nivel:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
      strsql = "Select * From NivelConfiden Order By NCcod"
      sInput = ofv.GenerarCombo("Usuarios",rs(4).name,rs(4),strsql,cn,0,1,"")
    else
      strsql = "Select * From NivelConfiden where NCcod = " & rs(4)
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    set cn2 = ofv.conectar(ofv.strconn2)

'Puesto
    celdas = ofv.GenCelda("","tt","","","","","","Puesto/Perfil:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
      strsql = "Select * From tipoperfil Order By descri"
      sInput = ofv.GenerarCombo("Usuarios",rs(27).name,rs(27),strsql,cn2,0,1,"")
    else
      strsql = "Select * From tipoperfil where tipoperfil = '" & rs(27) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


'Empresa

  if Session("EMPMUN") = "S" then

    celdas = ofv.GenCelda("","tt","","","","","","Empresa/U.de.Neg.:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
      strsql = "Select * From clienteplan Order By cliente"
      sInput = ofv.GenerarCombo("Usuarios",rs(29).name,rs(29),strsql,cn2,0,1,"")
    else
      strsql = "Select * From clienteplan where id = 0" & rs(29)
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  end if

  call ofv.cerrarconn(cn2)

'Legajo
    celdas = ofv.GenCelda("","tt","","","","","","Legajo:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")


   if bModificar  then
      sInput = ofv.GenerarInput(rs(28).name,rs(28) & " ","text","Usuarios","100","dt","")
    else
      sInput = ofv.GenerarInput("",rs(28),"readonly","","100","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","tt","","","","","","Estado:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    select case cint(rs("opcion"))
           case 0
                xestado = "<b><font face='wingdings' size=3 color=#008000>&#0252;</font></b>Activo - Normal"
           case 1
                xestado = "<b><font face='wingdings' size=3 color=#f7aa22>&#0252;</font></b>Activo - Password Forzada por el Administrador"
           case 2
                xestado = "<b><font face='wingdings' size=3 color=#993333>&#0251;</font></b>Bloqueado por Intentos Fallidos"
           case 3
                xestado = "<b><font face='wingdings' size=3 color=#aa0000>I</font></b><font color=#aa0000>Bloqueado por el Administrador</font>"
           case 4
                xestado = "<b><font face='wingdings' size=3 color=#f7aa22>&#0234;</font></b>Activo - <font color=#993333>Password expirada</font>"
    end select




    xestado = ucase(xestado)

    if bModificar and (cint(rs("opcion")) < 3 or cint(rs("opcion")) = 4 ) then

        scomK = rs("opcion") & chr(9) & "3"
        scomD = xestado & "&nbsp;NO BLOQUEADO" & chr(9) & "BLOQUEADO"

      '  sInput = "<font class=at><select name=" & rs("opcion").name & " onchange=submit()>"
      '  sInput = sInput & "<option value=" & rs("opcion") & ">No Bloqueado<option value=3 "
      '  sInput = sInput & ">Bloqueado</select></font>"

        sInput = ofv.gencomboFX("Usuarios",rs("opcion").name,rs("opcion"),scomK,scomD," ")

    else
      sInput = xestado
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","","","","","Email:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs("email").name,rs("email") & " ","text","Usuarios","100","dt","")
    else
      sInput = ofv.GenerarInput("",rs("email"),"readonly","","100","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    scomK = "0" & chr(9) & "1"
    scomD = "DEFAULT" & chr(9) & "ESPECIAL"

    celdas = ofv.GenCelda("","tt","","","","","","Acceso a Datos:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

      if cint(rs("dbdef")) = 1 then
        sSelect = ""
      else
        sSelect = "Selected"
      end if

      if bModificar then
      '  sInput = "<font class=at><select name=" & rs("dbdef").name & " onchange=submit()>"
      '  sInput = sInput & "<option value=1>ESPECIAL<option value=0 " & sSelect
      '  sInput = sInput & ">DEFAULT</select></font>"

         sInput = ofv.gencomboFX("Usuarios",rs("dbdef").name,rs("dbdef"),scomK,scomD," ")
      else
         if cint(rs("dbdef")) = 1 then
            xestado = "ESPECIAL"
         else
            xestado = "DEFAULT"
         end if
         sInput = ofv.GenerarInput("",xestado,"readonly","","10","dt","")
      end if

      celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=Usuarios"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

      if cint(rs("dbdef")) = 1 then

    celdas = ofv.GenCelda("","tt","","","","","","Usuario DB:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    xdbuid = ""
    if not isnull(rs("dbuid")) then xdbuid = trim(rs("dbuid"))

    if bModificar then
      sInput = ofv.GenerarInput(rs("dbuid").name,xdbuid,"text","accdat","100","dt","")
    else
      sInput = ofv.GenerarInput("",xdbuid,"readonly","","100","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","","","","","Password DB:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs("dbpwd").name,"********","text","accdat","100","dt","")
    else
      sInput = ofv.GenerarInput("","********","readonly","","100","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=Usuarios"
      formulario1 = ofv.GenForm("accdat","",sAccion,"","post","",filas,"si")
    else
      formulario1 = ofv.GenForm("","","","","post","",filas,"si")
    end if

      else
        formulario1 = ""
      end if


    celdas = ofv.GenCelda("","tt","","","","","","Grupo:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    xcampo = rs(3).name
    xcampoval = rs(3)
    xClave = rs(0)

    if bModificar then
      strsql = "Select * From Grupos Order By Descripcion"
      sInput = ofv.GenerarCombo("Usuarios2",rs(3).name,rs(3),strsql,cn,0,1,"Grupos")
    else
      if rs(2) <> "" then
        Strsql = "Select descripcion From Grupos Where ID = '" & rs(3) & "'"
        set rsGrupo = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsGrupo.eof then sGrupo = rsGrupo(0)
        call ofv.cerrarconsulta(rsGrupo)
      end if
      sInput = ofv.GenerarInput("",sGrupo,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = ofv.GenRow("","","","","10","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=Usuarios"
      formulario = ofv.GenForm("Usuarios2","",sAccion,"","post","",filas,"si")
    else
      formulario = ofv.GenForm("","","","","post","",filas,"si")
    end if

 ' Asignacion del sector

  set cn2 = ofv.conectar(ofv.strconn2)

    filas = ""

''' rs(29) empresa

    SUN = rs(29)

   if Session("EMPMUN") = "S" then

    celdas = ofv.GenCelda("","tt","","","","","","Empresa:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if bModificar then
       strsql = "Select * From clienteplan Order By cliente"
       sInput = ofv.GenerarCombo("Usuarios3",rs(29).name,rs(29),strsql,cn2,0,1,"un")
    else
        sGrupo = "S/D"
        Strsql = "Select cliente From clienteplan Where id = 0" & rs(29)
        set rsGrupo = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsGrupo.eof then sGrupo = rsGrupo(0)
        call ofv.cerrarconsulta(rsGrupo)
        sInput = sGrupo
    end if



    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

   end if


    celdas = ofv.GenCelda("","tt","","","","","","Site/Sucursal:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

   SSITE = "0" & rs(31)

    if bModificar then
      strsql = "Select * From SITES where SITEun =" & SUN & " Order By Descri"
      sInput = ofv.GenerarCombo("Usuarios3",rs(31).name,rs(31),strsql,cn2,0,1,"SITE")
    else
      sGrupo = "Sin SITE/SUCURSAL ASIGNADA"
      Strsql = "Select descri From SITES Where OBJID = 0" & rs(31)
      set rsGrupo = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsGrupo.eof then sGrupo = rsGrupo(0)
      call ofv.cerrarconsulta(rsGrupo)

      sInput = sGrupo
    end if








    celdas = celdas & ofv.GenCelda("","vt ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    celdas = ofv.GenCelda("","tt","","","","","","Sector:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")




    if bModificar then
      strsql = "Select * From eco,ECOSITE where ecoun =" & SUN & " and eco.eco = ECOSITE.ECO "
      STRSQL = STRSQL & " and ECOSITE.ecosite = " & SSITE & " Order By Descri"
      sInput = ofv.GenerarCombo("Usuarios3",rs(21).name,rs(21),strsql,cn2,0,1,"sector")
    else
      sGrupo = "Sin sector asignado"

        Strsql = "Select descri From eco Where eco = 0" & rs(21)
        set rsGrupo = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsGrupo.eof then sGrupo = rsGrupo(0)
        call ofv.cerrarconsulta(rsGrupo)

      sInput = ofv.GenerarInput("",sGrupo,"readonly","","40","dt","")
    end if



    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





  call ofv.cerrarconn(cn2)



    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=Usuarios"
      formulario2 = ofv.GenForm("Usuarios3","",sAccion,"","post","",filas,"si")
    else
      formulario2 = ofv.GenForm("","","","","post","",filas,"si")
    end if

 ' Rangos Horarios

    celdas = ofv.GenCelda("","tt","","","","","","Rango Horario:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")



    if bModificar then
      strsql = "Select * From Rangos where activo = 'SI' "
      sInput = ofv.GenerarCombo("Usuarios4",rs(22).name,rs(22),strsql,cn,1,2,"rango")
    else
      sGrupo = "Sin rango asignado"
      if rs(22) <> "" then
        sGrupo = " "
        Strsql = "Select descripcion From Rangos Where codigo = 0" & rs(22)
        set rsGrupo = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsGrupo.eof then sGrupo = rsGrupo(0)
        call ofv.cerrarconsulta(rsGrupo)
      end if
      sInput = ofv.GenerarInput("",sGrupo,"readonly","","40","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = ofv.GenRow("","","","","10","","",celdas,"si")

    scomK = "D" & chr(9) & "L" & chr(9) & "R"
    scomD = "DEFAULT" & chr(9) & "LOCAL" & chr(9) & "REMOTO"


    celdas = ofv.GenCelda("","tt","","","","","","Acceso:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
       sInput = ofv.gencomboFX("Usuarios4",rs(23).name,rs(23),scomK,scomD," ")
    else
      sInput = ofv.GenerarInput("",rs(23),"readonly","","90","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    scomK = "D" & chr(9) & "A" & chr(9) & "B"
    scomD = "DEFAULT" & chr(9) & "ALTA" & chr(9) & "BAJA"

    celdas = ofv.GenCelda("","tt","","","","","","Resolucion:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if bModificar then
       sInput = ofv.gencomboFX("Usuarios4",rs(24).name,rs(24),scomK,scomD," ")
    else
      sInput = ofv.GenerarInput("",rs(24),"readonly","","90","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=Usuarios"
      formulario3 = ofv.GenForm("Usuarios4","",sAccion,"","post","",filas,"si")
    else
      formulario3 = ofv.GenForm("","","","","post","",filas,"si")
    end if



    filas = ofv.GenRow("","","","","","","",formulario0 & formulario1 & formulario & formulario2 & formulario3,"si")

  if OpcAgregar <> "si" then
    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",filas,"si")
  END IF
    sHTML = sHTML & "</center>"
    sHTML = sHTML & "<br>"
    rs.movenext
  end if
'    sHTML = sHTML & "</div>"

'    sHTML = sHTML & "<div style='position:absolute;top:150;left:10;visibility:visible;'>"
  if OpcAgregar = "si" then
    sHTML = sHTML & "<center>"
    celdas = ofv.GenCelda("","cc colspan=3","","","","","","NUEVO USUARIO","si")
    filas  = ofv.GenRow("","","","","10","","",celdas,"si")

    sAccion = "asp/Agregar.asp?Tabla=Usuarios&volver=" & vuelta
    filas = filas & ofv.GenForm("Agregar","",sAccion,"","post","","","no")

    celdas = ofv.GenCelda("","tt","","","","","","ID:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sInput = ofv.GenerarInput("ID","","agregar","","100","dt","")
    sInput = sInput & ofv.GenerarInput("dbdef","0","hidden","","10","dt","")
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","","","","","Descripcion:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sInput = ofv.GenerarInput("Descripcion","","agregar","","100","dt","")
    sInput = sinput & ofv.GenerarInput("acceso","D","hidden","","100","dt","")
    sInput = sinput & ofv.GenerarInput("resol","D","hidden","","100","dt","")
    sInput = sinput & ofv.GenerarInput("dominio","1","hidden","","100","dt","")
    sInput = sinput & ofv.GenerarInput("site","1","hidden","","100","dt","")
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    celdas = celdas & ofv.GenFForm()
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")

    sInput = ofv.GenerarInput("","Cancelar","submit","","10","bt","")
    sInput = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "UsuariosX.asp?OpcAgregar=no"
    celdas = celdas & ofv.GenForm("Cancelar","",sAccion,"","post","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center>"


    sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")

  else

       restval = ""

    if bModificar then
       sInput = ofv.GenerarInput(xcampo,xcampoval,"hidden","","40","dt","")
       sInput = sInput & ofv.GenerarInput("botrv","DEFAULT DE GRUPO","submit","","40","bt","")
       celdas = ofv.GenCelda("","","","","","","",sInput,"si")
       sAccion = "asp/ActualizarPerfil.asp?ID=" & xClave & "&volver=" & vuelta
       restval = ofv.GenForm("UsuariosRV","",sAccion,"","post","",celdas,"si")

    end if


   celdas = ofv.GenCelda("","ttt colspan=7","","","","","","OPCIONES","si")
    filas  = ofv.GenRow("","","","","10","","",celdas,"si")


' Consulta Permisos del Usuario
    sInput = ofv.GenerarInput("","PERMISOS","submit","","10","bt","")
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "CoPermisos.asp?ID=" & nID & "&t=U"
    form3 = ofv.GenForm("pergrp","",sAccion,"_blank","post","",celdas,"si")

' Perfil del Usuario
    sInput = ofv.GenerarInput("boton","PERFIL","submit","","10","bt","")
  ''  sAccion = "GrpUsrMar.asp?Opcion=" & nID & "&t=U&o=A&volver=" & vuelta2

    sAccion = "Aplicaciones.asp?Opcion=" & nID & "&t=U&o=A&volver=" & vuelta2
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    bperf = ofv.GenForm("init","",sAccion,"","post","",celdas,"si")

' Esquema de Password
    sInput = ofv.GenerarInput("boton","ACCESO","submit","","10","bt","")
 ''   sAccion = "GrpUsrMar.asp?Opcion=" & nID & "&t=U&o=S&volver=" & vuelta2

    sAccion = "UsrSkrdef.asp?Opcion=" & nID & "&t=U&o=S&volver=" & vuelta2
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    bpsec = ofv.GenForm("init","",sAccion,"","post","",celdas,"si")

' EMPRESAS
    sInput = ofv.GenerarInput("boton","EMPRESAS","submit","","10","bt","")
    sAccion = "USREMPRE.asp?Opcion=" & nID & "&t=U&o=A&volver=" & vuelta2 & "&NOMBRE=" & REPLACE(XNOMBRE," ","_")
    celdas = ofv.GenCelda("","","","","","","",sInput,"si")
    bpsec2 = ofv.GenForm("init","",sAccion,"","post","",celdas,"si")


    filas = filas & ofv.GenRow("","","","","40","","",form3 & bperf & bpsec & bpsec2 & RESTVAL,"si")

    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","tbp","","10%","","0","","2","0","","",filas,"si")
    sHTML = sHTML & "</center>"


  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sAccion = "UsuariosxBus.asp?OpcBuscar=U"
  buscar  = ofv.GenBoton(4, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")



    if bModificar then

       sAccion = "UsuariosX.asp?OpcAgregar=si"
       Agregar = ofv.BotonAgregar(sAccion)

    end if

    sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,"")


  end if



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>