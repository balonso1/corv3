<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ID = request.querystring("Opcion")
  t = request.querystring("t")
  sMenu = ofv.convertircar(request.querystring("Menu"),"_"," ")
  submenu = ofv.convertircar(request.querystring("Submenu"),"_"," ")
  sAplicacion = ofv.convertircar(request.querystring("Aplicacion"),"_"," ")
  Nombre = ofv.convertircar(request.querystring("Nombre"),"_"," ")
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta = vuelta & "&Opcion=" & ID & "&Aplicacion=" & sAplicacion



  sHTML = ofv.FormHeader(session("FormN"))

  set cx =  ofv.conectar(ofv.strconn1)
  smenuD = smenu
  strsql = "SELECT Menu From Menux Where Menuatr = '" & smenu & "'  And mcodapli = '"
  strsql = strsql & sAplicacion & "' "
  set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
  if not rx.eof then smenuD = Ucase(rx(0))
  call ofv.cerrarconsulta(rx)

  set cn = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos Session("Form") ,""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  strsql = "SELECT * From Atrm Where (tipo = 'F' or tipo = 'N') And aplicacion = '"
  strsql = strsql & sAplicacion & "' and depen = '" & sMenu & "'  order by atributo"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  xgrupo = " "
  xgrp = "S"
  xusr = "S" 

  if t <> "G" then
     strsql = "SELECT grupo From usuarios Where ID = '" & ID & "' "
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if not rsu.eof then xgrupo = rsu(0) 
     call ofv.cerrarconsulta(rsu)

     strsql = "SELECT * From AtrmXusuario Where usuario = '" & ID & "' "
     strsql = strsql & "And (tipo = 'F' or tipo = 'N')  and depen = '" & sMenu & "' And aplicacion = '" & sAplicacion & "'  Order By atributo"
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsu.eof then xusr = "N"
  else

      xgrupo = ID
      xusr = "N" 

  end if

     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And (tipo = 'F' or tipo = 'N')  and depen = '" & sMenu & "' And aplicacion = '" & sAplicacion & "' Order By atributo"
     set rsg = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsg.eof then xgrp = "N"

    scomK = "0" & chr(9) & "1"
    scomD = "NO" & chr(9) & "SI"



  redim aAtributos(3)
  aAtributos(0) = "AGREGAR"
  aAtributos(1) = "MODIFICAR"
  aAtributos(2) = "ELIMINAR"
  aAtributos(3) = "CONSULTAR"

  redim aAtrm(3)
  aAtrm(0) = 1
  aAtrm(1) = 1
  aAtrm(2) = 1
  aAtrm(3) = 1

  redim aAtrmD(3)
  aAtrmD(0) = 1
  aAtrmD(1) = 1
  aAtrmD(2) = 1
  aAtrmD(3) = 1

  celdas = ofv.GenCelda("","ttr colspan=5","center","","","","","PANTALLAS DE " & sMenuD & "&nbsp;&nbsp;&nbsp;" & session("GRPUSRDSC"),"si")
  formulario = "<thead>" & ofv.GenRow("","","","","22","","",celdas,"si") & "</thead>"

  celdaz = ofv.GenCelda("","tt","","","","","","&nbsp;","si")
  For iz = 0 to 3
    celdaz = celdaz & ofv.GenCelda("","tt","","","","","",aAtributos(iz),"si")
  next
  formulario = formulario & ofv.GenRow("","","","","20","","",celdaz,"si")

  Menu = "'Form " & sMenu & "'"
  Menu2 = "'Nodo " & sMenu & "'"


  do until rs2.eof

 ' if not rs2.eof then

     xdef = "D"
     xcolor = ""

     aAtrm(0) = 1
     aAtrm(1) = 1
     aAtrm(2) = 1
     aAtrm(3) = 1

     aAtrmD(0) = 1
     aAtrmD(1) = 1
     aAtrmD(2) = 1
     aAtrmD(3) = 1

      xdescri = ""
      
      xdesatr = " <font color=#993333 face=tahoma size=1>(" & rs2("Alias") & ")</font>"
      menux = ofv.convertircar(Menu2,"'","")
      if rs2("tipo") <> Menux then
         strsql = "select menu from menux where menuatr = '" & trim(rs2("Alias")) & "' "
         set rsm = ofv.crearconsultaEx(StrSql,cx,1,parametros)
         if not rsm.eof then
            xdescri = Ucase(rsm(0)) & " - "
         end if 
         call ofv.cerrarconsulta(rsm)
      else
         xdescri = "Nodo "
      end if

    if rs2("tipo") = "N" then
       xdescri = "NODO: "
       xdesc2 = "<br>&nbsp;&nbsp;&nbsp;&nbsp;Dato: " & rs2("nodo") & " Valor: " & rs2("valor") & "<br>"
    else
       xdescri = "FORM: "
       xdesc2 = ""
    end if
    xdescri = xdescri & rs2("observaciones") & xdesatr & xdesc2

    celdaz = ofv.GenCelda("","ut","","","","","",xdescri,"si")

    if xgrp <> "N" then
       if rsg("atributo") = rs2("atributo") then
          if not rsg("grpagregar") or not rsg("grpmodificar") or not rsg("grpeliminar") or not rsg("grpconsultar")  then 
             aAtrm(0) = abs(cint(rsg("grpagregar")))
             aAtrm(1) = abs(cint(rsg("grpmodificar")))
             aAtrm(2) = abs(cint(rsg("grpeliminar")))
             aAtrm(3) = abs(cint(rsg("grpconsultar")))

             aAtrmD(0) = abs(cint(rsg("grpagregar")))
             aAtrmD(1) = abs(cint(rsg("grpmodificar")))
             aAtrmD(2) = abs(cint(rsg("grpeliminar")))
             aAtrmD(3) = abs(cint(rsg("grpconsultar"))) 

             xdef = "G"
             xcolor = "#993333"
          end if
          rsg.movenext
          if rsg.eof then xgrp = "N"
        end if
     end if

    if xusr <> "N" then
       if rsu("atributo") = rs2("atributo") then
             aAtrm(0) = abs(cint(rsu("usragregar")))
             aAtrm(1) = abs(cint(rsu("usrmodificar")))
             aAtrm(2) = abs(cint(rsu("usreliminar")))
             aAtrm(3) = abs(cint(rsu("usrconsultar")))
             xcolor = "#6699cc"
          rsU.movenext
          if rsU.eof then xusr = "N"
        end if
     end if






    fila = ""
    celdas = ""

    sForm = "AtributosGrupos" & rs2("ID")

    For iz = 0 to 3
      if aAtrm(iz) then
        sSelect = ""
      else
        sSelect = "Selected"
      end if

      if bModificar then

         sInput = ofv.gencomboFX(sForm,aAtributos(iz),aAtrm(iz),scomK,scomD," ")

      else
      if aAtrm(iz) = 1 then
         xestado = "SI"
      else
         xestado = "NO"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
      end if

      xcolor2 = ""
      if aAtrm(iz) = 0 then
         if (aAtrm(iz) <> aAtrmD(iz)) or xdef = "G" then
            xcolor2 = xcolor
         end if   
      end if     

      sinput1 = ""
      sInput2 = ofv.GenerarInput(aAtributos(iz) & "D",aAtrmD(iz),"hidden","","2","dt","")
      if iz = 3 then
         sInput1 = ofv.GenerarInput("default",xdef,"hidden","","2","dt","")
      end if
    celdas = celdas & ofv.GenCelda("","","center","","",xcolor2,"",sInput + sinput2 + sinput1,"si")

    next


      if bModificar then
        sAccion = "asp/ActualizarAtributo.asp?volver=" & vuelta & "&ID=" & rs2("ID")
        sAccion = sAccion & "&menu=" & ofv.convertircar(sMenu," ","_")
        sAccion = sAccion & "&submenu=" & ofv.convertircar(submenu," ","_")
        sAccion = sAccion & "&grupousuario=" & ID & "&t=" & t  & "&Nombre="
        sAccion = sAccion & ofv.convertircar(Nombre," ","_")
        fila = fila & ofv.GenForm(sForm,"",sAccion,"","post","",celdas,"si")
      else
        fila = fila & ofv.GenForm("","","","","post","",celdas,"si")
      end if

    formulario = formulario & ofv.GenRow("","","","","","","",celdaz & fila,"si")

    rs2.movenext
  loop

'end if


  tabla = ofv.GenTabla("","","","90%","","","0","0","0","","",formulario,"si")
  sHTML = sHTML & "<center>" & tabla & "</center>"


  sAccion = "SubMenues.asp?Menu=" & ofv.convertircar(submenu," ","_") & "&Opcion=" & ID
  sAccion = sAccion & "&Aplicacion=" & ofv.convertircar(sAplicacion," ","_") & "&t=" & t
  sAccion = sAccion & "&Nombre=" & ofv.convertircar(Nombre," ","_")

  Agregar = ""
  Buscar  = ""


  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"


  if t <> "G" then call ofv.cerrarconsulta(rsU)
  call ofv.cerrarconsulta(rsG)
  call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>