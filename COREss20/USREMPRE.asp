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
  Nombre = ofv.convertircar(request.querystring("Nombre"),"_"," ")
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)



  sHTML = ofv.FormHeader(Session("FormN"))

  ofv.ObtenerAtributos Session("Form") ,""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cx =  ofv.conectar(ofv.strconn1)

  set cn = ofv.conectar(ofv.strconn0)

  strsql = "SELECT * From CLIENTEPLAN  order by ID"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  xgrupo = " "
  xgrp = "S"
  xusr = "S" 

  if t <> "G" then
     strsql = "SELECT grupo From usuarios Where ID = '" & ID & "' "
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if not rsu.eof then xgrupo = rsu(0) 
     call ofv.cerrarconsulta(rsu)

     strsql = "SELECT * From EMPREATRM Where ATRMTIPO = 'U' AND ATRMUSRGRP = '" & ID & "' "
     strsql = strsql & "  Order By EMPRE"
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsu.eof then xusr = "N"

     XDESCT = "USUARIO: " 

  else

      xgrupo = ID
      xusr   = "N" 
      XDESCT = "GRUPO: " 

  end if

     strsql = "SELECT * From EMPREATRM Where ATRMTIPO = 'G' AND ATRMUSRGRP = '" & xgrupo & "' "
     strsql = strsql & "  Order By EMPRE"
     set rsg = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsg.eof then xgrp = "N"


  redim aAtributos(3)
  aAtributos(0) = "CONSULTAR"
  aAtributos(1) = "ACTUALIZAR"
  aAtributos(2) = "NO"
  aAtributos(3) = "NO"

  'aAtributos(0) = "AGREGAR"
  'aAtributos(1) = "MODIFICAR"
  'aAtributos(2) = "ELIMINAR"
  'aAtributos(3) = "CONSULTAR"

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

  XDESCT2 = "Grupo:"

  XDESCT = ""


  IF T = "U" THEN
     XDESCT2 = "Usuario:"
     XDESCT = XDESCT & "<big>GRUPO:&nbsp;" & xgrupo & "</big><br><br>" 
  END IF

  XDESCT = XDESCT & "ACCESO&nbsp;A&nbsp;EMPRESAS&nbsp;DEL&nbsp;" & XDESCT2 & "&nbsp;" & ID & "&nbsp;-&nbsp;" & Nombre & "</B>"

  celdas = ofv.GenCelda("","ttt colspan=5","","","","","",XDESCT,"si")
  formulario = "<thead>" & ofv.GenRow("","","","","10","","",celdas,"si") & "</thead>"

  celdaz = ofv.GenCelda("","ttr","","","","","","Empresa&nbsp;/&nbsp;Unidad&nbsp;de&nbsp;Negocios","si")
  For iz = 0 to 1
    celdaz = celdaz & ofv.GenCelda("","ttr","","","","","",aAtributos(iz),"si")
  next
  formulario = formulario & ofv.GenRow("","","","","25","","",celdaz,"si")

  Menu = "'Form " & sMenu & "'"
  Menu2 = "'Nodo " & sMenu & "'"


  do until rs2.eof

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
      
      xdesatr = " <font color=#993333 face=tahoma size=1>(" & rs2(0) & ")</font>"
      menux = ofv.convertircar(Menu2,"'","")
      xdescri = Ucase(rs2(1)) & " - "

      xdescri = xdescri & xdesatr & xdesc2

    celdaz = ofv.GenCelda("","ttt","","70%","","","",xdescri,"si")

    if xgrp <> "N" then
       if rsg(0) = rs2(0) then
          if not rsg(5) or not rsg(6)  then 
             aAtrm(0) = abs(cint(rsg(5)))
             aAtrm(1) = abs(cint(rsg(6)))


             aAtrmD(0) = abs(cint(rsg(5)))
             aAtrmD(1) = abs(cint(rsg(6)))


             xdef = "G"
             xcolor = "#993333"
          end if
          rsg.movenext
          if rsg.eof then xgrp = "N"
        end if
     end if

    if xusr <> "N" then
       if rsu(0) = rs2(0) then
             aAtrm(0) = abs(cint(rsu(5)))
             aAtrm(1) = abs(cint(rsu(6)))

             xcolor = "#6699cc"
          rsU.movenext
          if rsU.eof then xusr = "N"
        end if
     end if



    fila = ""
    celdas = ""

    sForm = "AtributosGrupos" & rs2(0)

    For iz = 0 to 1
      if aAtrm(iz) then
        sSelect = ""
         sSel    = 1
      else
        sSelect = "Selected"
         sSel    = 0
      end if

      if bModificar then

         IF IZ = 1 THEN
            IF  aAtrm(0) = 1 THEN
                sInput = "<font class=at><select name=" & aAtributos(iz) & " onchange=submit()>"
                sInput = sInput & "<option value=1>SI  <option value=0 " & sSelect
                sInput = sInput & ">NO  </select></font>"

                sInput = ofv.combosino(sForm,sSel,aAtributos(iz))
                sInput = replace(sInput,"document.all.","document." & sForm & ".all.")


            ELSE
                xcolor = ""
                SINPUT = "---"
                sInput = sInput & ofv.GenerarInput(aAtributos(iz),aAtrm(iz),"HIDDEN","","2","dt","")
            END IF
         ELSE
            sInput = "<font class=at><select name=" & aAtributos(iz) & " onchange=submit()>"
            sInput = sInput & "<option value=1>SI  <option value=0 " & sSelect
            sInput = sInput & ">NO  </select></font>"

            sInput = ofv.combosino(sForm,sSel,aAtributos(iz))
            sInput = replace(sInput,"document.all.","document." & sForm & ".all.")


         END IF  



      else
      if aAtrm(iz) = 1 then
         xestado = "SI"
      else
         IF IZ = 1 THEN
            IF  aAtrm(0) = 1 THEN
                xestado = "NO"
            ELSE
                xcolor = ""
                xestado = "---"
            END IF
         ELSE
            xestado = "NO"
         END IF  
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
      end if

      xcolor2 = session("APLBGD")
      if aAtrm(iz) = 0 then
         if (aAtrm(iz) <> aAtrmD(iz)) or xdef = "G" then
            xcolor2 = xcolor
         end if   
      end if     

      sinput1 = ""
      sInput2 = ofv.GenerarInput(aAtributos(iz) & "D",aAtrmD(iz),"hidden","","2","dt","")
      if iz = 0 then
         sInput1 = ofv.GenerarInput("default",xdef,"hidden","","2","dt","")
      end if
    celdas = celdas & ofv.GenCelda("","ut style='border-color:" & xcolor2 & ";border-style:solid;border-width:2;' ","center","15%","","","",sInput + sinput2 + sinput1,"si")

    next


      if bModificar then
        sAccion = "asp/ActualizarATRMEMPRESA.asp?volver=" & vuelta & "&ID=" & rs2(0)
        sAccion = sAccion & "&grupousuario=" & ID & "&t=" & t  & "&Nombre="
        sAccion = sAccion & ofv.convertircar(Nombre," ","_")
        fila = fila & ofv.GenForm(sForm,"",sAccion,"","post","",celdas,"si")
      else
        fila = fila & ofv.GenForm("","","","","post","",celdas,"si")
      end if

    formulario = formulario & ofv.GenRow("","","","","","","",celdaz & fila,"si")
    rs2.movenext
  loop
  celdas = ofv.GenCelda("","abc colspan=5","center","","","","","&nbsp;","si")
  fila = ofv.GenRow("","","","","","","",celdas,"si")

  sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
  celdas = ofv.GenCelda("","abc colspan=5","center","","","","",sInput,"si")
  fila = fila & ofv.GenRow("","","","","","","",celdas,"si")
  IF T = "U" THEN 
     sAccion = "USUARIOSX.asp"
  ELSE
     sAccion = "GRUPOSX.asp"
  END IF

'''  formulario = formulario & ofv.GenForm("Volver","",sAccion,"","post","",fila,"si")

  tabla = ofv.GenTabla("","aaa style='border-color:#dddddd;border-style:solid;border-width:1;' ","","70%","","","0","0","0","","",formulario,"si")

  sHTML = sHTML & "<center><BR>" & tabla & "</center>"


  Agregar = ""
  Buscar  = ""

  IF T = "U" THEN 
     sAccion = "USUARIOSX.asp"
  ELSE
     sAccion = "GRUPOSX.asp"
  END IF
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"

  if t <> "G" then call ofv.cerrarconsulta(rsU)
  call ofv.cerrarconsulta(rsG)
  call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>