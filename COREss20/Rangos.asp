<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  'Encabezados
  formulario =  ofv.Encabezados("#","Descripcion","activo","rest.Dias","rest.Horaria","")

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  set cn = ofv.conectar(ofv.strconn0)



  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  StrSql = "Select * From Rangos order by codigo"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  do until rs.eof 
    sInput = ofv.GenerarInput(rs(0).name,rs(0),"hidden","niveles" & rs(0),"10","dt","")
    celdas = ofv.GenCelda("","tt colspan=2","","","","","",sinput & rs(1),"si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text","Rangos" & rs(0),"60","dt","")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","60","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    if bModificar and clng(rs(1)) > 1 then
      sInput = "<font class=at><select onchange=Rangos" & rs(0) & ".submit() "
      sInput = sInput & "name=" & rs(5).name & " ><option value=NO >NO  "
      if  rs(5) = "SI" then
        sInput = sInput & "<option value=SI selected >SI  "
      else
        sInput = sInput & "<option value=SI >SI  "
      end if
      sInput = sInput & "</select></font>"
    else
      if rs(5) = "SI" then
         xestado = "SI"
      else
         xestado = "NO"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

      if  rs(5) = "SI" then

    if bModificar  and clng(rs(1)) > 1 then
      sInput = "<font class=at><select onchange=Rangos" & rs(0) & ".submit() "
      sInput = sInput & "name=" & rs(3).name & " >"
      sInput = sInput & "<option value=0 >Sin Restriccion"
      sInput = sInput & "<option value=1 >Sin Sabados y Domingos"
      sInput = sInput & "<option value=2 >Sin Sabados, Domingos y Feriados"


      select case cint(rs(3))
             case 0
                  sInput = replace(sInput,"value=0","value=0 selected")
             case 1
                  sInput = replace(sInput,"value=1","value=1 selected")
             case 2
                  sInput = replace(sInput,"value=2","value=2 selected")

      end select


      sInput = sInput & "</select></font>"
    else

      select case cint(rs(3))
             case 0
                  sInput = "Sin Restriccion"
             case 1
                  sInput = "Sin Sabados y Domingos"
             case 2
                  sInput = "Sin Sabados, Domingos y Feriados"

      end select

      sInput = ofv.GenerarInput("",sInput,"readonly","","35","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    if bModificar  and clng(rs(1)) > 1 then
      sInput = "<font class=at><select onchange=Rangos" & rs(0) & ".submit() "
      sInput = sInput & "name=" & rs(4).name & " ><option value=0 >NO  "
      if  clng(rs(4)) = 1 then
        sInput = sInput & "<option value=1 selected >SI  "
      else
        sInput = sInput & "<option value=1 >SI  "
      end if
      sInput = sInput & "</select></font>"
    else
      if clng(rs(4)) = 1 then
         xestado = "SI"
      else
         xestado = "NO"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

      if  clng(rs(4)) = 1 then
          sInput = "<a class=nt href=RangosHor.asp?Id=" & rs(0) & "&Opc=N ><b>Horarios&nbsp;>></b></a>"
          celdas = celdas & ofv.GenCelda("","cc ","","","","","",sInput,"si")
      else
          celdas = celdas & ofv.GenCelda("","aaa ","","","","","","&nbsp;","si")
      end if


    else

    celdas = celdas & ofv.GenCelda("","aaa colspan=3","","","","","","&nbsp;","si")

    end if

    if clng(rs(1)) > 1 then
      if  rs(5) = "SI" then
       filas = ofv.GenRow("","aaa style='background-color:#aaaaaa;' ","","","","","",celdas,"si")
      else
       filas = ofv.GenRow("","","","","","","",celdas,"si")
      end if
    else
       filas = ofv.GenRow("","aaa style='background-color:#336699;' ","","","","","",celdas,"si")
    end if

    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=Rangos"
      formulario = formulario & ofv.GenForm("Rangos" & rs(0),"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    rs.movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no","",Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

%>