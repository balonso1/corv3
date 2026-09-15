<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sel = request.querystring("ID")
  opc = request.querystring("Opc")


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," ") & " - Detalle ")

  'Encabezados
  formulario =  ofv.Encabezados("Rango","hora desde","hora hasta")

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  set cn = ofv.conectar(ofv.strconn0)



  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  dim Ahora(6)
  dim iHora

  for iHora=1 to 6
      Ahora(iHora) = "NO"
  next

     res="OK"

  if Opc = "A" then
     Ahora(1) = request.form("HD1")
     Ahora(2) = request.form("HH1")
     Ahora(3) = request.form("HD2")
     Ahora(4) = request.form("HH2")
     Ahora(5) = request.form("HD3")
     Ahora(6) = request.form("HH3")


     for ii=1 to 7
         if res<>"OK" then exit for

         select case ii

                case 1      
                      if Ahora(1) = "NO" and Ahora(2) <> "NO" or Ahora(3) = "NO" and Ahora(4) <> "NO" or Ahora(5) = "NO" and Ahora(6) <> "NO" or Ahora(1) <> "NO" and Ahora(2) = "NO" or Ahora(3) <> "NO" and Ahora(4) = "NO" or Ahora(5) <> "NO" and Ahora(6) = "NO" then
                         res = "AA"
                         wmsg = "Error 1 - Rango 1,2 o 3 Incorrectamente informado"
                      end if

                case 2
                      if (Ahora(1) = "NO" and (Ahora(3) <> "NO" or Ahora(5) <> "NO" )) or (Ahora(3) = "NO" and Ahora(5) <> "NO" ) then
                         res = "BB"
                         wmsg = "Error 2 - Rango 1,2 o 3 incompleto"
                      end if


                case 3
                      if Ahora(1) <> "NO" then
                         if not (Ahora(2) > Ahora(1)) then
                            res = "C1"
                            wmsg = "Error 3 - Rango 1 hora hasta no mayor a hora desde"
                         end if 
                      end if

                case 4
                      if Ahora(3) <> "NO" then
                         if not (Ahora(4) > Ahora(3)) then
                            res = "C2"
                            wmsg = "Error 4 - Rango 2 hora hasta no mayor a hora desde"
                         end if 
                      end if

                case 5
                      if Ahora(5) <> "NO" then
                         if not (Ahora(6) > Ahora(5)) then
                            res = "C3"
                            wmsg = "Error 5 - Rango 3 hora hasta no mayor a hora desde"
                         end if 
                      end if

                case 6
                      if Ahora(1) <> "NO" and Ahora(3) <> "NO" then
                         if not (Ahora(3) > Ahora(2)) then
                            res = "D1"
                            wmsg = "Error 6 - Rango 2 debe ser mayor a Rango 1 "
                         end if 
                      end if

                case 7
                      if Ahora(3) <> "NO" and Ahora(5) <> "NO" then
                         if not (Ahora(5) > Ahora(4)) then
                            res = "D1"
                            wmsg = "Error 7 - Rango 3 debe ser mayor a Rango 2 "
                         end if 
                      end if



         end select 



     next

     if res="OK" then
        StrSql = "Update Rangos set "
        StrSql = StrSql & " HD1 = '" & Ahora(1) & "', "
        StrSql = StrSql & " HH1 = '" & Ahora(2) & "', "
        StrSql = StrSql & " HD2 = '" & Ahora(3) & "', "
        StrSql = StrSql & " HH2 = '" & Ahora(4) & "', "
        StrSql = StrSql & " HD3 = '" & Ahora(5) & "', "
        StrSql = StrSql & " HH3 = '" & Ahora(6) & "' "
        StrSql = StrSql & " where Id = " & sel
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        wmsg = "Rango Actualizado"


     end if



  else

     StrSql = "Select * From Rangos where Id = " & sel
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
       for iHora=1 to 6
           Ahora(iHora) = rs(iHora + 5)
       next
     end if
     call ofv.cerrarconsulta(rs)


  end if  

  call ofv.cerrarconn(cn)


      scomb1 = "<font class=at><select name=" 

      scomb = " > "
      scomb = scomb & "<option value='NO'>Sin Especificar</option>"
      scomb = scomb & "<option value='00:30'>00:30</option>"
      scomb = scomb & "<option value='01:00'>01:00</option>"
      scomb = scomb & "<option value='01:30'>01:30</option>"
      scomb = scomb & "<option value='02:00'>02:00</option>"
      scomb = scomb & "<option value='02:30'>02:30</option>"
      scomb = scomb & "<option value='03:00'>03:00</option>"
      scomb = scomb & "<option value='03:30'>03:30</option>"
      scomb = scomb & "<option value='04:00'>04:00</option>"
      scomb = scomb & "<option value='04:30'>04:30</option>"
      scomb = scomb & "<option value='05:00'>05:00</option>"
      scomb = scomb & "<option value='05:30'>05:30</option>"
      scomb = scomb & "<option value='06:00'>06:00</option>"
      scomb = scomb & "<option value='06:30'>06:30</option>"
      scomb = scomb & "<option value='07:00'>07:00</option>"
      scomb = scomb & "<option value='07:30'>07:30</option>"
      scomb = scomb & "<option value='08:00'>08:00</option>"
      scomb = scomb & "<option value='08:30'>08:30</option>"
      scomb = scomb & "<option value='09:00'>09:00</option>"
      scomb = scomb & "<option value='09:30'>09:30</option>"
      scomb = scomb & "<option value='10:00'>10:00</option>"
      scomb = scomb & "<option value='10:30'>10:30</option>"
      scomb = scomb & "<option value='11:00'>11:00</option>"
      scomb = scomb & "<option value='11:30'>11:30</option>"
      scomb = scomb & "<option value='12:00'>12:00</option>"
      scomb = scomb & "<option value='12:30'>12:30</option>"
      scomb = scomb & "<option value='13:00'>13:00</option>"
      scomb = scomb & "<option value='13:30'>13:30</option>"
      scomb = scomb & "<option value='14:00'>14:00</option>"
      scomb = scomb & "<option value='14:30'>14:30</option>"
      scomb = scomb & "<option value='15:00'>15:00</option>"
      scomb = scomb & "<option value='15:30'>15:30</option>"
      scomb = scomb & "<option value='16:00'>16:00</option>"
      scomb = scomb & "<option value='16:30'>16:30</option>"
      scomb = scomb & "<option value='17:00'>17:00</option>"
      scomb = scomb & "<option value='17:30'>17:30</option>"
      scomb = scomb & "<option value='18:00'>18:00</option>"
      scomb = scomb & "<option value='18:30'>18:30</option>"
      scomb = scomb & "<option value='19:00'>19:00</option>"
      scomb = scomb & "<option value='19:30'>19:30</option>"
      scomb = scomb & "<option value='20:00'>20:00</option>"
      scomb = scomb & "<option value='20:30'>20:30</option>"
      scomb = scomb & "<option value='21:00'>21:00</option>"
      scomb = scomb & "<option value='21:30'>21:30</option>"
      scomb = scomb & "<option value='22:00'>22:00</option>"
      scomb = scomb & "<option value='22:30'>22:30</option>"
      scomb = scomb & "<option value='23:00'>23:00</option>"
      scomb = scomb & "<option value='23:30'>23:30</option>"
      scomb = scomb & "<option value='24:00'>24:00</option>"
      scomb = scomb & "</select></font>"



  filas = ""


    sInput = 1
    celdas = ofv.GenCelda("","tt colspan=2","center","","","","",sinput,"si")

    if bModificar then
      sInput = scomb1 & "HD1" & replace(scomb,"'" & Ahora(1) & "'","'" & Ahora(1) & "' selected")  
    else
      sInput = ofv.GenerarInput("",Ahora(1),"readonly","","10","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","30%","","","",sInput,"si")

    if bModificar then
      sInput = scomb1 & "HH1" & replace(scomb,"'" & Ahora(2) & "'","'" & Ahora(2) & "' selected")  
    else
      sInput = ofv.GenerarInput("",Ahora(2),"readonly","","10","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","30%","","","",sInput,"si")

    filas = filas &  ofv.GenRow("","","","","","","",celdas,"si")


    sInput = 2
    celdas = ofv.GenCelda("","tt colspan=2","center","","","","",sinput,"si")

    if bModificar then
      sInput = scomb1 & "HD2" & replace(scomb,"'" & Ahora(3) & "'","'" & Ahora(3) & "' selected")  
    else
      sInput = ofv.GenerarInput("",Ahora(3),"readonly","","10","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    if bModificar then
      sInput = scomb1 & "HH2" & replace(scomb,"'" & Ahora(4) & "'","'" & Ahora(4) & "' selected")  
    else
      sInput = ofv.GenerarInput("",Ahora(4),"readonly","","10","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    filas = filas &  ofv.GenRow("","","","","","","",celdas,"si")

    sInput = 3
    celdas = ofv.GenCelda("","tt colspan=2","center","","","","",sinput,"si")

    if bModificar then
      sInput = scomb1 & "HD3" & replace(scomb,"'" & Ahora(5) & "'","'" & Ahora(5) & "' selected")  
    else
      sInput = ofv.GenerarInput("",Ahora(5),"readonly","","10","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    if bModificar then
      sInput = scomb1 & "HH3" & replace(scomb,"'" & Ahora(6) & "'","'" & Ahora(6) & "' selected")  
    else
      sInput = ofv.GenerarInput("",Ahora(6),"readonly","","10","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    filas = filas &  ofv.GenRow("","","","","","","",celdas,"si")

    for ii=1 to 2
        sInput = "&nbsp;"
        celdas = ofv.GenCelda("","aaa  colspan=4","","","","","",sinput,"si")
        filas = filas &  ofv.GenRow("","","","","","","",celdas,"si")
    next

    if Opc = "A" then

           wmsg = "<font color=#993333 ><b>** - " & wmsg & "</b></font>"

           sInput = "&nbsp;"
           celdas = ofv.GenCelda("","cc  colspan=4","","","","","",wmsg,"si")
           filas = filas &  ofv.GenRow("","","","","","","",celdas,"si")

       for ii=1 to 2
           sInput = "&nbsp;"
           celdas = ofv.GenCelda("","aaa  colspan=4","","","","","",sinput,"si")
           filas = filas &  ofv.GenRow("","","","","","","",celdas,"si")
       next
  
    end if


    sInput = "&nbsp;"
    celdas = ofv.GenCelda("","aaa  colspan=2","","","","","",sinput,"si")

    if bModificar then
      sInput = ofv.GenerarInput("","Actualizar","submit","","10","dt","")
    else
      sInput = "&nbsp;"
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

 '   sInput = "<a href=Rangos.asp ><b>Volver</b></a>"
 '   celdas = celdas & ofv.GenCelda("","tt style='border-style:outset;border-width:2;' ","","","","","",sInput,"si")

 '   sInput = ofv.GenerarInput("","Volver","button","","10","dt  onclick=" & chr(34) & "document.rangosv.submit(); return true;" & chr(34) & " ","")

    sInput = ofv.GenerarInput("","Volver","button","","10","dt","document.rangosv.submit(); return true;")
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    filas = filas &  ofv.GenRow("","","","","","","",celdas,"si")


    if bModificar then
      sAccion = "RangosHor.asp?ID=" & sel & "&Opc=A"
      formulario = formulario & ofv.GenForm("Rangos","",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if




  sHTML = sHTML  & ofv.GenTabla("","","","100%","","0","","0","0","","",formulario,"si")

  sAccion = "Rangos.asp"
  sHTML = sHTML  & ofv.GenForm("rangosv","",sAccion,"","post","",ofv.GenerarInput("","uno","hidden","","10","dt",""),"si")





  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

%>