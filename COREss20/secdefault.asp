<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%



ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  vuelta = ofv.vueltaasp("../",0,0)

  ofv.ObtenerAtributos request.querystring("seg"),""

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  formulario = ""

  set cn = ofv.conectar(ofv.strconn0)

  Strsql = "Select * From skrdef"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  sform = "Defa"

  if not rs.eof then

  celdas = ofv.GenCelda("","ttt colspan=2","","","","","","PASSWORDS:","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","tt","","","","","","Caracteres minimos:","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(0).name,rs(0),"text",sForm,"5","dtn","SI")
    else
      sInput = ofv.GenerarInput("",rs(0),"readonly","","5","dtn","")
    end if
  celdas = celdas & ofv.GenCelda("","vti","","","","","",sInput,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","tt","","","","","","Caracteres maximos:","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"5","dtn","SI")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","5","dtn","")
    end if
  celdas = celdas & ofv.GenCelda("","vti","","","","","",sInput,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","tt","","","","","","Esquema de Seguridad:","si")
      if cint(rs(2)) = 1 then
        sSelect = ""
      else
        sSelect = "Selected"
      end if

      if bModificar then
         scomK = "0" & chr(9) & "1"
         scomD = "Normal" & chr(9) & "Detallado"
         sInput = ofv.gencomboFX(sForm,rs(2).name,rs(2),scomK,scomD," ")

      else
      if cint(rs(2)) = 1 then
         xestado = "Detallado"
      else
         xestado = "Normal"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","10","dt","")
      end if

  celdas = celdas & ofv.GenCelda("","vti","","","","","",sInput,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

      if cint(rs(2)) = 1 then
         celdas = ofv.GenCelda("","ttt colspan=2","","","","","","<br>HISTORIAL:","si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         celdas = ofv.GenCelda("","tt","","","","","","Cantidad:","si")
         if bModificar then
            sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"5","dtn","SI")
         else
            sInput = ofv.GenerarInput("",rs(3),"readonly","","5","dtn","")
         end if
         celdas = celdas & ofv.GenCelda("","vti","","","","","",sInput,"si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         celdas = ofv.GenCelda("","ttt colspan=2","","","","","","<br>INTENTOS FALLIDOS:","si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         celdas = ofv.GenCelda("","tt","","","","","","Cantidad:","si")
         if bModificar then
            sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"5","dtn","SI")
         else
            sInput = ofv.GenerarInput("",rs(4),"readonly","","5","dtn","")
         end if
         celdas = celdas & ofv.GenCelda("","vti","","","","","",sInput,"si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         celdas = ofv.GenCelda("","ttt colspan=2","","","","","","<br>EXPIRACION:","si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         celdas = ofv.GenCelda("","tt","","","","","","Expiran Passwords:","si")
         if cint(rs(5)) = 1 then
            sSelect = ""
         else
            sSelect = "Selected"
         end if

         if bModificar then

            sInput = ofv.combosino(sForm,rs(5),rs(5).name)

         else
         if cint(rs(5)) = 1 then
            xestado = "SI"
         else
            xestado = "NO"
         end if   
         sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
         end if

         celdas = celdas & ofv.GenCelda("","vti","","","","","",sInput,"si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         if cint(rs(5)) = 1 then

            celdas = ofv.GenCelda("","tt","","","","","","Dias de Expiracion:","si")
            if bModificar then
               sInput = ofv.GenerarInput(rs(6).name,rs(6),"text",sForm,"5","dtn","SI")
            else
               sInput = ofv.GenerarInput("",rs(6),"readonly","","5","dtn","")
            end if
            celdas = celdas & ofv.GenCelda("","vti","","","","","",sInput,"si")
            formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")
         end if 


      end if 



    if bModificar then
       sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=skrdef"
       celdas = ofv.GenForm(sForm,"",sAccion,"","post","",formulario,"si")
    else
       celdas = ofv.GenForm("","","","","post","",formulario,"si")
    end if

    end if

  tabla = ofv.GenTabla("","aa ","","60%","","0","0","0","0","","",celdas,"si")

  sHTML = sHTML & "<center>" & tabla & "</center>"

  'Botonera
   Agregar  = ""
   Buscar   = ""
   Retornar = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>