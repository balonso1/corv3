<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%



ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(session("FormN"))




  nid = request.querystring("opcion")
  t = request.querystring("t")
  vuelta = ofv.vueltaasp("../",0,0) & "_opcion=" & nid & "_t=" & t

  ofv.ObtenerAtributos request.querystring("USUARIOS"),""

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  celdas = ofv.GenCelda("","ttr colspan=2","center","","","","","ESQUEMA DE PASSWORDS","si")
  formulario = ofv.GenRow("","","","","22","","",celdas,"si")


  set cn = ofv.conectar(ofv.strconn0)

  Strsql = "Select * From Usuarios where ID ='" & nid & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  sform = "Defa"

  if not rs.eof then

  celdas = ofv.GenCelda("","tt","","40%","","","","Seguridad:","si")
      if cint(rs("defskr")) = 1 then
        sSelect = ""
      else
        sSelect = "Selected"
      end if

      scomK = "0" & chr(9) & "1"
      scomD = "DEFAULT" & chr(9) & "ESPECIAL"

      if bModificar then

         sInput = ofv.gencomboFX(sForm,"defskr",rs("defskr"),scomK,scomD," ")

      '  sInput = "<font class=at><select name=" & rs("defskr").name & " onchange=submit()>"
      '  sInput = sInput & "<option value=1>ESPECIAL<option value=0 " & sSelect
      '  sInput = sInput & ">DEFAULT</select></font>"


      else
      if cint(rs("defskr")) = 1 then
         xestado = "Especial"
      else
         xestado = "Default"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","10","dt","")
      end if

  celdas = celdas & ofv.GenCelda("","uti","center","60%","","","",sInput,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

     if cint(rs("defskr")) = 1 then

  celdas = ofv.GenCelda("","ttt colspan=2","","","","","","PASSWORDS:","si")
  formulario = formulario & ofv.GenRow("","","","","25","","",celdas,"si")

  celdas = ofv.GenCelda("","tt","","","","","","Caracteres minimos:","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs("passmin").name,rs("passmin"),"text",sForm,"5","dtn","SI")
    else
      sInput = ofv.GenerarInput("",rs("passmin"),"readonly","","5","dtn","")
    end if
  celdas = celdas & ofv.GenCelda("","uti","center","","","","",sInput,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","tt","","","","","","Caracteres maximos:","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs("passmax").name,rs("passmax"),"text",sForm,"5","dtn","SI")
    else
      sInput = ofv.GenerarInput("",rs("passmax"),"readonly","","5","dtn","")
    end if
  celdas = celdas & ofv.GenCelda("","uti","center","","","","",sInput,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","tt","","","","","","Esquema de Seguridad:","si")
      if cint(rs("tipoesq")) = 1 then
        sSelect = ""
      else
        sSelect = "Selected"
      end if

      if bModificar then

         scomK = "0" & chr(9) & "1"
         scomD = "Normal" & chr(9) & "Detallado"
         sInput = ofv.gencomboFX(sForm,"tipoesq",rs("tipoesq"),scomK,scomD," ")

     '   sInput = "<font class=at><select name=" & rs("tipoesq").name & " onchange=submit()>"
     '   sInput = sInput & "<option value=1>Detallado  <option value=0 " & sSelect
     '   sInput = sInput & ">Normal  </select></font>"

      else
      if cint(rs("tipoesq")) = 1 then
         xestado = "Detallado"
      else
         xestado = "Normal"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","10","dt","")
      end if

  celdas = celdas & ofv.GenCelda("","uti","center","","","","",sInput,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

      if cint(rs("tipoesq")) = 1 then
         celdas = ofv.GenCelda("","ttt colspan=2","","","","","","HISTORIAL:","si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         celdas = ofv.GenCelda("","tt","","","","","","Cantidad:","si")
         if bModificar then
            sInput = ofv.GenerarInput(rs("passhist").name,rs("passhist"),"text",sForm,"5","dtn","SI")
         else
            sInput = ofv.GenerarInput("",rs("passhist"),"readonly","","5","dtn","")
         end if
         celdas = celdas & ofv.GenCelda("","uti","center","","","","",sInput,"si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         celdas = ofv.GenCelda("","ttt colspan=2","","","","","","INTENTOS FALLIDOS:","si")
         formulario = formulario & ofv.GenRow("","","","","25","","",celdas,"si")

         celdas = ofv.GenCelda("","tt","","","","","","Cantidad:","si")
         if bModificar then
            sInput = ofv.GenerarInput(rs("intfall").name,rs("intfall"),"text",sForm,"5","dtn","SI")
         else
            sInput = ofv.GenerarInput("",rs("intfall"),"readonly","","5","dtn","")
         end if
         celdas = celdas & ofv.GenCelda("","uti","center","","","","",sInput,"si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         celdas = ofv.GenCelda("","ttt colspan=2","","","","","","EXPIRACION:","si")
         formulario = formulario & ofv.GenRow("","","","","25","","",celdas,"si")

         celdas = ofv.GenCelda("","tt","","","","","","Expiran Passwords:","si")
         if cint(rs("expira")) = 1 then
            sSelect = ""
         else
            sSelect = "Selected"
         end if

         if bModificar then

            sInput = ofv.combosino(sForm,rs("expira"),"expira")

           ' sInput = "<font class=at><select name=" & rs("expira").name & " onchange=submit()>"
           ' sInput = sInput & "<option value=1>SI  <option value=0 " & sSelect
           ' sInput = sInput & ">NO  </select></font>"

         else
         if cint(rs("expira")) = 1 then
            xestado = "SI"
         else
            xestado = "NO"
         end if   
         sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
         end if

         celdas = celdas & ofv.GenCelda("","uti","center","","","","",sInput,"si")
         formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

         if cint(rs("expira")) = 1 then

            celdas = ofv.GenCelda("","tt","","","","","","Dias de Expiracion:","si")
            if bModificar then
               sInput = ofv.GenerarInput(rs("diasexp").name,rs("diasexp"),"text",sForm,"5","dtn","SI")
            else
               sInput = ofv.GenerarInput("",rs("diasexp"),"readonly","","5","dtn","")
            end if
            celdas = celdas & ofv.GenCelda("","uti","center","","","","",sInput,"si")
            formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")
         end if 


      end if 
    end if 



    if bModificar then
       sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=Usuarios"
       celdas = ofv.GenForm(sForm,"",sAccion,"","post","",formulario,"si")
    else
       celdas = ofv.GenForm("","","","","post","",formulario,"si")
    end if

    end if

  tabla = ofv.GenTabla("","","","60%","","0","0","0","0","","",celdas,"si")

'''  sHTML = sHTML & "<center>" & tabla & "</center></body></html>"

  sHTML = sHTML & "<center>" & tabla & "</center>"

  Agregar = ""
  Buscar  = ""

  sAccion = "Usuariosx.asp"
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>