<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%



ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  CantRegMostrar = 11
  if request.querystring("CantRegMover") = "" then
     CantRegMover = 0
  else 
     CantRegMover = cdbl(request.querystring("CantRegMover"))
  end if




  sid = request.querystring("codapli")
  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_codapli=" & sid

  ofv.ObtenerAtributos Session("Form"),""

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn1)

  Strsql = "Select * From CorAplic WHERE codapli = '" & sid & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then 
     sapldes = ucase(replace(rs(2),"*","_"))

  celdas = ofv.GenCelda("","TTT colspan=3","center","","","","","PROPIEDADES DE " & sapldes,"si")
  formulario = "<thead>" & ofv.GenRow("","","","","10","","",celdas,"si") & "</thead>"

  celdas = ofv.GenCelda("","aaa  colspan=3","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

     sAplicacion = rs(1)
     sform = "Apl" & rs(1)
     filas = ""

    celdas = ofv.GenCelda("","tt","","30%","","","","Version: ","si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"5","dt","")
    else
      sInput = rs(4)
    end if
    celdas = celdas & ofv.GenCelda("","UTI colspan=2","","70%","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
      

    celdas = ofv.GenCelda("","tt","","30%","","","","Icono: ","si")

    if bModificar  then
      'Icono
      sImagenValor = UCASE(rs(6) & "")
      sInput = ofv.GenerarCombo3(sForm,ofv.Getfiles("Imagenes"),"iconoapli",sImagenValor)
    else
      sinput = rs(6)
    end if
    celdas = celdas & ofv.GenCelda("","UTI","","50%","","","",sInput,"si")
    sinput = "<img border=0 src='imagenes/" & rs(6) & "' >"
    celdas = celdas & ofv.GenCelda("","UT","","20%","","","",sInput,"si")

   filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","30%","","","","Directorio: ","si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"90","dt","")
    else
      sInput = rs(7)
    end if
    celdas = celdas & ofv.GenCelda("","UTI colspan=2","","70%","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","30%","","","","Imagen: ","si")

    if bModificar  then
      'Imagen
      sImagenValor = UCASE(rs(8) & "")
      sInput = ofv.GenerarCombo3(sForm,ofv.Getfiles("Imagenes"),"imagenapli",sImagenValor)
    else
      sinput = rs(8)
    end if
    celdas = celdas & ofv.GenCelda("","UTI","","50%","","","",sInput,"si")
    sinput = "<img border=0 src='imagenes/" & rs(8) & "' >"
    celdas = celdas & ofv.GenCelda("","UT","","20%","","","",sInput,"si")

   filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","30%","","","","Ventana: ","si")

    if bModificar  then
      sInput = "<font class=at><select onchange=" & sForm & ".submit() "
      sInput = sInput & "name=targapli ><option value='Menus'>DEFAULT"
      if  rs(10) = "_blank" then
        sInput = sInput & "<option value='_blank' selected>NUEVA"
      else
        sInput = sInput & "<option value='_blank' >NUEVA"
      end if
      sInput = sInput & "</select></font>"
    else
      if  rs(10) = "_blank" then
        sInput = "NUEVA"
      else
        sInput = "DEFAULT"
      end if
    end if
    celdas = celdas & ofv.GenCelda("","UTI colspan=2","","70%","","","",sInput,"si")


   filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

   if  rs(10) = "_blank" then
    celdas = ofv.GenCelda("","tt","","30%","","","","Modulo de Inicio: ","si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(11).name,rs(11),"text",sForm,"90","dt","")
    else
      sInput = rs(11)
    end if
    celdas = celdas & ofv.GenCelda("","UTI colspan=2","","70%","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

   end if

    celdas = ofv.GenCelda("","tt","","30%","","","","Titulo: ","si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(12).name,rs(12),"text",sForm,"90","dt","")
    else
      sInput = rs(12)
    end if
    celdas = celdas & ofv.GenCelda("","UTI colspan=2","","70%","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","30%","","","","Titulo Abreviado: ","si")

    if bModificar then
      if rs(14) = "" then 
         titabr = rs(12)
      else
         titabr = rs(14)
      end if   
      sInput = ofv.GenerarInput(rs(14).name,titabr,"text",sForm,"20","dt","")
    else
      sInput = rs(14)
    end if
    celdas = celdas & ofv.GenCelda("","UTI colspan=2","","70%","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if bModificar then
       sAccion = "asp/ActualizarPropApl.asp?volver=" & vuelta & "&ID=" & rs(0) & "&tabla=Coraplic" 
      fila = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      fila = ofv.GenForm("","","","","post","",filas,"si")
    end if


    formulario = formulario & fila 

   end if

 


  celdas = ofv.GenCelda("","aaa colspan=3","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")


  tabla = ofv.GenTabla("","","","80%","","0","0","0","0","","",formulario,"si")

  sHTML = sHTML & "<center><br>" & tabla & "</center>"


  sAccion = "MastApl.asp?seg=" & session("Form")
  retornar = ofv.BotonVolver(sAccion)
  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no","no",retornar,"")

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



end if

response.write sHTML

%>