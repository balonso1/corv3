<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sHTML = ofv.FormHeader("EVENTOS")

  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
'  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

 CantRegMostrarx=10

   Carga=cdbl(request.querystring("CantRegMostrar"))

   if Carga = 0 then
      Session("Subtema")=cdbl(request.querystring("Subtema"))
      Session("Subtemadesc")=request.querystring("Subtemadesc")
      Session("Subtemadesc")=ofv.ConvertirCar(Session("Subtemadesc"),"*"," ")
      Session("STMoverx")=cdbl(request.querystring("CantRegMover"))
      CantRegMoverx=0
   else
      CantRegMoverx=cdbl(request.querystring("CantRegMover"))
   end if

  if request.querystring("seg") <> "" then
     Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn4)

   Strsql="Select * From Evento Where Codprodserv =" & Session("Subtema") & " Order By Indc "
'response.write strsql
   set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)

' Valida si no hay registros
  if rs.eof then
     sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>NO HAY ELEMENTOS ASOCIADOS, INGRESE EVENTOS PRIMERO </b><br><br></FONT></P>"
  Else 
    Celda = ofv.GenCelda("","tx","","10%","10","","","Tema","si")
    Celda = Celda & ofv.GenCelda("","cc","","90%","10","","",Session("Temadesc"),"si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","tx","","10%","10","","","Subtema","si")
    Celda = Celda & ofv.GenCelda("","cc","","90%","10","","",Session("Subtemadesc"),"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","93%","","0","","0","0","","",Fila,"si")

'Encabezados
    Formulario = ofv.Encabezados("SubTema","Secuencia","Orden")

    rs.move(CantRegMoverx)
    posi=0
    vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
    Do While Not rs.EOF and posi < CantRegMostrarx
      sForm= "form" & posi

'Eliminar
      if bSuprimir then
        sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
        sHref = sHref & "&Tabla=Evento&volver=" & vuelta & "&Eliminar=NSNC"
        Celda = ofv.BotonEliminar(rs(1),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

'Descripcion
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt MAXLENGTH=50","")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

'secuencia
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"10","dt MAXLENGTH=5","SI")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

'Orden
      sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"10","dt MAXLENGTH=5","SI")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

      Fila = ofv.genrow("","","","","","","",Celda,"no")
 
      sAccion = ""
      if bModificar then
        sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
        sAccion = sAccion & "&tabla=evento&volver=" & vuelta
      end if
      Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post","",Fila,"si")

' Bifurcacion a Procesos

      sForm1 = "Boton" & posi
      sInput = ofv.GenerarInput("","Procesos","submit",sForm1,"10","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
      Celda = Celda & ofv.genfrow()
      sAccion = "Procesos.asp?cantregmostrar=0" & "&Evento=" & rs(0) & "&CantRegMover=" & CantRegMoverx
      sAccion = sAccion  & "&Eventodesc=" & ofv.ConvertirCar(rs(1)," ","*") 
      Formulario = Formulario & ofv.GenForm(SForm1,"",sAccion,"","post","",celda,"si")
      posi=posi+1
      rs.Movenext
    Loop

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
  End If

'Botonera
'Volver pagina anterior
    sAccion = "Subtemas.asp?cantregmostrar=10" & "&CantRegMover=" & Session("STMoverx") 
    sInput = ofv.GenerarInput("","Volver","submit","Volver","10","bt","")
    Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
    volver = ofv.GenForm("Volver","",sAccion,"","post","",celda,"si")

    Agregar = ""
    if bAgregar then
      sAccion = "AgregarEventos.asp?Regisvuelta=" & CantRegMoverx &  "&Opcion=si&tabla=EVENTO&ID=" & Session("Subtema") 
      sAccion = sAccion & "&str=" & ofv.convertircar(Strsql," ","_")
      Agregar = ofv.BotonAgregar(sAccion)
    end if

   str1 = "Select descri from Evento Where Codprodserv =" & Session("Subtema") & " Order By Indc "    ' orden pagina
   str2 = "Select descri from Evento Where Codprodserv =" & Session("Subtema") & " Order By descri "   ' orden alfabetico
    Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

    sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,volver,Buscar)

    call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)

    sHTML = sHTML & "</body></html>"& chr(13)

End If ' fin del usuario

Response.Write sHTML

%>