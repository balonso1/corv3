<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sHTML = ofv.FormHeader("SUBTEMAS")

  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
'  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  CantRegMostrarx=10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  Carga=cdbl(request.querystring("CantRegMostrar"))
  if Carga = 0 then
      Session("Tema")=cdbl(request.querystring("Tema"))
      Session("Temadesc")=request.querystring("Temadesc")
      Session("Temadesc")=ofv.ConvertirCar(Session("Temadesc"),"*"," ")
      Session("TMoverx")=cdbl(request.querystring("CantRegMover"))
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

' Carga de Vector para rutina de verificacion de clave repetida
  Strsql="Select * From Prodserv Where CodTema=" & Session("Tema") & " Order By Indb "
  Clave = ofv.vectorclave(Strsql,3) 

' Controlo si la clave esta repetida
  sHTML = sHTML & ofv.ValClaveRep(Clave)

  Strsql="Select * From Prodserv Where CodTema=" & Session("Tema") & " Order By Indb "
  set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)

' Valida si no hay registros
  if rs.eof then
     sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>NO HAY ELEMENTOS ASOCIADOS, INGRESE SUBTEMAS PRIMERO </b><br><br></FONT></P>"
  Else 
    Celda = ofv.GenCelda("","tx","","10%","10","","","Tema","si")
    Celda = Celda & ofv.GenCelda("","cc","","90%","10","","",Session("Temadesc"),"si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","","","93%","","0","","0","0","","",Fila,"si")

'Encabezados
    Formulario = ofv.Encabezados("SubTema","Orden")

    rs.move(CantRegMoverx)
    posi=0
    vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
    Do While Not rs.EOF and posi < CantRegMostrarx
      sForm= "form" & posi

'Eliminar
      if bSuprimir then
        sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
        sHref = sHref & "&Tabla=Prodserv&volver=" & vuelta & "&Eliminar=NSNC"
        Celda = ofv.BotonEliminar(rs(1),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

'Tema
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"110","dt MAXLENGTH=50","")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

'Orden
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"10","dt MAXLENGTH=5","SIR")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

      Fila = ofv.genrow("","","","","","","",Celda,"no")

      sAccion = ""
      if bModificar then
        sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
        sAccion = sAccion & "&tabla=Prodserv&volver=" & vuelta
      end if
      Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post","",Fila,"si")

' Bifurcacion a Eventos

      sForm1 = "Boton" & posi
      sInput = ofv.GenerarInput("","Eventos","submit",sForm1,"10","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
      Celda = Celda & ofv.genfrow()
      sAccion = "Eventos.asp?cantregmostrar=0" & "&Subtema=" & rs(0) & "&CantRegMover=" & CantRegMoverx

      sAccion = sAccion  & "&SubTemadesc=" & ofv.ConvertirCar(rs(1)," ","*") 
      Formulario = Formulario & ofv.GenForm(SForm1,"",sAccion,"","post","",celda,"si")
      posi=posi+1
      rs.Movenext
    Loop

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
  End If

'Botonera
    sAccion = "Circuitos.asp?cantregmostrar=10" & "&CantRegMover=" & Session("TMoverx") 
    sInput = ofv.GenerarInput("","Volver","submit","Volver","10","bt","")
    Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
    svolv = ofv.GenForm("Volver","",sAccion,"","post","",celda,"si")

    Agregar = ""
    if bAgregar then
      sAccion = "AgregarSubtemas.asp?Regisvuelta=" & CantRegMoverx &  "&Opcion=si&tabla=PRODSERV&ID=" & Session("Tema") 
      sAccion = sAccion & "&str=" & ofv.convertircar(Strsql," ","_")
      Agregar = ofv.BotonAgregar(sAccion)
    end if

    str1 = "Select descri From Prodserv Where Codtema =" & Session("Tema") & " Order By Indb "
    str2 = "Select descri From Prodserv Where Codtema =" & Session("Tema") & " Order By descri " 
    Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

    sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,svolv,Buscar)

    call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)

    sHTML = sHTML & "</body></html>"& chr(13)

  End If   ' if de usuario

Response.Write sHTML

%>