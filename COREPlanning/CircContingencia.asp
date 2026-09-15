<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
   response.expires=0

   CantRegMostrarx=10
   Carga=cdbl(request.querystring("CantRegMostrar"))
   if Carga = 0 then
      Session("Proceso")=cdbl(request.querystring("Proceso"))
      Session("Procesodesc")=request.querystring("Procesodesc")
      Session("Procesodesc")=ofv.ConvertirCar(Session("Procesodesc"),"*"," ")
      Session("PRMoverx")=cdbl(request.querystring("CantRegMover"))
      CantRegMoverx=0
   else
      CantRegMoverx=cdbl(request.querystring("CantRegMover"))
   end if

   sHTML = ofv.FormHeader("CIRCUITO DE CONTINGENCIA")

   if request.querystring("seg") <> "" then
      Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
   end if
   ofv.ObtenerAtributos Session("Form"),""
   bAgregar = ofv.AAgregar
   bModificar = ofv.AModificar
   bSuprimir = ofv.AEliminar
   bConsultar = ofv.AConsultar

   set cn = ofv.conectar(ofv.strconn4)

   Celda = ofv.GenCelda("","tx","","10%","10","","","Tema","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Temadesc"),"si")
   Celda = Celda & ofv.GenCelda("","tx","","10%","10","","","Subtema","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Subtemadesc"),"si")
   Fila = ofv.genrow("","","","","","","",Celda,"si")

   Celda = ofv.GenCelda("","tx","","10%","10","","","Evento","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Eventodesc"),"si")
   Celda = Celda & ofv.GenCelda("","tx","","10%","10","","","Proceso","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Procesodesc"),"si")
   Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")

' Valida si hay contingencias
   Strsql = "Select Tipocont.*, Asociacion.descri From Tipocont, Asociacion" 
   Strsql = Strsql & " Where Asociacion.codasoc = Tipocont.codasoc Order By Tipocont.descri"
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

   If rs.Eof then 
      sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>HA OCURRIDO UN ERROR, VERIFIQUE SI HAY * TIPOS DE CONTINGENCIA * CARGADOS</b><br><br></font></p>"
      call ofv.cerrarconsulta(rs)

' Boton para Volver a pagina anterior solo en caso de error
      sInput = ofv.GenerarInput("","Volver","submit","volver","10","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
      sAccion = "Procesos.asp?cantregmostrar=10&CantRegMover=" & Session("PRMoverx")
      sHTML = sHTML & ofv.GenForm("volver","",sAccion,"","post","",celda,"si")

else
      call ofv.cerrarconsulta(rs)
     Strsql = "Select Contingencia.*, tipocont.descri From Contingencia,tipocont Where       Contingencia.tipocontin=Tipocont.tipocontin and Contingencia.Codproceso =" & Session("Proceso")
'     Strsql = "Select * From Contingencia Where Codproceso =" & Session("Proceso") 

      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

' detecto si hay registros
      If rs.eof then
         sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>NO HAY ELEMENTOS ASOCIADOS, INGRESE CONTINGENCIAS PRIMERO </b><br><br></font></p>"
      Else 

'Encabezados
    Formulario = ofv.Encabezados("Contingencias","","")

    rs.move(CantRegMoverx)
    posi=0
    vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
    Do While Not rs.EOF and posi < CantRegMostrarx
      sForm= "form" & posi

'Eliminar
      if bSuprimir then
        sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
        sHref = sHref & "&Tabla=Contingencia&volver=" & vuelta & "&Eliminar=NSNC"
        Celda = ofv.BotonEliminar(rs(3),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

'Generar combo Contingencia
      Strsql = "Select * From Tipocont" 
      sID = "document." & sForm 
      sCombo = ofv.GenerarCombo(sID,rs(1).name,rs(1),Strsql,cn,0,1,"")
      Celda = Celda & ofv.GenCelda("","","","","","","",sCombo,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"no")
      sAccion = ""
      if bModificar then
         sAccion= "asp/actualizar.asp?fname="& rs(0) & "&tabla=Contingencia&volver=" & vuelta
      end if
      Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post","",Fila,"si")

'obtener los valores para Contindesc Asoc Asocdesc
      Strsql = "Select Tipocont.*, Asociacion.descri From Tipocont, Asociacion" 
      Strsql = Strsql & " Where Asociacion.codasoc = Tipocont.codasoc and Tipocont.Tipocontin = " & rs(1) & " Order By Tipocont.descri"
      set rsc = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       
'Bifurcacion a Circuito Alternativo
      sForm1 = "Botonalt" & posi
      sInput = ofv.GenerarInput("","Alternativo","submit",sForm1,"5","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")

      sAccion = "Alternativo.asp?cantregmostrar=0" & "&Contingencia=" & rs(0) & "&CantRegMover=" & CantRegMoverx
      sAccion = sAccion  & "&Contindesc=" & ofv.ConvertirCar(rsc(1)," ","*") 
      Formulario = Formulario & ofv.GenForm(SForm1,"",sAccion,"","post","",celda,"si")

'Bifurcacion a Circuito Correctivo 
      sForm1 = "BotonCor" & posi
      sInput = ofv.GenerarInput("","Correctivo","submit",sForm1,"5","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
      Celda = Celda & ofv.genfrow()
      sAccion = "Correctivo.asp?cantregmostrar=0" & "&Contingencia=" & rs(0) & "&CantRegMover=" & CantRegMoverx
      sAccion = sAccion  & "&Contindesc=" & ofv.ConvertirCar(rsc(1)," ","*") 
      sAccion = sAccion  & "&Asociacion=" & rsc(2) 
      sAccion = sAccion  & "&Asociadesc=" & ofv.ConvertirCar(rsc(3)," ","*") 
      Formulario = Formulario & ofv.GenForm(SForm1,"",sAccion,"","post","",celda,"si")
      call ofv.cerrarconsulta(rsc)
      posi=posi+1
      rs.Movenext
   Loop

   sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
End If

'Botonera
'Volver pagina anterior
   sAccion = "Procesos.asp?cantregmostrar=10" & "&CantRegMover=" & Session("PRMoverx") 
   sInput = ofv.GenerarInput("","Volver","submit","Volver","10","bt","")
   Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
   volver = ofv.GenForm("Volver","",sAccion,"","post","",celda,"si")

   Agregar = ""
   if bAgregar then
     sAccion = "AgregarCircContingencia.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=CONTINGENCIA"   
     sAccion = sAccion & "&ID=" & Session("Proceso") & "&str=" & ofv.convertircar(Strsql," ","_")

      Agregar = ofv.BotonAgregar(sAccion)
   end if

' Combo para busqueda alfabetica
     str1="Select TIPOCONT.Descri From CONTINGENCIA,TIPOCONT Where CONTINGENCIA.Codproceso =" & Session("Proceso") & " And CONTINGENCIA.Tipocontin = TIPOCONT.Tipocontin "
     str2= str1 & " Order By descri"

   Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

' Boton de filtro de critico
   vuelta1= ofv.vueltaasp("",CantRegMostrarx,0) 'Cantregmoverx=0 para que siempre empiece desde el principio
   filtro= ofv.Combofiltrar(cantregmostrarx,CantRegMoverx,vuelta1,Session("Filtrar"))

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,volver,buscar,"no")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if
  sHTML = sHTML & "</body></html>"& chr(13)

End If ' fin del usuario
Response.Write sHTML
 
%>