<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
   response.expires=0

   CantRegMostrarx=9
   Carga=cdbl(request.querystring("CantRegMostrar"))
   if Carga = 0 then
      Session("Contingencia") = cdbl(request.querystring("Contingencia"))
      Session("Contindesc") = request.querystring("Contindesc")
      Session("Contindesc") = ofv.ConvertirCar(Session("Contindesc"),"*"," ")
      Session("CNMoverx") = cdbl(request.querystring("CantRegMover"))
      CantRegMoverx=0
   else
      CantRegMoverx=cdbl(request.querystring("CantRegMover"))
   end if

   sHTML = ofv.FormHeader("CIRCUITO ALTERNATIVO")

   if request.querystring("seg") <> "" then
      Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
   end if
   ofv.ObtenerAtributos Session("Form"),""
   bAgregar = ofv.AAgregar
   bModificar = ofv.AModificar
   bSuprimir = ofv.AEliminar
   bConsultar = ofv.AConsultar

   set cn = ofv.conectar(ofv.strconn4)

   Celda = ofv.GenCelda("","tx","","10%","10","","","Tema:","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Temadesc"),"si")
   Celda = Celda & ofv.GenCelda("","tx","","10%","10","","","Subtema:","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Subtemadesc"),"si")
   Fila = ofv.genrow("","","","","","","",Celda,"si")

   Celda = ofv.GenCelda("","tx","","10%","10","","","Evento:","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Eventodesc"),"si")
   Celda = Celda & ofv.GenCelda("","tx","","10%","10","","","Proceso:","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Procesodesc"),"si")
   Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   Celda = ofv.GenCelda("","tx","","10%","10","","","Contigencia:","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Contindesc"),"si")
   Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")

'Valida si hay Entidades de Contingenacia
   Strsql="Select * From Eco Order By descri"
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

   If rs.Eof then 
      sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>HA OCURRIDO UN ERROR, VERIFIQUE SI HAY * ENTIDADES DE CONTINGENCIA * CARGADAS</b><br><br></font></p>"
      call ofv.cerrarconsulta(rs)

' Boton para Volver a pagina anterior solo en caso de error
      sInput = ofv.GenerarInput("","Volver","submit","volver","10","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
      sAccion = "CircContingencia.asp?cantregmostrar=10&CantRegMover=" & Session("PRMoverx")
      sHTML = sHTML & ofv.GenForm("volver","",sAccion,"","post","",celda,"si")
   else
      call ofv.cerrarconsulta(rs)
      Strsql="select Circalt.*, Eco.descri From Circalt, Eco Where Circalt.eco = Eco.eco and Circalt.Codcontin =" & Session("Contingencia") 
'      Strsql="Select * From Circalt Where Codcontin =" & Session("Contingencia") 
'response.write strsql
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

' detecto si hay registros
      If rs.eof then
         sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>NO HAY ELEMENTOS ASOCIADOS, INGRESE CIRCUITO ALTERNATIVO PRIMERO </b><br><br></font></p>"
      Else 
'Encabezados
         Formulario = ofv.Encabezados("Entidad de Contingencia")
  
         rs.move(CantRegMoverx)
         posi=0
         vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
         Do While Not rs.EOF and posi < CantRegMostrarx
            sForm= "form" & posi
'Eliminar
            if bSuprimir then
               sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
               sHref = sHref & "&Tabla=Circalt&volver=" & vuelta & "&Eliminar=NSNC"
               Celda = ofv.BotonEliminar(rs(3),sHref)
            else
               Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
            end if

'Generar combo Contingencia
            Strsql = "Select * From Eco Order By descri" 
            sID = "document." & sForm 
            sCombo = ofv.GenerarCombo(sID,rs(1).name,rs(1),Strsql,cn,0,1,"")
            Celda = Celda & ofv.GenCelda("","","","","","","",sCombo,"si")
            Fila = ofv.genrow("","","","","","","",Celda,"no")
            sAccion = ""
            if bModificar then
               sAccion= "asp/actualizar.asp?fname="& rs(0) & "&tabla=Circalt&volver=" & vuelta
            end if 
            Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post","",Fila,"si")

'Bifurcacion a Procedimientos
            sForm1 = "Boton" & posi
            sInput = ofv.GenerarInput("","Procedimiento","submit",sForm1,"5","bt","")
            Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
            Celda = Celda & ofv.genfrow()
            sAccion = "ProcAlternativo.asp?cantregmostrar=0" & "&Circalt=" & rs(0)
            sAccion = sAccion & "&CantRegMover=" & CantRegMoverx
            sAccion = sAccion & "&CircAltdesc=" & ofv.ConvertirCar(rs(3)," ","*") 
            Formulario = Formulario & ofv.GenForm(SForm1,"",sAccion,"","post","",celda,"si")
            posi=posi+1
            rs.Movenext
         Loop
         sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
      End If
'Botonera
'Volver pagina anterior
      sAccion = "CircContingencia.asp?cantregmostrar=10" & "&CantRegMover=" & Session("CNMoverx") 
      sInput = ofv.GenerarInput("","Volver","submit","Volver","10","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
      volver = ofv.GenForm("Volver","",sAccion,"","post","",celda,"si")

      Agregar = ""
      if bAgregar then
         sAccion = "AgregarAlternativo.asp?Regisvuelta=" & cantregmoverx 
         sAccion = sAccion & "&Opcion=si&tabla=CIRCALT" & "&ID=" & Session("Contingencia")
         sAccion = sAccion & "&str=" & ofv.convertircar(Strsql," ","_")
         Agregar = ofv.BotonAgregar(sAccion)
      end if

' Combo para busqueda alfabetica
      str1 = "Select eco.descri From Circalt,eco Where Codcontin =" & Session("Contingencia") & " and circalt.eco = eco.eco "  ' orden pagina
      str2 = "Select eco.descri From Circalt,eco Where Codcontin =" & Session("Contingencia") & " and circalt.eco = eco.eco order by eco.descri "' orden alfabetico

      Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

      sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,volver,buscar,"no")

      call ofv.cerrarconsulta(rs)
      call ofv.cerrarconn(cn)
   end if
   sHTML = sHTML & "</body></html>"& chr(13)

end if

Response.Write sHTML

%>