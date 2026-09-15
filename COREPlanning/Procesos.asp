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
      Session("Evento")=request.querystring("Evento")
      Session("Eventodesc")=request.querystring("Eventodesc")
      Session("Eventodesc")=ofv.ConvertirCar(Session("Eventodesc"),"*"," ")
      Session("EVMoverx")=cdbl(request.querystring("CantRegMover"))
      CantRegMoverx=0
      Session("Filtrar")= "no"  ' Esto es para que cuando apreto boton volver muestre siempre TODOS
   else
      CantRegMoverx=cdbl(request.querystring("CantRegMover"))
   end if

   Redim Preserve Perfiles(1)
   Perfiles(1) = "N"
   Redim Preserve Perfiles(2)
   Perfiles(2) = "S"

   Redim Preserve Perfiles2(1)
   Perfiles2(1) = "A"
   Redim Preserve Perfiles2(2)
   Perfiles2(2) = "M"
   Redim Preserve Perfiles2(3)
   Perfiles2(3) = "B"

   sHTML = ofv.FormHeader("PROCESOS")

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
   Fila =  ofv.genrow("","","","","","","",Celda,"si")
   Celda = ofv.GenCelda("","tx","","10%","10","","","Evento","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Eventodesc"),"si")
   Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")

'establesco el filtro solo cuando cambia el combofiltrar
   Filtrar=request.querystring("Filtrar")

   nucleo= " From Proceso Where Codevento =" & Session("Evento") 

   If Filtrar = "cri" or Filtrar = "nocri" or Filtrar = "no" then
      Session("Filtrar")=request.querystring("Filtrar")
   end if
   If Session("Filtrar") = "cri" then ' Muestro la tabla filtrada tambien por critico
      Strsql = " Select * " & nucleo & " and critico= 's' Order By Prosecuen "
      Str1 = "Select descri " & nucleo & " and critico='s' Order By Prosecuen "
      str2 = "Select descri " & nucleo & " and critico='s' Order By Descri " 

   ElseIf Session("Filtrar") = "nocri" then ' Muestro la tabla filtrada tambien por critico
      Strsql = "Select * " & nucleo & " and critico='n' Order By Prosecuen "
      Str1 = "Select descri " & nucleo & " and critico='n' Order By Prosecuen "
      str2 = "Select descri " & nucleo & " and critico='n' Order By Descri " 
   else
      Strsql="Select * " & nucleo & " Order By Prosecuen "
      Str1="Select descri " & nucleo & " Order By Prosecuen "
      str2="Select descri " & nucleo & " Order By Descri "
   End If

   set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)

' Valida si no hay registros
  if rs.eof then
     sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>NO HAY ELEMENTOS ASOCIADOS, INGRESE PROCESOS PRIMERO </b><br><br></FONT></P>"
  Else 

'Encabezados
    Formulario = ofv.Encabezados("Procesos","Sec.","Cri.","Pri.")

    rs.move(CantRegMoverx)
    posi=0
    vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
    Do While Not rs.EOF and posi < CantRegMostrarx
      sForm= "form" & posi

'Eliminar
      if bSuprimir then
        sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
        sHref = sHref & "&Tabla=Proceso&volver=" & vuelta & "&Eliminar=NSNC"
        Celda = ofv.BotonEliminar(rs(1),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

'Descripcion
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt MAXLENGTH=50","")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

'secuencia
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"5","dt MAXLENGTH=5","SI")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

'Criterio
      Combo = "<font class=at><select onchange=document." & sform & ".submit() name=" & rs(3).name & ">"
      For n = 1 To Ubound(Perfiles)
         If Perfiles(n) = rs(3) Then
            Combo = Combo & "<option value=" & Perfiles(n) & " selected>" & Perfiles(n)
         Else
            Combo = Combo & "<option value=" & Perfiles(n) & ">" & Perfiles(n)
         End If
      Next
      Combo = Combo & "</font></select></font>"
      Celda = Celda & ofv.GenCelda("","","","","","","",Combo,"si")

'Prioridad
      Combo = "<font class=at><select onchange=document." & sForm & ".submit() name=" & rs(4).name & ">"
      For n = 1 To Ubound(Perfiles2)
         If Perfiles2(n) = rs(4) Then
            Combo = Combo & "<option value=" & Perfiles2(n) & " selected>" & Perfiles2(n)
         Else
            Combo = Combo & "<option value=" & Perfiles2(n) & ">" & Perfiles2(n)
         End If
      Next
      Combo = Combo & "</font></select></font>"
      Celda = Celda & ofv.GenCelda("","","","","","","",Combo,"si")

      Fila = ofv.genrow("","","","","","","",Celda,"no")
 
      sAccion = ""
      if bModificar then
         sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
         sAccion = sAccion & "&tabla=PROCESO&volver=" & vuelta
      end if
      Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post","",Fila,"si")

' Bifurcacion a Procesos

      sForm1 = "Boton" & posi
      sInput = ofv.GenerarInput("","Contingencias","submit",sForm1,"10","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
      Celda = Celda & ofv.genfrow()
      sAccion = "CircContingencia.asp?cantregmostrar=0" & "&Proceso=" & rs(0) & "&CantRegMover=" & CantRegMoverx
      sAccion = sAccion  & "&Procesodesc=" & ofv.ConvertirCar(rs(1)," ","*") 
      Formulario = Formulario & ofv.GenForm(SForm1,"",sAccion,"","post","",celda,"si")
      posi=posi+1
      rs.Movenext
   Loop

   sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
End If

'Botonera
'Volver pagina anterior
   sAccion = "Eventos.asp?cantregmostrar=10" & "&CantRegMover=" & Session("STMoverx") 
   sInput = ofv.GenerarInput("","Volver","submit","Volver","10","bt","")
   Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
   volver = ofv.GenForm("Volver","",sAccion,"","post","",celda,"si")

   Agregar = ""
   if bAgregar then
     sAccion = "AgregarProcesos.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=PROCESO"   
     sAccion = sAccion & "&ID=" & Session("Evento") & "&str=" & ofv.convertircar(Strsql," ","_")
      Agregar = ofv.BotonAgregar(sAccion)
   end if


'Combo para busqueda alfabetica
   Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

' Boton de filtro de critico
   vuelta1= ofv.vueltaasp("",CantRegMostrarx,0) 'Cantregmoverx=0 para que siempre empiece desde el principio
   filtro= ofv.Combofiltrar(cantregmostrarx,CantRegMoverx,vuelta1,Session("Filtrar"))

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,volver,buscar,filtro)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"& chr(13)

End If ' fin del usuario

Response.Write sHTML

%>