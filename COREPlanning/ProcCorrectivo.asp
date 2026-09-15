<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
   response.expires=0

   CantRegMostrarx=8
   Carga=cdbl(request.querystring("CantRegMostrar"))
   if Carga = 0 then
      Session("CircCorr")=cdbl(request.querystring("CircCorr"))
      Session("Niveldesc")=request.querystring("Niveldesc")
      Session("Niveldesc")=ofv.ConvertirCar(Session("Niveldesc"),"*"," ")
      Session("CCMoverx")=cdbl(request.querystring("CantRegMover"))
      CantRegMoverx=0
   else
      CantRegMoverx=cdbl(request.querystring("CantRegMover"))
   end if


   sHTML = ofv.FormHeader("PROCEDIMIENTO CIRCUITO CORRECTIVO")

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

   Celda = ofv.GenCelda("","tx","","10%","10","","","Contingencia:","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Contindesc"),"si")
   Celda = Celda & ofv.GenCelda("","tx","","10%","10","","","Asociacion:","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Asociadesc"),"si")
   Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   Celda = ofv.GenCelda("","tx","","10%","10","","","Nivel:","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("Niveldesc"),"si")
   Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")

   Strsql = "Select * From Accioncorr Where Codcircorr ="
   Strsql = Strsql & Session("CircCorr") & " Order by Prioridad" 
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


   if rs.eof then
      sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>NO HAY ELEMENTOS ASOCIADOS</b><br><br></FONT></P>"
   else 
      rs.move(CantRegMoverx)

      posi=0
      vuelta= ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

'Encabezados
      Celda = ofv.GenCelda("","tt","","1%","10","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","tt","","1%","10","","","Sec","si")
      Celda = Celda & ofv.GenCelda("","tt","","2%","10","","","Accion","si")

'Agregar Accion
    if (Not rs.EOF) and ( posi< CantRegMostrarx) then   
      if bAgregar then
         sHref = "AgregarAccionesCorr.asp?ID=" & ofv.convertircar(rs(0)," ","_")
         sHref = sHref & "&Tabla=TIPOACCCORR&Proceso=si&vuelta=" &  ofv.vueltaasp("",CantRegMostrarx,Cantregmoverx)
         sImage = ofv.GenImage("","","15","10","iconos/ICorAcciones.jpg","0","")
         sLink = ofv.GenLink("","",sHref,"","Agregar&nbsp;Accion","",sImage,"si")
         Celda = Celda & ofv.GenCelda("","tt align=center","","50%","10","","",sLink,"si")
'         Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
      else
         Celda = Celda & ofv.GenCelda("","tt","","1%","10","","","&nbsp;","si")
      end if
      Formulario =  ofv.genrow("","","","","","","",Celda,"si")
    end if
   Do While ((Not rs.EOF) and ( posi< CantRegMostrarx))
      sForm= "form" & posi
'Obtengo la descripcion de la accion correctiva seleccionada
      StrNivel="Select descri From Tipoacccorr Where Tipoaccioncorr =" & rs(1)
      set rx = ofv.crearconsultaEx(StrNivel,cn,1,parametros)
      NomAccionCor= rx(0)
      call ofv.cerrarconsulta(rx)

'Eliminar
      if bSuprimir then
         sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
         sHref = sHref & "&Tabla=Accioncorr&volver=" & vuelta & "&Eliminar=NSNC"
         Celda = ofv.BotonEliminar(NomAccionCor,sHref)
      else
         Celda = ofv.GenCelda("","st","","","","","","&nbsp;","si")
      end if

'Generar Input Sec
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"2","dt MAXLENGTH=5","SI")
      Celda = Celda & ofv.GenCelda("","","","1%","10","","",sInput,"si")

'Generar Combo Accion Correctiva

      Strsql="Select * From TIPOACCCORR Order By descri"
      sID = "document." & sForm 
      sCombo = ofv.GenerarCombo(sID,rs(1).name,rs(1),Strsql,cn,0,1,"")
      Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")

      Fila = ofv.genrow("","","","","","","",Celda,"si")

      sAccion = ""
      if bModificar then
         sFuncion = "onsubmit=" & chr(34) & " if (isNaN(this." & rs(2).name & ".value) || this." & rs(2).name & ".value < 0 ) {alert('Secuencia debe ser numerico'); this." & rs(2).name & ".focus(); this." & rs(2).name & ".value=this." & rs(2).name & ".defaultValue; this." & rs(2).name & ".select();return false;} else { if (this." & rs(2).name & ".value.length==0 ){ alert( 'Secuencia no puede ser nulo.');this." & rs(2).name & ".focus();return false;} ;};" & chr(34)
         sAccion = "asp/actualizar.asp?fname=" & rs(0) & "&tabla=Accioncorr&volver=" & vuelta
      end if      
      Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post",sFuncion,Fila,"si")

      posi=posi+1
      rs.Movenext
   Loop
 
   sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
   End If ' no tiene registros
'Botonera
'Volver pagina anterior
   sAccion = "Correctivo.asp?cantregmostrar=1" & "&CantRegMover=" & Session("CCMoverx") 
   sInput = ofv.GenerarInput("","Volver","submit","Volver","10","bt","")
   Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
   volver = ofv.GenForm("Volver","",sAccion,"","post","",celda,"si")

   Agregar = ""
   if bAgregar then
     sAccion = "AgregarProcCorrectivo.asp?Regisvuelta=" & CantRegMoverx
     sAccion = sAccion & "&Opcion=si&tabla=ACCIONCORR&ID=" & Session("CircCorr")   
     sAccion = sAccion & "&str=" & ofv.convertircar(Strsql," ","_")
     Agregar = ofv.BotonAgregar(sAccion)
   end if

' Combo para busqueda alfabetica
   str = "Select TIPOACCCORR.Descri From TIPOACCCORR,ACCIONCORR Where ACCIONCORR.Codcircorr ="
   str = str & Session("CircCorr") & " and Tipoacccorr.Tipoaccioncorr = ACCIONCORR.Tipoaccioncorr "
   str1 = str & " Order by ACCIONCORR.Prioridad"  ' orden pagina
   str2 = str & " Order By TIPOACCCORR.Descri "  ' orden alfabetico

   Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

   sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,volver,buscar,"no")

   call ofv.cerrarconsulta(rs)
   call ofv.cerrarconn(cn)

   sHTML = sHTML & "</body></html>"& chr(13)

End If 

Response.Write sHTML 

%>

