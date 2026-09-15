<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
   response.expires=0

   CantRegMostrarx=3
   Carga=cdbl(request.querystring("CantRegMostrar"))
   if Carga = 0 then
      Session("CirCalt")=cdbl(request.querystring("CirCalt"))
      Session("CirCaltdesc")=request.querystring("CirCaltdesc")
      Session("CirCaltdesc")=ofv.ConvertirCar(Session("CirCaltdesc"),"*"," ")
      Session("CAMoverx")=cdbl(request.querystring("CantRegMover"))
      CantRegMoverx=0
   else
      CantRegMoverx=cdbl(request.querystring("CantRegMover"))
   end if

   sHTML = ofv.FormHeader("PROCEDIMIENTO CIRCUITO ALTERNATIVO")

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
   Celda = Celda & ofv.GenCelda("","tx","","10%","10","","","Entidad:","si")
   Celda = Celda & ofv.GenCelda("","cc","","40%","10","","",Session("CircAltdesc"),"si")
   Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")


   Strsql="Select * From Accionalt Where Codcircalt =" & Session("CircAlt") & " Order by Prioridad" 
   set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)
   If rs.eof then
      sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>NO HAY ELEMENTOS ASOCIADOS</b><br><br></FONT></P>"
   Else 
      Formulario = ""
      rs.move(CantRegMoverx)
      posi=0
      vuelta= ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
        Fila = ofv.genrow("","st","","10","","","","","si")
      Do While ((Not rs.EOF) and ( posi< CantRegMostrarx))
        sForm= "form" & posi

' agrego una linea para separacion
        Celda = ofv.GenCelda("","tx","","1%","10","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","st","","20%","10","","","Accion","si")
'Agregar Accion
        if bAgregar then
           sHref = "AgregarAccionesAlt.asp?ID=" & ofv.convertircar(rs(0)," ","_")
           sHref = sHref & "&Tabla=TIPOACCALT&Proceso=si&vuelta=" & ofv.vueltaasp("",CantRegMostrarx,Cantregmoverx)
' arreglar imagen y celda 
           sImage = ofv.GenImage("","st","15","15","iconos/ICorAcciones.jpg","0","")
           sLink = ofv.GenLink("","st",sHref,"","Agregar&nbsp;Accion","",sImage,"si")
           Celda = Celda & ofv.GenCelda("","st align=center","","2%","10","","",sLink,"si")
        else
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","","&nbsp;","si")
        end if
        Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Informacion de Apoyo","si")
        Fila =  fila & ofv.genrow("","","","","","","",Celda,"si")

'Obtengo la descripcion de la accion alternativa seleccionada
      StrNivel="Select descri From Tipoaccalt Where Tipoaccionalt =" & rs(1)
      set rx = ofv.crearconsultaEx(StrNivel,cn,1,parametros)
      NomAccionAlt= rx(0)
      call ofv.cerrarconsulta(rx)

'Eliminar
        if bSuprimir then
           sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
           sHref = sHref & "&Tabla=Accionalt&volver=" & vuelta & "&Eliminar=NSNC"
           Celda = ofv.BotonEliminar(NomAccionAlt,sHref)
        else
           Celda = ofv.GenCelda("","st","","","","","","&nbsp;","si")
        end if

'Generar Combo Accion
        Strsql="Select * From Tipoaccalt Order By descri"
        sID = "document." & sForm 
        sCombo = ofv.GenerarCombo(sID,rs(1).name,rs(1),Strsql,cn,0,1,"")
        Celda = Celda & ofv.GenCelda("","","","1% colspan=2","","","",sCombo,"si")

'Generar combo Info Apoyo
        Strsql="Select * From Iar Order By descri"
        sID = "document." & sForm 
        sCombo = ofv.GenerarCombo(sID,rs(3).name,rs(3),Strsql,cn,0,1,"sb0")
        Celda = Celda & ofv.GenCelda("","","","","","","",sCombo,"si")
        Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Titulos de la sgunda fila
        Celda = ofv.GenCelda("","st","","1%","10","","","Sec","si")
        Celda = Celda & ofv.GenCelda("","st","","2% colspan=2","10","","","Disparador","si")  
        Celda = Celda & ofv.GenCelda("","st","","1%","10","","","Entidad Continuadora","si")
        Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Input
        sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"1","dt MAXLENGTH=5","SI")
        Celda = ofv.GenCelda("","st","","1%","20","","",sInput,"si")

'Generar combo disparador
        Strsql="Select * From Tipocont where tipocontin > 400 Order By descri"
        sID = "document." & sform
        sCombo = ofv.GenerarCombo(sID,rs(4).name,rs(4),Strsql,cn,0,1,"sb0") & chr(13)
        Celda = Celda & ofv.GenCelda("","","","1% colspan=2","","","",sCombo,"si")
       
'Generar combo informacion de apoyo
        Strsql="Select * From Eco Order By descri"
        sID = "document." & sform
        sCombo = ofv.GenerarCombo(sID,rs(5).name,rs(5),Strsql,cn,0,1,"sb0") & chr(13)
        Celda = Celda & ofv.GenCelda("","st","","1%","10","","",sCombo,"si")
        Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

        sAccion = ""
        if bModificar then
           sFuncion = "onsubmit=" & chr(34) & " if (isNaN(this." & rs(2).name & ".value)) {alert('Secuencia debe ser numerico'); this." & rs(2).name & ".focus(); this." & rs(2).name & ".select();return false;} else { if (this." & rs(2).name & ".value.length==0){ alert( 'Secuencia no puede ser nula');this." & rs(2).name & ".focus();return false;};};" & chr(34)
           sAccion = "asp/actualizar.asp?fname="& rs(0) & "&tabla=Accionalt&volver=" & vuelta
        end if      
        Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post",sFuncion,Fila,"si")
	Fila = "" 
        posi=posi+1
        rs.Movenext
      Loop
 
      sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
   End If

'Botonera
'Volver pagina anterior
   sAccion = "Alternativo.asp?cantregmostrar=3" & "&CantRegMover=" & Session("CAMoverx") 
   sInput = ofv.GenerarInput("","Volver","submit","Volver","10","bt","")
   Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
   volver = ofv.GenForm("Volver","",sAccion,"","post","",celda,"si")

   Agregar = ""
   if bAgregar then
     sAccion = "AgregarProcAlternativo.asp?Regisvuelta=" & CantRegMoverx
     sAccion = sAccion & "&Opcion=si&tabla=ACCIONALT"   
     sAccion = sAccion & "&ID=" & Session("CircAlt") & "&str=" & ofv.convertircar(Strsql," ","_")
     Agregar = ofv.BotonAgregar(sAccion)
   end if

' Combo para busqueda alfabetica
   str = "Select TIPOACCALT.Descri  From ACCIONALT,TIPOACCALT Where ACCIONALT.Codcircalt ="
   str = str & Session("CircAlt") & " And ACCIONALT.tipoaccionalt = TIPOACCALT.tipoaccionalt "
   str1 = str & " Order by Prioridad"  ' orden pagina
   str2 = str & " Order By TIPOACCALT.Descri "  ' orden alfabetico

   Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

   sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,volver,buscar,"no")

   call ofv.cerrarconsulta(rs)
   call ofv.cerrarconn(cn)

   sHTML = sHTML & "</body></html>"& chr(13)

End If 

Response.Write sHTML 

%>