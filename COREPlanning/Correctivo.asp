<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
   response.expires=0

   CantRegMostrarx=1
   Carga=cdbl(request.querystring("CantRegMostrar"))
   if Carga = 0 then
      Session("Contingencia")=cdbl(request.querystring("Contingencia"))
      Session("Contindesc")=request.querystring("Contindesc")
      Session("Contindesc")=ofv.ConvertirCar(Session("Contindesc"),"*"," ")
      Session("Asociacion")=cdbl(request.querystring("Asociacion"))
      Session("Asociadesc")=request.querystring("Asociadesc")
      Session("Asociadesc")=ofv.ConvertirCar(Session("Asociadesc"),"*"," ")
      Session("CNMoverx")=cdbl(request.querystring("CantRegMover"))
      CantRegMoverx=0
   else
      CantRegMoverx=cdbl(request.querystring("CantRegMover"))
   end if

   sHTML = ofv.FormHeader("CIRCUITO CORRECTIVO")

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

   sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")

' sHTML = sHTML & "<br size=1 >"

   set cn=ofv.conectar(ofv.strconn4)

   Strsql = "Select * From Niveles Where codasoc =" &  Session("Asociacion") & " Order By secuen"
'response.write strsql
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'verifico que alla niveles asociados
   If rs.Eof then 
      sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>HA OCURRIDO UN ERROR, VERIFIQUE SI HAY * NIVELES * CARGADOS</b><br><br></FONT></P>"
      call ofv.cerrarconsulta(rs)

' Boton para Volver a pagina anterior solo en caso de error
      sInput = ofv.GenerarInput("","Volver","submit","volver","10","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
      sAccion = "CircContingencia.asp?cantregmostrar=1&CantRegMover=" & Session("CNMoverx")
      sHTML = sHTML & ofv.GenForm("volver","",sAccion,"","post","",celda,"si")
   Else  

      Strsql = "Select * From Circorr Where Codcontin =" & Session("Contingencia") 
'response.write strsql
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      if rs.eof then
         sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>NO HAY ELEMENTOS ASOCIADOS</b><br><br></FONT></P>"
      else 
         Formulario = ""
         rs.move(CantRegMoverx)
         posi=0
         vuelta= ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
         Do While ((Not rs.EOF) and ( posi< CantRegMostrarx))
            sForm= "form" & posi

' Obtengo la descripcion del nivel seleccionado
            StrNivel="Select descri From Niveles Where codasoc =" &  Session("Asociacion")
            StrNivel= StrNivel & " and Codnivel=" & rs(1) & " Order By secuen"
            set rx = ofv.crearconsultaEx(StrNivel,cn,1,parametros)
            NomNivel= rx(0)
            call ofv.cerrarconsulta(rx)

' agrego una linea para separacion
            Fila = ofv.genrow("","st","","20","","","","","si")

'Titulos de primer linea
            Celda = ofv.GenCelda("","tx","","1%","10","","","&nbsp;","si")
            Celda = Celda & ofv.GenCelda("","st","","20%","10","","","Nivel","si")
            Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Centro Procesamiento","si")
            Fila =  fila & ofv.genrow("","","","","","","",Celda,"si")

 'Eliminar
            if bSuprimir then
               sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
               sHref = sHref & "&Tabla=Circorr&volver=" & vuelta & "&Eliminar=NSNC"
               Celda = ofv.BotonEliminar(NomNivel,sHref)
            else
               Celda = ofv.GenCelda("","st","","","","","","&nbsp;","si")
            end if

'Generar Combo Nivel
            Strsql="Select * From Niveles Where codasoc =" &  Session("Asociacion") 
            Strsql= Strsql & " Order By secuen"
            sID = "document." & sForm 
            sCombo = ofv.GenerarCombo(sID,rs(1).name,rs(1),Strsql,cn,0,1,"")
            Celda = Celda & ofv.GenCelda("","st","","1%","","","",sCombo,"si")

'Generar combo Centro de Procesamiento
            Strsql="Select * From CENPROCCORR Order By descri"
            sID = "document." & sForm 
            sCombo = ofv.GenerarCombo(sID,rs(7).name,rs(7),Strsql,cn,0,1,"sb0")
            Celda = Celda & ofv.GenCelda("","","","","","","",sCombo,"si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Titulos de la segunda fila
            Celda = ofv.GenCelda("","st","","1%","10","","","Plazo","si")
            Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Proveedor Principal","si")  
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","","Proveedor Alternativo","si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Input Plazo
            sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sform,"2","dt MAXLENGTH=5","SI")
            Celda = ofv.GenCelda("","st","","1%","20","","",sInput,"si")

'Generar combo Proveedor Principal
            Strsql="Select * From Provcorr Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(3).name,rs(3),Strsql,cn,0,1,"sb0")
            Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")
       
'Generar combo alternativo
            Strsql="Select * From Provcorr Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(4).name,rs(4),Strsql,cn,0,1,"sb0")
            Celda = Celda & ofv.GenCelda("","","","1%","10","","",sCombo,"si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Titulos de la tercera fila
            Celda = ofv.GenCelda("","st","","1%","10","","","&nbsp;","si")
            Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Equipo Principal","si")  
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","","Equipo Alternativo","si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Columana vacia
            Celda = ofv.GenCelda("","st","","1%","20","","","&nbsp;","si")

'Generar combo Equipo Principal
            Strsql="Select * From Equipcorr Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(5).name,rs(5),Strsql,cn,0,1,"sb0")
            Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")
       
'Generar combo Equipo Alternativo
            Strsql="Select * From Equipcorr Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(6).name,rs(6),Strsql,cn,0,1,"sb0")
            Celda = Celda & ofv.GenCelda("","","","1%","10","","",sCombo,"si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")
 
'Generar Titulos de la cuarta fila
            Celda = ofv.GenCelda("","st","","1%","10","","","&nbsp;","si")
            Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Info. Respaldo 1","si")  
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","","Info. Respaldo 2","si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Columana vacia
            Celda = ofv.GenCelda("","st","","1%","20","","","&nbsp;","si") 

'Generar combo Inf. Respaldo 1
            Strsql="Select * From INFRESPCORR Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(8).name,rs(8),Strsql,cn,0,1,"sb0")
            Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")
       
'Generar combo Inf. Respaldo 2
            Strsql="Select * From INFRESPCORR Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(9).name,rs(9),Strsql,cn,0,1,"sb0")
            Celda = Celda & ofv.GenCelda("","","","1%","10","","",sCombo,"si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Titulos de la quinta fila
            Celda = ofv.GenCelda("","st","","1%","10","","","&nbsp;","si")
            Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Info. Respaldo 3","si")  
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","","Info. Respaldo 4","si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Columana vacia
            Celda = ofv.GenCelda("","st","","1%","20","","","&nbsp;","si") 

'Generar combo Inf. Respaldo 3
            Strsql="Select * From INFRESPCORR Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(10).name,rs(10),Strsql,cn,0,1,"sb0")
            Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")
       
'Generar combo Inf. Respaldo 4
            Strsql="Select * From INFRESPCORR Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(11).name,rs(11),Strsql,cn,0,1,"sb0")
            Celda = Celda & ofv.GenCelda("","","","1%","10","","",sCombo,"si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

            sAccion = ""
            if bModificar then
               sFuncion = "onsubmit=" & chr(34) & " if (isNaN(this." & rs(2).name & ".value)) {alert('Plazo debe ser numerico'); this." & rs(2).name & ".focus(); this." & rs(2).name & ".select();return false;} else { if (this." & rs(2).name & ".value.length==0 || this." & rs(2).name & ".value < 0 ){ alert( 'Plazo no puede ser nulo ni negativo');this." & rs(2).name & ".focus();return false;} ;};" & chr(34)
               sAccion = "asp/actualizar.asp?fname="& rs(0) & "&tabla=Circorr&volver=" & vuelta
            end if      
            Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post",sFuncion,Fila,"si")
            sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

' Boton para ir a Procedimiento
            sInput = ofv.GenerarInput("","Procedimiento","submit","BotonProc","10","bt","")
            Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
            sAccion = "ProcCorrectivo.asp?cantregmostrar=0&CantRegMover=" & CantRegMoverx
            sAccion = sAccion & "&CircCorr=" & rs(0)
            sAccion = sAccion  & "&Niveldesc=" & ofv.ConvertirCar(NomNivel," ","*") 
            Formulario = ofv.GenForm("BotonProc","",sAccion,"","post","",celda,"si")
            sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

            posi=posi+1
            rs.Movenext
         Loop
      end if  ' cierro si no hay registros 

'Botonera
'Volver pagina anterior
      sAccion = "CircContingencia.asp?cantregmostrar=1" & "&CantRegMover=" & Session("CNMoverx") 
      sInput = ofv.GenerarInput("","Volver","submit","Volver","10","bt","")
      Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
      volver = ofv.GenForm("Volver","",sAccion,"","post","",celda,"si")

      Agregar = ""
      if bAgregar then
         sAccion = "AgregarCorrectivo.asp?Regisvuelta=" & CantRegMoverx
         sAccion = sAccion & "&Opcion=si&tabla=CIRCORR"   
         sAccion = sAccion & "&ID=" & Session("Contingencia")
         sAccion = sAccion & "&str=" & ofv.convertircar(Strsql," ","_")
         sAccion = sAccion & "&Asociacion=" & Session("Asociacion")
         Agregar = ofv.BotonAgregar(sAccion)
      end if

' Combo para busqueda alfabetica
      str = "Select NIVELES.Descri  From CIRCORR,NIVELES Where CIRCORR.Codcontin ="
      str = str & Session("Contingencia") & " And CIRCORR.Codnivel = NIVELES.Codnivel "
      str1 = str  ' orden pagina
      str2 = str & " Order By NIVELES.Descri "  ' orden alfabetico

      Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

      sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,volver,buscar,"no")

      call ofv.cerrarconsulta(rs)
      call ofv.cerrarconn(cn)
   end if  ' cierro testeo de Nivel
   sHTML = sHTML & "</body></html>"& chr(13)
End If 

Response.Write sHTML 

%>





