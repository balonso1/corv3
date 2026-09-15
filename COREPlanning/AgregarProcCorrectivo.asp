<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
   response.expires=0

   Tabla = request.querystring("Tabla")
   vuelta = request.querystring("vuelta")
   Regisvuelta = cdbl(request.querystring("Regisvuelta"))
   ID = request.querystring("ID")
'response.write regisvuelta
   set cn=ofv.conectar(ofv.strconn4)

   sHTML = ofv.FormHeader("AGREGAR PROCEDIMIENTO CIRCUITO CORRECTIVO")

   sForm = "FormAgregar" 
   Strsql = "Select * From " & tabla
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

'Encabezados
   Formulario = ofv.Encabezados("Sec","Accion")

'Generar Input Sec
      Celda = ofv.gencelda("","","","","","","","","si")
      sInput = ofv.GenerarInput(rs(2).name,"","agregar","AgregarCir","3","dt MAXLENGTH=5","")
      Celda = Celda & ofv.GenCelda("","","","1%","10","","",sInput,"si")

'Generar combo Accion Proc. Correctivo
      
      Strsql="Select * From TIPOACCCORR Order By descri"
      sID = "document." & sForm1
      sCombo = ofv.GenerarCombo(sID,rs(1).name,"",Strsql,cn,0,1,"no")
      Celda =  Celda & ofv.GenCelda("","","","","","","",sCombo,"si")

'Guardo el Id para actualizar
      sInput = ofv.GenerarInput(rs(3).name,ID,"hidden","","10","","")
      Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

      Fila = ofv.genrow("","","","","","","",Celda,"si")

      sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name & ".options[" & rs(1).name & ".selectedIndex].value=='nsnc') { alert( 'Accion no puede ser nula');this." & rs(1).name & ".focus();return false;} else {if (isNaN(this." & rs(2).name & ".value) || this." & rs(2).name & ".value < 0 ) {alert('Secuencia debe ser numerico'); this." & rs(2).name & ".focus(); this." & rs(2).name & ".defaultValue; this." & rs(2).name & ".select();return false;}   else { if (this." & rs(2).name & ".value.length==0 ){ alert( 'Secuencia no puede ser nulo');this." & rs(2).name & ".focus();return false;};};};" & chr(34)
      sAccion = "asp/Agregar.asp?Tabla=" & Tabla
      sAccion = sAccion & "&volver=../ProcCorrectivo.asp?CantRegMostrar=8_CantregMover="
      sAccion = sAccion & Regisvuelta  

      Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post",sFuncion,Fila,"no")

   sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

   call ofv.cerrarconsulta(rs)
   call ofv.cerrarconn(cn)

'Aceptar y Cancelar
   sAccion = "ProcCorrectivo.asp?CantRegMover=" & Regisvuelta & "&CantRegMostrar=4"
   sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

   sHTML = sHTML & "</body></html>"
 
End If   ' if de usuario

   Response.Write sHTML
%>

