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

   set cn = ofv.conectar(ofv.strconn4)

   sHTML = ofv.FormHeader("AGREGAR PROCEDIMIENTO DE CIRCUITO ALTERNATIVO")
 
   Strsql="Select * From " & tabla
   set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)

   sForm = "sForm1"
    
        'Fila = ofv.genrow("","st","","","","","","","si")
        Celda = ofv.GenCelda("","tx","","1%","10","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","st","","20%","10","","","Accion","si")
        Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Informacion de Apoyo","si")
        Fila =  ofv.genrow("","","","","","","",Celda,"si")

        Celda = ofv.GenCelda("","st","","","","","","&nbsp;","si")

'Generar Combo Accion
        Strsql="Select * From Tipoaccalt Order By descri"
        sID = "document." & sForm 
        sCombo = ofv.GenerarCombo(sID,rs(1).name,"",Strsql,cn,0,1,"no")
        Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")

'Generar combo Info Apoyo
        Strsql="Select * From Iar Order By descri"
        sID = "document." & sForm 
        sCombo = ofv.GenerarCombo(sID,rs(3).name,"",Strsql,cn,0,1,"no0")
        Celda = Celda & ofv.GenCelda("","","","","","","",sCombo,"si")
        Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Titulos de la segunda fila
        Celda = ofv.GenCelda("","st","","1%","10","","","Sec","si")
        Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Disparador","si")  
        Celda = Celda & ofv.GenCelda("","st","","1%","10","","","Entidad Continuadora","si")
        Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Input
        sInput = ofv.GenerarInput(rs(2).name,"","agregar","form1","1","dt MAXLENGTH=5","")
        Celda = ofv.GenCelda("","st","","1%","20","","",sInput,"si")

'Generar combo disparador
        Strsql="Select * From Tipocont where tipocontin > 400 Order By descri"
        sID = "document." & sform
        sCombo = ofv.GenerarCombo(sID,rs(4).name,"",Strsql,cn,0,1,"no0") & chr(13)
        Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")
       
'Generar combo informacion de apoyo
        Strsql="Select * From Eco Order By descri"
        sID = "document." & sform
        sCombo = ofv.GenerarCombo(sID,rs(5).name,"",Strsql,cn,0,1,"no0") & chr(13)
        Celda = Celda & ofv.GenCelda("","st","","1%","10","","",sCombo,"si")
'        Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

' Guardo el Id para actualizar

      sInput = ofv.GenerarInput(rs(6).name,ID,"hidden","","10","","")
      Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

      Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

      sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name & ".options[" & rs(1).name & ".selectedIndex].value=='nsnc') { alert( 'Accion no puede ser nula');this." & rs(1).name & ".focus();return false;} else {if (isNaN(this." & rs(2).name & ".value) || this." & rs(2).name & ".value <0 ) {alert('Secuencia debe ser numerico'); this." & rs(2).name & ".focus(); this." & rs(2).name & ".select();return false;}  else { if (this." & rs(2).name & ".value.length==0){ alert( 'Secuencia no puede ser nula');this." & rs(2).name & ".focus();return false;};};};" & chr(34)
      sAccion = "asp/Agregar.asp?Tabla=" & Tabla
      sAccion = sAccion & "&volver=../ProcAlternativo.asp?CantRegMostrar=3_CantRegMover=" 
      sAccion = sAccion & Regisvuelta  
      Formulario = ofv.GenForm("form1","",sAccion,"","post",sFuncion,Fila,"no")

   sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

   call ofv.cerrarconsulta(rs)
   call ofv.cerrarconn(cn)

'Aceptar y Cancelar
   sAccion = "ProcAlternativo.asp?CantRegMover=" & Regisvuelta & "&CantRegMostrar=3"
   sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

   sHTML = sHTML & "</body></html>"
 
End If   ' if de usuario

   Response.Write sHTML
%>

