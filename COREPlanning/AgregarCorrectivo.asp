<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
   response.expires=0
   Asociacion = request.querystring("Asociacion")
   Tabla = request.querystring("Tabla")
   vuelta = request.querystring("vuelta")
   Regisvuelta = cdbl(request.querystring("Regisvuelta"))
   ID = request.querystring("ID")
   set cn= ofv.conectar(ofv.strconn4)

   sHTML = ofv.FormHeader("AGREGAR CIRCUITO CORRECTIVO")
 
   Strsql = "Select * From " & tabla
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   sForm = "FormAgregar"

'Titulos de primer linea
            Celda = ofv.GenCelda("","tx","","1%","10","","","&nbsp;","si")
            Celda = Celda & ofv.GenCelda("","st","","20%","10","","","Nivel","si")
            Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Centro Procesamiento","si")
            Fila =  fila & ofv.genrow("","","","","","","",Celda,"si")

   Celda = ofv.GenCelda("","st","","","","","","&nbsp;","si")

'Generar Combo Nivel
            Strsql="Select * From Niveles Where codasoc =" &  Session("Asociacion") 
            Strsql= Strsql & " Order By secuen"
            sID = "document." & sForm 
            sCombo = ofv.GenerarCombo(sID,rs(1).name,"",Strsql,cn,0,1,"no")
            Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")

'Generar combo Centro de Procesamiento
            Strsql="Select * From CENPROCCORR Order By descri"
            sID = "document." & sForm 
            sCombo = ofv.GenerarCombo(sID,rs(7).name,"",Strsql,cn,0,1,"no0")
            Celda = Celda & ofv.GenCelda("","","","","","","",sCombo,"si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Titulos de la segunda fila
            Celda = ofv.GenCelda("","st","","1%","10","","","Plazo","si")
            Celda = Celda & ofv.GenCelda("","st","","2%","10","","","Proveedor Principal","si")  
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","","Proveedor Alternativo","si")
            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

'Generar Input Plazo    
            sInput = ofv.GenerarInput(rs(2).name,"","agregar",sForm,"2","dt MAXLENGTH=5","")
            Celda = ofv.GenCelda("","st","","1%","20","","",sInput,"si")

'Generar combo Proveedor Principal
            Strsql="Select * From Provcorr Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(3).name,"",Strsql,cn,0,1,"no0")
            Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")
                        
'Generar combo alternativo
            Strsql="Select * From Provcorr Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(4).name,"",Strsql,cn,0,1,"no0")
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","",sCombo,"si")
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
            sCombo = ofv.GenerarCombo(sID,rs(5).name,"",Strsql,cn,0,1,"no0")
            Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")
       
'Generar combo Equipo Alternativo
            Strsql="Select * From Equipcorr Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(6).name,"",Strsql,cn,0,1,"no0")
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","",sCombo,"si")
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
            sCombo = ofv.GenerarCombo(sID,rs(8).name,"",Strsql,cn,0,1,"no0")
            Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")
       
'Generar combo Inf. Respaldo 2
            Strsql="Select * From INFRESPCORR Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(9).name,"",Strsql,cn,0,1,"no0")
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","",sCombo,"si")
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
            sCombo = ofv.GenerarCombo(sID,rs(10).name,"",Strsql,cn,0,1,"no0")
            Celda = Celda & ofv.GenCelda("","","","1%","","","",sCombo,"si")
       
'Generar combo Inf. Respaldo 4
            Strsql="Select * From INFRESPCORR Order By descri"
            sID = "document." & sform
            sCombo = ofv.GenerarCombo(sID,rs(11).name,"",Strsql,cn,0,1,"no0")
            Celda = Celda & ofv.GenCelda("","st","","1%","10","","",sCombo,"si")

'Guardo el Id para actualizar
   sInput = ofv.GenerarInput(rs(12).name,ID,"hidden","","10","","")
   Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


            Fila = Fila & ofv.genrow("","st","","","","","",Celda,"si")

   sFuncion =" onsubmit= " & chr(34) & "if (this." & rs(1).name & ".options[" & rs(1).name & ".selectedIndex].value=='nsnc') { alert( 'Accion no puede ser nula');this." & rs(1).name & ".focus();return false;} else {if (isNaN(this." & rs(2).name & ".value) || this." & rs(2).name & ".value < 0 ) {alert('Plazo debe ser numerico'); this." & rs(2).name & ".focus(); this." & rs(2).name & ".defaultValue; this." & rs(2).name & ".select();return false;} else { if (this." & rs(2).name & ".value.length==0 ){ alert( 'Plazo no puede ser nulo');this." & rs(2).name & ".focus();return false;};};};" & chr(34)
   sAccion = "asp/Agregar.asp?Tabla=" & Tabla 
   sAccion = sAccion & "&volver=../Correctivo.asp?CantRegMostrar=1_CantRegMover=" & Regisvuelta
   Formulario = ofv.GenForm(SForm,"",sAccion,"","post",sFuncion,Fila,"no")
   sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

   call ofv.cerrarconsulta(rs)
   call ofv.cerrarconn(cn)

'Aceptar y Cancelar
   sAccion = "Correctivo.asp?CantRegMover=" & Regisvuelta & "&CantRegMostrar=1"
   sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

   sHTML = sHTML & "</body></html>"
 
End If   ' if de usuario

Response.Write sHTML
%>
   
