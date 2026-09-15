<!-- #INCLUDE FILE="IncFile/FunDb.inc" -->
<!-- #INCLUDE FILE="IncFile/FunVarias.inc" -->
<%

ok=CheckUsuario()
if ok <> "Y" then
   sHTML = ok

else

Dim rs
Dim DBQValue
public dato
dim CantRegMos 
dim CantRegMov 
Dim cn
dim parametros
Dim StrSql
response.expires=0

wbrow = explorador()
if wbrow = "MSIE" then
   wfont = 7
   wspacing = 1
   wsize = 1
   wresto = 0
else
   wfont = 8
   wspacing = 0
   wsize = 2
   wresto = 3
end if

sHTML = FormHeader("ELIMINAR CENTRO DE PROCESAMIENTO")

Tabla = request.querystring("Tabla")
vuelta = request.querystring("vuelta")
ID = request.querystring("ID")
tipoId = request.querystring("tipoId")
tipoId = queesx(tipoId)
NOM = request.querystring("NOM")
NOM = convertircar(NOM,"*"," ")
	set cn=conectar(strconn2)
'response.write tabla & chr(13)
'response.write NOM & chr(13)
'response.write ID & chr(13)
'response.write tipoID & chr(13)
'response.write vuelta & chr(13)

' Obtengo la Clave
	strsql="select * from " & tabla 
	set rs = crearconsultaEx(StrSql,cn,1,volver)
	clave = rs(0).name
'	tipo = queesx(rs(0).type)
	call cerrarconsulta(rs)

' reviso que no figure en la tabla CIRCORR
	strsql="select * from CIRCORR Where codcentalt =" & tipoID & ID & tipoID
	set rs = crearconsultaEx(StrSql,cn,1,volver)
	if not rs.eof then  
	   sMensaje = "No puede eliminar el centro de procesamiento * " & NOM & " * porque esta asociado "
	   sMensaje = sMensaje & " a uno o mas circuitos correctivos. Debe eliminarlos previamente."
	   Boton = 1
	else
	   sMensaje = "Usted esta a punto de eliminar el centro de procesamiento * " & NOM & " * definitivamente."  
	   Boton = 2
	end if
	call cerrarconsulta(rs)

'Pantalla
	sHTML = sHTML & "<table width=100% marginwidth=0 cellspacing=1 cellpadding=2>"

	sHTML = sHTML & "<tr><td  height=40 width=10% ><p align=center><font color=#336699 face=verdana size=1.5>" & chr(13)
	sHTML = sHTML & "<b>" & sMensaje & "</b></font></p></td></tr>" & chr(13)
	sHTML = sHTML & "<tr><td height= 20 /td></tr>" & chr(13)
	sHTML = sHTML & "</table>" & chr(13)

	sHTML = sHTML & "<table width=100% marginwidth=0 cellspacing=1 cellpadding=2>" & chr(13)

	if Boton = 1 then
	  vuelta  = convertircar(vuelta,"_","&") 'CAMBIO _ x &
	  vuelta = right(vuelta,len(vuelta)-3) ' SACO EL ../
'response.write vuelta & chr(13)
' Boton para Volver a pagina anterior 
          sHTML = sHTML & "<form name=Volver" & posi & " action=" & vuelta
          sHTML = sHTML & " method=post>" &chr(13)  
          sHTML = sHTML & "<td height=40 align=center>" & chr(13)
          sHTML = sHTML & "<input class=bt type=submit size=10% value='Volver' ></td>" & chr(13)
          sHTML = sHTML & "</form>"
	else
' Boton para Eliminar 
          sHTML = sHTML & "<form name=Elim" & posi & " action=asp/EliminarRegistro.asp?Volver=" & vuelta & "&tabla=" & tabla & "&ID=" & ID & "&clave=" & clave & "&Eliminar=SI"
          sHTML = sHTML & " method=post>" &chr(13)  
          sHTML = sHTML & "<td height=40 align=center>" & chr(13)
          sHTML = sHTML & "<input class=bt type=submit size=10% value='Eliminar' ></td>" & chr(13)
          sHTML = sHTML & "</form>"
' Boton para Volver a pagina anterior 
          sHTML = sHTML & "<form name=Volver" & posi & " action=asp/EliminarRegistro.asp?Volver=" & vuelta & "&tabla=" & tabla & "&ID=" & ID & "&Eliminar=NO"
          sHTML = sHTML & " method=post>" &chr(13)  
          sHTML = sHTML & "<td height=50 align=center>" & chr(13)
          sHTML = sHTML & "<input class=bt type=submit size=10% value='Volver' ></td>" & chr(13)
          sHTML = sHTML & "</form>"
	end if

	call cerrarconn(cn)

	sHTML = sHTML & "</table></body></html>"
end if
	response.write sHTML

%>