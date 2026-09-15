<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%
ok=ofv.CheckUsuario()
response.buffer = true

response.expires=0

Tabla = request.querystring("Tabla")

sid = request.querystring("codapli")
sid2 = request.querystring("ATRM")
salias = request.querystring("alias")

Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")


on error resume next

sDupKey = 0

set cn = ofv.conectar(ofv.strconn1)




strsql = "select * from " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Insert Into " & tabla & " ("
For Each x In Request.form
  strsql = strsql & x & "," 
Next
strsql = mid(strsql,1,len(strsql)-1)
strsql = strsql & ") Values ("

ValorNulo = False
TipoValido = True
For Each x In Request.form
  vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
  s = ofv.QueEsx(rs(x).type)
  if rs(x).type = 200 then s = "'"
  ValorNulo = ofv.EsNulo(rs(x),vNuevoValor)
  if ValorNulo then exit for
'  TipoValido = ofv.ValidarTipo(rs(x),vNuevoValor)
'  if not TipoValido then exit for
  if vNuevoValor = "TRUE" then
     vNuevoValor = 1
  elseif vNuevoValor = "FALSE" then
     vNuevoValor = 0
  end if
  strsql = strsql & s & vNuevoValor & s & ","
Next

if not TipoValido or ValorNulo then
  on error resume next
  response.write ofv.ErrorTipo(volver,Request.Form(x),rs(x).name)
  if err.number <> 0 then 
     response.write ofv.ErrorTipo(volver,"","")
     response.flush
  end if
else
  strsql = mid(strsql,1,len(strsql)-1) & ")"
' response.write strsql
  call ofv.crearconsultaEx(StrSql,cn,1,volver)
  
  set errores = cn.errors
  if cn.errors.count = 0 then
     call ofv.cerrarconsulta(rs)
     call ofv.cerrarconn(cn)
     set cn = ofv.conectar(ofv.strconn0)
     strsql = "insert into atrm (atributo,alias,tipo,comprobar,aplicacion) "
     strsql = strsql & " values('" & sid2 & "','" & salias & "','A',1,'" & sid & "') "    
     call ofv.crearconsultaEx(StrSql,cn,1,volver)
     if cn.errors.count = 0 then
        response.clear
        response.redirect volver 
     else
        response.write ofv.ErrorDB(cn,volver)
        response.flush
     end if 
  else
    response.write ofv.ErrorDB(cn,volver)
    response.flush
  end if 
end if






call ofv.cerrarconn(cn)

%>