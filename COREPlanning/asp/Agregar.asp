<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

nCampo = request.querystring("Campo")
if nCampo = "" then nCampo = 0
nCampo = cdbl(nCampo)
Tabla = request.querystring("Tabla")
Volver = request.querystring("Volver")
' Lo tuve que invertir el 4/10 porque estaba mal pero antes funcionaba?
'Volver = ofv.convertircar(Volver,"&","_")
Volver = ofv.convertircar(Volver,"_","&")

Regisvuelta = cdbl(request.querystring("CantRegMover"))

set cn = ofv.conectar(ofv.strconn4)

strsql = "Select * From " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
i = 0
for each fld in rs.fields
  redim preserve vDesc(i)
  vDesc(i) = fld.name
  i = i + 1
next
sCampo = rs(nCampo).name

s = ofv.QueEsx(rs(nCampo).type)
sValor = s & UCase(request.form(1)) & s
if request.form(1) = "" and s <> "'" Then sValor = 0

strsql = "Insert into " & tabla & " ("
For Each x In Request.form
   strsql = strsql & x & "," 
Next
strsql = mid(strsql,1,len(strsql)-1)
strsql = strsql & ") Values ("

ValorNulo = False
TipoValido = True
For Each x In Request.form
  vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
  vNuevoValor = replace(vNuevoValor,"'","")
  vNuevoValor = replace(vNuevoValor,chr(34),"")
  nx = 0
  for i = 0 to UBound(vDesc)
    if x = vDesc(i) then
      nx = i
      exit for
    end if
  next
  s = ofv.QueEsx(rs(nx).type)
  ValorNulo = ofv.EsNulo(rs(nx),vNuevoValor)
  if ValorNulo then exit for
  TipoValido = ofv.ValidarTipo(rs(nx),vNuevoValor)
  if not TipoValido then exit for
  strsql = strsql & s & vNuevoValor & s & ","
Next

if not TipoValido or ValorNulo then
  on error resume next
  response.write ofv.ErrorTipo(volver,Request.Form(x),rs(nx).name)
  if err.number <> 0 then response.write ofv.ErrorTipo(volver,"","")
else
  strsql = mid(strsql,1,len(strsql)-1) & ")"
  call ofv.crearconsultaEx(StrSql,cn,1,volver)
  
  set errores = cn.errors
  if cn.errors.count = 0 then
    response.redirect volver 
  else
    response.write ofv.ErrorDB(cn,volver)
  end if 
end if

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>