<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0



codapli = request.querystring("codapli")
coddepen = request.querystring("coddepen")

Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
Volver = ofv.convertircar(Volver,"*","_")

set cn = ofv.conectar(ofv.strconn1)
tabla = "Menux"

strsql = "Select * From " & Tabla 
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Update " & Tabla & " Set "
bEjecutar = False

scompo = request.form("scompo")
sstr = request.form("sstr")
sdato = request.form("sdato")
svalor = request.form("svalor")

sAsp = ""

if scompo <> "" then
   sAsp = scompo
   if sstr <> "" then 
      sAsp = sAsp & "?" & sstr
   end if
   if sdato <> "" then
      if sstr = "" then
         sAsp = sAsp & "?"
      else  
         sAsp = sAsp & "&"
      end if  
      sAsp = sAsp & "NDAT=" & sdato & "&NVAL=" & svalor
   end if
end if

if sAsp <> "" then
   bEjecutar = True
   strsql = strsql & " ASP = '" & sAsp & "' where menuatr = '" & coddepen & "'  "
   strsql = strsql & " and mcodapli = '" & codapli & "' "
end if



  if bEjecutar then
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
  end if

  if cn.errors.count = 0 then
    response.redirect volver
  else
    response.write ofv.ErrorDB(cn,volver)
  end if 


call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>