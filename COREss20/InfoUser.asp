<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  set cn2 = ofv.conectar(ofv.strconn0)
  Strsqlx = "Select * From NivelConfiden order by NCCOD"
  set rs = ofv.crearconsultaEx(StrSqlx,cn2,1,parametros)
  redim preserve CNIV(0)
  redim preserve DNIV(0)
  ii = 0
  if not rs.eof Then
     do until rs.eof
        redim preserve CNIV(ii)
        redim preserve DNIV(ii)
        CNIV(ii) = cstr(rs(0))
        DNIV(ii) = rs(1)
        ii = ii + 1
     rs.movenext
     loop

  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn2)



if session("USRES") <> "B" then

  shtml = ofv.MenuHeader(""," bacKground='imagenes/corf5.jpg' ")
else
  shtml = ofv.MenuHeader("","")
end if



''logoff
shtml = shtml & "<SCRIPT language=javascript >"
shtml = shtml & " function mostrar() {"
sHTML = sHTML  & " alert(window.top.CorUsr + ' ' + window.top.CorNetUsr  + ' ' + window.top.CorDomain + ' ' + window.top.CorLogid);  "
''sHTML = sHTML  & " alert(document.all.logoff.action);  "
shtml = shtml & "}"
shtml = shtml & "</SCRIPT>"


  sHTML = sHTML & "<br>"

  sStyle = "style={border-style:outset;border-width:thin;}"
  sImage = ofv.GenImage("","","","","imagenes/COREssuite20.jpg","0","")
  sInput = "&nbsp;<font face=verdana color=#000088 size=2 ><b>"
  sInput = sInput & "COR <font face=helvetica color=#f75500 size=4><i>e</i></font>Solution Suite 2.8</b></font>" 

  Celda = ofv.GenCelda("","","","5%","30","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","aaa " & sStyle,"center","9%","30","","",sImage,"si")
  Celda = Celda & ofv.GenCelda("","aaa valign=middle ","","","30","","",sInput,"si")




  sInput = ofv.GenerarInput("boton","Logon","submit","","20","bt","")
  sAccion = "inicio.asp?Opcion=ISET&ddat="

'  sAccion = sAccion & "&USR=" & REPLACE(Session("Usuario") & "*"," ","_")
'  sAccion = sAccion & "&NETUSR=" & REPLACE(Session("UsuarNT") & "*"," ","_")
'  sAccion = sAccion & "&DMN=" & REPLACE(Session("DOMINIO") & "*"," ","_")
'  sAccion = sAccion & "&LID=" & REPLACE(SESSION("LOGIDUSRDET") & "*"," ","_")

   Formulario = ofv.genform("logoff","",sAccion,"Master","post","",sInput,"si")
'  Formulario = ofv.genform("logoff","",sAccion,"","post","",sInput,"si")
  Celda = Celda & ofv.gencelda("","","center","1%","","","",Formulario,"si")

  Celda = Celda & ofv.gencelda("","","center","1%","","","","&nbsp;","si")


  sInput = ofv.GenerarInput("boton","Ayuda","submit","","20","bt","")
  sAccion = "../COREss20/Ayuda/NormasHelp.asp"
  Formulario = ofv.genform("ayudar","",sAccion,"info","post","",sInput,"si")
  Celda = Celda & ofv.gencelda("","","center","1%","","","",Formulario,"si")



  Celda = Celda & ofv.GenCelda("","","","10%","30","","","&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML & ofv.GenTabla("","","","100% border=0","","0","4","-1","-2","","",Fila,"si")

  Celda = ofv.GenCelda("","abc colspan=4","","100%","","","","&nbsp;","si")

  LineaVacia = ofv.GenRow("","","","","","","",Celda,"si")

  Fila = ""

  Fila = Fila & LineaVacia

  Fila = Fila & LineaVacia

  Fila = Fila & LineaVacia

  Celda = ofv.GenCelda("","","","2%","","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","cc colspan=2","","","","","","<img src='imagenes/corpnt2.jpg' border=0 align=middle >&nbsp;&nbsp;<b>USUARIO Y PROPIEDADES DE LA SESION</b>","si")
  Celda = Celda & ofv.GenCelda("","","","","2%","","","&nbsp;","si")
  Lineatit = ofv.GenRow("","","","","","","",Celda,"si")
  
  Fila =  Fila & Lineatit



  'Usuario
  Celda = ofv.GenCelda("","","","2%","","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","tt","","30%","","","","Usuario: ","si")
 '' Celda = Celda & ofv.GenCelda("","cc","","66%","","","","<a class=cc href='javascript:void(0);' onclick='mostrar();' >" & Ucase(Session("SegUsuario")) & "</a>","si")
 '' Celda = Celda & ofv.GenCelda("","cc","","66%","","","","<a class=cc href='inicio2.asp'  >" & Ucase(Session("SegUsuario")) & "</a>","si")
  Celda = Celda & ofv.GenCelda("","cc","","66%","","","",Ucase(Session("SegUsuario")),"si")
  Celda = Celda & ofv.GenCelda("","","","","2%","","","&nbsp;","si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



  'Codigo de Sesion
  Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","tt","","","","","","Codigo de Sesion: ","si")
  Celda = Celda & ofv.GenCelda("","cc","","","","","",Ucase(Session.SessionId),"si")
  Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  'Browser
  Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","tt","","","","","","Explorador: ","si")
  Celda = Celda & ofv.GenCelda("","cc","","","","","",Ucase(ofv.explorador) & "&nbsp;-&nbsp;" & Ucase(Request.ServerVariables("HTTP_USER_AGENT")),"si")
  Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



  'Fecha
  Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","tt","","","","","","Fecha del Sistema: ","si")
  Celda = Celda & ofv.GenCelda("","cc","","","","","",Ucase(ofv.fechador()),"si")
  Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



  Fila = Fila & LineaVacia



  sHTML = sHTML & "<center>" & ofv.GenTabla("","","","100%","","0","4","2","1","","",Fila,"si")

shtml = shtml & "<SCRIPT language=javascript >"
sHTML = sHTML  & " var xaction = document.all.logoff.action; "
sHTML = sHTML  & " document.all.logoff.action = xaction + window.top.CorUsr + '___' + window.top.CorNetUsr  + '___' + window.top.CorDomain + '___' + window.top.CorLogid;  "
shtml = shtml & "</SCRIPT>"

  sHTML = sHTML & "</center></BODY></html>"
end if

response.write shtml

%>