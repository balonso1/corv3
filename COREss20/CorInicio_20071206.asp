<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

  Opcion = Request.querystring("Opcion")


shtml = "<html><head><bgsound src=" & chr(34) & session("sound") &  chr(34) & " >"
shtml = shtml & "<style></style><title>COR ESolution Suite V3 - " & ofv.fechador() & " </title>" & chr(13)
shtml = shtml & "<public>" & chr(13)
shtml = shtml & "<PROPERTY NAME='CorType' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorUsr' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorNetUsr' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorDomain' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorLogid' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMMnu' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMtopm' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMtopf' />" & chr(13)

shtml = shtml & "<PROPERTY name='CorUsrN' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorUsrE' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorEmpL' />" & chr(13)
'''shtml = shtml & "<PROPERTY name='CorLogid' />" & chr(13)
shtml = shtml & "<PROPERTY name='xfwin' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMcol1' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMcol2' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMcol3' />" & chr(13)

shtml = shtml & "<method name='verfrm2' />" & chr(13)
shtml = shtml & "<method name='verfrm1' />" & chr(13)
shtml = shtml & "<method name='verfrm0' />" & chr(13)
shtml = shtml & "<method name='verfrmU' />" & chr(13)
shtml = shtml & "<method name='ldfobj' />" & chr(13)
shtml = shtml & "<method name='verestilo' />" & chr(13)
shtml = shtml & "<method name='verfuncion' />" & chr(13)

shtml = shtml & "<method name='xopcom' />" & chr(13)
shtml = shtml & "<method name='xclcom' />" & chr(13)
shtml = shtml & "<method name='xclcomCD' />" & chr(13)
shtml = shtml & "<method name='xclcomVD' />" & chr(13)
shtml = shtml & "<method name='xclcomS' />" & chr(13)
shtml = shtml & "<method name='xclrstr' />" & chr(13)

shtml = shtml & " </public>" & chr(13)

shtml = shtml & "<SCRIPT language=javascript >" & chr(13)
sHTML = sHTML & " window.name   = 'superm';" & chr(13)
sHTML = sHTML & " var CorType   = 'M';" & chr(13)
sHTML = sHTML & " var CorMMnu   = 'S';" & chr(13)
sHTML = sHTML & " CorMtopm      = window;" & chr(13)
sHTML = sHTML & " CorMtopf      = window;" & chr(13)

sHTML = sHTML & " var CorUsr    = '" & REPLACE(Session("Usuario") & "*"," ","_") & "';  " & chr(13)
sHTML = sHTML & " var CorNetUsr = '" & REPLACE(Session("UsuarNT") & "*"," ","_") & "';  " & chr(13)
sHTML = sHTML & " var CorDomain = '" & REPLACE(Session("DOMINIO") & "*"," ","_") & "';  " & chr(13)
sHTML = sHTML & " var CorLogid  = '" & REPLACE(SESSION("LOGIDUSRDET") & "*"," ","_") & "';  " & chr(13)
sHTML = sHTML & " var CorUsrN  = '" & Session("SegUsuario") & "';  " & chr(13)
sHTML = sHTML & " var CorUsrE  = '" & session("cliemp") & "';  " & chr(13)
sHTML = sHTML & " var CorEmpL  = '" & session("clilogo") & "';  " & chr(13)
sHTML = sHTML & " var xfwin    = window;" & chr(13)
sHTML = sHTML & " var CorMcol1   = '" & Session("APLSHW") & "';" & chr(13)
sHTML = sHTML & " var CorMcol2   = '" & Session("CMFMENU") & "';" & chr(13)
sHTML = sHTML & " var CorMcol3   = '" & Session("CMFPNBOT") & "';" & chr(13)
shtml = shtml & "</SCRIPT>" & chr(13)



shtml = shtml & "<SCRIPT language=javascript >" & chr(13)

XFSCRP = "" & chr(13)
XFSCRP = XFSCRP & " var xfcomv  =  0;" & chr(13)
XFSCRP = XFSCRP & " var xfnomc  = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xfnoms  = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xffrm   = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xffrmot = 'N';" & chr(13)
'''''XFSCRP = XFSCRP & " var xfwin   = window;" & chr(13)
XFSCRP = XFSCRP & " var xfdivc  = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xftdc   = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xfnfld;" & chr(13)
'''XFSCRP = XFSCRP & "    alert(xfcomv + ' ' + xfwin);"
XFSCRP = XFSCRP & " " & chr(13)
XFSCRP = XFSCRP & "function xopcom(xnomc,xnoms,xfrm,xfrmot) {" & chr(13)
XFSCRP = XFSCRP & "    "
''XFSCRP = XFSCRP & "    xfwin    = document.frames(0);" & chr(13)

'''XFSCRP = XFSCRP & "    alert(xfcomv + ' ' + xfwin + ' 2');"

XFSCRP = XFSCRP & "    if (xfcomv==1) {xclcomS();}"


XFSCRP = XFSCRP & "    xfcomv   =   1;"
XFSCRP = XFSCRP & "    xfnomc   = xnomc;"
XFSCRP = XFSCRP & "    xfnoms   = xnoms;"
XFSCRP = XFSCRP & "    xffrm    = xfrm;"
XFSCRP = XFSCRP & "    xffrmot  = xfrmot;"
XFSCRP = XFSCRP & "    xfdivc   = 'DF@' + xfnoms;"
XFSCRP = XFSCRP & "    xftdc    = 'TD@' + xfnoms;"
XFSCRP = XFSCRP & " "

XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    if (xffrmot=='1') {xfnfld=xfwin.document;}"
XFSCRP = XFSCRP & "    else {xfnfld=xfwin.document.all[xffrm];}"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    "


XFSCRP = XFSCRP & "    xfwin.document.all[xnoms].value=xfnfld.all[xnomc].value;"
XFSCRP = XFSCRP & "    xfwin.document.all[xfdivc].style.display='block';"

XFSCRP = XFSCRP & ""
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcom() {" & chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    "
''''XFSCRP = XFSCRP & "    xfnfld2 = xfwin.document.all['uuu'].value;"
XFSCRP = XFSCRP & "    if (xffrmot=='1') {xfnfld=xfwin.document;}"
XFSCRP = XFSCRP & "    else {xfnfld=xfwin.document.all[xffrm];}"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1') {xclcomCD();}"
XFSCRP = XFSCRP & "    else {if (xffrmot=='2') "
XFSCRP = XFSCRP & "             {xfnfld.all[xfnomc].value=xfwin.document.all[xfnoms].value;"
XFSCRP = XFSCRP & "              xfwin.document.all[xffrm].submit();}"
XFSCRP = XFSCRP & "          else {xclcomVD();}}"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcomCD() {" & chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    "
''''XFSCRP = XFSCRP & "    xfnfld2 = xfwin.document.all['uuu'].value;"

XFSCRP = XFSCRP & "    if (xffrmot=='1') {xfnfld=xfwin.document;}"
XFSCRP = XFSCRP & "    else {xfnfld=xfwin.document.all[xffrm];}"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "     xfnfld.all[xfnomc].value=xfwin.document.all[xfnoms].value;"
XFSCRP = XFSCRP & "     xfnfld.all[xftdc].innerHTML="
XFSCRP = XFSCRP & "     xfwin.document.all[xfnoms].options[xfwin.document.all[xfnoms].selectedIndex].text;"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    xclcomS();"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcomVD() {" & chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1') {xfnfld=xfwin.document;}"
XFSCRP = XFSCRP & "    else {xfnfld=xfwin.document.all[xffrm];}"
XFSCRP = XFSCRP & "    "
''''XFSCRP = XFSCRP & "    xfnfld2 = xfwin.document.all['uuu'].value;"
''''XFSCRP = XFSCRP & "    alert(xfwin.document.all[xfnoms].options.length + ' ' + xfwin.document.all[xfnoms].selectedIndex);"
XFSCRP = XFSCRP & "    if (xfwin.document.all[xfnoms].options[xfwin.document.all[xfnoms].selectedIndex].value=='nsnc') "
XFSCRP = XFSCRP & "       { "
XFSCRP = XFSCRP & "         alert('La seleccion es incorrecta ...');"
XFSCRP = XFSCRP & "         xfwin.document.all[xfnoms].value=xfnfld.all[xfnomc].value;"
XFSCRP = XFSCRP & " "
''XFSCRP = XFSCRP & "         for (i = 0; i < xfwin.document.all[xfnoms].options.length; i++) "
''XFSCRP = XFSCRP & "             {if (xfwin.document.all[xfnoms].options[i].defaultSelected== true) "
''XFSCRP = XFSCRP & "                 {xfwin.document.all[xfnoms].options[i].selected=true;} "
''XFSCRP = XFSCRP & "             } "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "       } "
XFSCRP = XFSCRP & "    else "
XFSCRP = XFSCRP & "       {xfnfld.all[xfnomc].value=xfwin.document.all[xfnoms].value;"
XFSCRP = XFSCRP & "        xfwin.document.all[xffrm].submit();}"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcomS() {" & chr(13)
XFSCRP = XFSCRP & " "
'''XFSCRP = XFSCRP & "    alert(xfcomv + ' ' + xfwin + ' 2' );"
XFSCRP = XFSCRP & "    xfwin.document.all[xfdivc].style.display='none';"
XFSCRP = XFSCRP & "    xfcomv     =   0;"
XFSCRP = XFSCRP & "    xfnomc     = 'N';"
XFSCRP = XFSCRP & "    xfnoms     = 'N';"
XFSCRP = XFSCRP & "    xffrm      = 'N';"
XFSCRP = XFSCRP & "    xffrmot    = 'N';"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

''XFSCRP = XFSCRP & "</script>" & chr(13) & chr(13)

shtml = shtml & XFSCRP

 Whtml = ""

''''''''''

shtml = shtml & "function getIFD(xelID){  "
shtml = shtml & " var xxrv = null;  "
''shtml = shtml & " // if contentDocument exists, W3C compliant (Mozilla)  "
shtml = shtml & " if (document.getElementById(xelID).contentDocument){  "
shtml = shtml & "   xxrv = document.getElementById(xelID).contentDocument;  "
shtml = shtml & " } else {  "
''shtml = shtml & "   // IE  "
shtml = shtml & "   xxrv = document.frames[xelID].document;  "
shtml = shtml & " }  "
shtml = shtml & " return xxrv;  "
shtml = shtml & "  }  " & chr(13)



'''''''''''


shtml = shtml & " function ldfobj() { "
''shtml = shtml & "   document.frames(0).document.getElementById('DX4').innerHTML = document.frames(5).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('info'); "
shtml = shtml & " var XFRD2 = getIFD('fesquemas'); "

shtml = shtml & "   XFRD1.getElementById('DX4').innerHTML = XFRD2.body.innerHTML; "

shtml = shtml & " } " & chr(13)

shtml = shtml & " function verfrmU() { "
shtml = shtml & "  var MDUSR = '" & Whtml & "'; "
shtml = shtml & "  return  MDUSR; "
shtml = shtml & " } " & chr(13)

 Whtml = ""

shtml = shtml & " function verfrm0() { "
''shtml = shtml & "  return  document.frames(1).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('Footer'); "

shtml = shtml & "  return XFRD1.body.innerHTML; "


shtml = shtml & " } " & chr(13)

shtml = shtml & " function verfrm1() { "
''shtml = shtml & "  return  document.frames(3).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('Menus'); "

shtml = shtml & "  return XFRD1.body.innerHTML; "


shtml = shtml & " } " & chr(13)

shtml = shtml & " function verfrm2() { "
''shtml = shtml & "  return  document.frames(4).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('Master'); "

shtml = shtml & "  return XFRD1.body.innerHTML; "

shtml = shtml & " } " & chr(13)


shtml = shtml & " function verestilo() { "
''shtml = shtml & "  return  document.frames(7).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('estilo'); "

shtml = shtml & "  return XFRD1.getElementById('estil').innerHTML; "

shtml = shtml & " } " & chr(13)

shtml = shtml & " function verfuncion() { "
''shtml = shtml & "  return  document.frames(7).document.body.innerHTML; "

shtml = shtml & " return getIFD('estilo'); "

shtml = shtml & " } " & chr(13)



shtml = shtml & " function xclrstr() { "
'''shtml = shtml & "  xfwin    = document.frames(0);" & chr(13)

shtml = shtml & " if (document.getElementById('info').contentDocument){  "
shtml = shtml & "   xfwin = document.getElementById('info');  "
shtml = shtml & " } else {  "
shtml = shtml & "   xfwin = document.frames['info'];  "
shtml = shtml & " }  "

shtml = shtml & " " & chr(13)
shtml = shtml & "  xfcomv  =  0;" & chr(13)
shtml = shtml & "  xfnomc  = 'N';" & chr(13)
shtml = shtml & "  xfnoms  = 'N';" & chr(13)
shtml = shtml & "  xffrm   = 'N';" & chr(13)
shtml = shtml & "  xffrmot = 'N';" & chr(13)
shtml = shtml & "  xfdivc  = 'N';" & chr(13)
shtml = shtml & "  xftdc   = 'N';" & chr(13)
shtml = shtml & " } " & chr(13)


shtml = shtml & "</SCRIPT>"

shtml = shtml & "<SCRIPT FOR=window EVENT=onunload>"
''shtml = shtml & "    alert('La sesion de Cor E-Solution Suite se cerrara.' + window. event.clientX  + ' ' + window. event.clientY );"
sHTML = sHTML  & " if (window. event.clientX < 0) { "


sHTML = sHTML & " myWin= open(" & chr(34) & "CorCierra.ASP?USR=" & REPLACE(Session("Usuario") & "*"," ","_")
sHTML = sHTML & "&NETUSR=" & REPLACE(Session("UsuarNT") & "*"," ","_")
sHTML = sHTML & "&DMN=" & REPLACE(Session("DOMINIO") & "*"," ","_")
sHTML = sHTML & "&LID=" & REPLACE(SESSION("LOGIDUSRDET") & "*"," ","_")
sHTML = sHTML & chr(34) & "," & chr(34) & "Cons" & chr(34) & "," & chr(34)
sHTML = sHTML & "menubar=no,toolbar=no,height=300,width=400" & chr(34) & " ); "
sHTML = sHTML & " } "
shtml = shtml & "</SCRIPT>"
shtml = shtml & "</head>"

 Novedform = "CLIMN16317206150930885546227"
 Session("CorApli") = "COREM20"

 OFV.ObtenerAtributosNOdo Novedform, "", "TIPO", CSTR("10")
 bConsultar = OFV.AConsultar
    session("SMPLMENU") = "N"
 IF bConsultar THEN
    session("SMPLMENU") = "N"
 else
    session("SMPLMENU") = "S"
 end if

 IF OPCION="LC" THEN
    session("SMPLMENU")="S"
 END IF
 Session("CorApli") = ""

  IF OPCION="LC" THEN
   XVAR=REQUEST.QUERYSTRING
   XVARN=INSTR(XVAR,"&")
   XVAR=RIGHT(XVAR,LEN(XVAR)-XVARN)
  END IF
'response.write xvar
shtml = shtml & "<frameset rows='0,*,0,0,0,0,0,0' FRAMESPACING=0  FRAMEBORDER=0 BORDER=0 >"

 shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='estilo'    ID='estilo'    sRC='estilo.asp'    SCROLLING=no STYLE='display:none;' >"

  
  if opcion="LC" THEN
   shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='info'      ID='info'      SRC='corbase.htm' SCROLLING=auto >"
  ELSE
   shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='info'      ID='info'      SRC='inicio.asp?opcion=" & opcion & "' SCROLLING=auto >"
  END IF

 shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Footer'    ID='Footer'    SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
 shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Estad'     ID='Estad'     SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
 shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Menus'     ID='Menus'     SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
IF OPCION="LC" THEN
   shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Master'    ID='Master'    SRC='http://ws001z5002.blp.com.ar/desarrollo/dvara/Intranet/Formularios.php?" & XVAR &"'  SCROLLING=no STYLE='display:none;' >"
 ' shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Master'    ID='Master'    SRC='http://ws001z5002.blp.com.ar/desarrollo/dvara/Intranet/Formularios.php?form=Cartas_Documento&separador=1&idSeccion=2&nodo=5'  SCROLLING=no STYLE='display:none;' >"
ELSE
 shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Master'    ID='Master'    SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
END IF
 shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='fesquemas' ID='fesquemas' SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
 shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='expir'     ID='expir'     sRC='refrescar.asp' SCROLLING=no STYLE='display:none;' >"



shtml = shtml & "<noframes><body><p>Esta página usa marcos, pero su explorador no los admite.</p></body></noframes>"
shtml = shtml & "</frameset></html>"






response.write shtml

%>