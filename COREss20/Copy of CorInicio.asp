<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

  Opcion = Request.querystring("Opcion")





shtml = "<html><head><bgsound src=" & chr(34) & session("sound") &  chr(34) & "><style></style><title>COR ESolution Suite 2.8.7 - " & ofv.fechador() & " </title>"
'shtml = shtml & " cliemp=" & session("cliemp") & " clilogo=" & session("clilogo") & " cliente=" & session("cliente")
shtml = shtml & "<public>"
shtml = shtml & "<PROPERTY name='CorUsr' />"
shtml = shtml & "<PROPERTY name='CorNetUsr' />"
shtml = shtml & "<PROPERTY name='CorDomain' />"
shtml = shtml & "<PROPERTY name='CorLogid' />"
shtml = shtml & " </public>"

shtml = shtml & "<SCRIPT language=javascript >"
sHTML = sHTML  & " var CorUsr    ='" & REPLACE(Session("Usuario") & "*"," ","_") & "';  "
sHTML = sHTML  & " var CorNetUsr ='" & REPLACE(Session("UsuarNT") & "*"," ","_") & "';  "
sHTML = sHTML  & " var CorDomain ='" & REPLACE(Session("DOMINIO") & "*"," ","_") & "';  "
sHTML = sHTML  & " var CorLogid  ='" & REPLACE(SESSION("LOGIDUSRDET") & "*"," ","_") & "';  "
shtml = shtml & "</SCRIPT>"

shtml = shtml & "<SCRIPT FOR=window EVENT=onunload>"
''shtml = shtml & "    alert('La sesion de Cor E-Solution Suite se cerrara.' + window. event.clientX  + ' ' + window. event.clientY );"
sHTML = sHTML  & " if (window. event.clientX < 0) { "

   '      Session("Usuario") = UCase(Request.Form("Usuarnt"))
   '      Session("UsuarNT") = UCase(Request.Form("Usuarnt"))
   '      Session("DOMINIO") = UCase(Request.Form("DOMAINnt"))

sHTML = sHTML & " myWin= open(" & chr(34) & "CorCierra.ASP?USR=" & REPLACE(Session("Usuario") & "*"," ","_")
sHTML = sHTML & "&NETUSR=" & REPLACE(Session("UsuarNT") & "*"," ","_")
sHTML = sHTML & "&DMN=" & REPLACE(Session("DOMINIO") & "*"," ","_")
sHTML = sHTML & "&LID=" & REPLACE(SESSION("LOGIDUSRDET") & "*"," ","_")
sHTML = sHTML & chr(34) & "," & chr(34) & "Cons" & chr(34) & "," & chr(34)
sHTML = sHTML & "menubar=no,toolbar=no,height=300,width=400" & chr(34) & " ); "
sHTML = sHTML & " } "
shtml = shtml & "</SCRIPT>"
shtml = shtml & "</head>"



shtml = shtml & "<frameset rows='1,32,*' FRAMESPACING=0  FRAMEBORDER=0 BORDER=0 >"
   shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC='corbase.htm' NAME='Estad' SCROLLING=no>"
   shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC='Front.asp' NAME='Menus' SCROLLING=no>"
   shtml = shtml & "<frameset rows='*,29' FRAMESPACING=0  FRAMEBORDER=0 BORDER=0 >"
      shtml = shtml & "<frameset cols='14%,*' FRAMESPACING=0  FRAMEBORDER=0 BORDER=0 >"
         shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Master' SCROLLING=no src='Front.asp' >"
         shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='info' src='inicio.asp?opcion=" & opcion & "'  SCROLLING=auto>"
      shtml = shtml & "</frameset>"
      shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Footer' sRC='footer.asp?opcion=I' SCROLLING=no>"
  shtml = shtml & "</frameset>"
''shtml = shtml & "</frameset>"
shtml = shtml & "<noframes><body><p>Esta página usa marcos, pero su explorador no los admite.</p></body></noframes>"
shtml = shtml & "</frameset></html>"

response.write shtml

%>