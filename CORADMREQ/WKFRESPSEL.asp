<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0


  WKFRESP = REQUEST.QUERYSTRING("WKFRESP")

  IF WKFRESP <> "" THEN  session("WKFRESP") = WKFRESP
     


sHTML = "<html>" & chr(13)
sHTML = sHTML & "<frameset rows=150,* FRAMESPACING=0 FRAMEBORDER=0 BORDER=0>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC=" & chr(34) & "WKFRESPSELHC.asp"
sHTML = sHTML & chr(34) & " NAME=TOPE SCROLLING=auto>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=BUSCA "
sHTML = sHTML & "src=" & chr(34) & "WKFRESPSELSC.asp"& chr(34) & " SCROLLING=auto></frameset>"
sHTML = sHTML & "<body><head><title>COR EManager - Edicion de Esquemas Normativos</title></head></body></html>"

response.write sHTML

%>