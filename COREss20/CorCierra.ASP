<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
%>
<html>

<head>

<style>
</style>
<title>COR ESolution Suite V3</title>
</head>
<body bgcolor=#f9f9f9 style='border-style:solid;border-width:2;border-color:#336699;' >
<center>
<br>
<table width=80% style='filter:shadow(color=#dddddd,direction=125);' >
<tr>
<td align=center>
<img src=imagenes/COREssuite20.jpg border=0 >
</td>
</tr>
<tr>
<td align=center>

<b><font face=tahoma color=#6699cc size=5>
COR <font face=helvetica color=#f75500 size=6>
e</font>Solution Suite V3</font></b>
</td>
</tr>

</table>
<br>
<table bgcolor=#ffffff width=80% style='border-style:ridge;border-width:2;border-color:#6699cc;' cellpadding=2 >


<tr>
<td align=center>

<font face=tahoma color=#336699 size=4>
<u>Cierre de la Sesion</u></font>
</td>
</tr>
<%

  IF NOT (Session("Usuario") & "") <> "" THEN

         Session("Usuario")     = REPLACE(REPLACE(request.querystring("USR"),"_"," "),"*","")
         Session("UsuarNT")     = REPLACE(REPLACE(request.querystring("NETUSR"),"_"," "),"*","")
         Session("DOMINIO")     = REPLACE(REPLACE(request.querystring("DMN"),"_"," "),"*","")
         SESSION("LOGIDUSRDET") = REPLACE(REPLACE(request.querystring("LID"),"_"," "),"*","")

  END IF
response.write "<tr><td align=center style='filter:shadow(color=#dddddd,direction=125);' >"
response.write "<table style='border-style:solid;border-width:2;border-color:#dddddd;' cellpadding=2 >"

response.write "<tr><td ><font face=verdana color=#000088 size=2>Usuario:&nbsp;</font>"
response.write "</td>"
response.write "<td ><font face=verdana color=#aaaaaa size=1>" & Session("Usuario") & "</font>"
response.write "</td></tr>"

response.write "<tr><td ><font face=verdana color=#000088 size=2>Usuario de red:&nbsp;</font>"
response.write "</td>"
response.write "<td ><font face=verdana color=#aaaaaa size=1>" & Session("UsuarNT") & "</font>"
response.write "</td></tr>"

response.write "<tr><td ><font face=verdana color=#000088 size=2>Dominio:&nbsp;</font>"
response.write "</td>"
response.write "<td ><font face=verdana color=#aaaaaa size=1>" & Session("DOMINIO") & "</font>"
response.write "</td></tr>"

response.write "<tr><td ><font face=verdana color=#000088 size=2>Fecha&nbsp;y&nbsp;hora:&nbsp;</font>"
response.write "</td>"

xdia =  datepart("d",date()) & "-" & datepart("m",date()) & "-" & datepart("yyyy",date()) & "&nbsp;&nbsp;" & time()

''xdia = formatdatetime(date(),2)

response.write "<td ><font face=verdana color=#aaaaaa size=1>" & xdia & "</font>"
response.write "</td></tr>"

response.write "<tr><td ><font face=verdana color=#000088 size=2>Logid:&nbsp;</font>"
response.write "</td>"
response.write "<td ><font face=verdana color=#aaaaaa size=1>" & SESSION("LOGIDUSRDET") & "</font>"
response.write "</td></tr>"

response.write "</table>"
response.write "<br>"
response.write "</td></tr>"

  set cn2 = ofv.conectar(ofv.strconn2)

  XXXX = GETUSRLOGUPD("LOGOFF")

  call ofv.cerrarconn(cn2)

%>

</table>




</body>


</html>