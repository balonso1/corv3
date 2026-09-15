<!-- #INCLUDE FILE="FunVarias.inc" -->
<%

function Combo (byref rs,StrNomObj,CampoValor,TextoMostrar,OtroCampo,varss)
dim str
Dim Dato
dim nulo
str=""
if not ConsultaVacia(Rs) Then

rutina=" onchange ='ss(" &chr(34) &varss & chr(34) & ")'"
 
 str= "<SELECT NAME=" & StrNomObj  & rutina  & "  style={font-size:8pt ;font-family:Times New Roman}>" & chr(13)
 str=str & "<OPTION VALUE=" & CarTodos &" SELECTED>Todos" & chr(13)

 rs.movefirst
 Dato=OtroCampo

 

 Do While Not rs.EOF 
   if len(otroCampo) > 0 then
     Dato=Rs(OtroCampo)
   end if
   str= str &   "<option value=" & chr(34) &ConvertirCar(rs(CampoValor)," ","-") & chr(34) & ">" & TextoMostrar & " " & Dato & chr(13)
   rs.movenext 
 loop  

 str=str & "</SELECT>" & chr(13) & chr(13)


else
rutina=" onchange ='ss(" &chr(34) &varss & chr(34) & ")'"
 
 str= "<SELECT NAME=" & StrNomObj  & rutina  & "  style={font-size:8pt ;font-family:Times New Roman}>" & chr(13)
 str=str & "<OPTION VALUE=" & CarTodos &" SELECTED>Todos" & chr(13)
 str=str & "</SELECT>" & chr(13) & chr(13)

End if


Combo=str
end function


function MostrarInmuebles(byref Rs,strnombreObj,CantRegMostrar,CantRegMover,titulo,Url,parametros)
dim str
dim f
dim i 
dim DirJpg
str= "<html><head><title>Seleccion de Rubros</title><style>a.link {text-decoration:none; color:#fdfdfd;} a {text-decoration:none; color:#fdfdfd;}"&_
"th {font-size:10pt ;font-family:garamond; border-style:outset; border-width:thin; color:#800000; background-color:#DEB887}"&_
"</style></head><body background='../../../repetidosjpg/back32.gif' style={font-size:8pt ;font-family:Times New Roman}>"
 
str= str & "<table HEIGHT=22 WIDTH=100% background='../../../repetidosjpg/back21.gif'><tr><td align=center ><font size=1 color=#ffffff face=verdana,arial><b>" & convertircar(titulo,"-"," ") & "</b></font></td></tr></table>"

if not ConsultaVacia(Rs) and cantregmover >=0 Then

 str= str & " <TABLE BORDER=0   RULES=ROWS  ALIGN=CENTER CELLPADDING=2 NAME=" & StrNombreObj & " STYLE=" & chr(34) & "position:absolute left:0; top:22" & chr(34) & " >" 

 str=str & chr(13) & " <TBODY style={font-size:10pt ;font-family:GARAMOND}>"


 rs.move(CantRegMover)
 str= str & chr(13) & "<TH  > Mas!!! </TH> "
 for f=1 to rs.fields.count -1
   str= str  & " <TH  >" & Rs(f).name & "</TH>"
  next
 
 i=0


 Do While ((Not rs.EOF) and ( i< CantRegMostrar))
   

  str=str & chr(13) & "<TR>" & chr(13)

  if isnull(rs(0)) then
   dirjpg=" <TD> &nbsp </TD>"
  else
   dirjpg="<td> <a href='../../aspcomun/vermas.asp?idrubro="& titulo &"&Id="& rs(0) &"&url=" & url & "&cantregmover=" & CantRegMover & "&cantregmostrar="& cantregmostrar & Devparametros(parametros) &"'> <img src='../../../RepetidosJpg/Trabul3a.gif'  BORDER=0 >  </td> </a>" 
  end if

 str=str & dirjpg
  
  for f=1 to rs.fields.count -1
   if f=2 then
    str= str &   "  <TD> " & Rs(f)   & "</TD>" 
   else
    str= str &   " <TD> " & Rs(f)  & "</TD>" 
   end if
  next
   
 str=str & "</TR>"
 rs.movenext 
 i=i+1 
 loop  

str=str & CHR(13) & "</Table>"


if CantRegMover <> 0  then
 if rs.eof then
   str=str & chr(13) & "<a href=buscar.asp?cantregmostrar=" & cstr(CantRegMostrar) & "&CantRegMover=" & cstr((cantregmover-cantregmostrar))& Devparametros(parametros) & "> <img border=0 height=17 width=17 src=../../../RepetidosJpg/bluprev2.gif>  </a>"
 else 
   str=str & chr(13) & "<a href=buscar.asp?cantregmostrar=" & cstr(CantRegMostrar) & "&CantRegMover=" & cstr((cantregmover-i)) & Devparametros(parametros) &"> <img border=0 height=17 width=17 src=../../../RepetidosJpg/bluprev2.gif>  </a>"
 end if
end if 


if not rs.eof then
  str=str & chr(13) & "<a href=buscar.asp?cantregmostrar=" & cstr(CantRegMostrar) & "&CantRegMover=" & cstr(cantregmover+i) & Devparametros(parametros) & "> <img border=0 height=17 width=17 src=../../../RepetidosJpg/blunext2.gif>  </a>"
end if 

str=str & "</body></html>"

MostrarInmuebles=str
else
 MostrarInmuebles="No Hay información que cumplan las condiciones dadas"
End if

end function

function VerMasInmuebles(byref Rs,titulo)
dim str
dim f
dim i 

titulo=convertircar(titulo,"-"," ")

str= "<table HEIGHT=22 WIDTH=100% background='../../repetidosjpg/back21.gif'><tr><td align=center ><font size=1 color=#ffffff face=verdana,arial><b>" & titulo & "</b></font></td></tr></table>"

if not ConsultaVacia(Rs)  Then

str=str & "<hr color=#990000 size=1 width='100%'>"
str = str & "<TABLE BORDER=0 > <TR><TD WIDTH=10% ></TD><TD WIDTH=91% >"

for f=0 to rs.fields.count -1
    str= str  & " <p style=" &chr(34) & " color:#00008B; text-align:left;font-size:10pt ;font-WEIGHT:500 ;font-family:VERDANA"& chr(34)&">"
    for i = f to f + 2
        if i < rs.fields.count  then
          if rs(i).name <> "E-Mail" then
            str= str  & Rs(i).name & ":  <b style=" &chr(34) & "color:#993333;font-size:10pt ;font-family:COMIC SANS MS"& chr(34)&">" & Rs(i) & "</b>  - " 
          else
             str= str  & Rs(i).name & ":  <b style=" &chr(34) & "color:#993333;font-size:10pt ;font-family:COMIC SANS MS"& chr(34)& " > <a href=mailto:" & Rs(i) &  ">" & Rs(i) & " </a> </b>  - " 
          end if    
        end if
    next
    f = i-1
    str= str  & " </p>"
next
 
str = str & " </TD></TR></TABLE>"

End if
VerMasInmuebles=str
end function

function VerFotos(Rs)
 i=0
 Do While ((Not rs.EOF) 
  Fotos(i)=rs(0)
  rs.movenext 
 loop  
 param= DevParametros(fotos)

 VerFotos="<a href=verfottos.asp?" &  mid(param,2,len(param)) & ">Fotos </a>"
 
end function
%>



 

