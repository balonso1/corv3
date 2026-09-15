<!-- #INCLUDE FILE="Decvar.inc" -->
<%


function DevParametros(parametros)
dim  str

str=""
for i =0 to ubound(parametros)
 str=str & "&P" & cstr(i) & "=" & parametros(i)
next
DevParametros=str

end function

function ConvCarTodos (dato)
 
 if trim(cstr(Dato))=CarTodos then
   ConvCarTodos="%"
 else
  ConvCarTodos=Dato
 end if

end function

function VariableVacia(Dato,Parametro)

if len(trim(cstr(dato)))=0 then
    VariableVacia=parametro
else
    Variablevacia=Dato
end if

end function

function ConvertirCar(cadena,car1,car2)

for i= 1 to len(cadena)

 if mid(cadena,i,1)=car1 then
   convertirCar=Convertircar & Car2
 else
   convertirCar=ConvertirCar & mid(cadena,i,1) 
 end if
next

end function

sub convertirArregloParam(byref p ,car1,car2)

 for n =0 to ubound(p)
  p(n)=convertircar(p(n),car1,car2)
 next

end sub 

function explorador()

  bc = Request.ServerVariables("HTTP_USER_AGENT")
  if  Instr(bc,"MSIE") >0 THEN
   explorador="MSIE"
  else
   explorador="NETSCAPE" 
  end if

end function



%>