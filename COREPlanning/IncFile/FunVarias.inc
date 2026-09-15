<!-- #INCLUDE FILE="Decvar.inc" -->

<%
function DescCampo(Gcampo)

dim IDCol(13)
dim DescCol(13)
IDCol(0) = "descri"
DescCol(0) = "Descripcion"
IDCol(1) = "tipoperfil"
DescCol(1) = "Perfil"
IDCol(2) = "codinter"
DescCol(2) = "Codigo interno"
IDCol(3) = "ecomaster"
DescCol(3) = "Dependencia"
IDCol(4) = "codperfil"
DescCol(4) = "Perfil"
IDCol(5) = "sisgen"
DescCol(5) = "Descripcion"
IDCol(6) = "tipogen"
DescCol(6) = "Descripcion"
IDCol(8) = "tipobjeto"
DescCol(8) = "Tipo"
IDCol(9) = "tipomodo"
DescCol(9) = "Tipo"
IDCol(10) = "inda"
DescCol(10) = "Orden"
IDCol(11) = "codasoc"
DescCol(11) = "Asociacion"
IDCol(12) = "secuen"
DescCol(12) = "Secuencia"
DescCampo = Gcampo

for yn=0 to ubound(IDCol)
if Gcampo = IDCol(yn) then
   DescCampo = DescCol(yn)
   exit for
end if
next

end function

function GenerarInput(Gnombre,Gvalor,Gtipo,Gform,Gsize,Gclase,Gnumb)

wbrow = explorador()
if wbrow = "MSIE" then
   wfont = 7
   wspacing = 1
   wsize = 1
   wresto = 0
else
   wfont = 8
   wspacing = 0
   wsize = 2
   wresto = 3
end if

    xHTML =  "<input " 
   if Ucase(Gtipo) = Ucase("readonly") or Ucase(Gtipo) = Ucase("agregar") or Ucase(Gtipo) = Ucase("textsb") then
    xGtipo = "TEXT" 
   else
    xGtipo = Ucase(Gtipo)
   end if
    if Ucase(Gtipo) = Ucase("readonly") then
    if wbrow = "MSIE" then
      xtag =  " readonly "
    else
      xtag =  " " 
    end if
   else
     xtag =  " " 
    end if

    xHTML = xHTML & " type=" & xGtipo & xtag

    if Ucase(Gtipo) = Ucase("text") then
    xHTML = xHTML & ValidarActualizacion(Gnumb,Gform,Gnombre)
    else
     if Ucase(Gtipo) = Ucase("textsb") then
       xHTML = xHTML & " onblur=" & chr(34) & "document." & Gform & ".submit(); return true;" & chr(34) & " "
     end if
    end if


    xHTML = xHTML & " size=" & (Gsize / wsize) - wresto & "% "
    xHTML = xHTML & " value='" & Gvalor & "'"
    if Gtipo <> "readonly" then xHTML = xHTML & " name=" & Gnombre
    xHTML = xHTML & " class=" & Gclase
    xHTML = xHTML & " >"

GenerarInput = xHTML

end function



function ValidarActualizacion(bNumber,Gform,Gcampo)

' Si bnumber = SIR valida numerico y repetido
' Si bnumber = SI valida numerico
' Si dnumber = NOR valida repetido
' Si bnumber = "" solo valida dato nulo igual que las anteriores 

  Gnombre = DescCampo(Gcampo)

  s = s & "onblur=" & chr(34) & "if (this.value.length == 0) "
  s = s & "{alert('" & Gnombre & " no debe ser nulo'); this.focus(); "
  s = s & "this.value=this.defaultValue;this.select(); return false;} "
  s = s & "else { if (this.value != this.defaultValue) {" 
  if bNumber = "SI" or bnumber = "SIR" then
    s = s & "if (isNaN(this.value) || this.value < 0) {alert('" & Gnombre & " debe ser numerico');"
    s = s & " this.focus();this.value=this.defaultValue;this.select();" 
    s = s & " return false;} else {"
  end if  
  if bNumber = "SI" or bnumber = "" then
    s = s & " document." & Gform & ".submit(); return true;} "
  end if
  if bnumber = "SIR" or bnumber = "NOR" then
    s = s & " valida(this);if (sale==0) {document." & Gform & ".submit();} } "
  end if 
 
  if bNumber = "SI" or bnumber = "SIR" then
    s = s & "}"
  end if
  s = s & " else {return false;} };" & chr(34)


  ValidarActualizacion = s

end function

function ValidarActualizacionCombo(Gform,Gcampo)

  Gnombre = DescCampo(Gcampo)

  s = s & " onchange=" & chr(34) & "if (this.options[this.selectedIndex].value=='nsnc') { "
  s = s & "alert('La seleccion es incorrecta para " & Gnombre & "');"
  s = s & "for (i = 0; i < this.options.length; i++) {if (this.options[i].defaultSelected== true){"
  s = s & "this.options[i].selected=true;}};"

  s = s & " return false;} else {"
  s = s & Gform & ".submit(); return true;};" & chr(34)

  ValidarActualizacionCombo = s

end function

function GenerarCombo(ASPx ,NombreCampox ,ValorCampo,SQL,Conexion,Indice,Descripcion,ID) 
GenerarCombo = GenerarComboEss("1",ASPx,NombreCampox,ValorCampo,"","","","","","","",SQL,Conexion,Indice,Descripcion,ID)  
end function

function GenerarComboSP(ASP,NombreCampo,ValorCampo,SQL,Conexion,Indice,Descripcion,ID) 
GenerarComboSP = GenerarComboEss("2",ASP,NombreCampo,ValorCampo,"","","","","","","",SQL,Conexion,Indice,Descripcion,ID)  
end function

function GenerarLista(SQL,Conexion,Campo)
GenerarLista = GenerarComboEss("3","",Campo,"","","","","","","","",SQL,Conexion,"","","")  
end function

function GenerarCombox(SQL,Conexion,Campo,Nombre)
GenerarCombox = GenerarComboEss("4","",Nombre,"",Campo,"","","","","","",SQL,Conexion,"","","")  
end function

function GenerarCombox2(SQL,Conexion,Campo1,Campo2,Nombre)
GenerarCombox2 = GenerarComboEss("5","",Nombre,"",Campo1,"","","","","","",SQL,Conexion,"",Campo2,"")  
end function

function GenerarCombox3(SQL,Conexion,Campo1,Campo2,Nombre)
GenerarCombox3 = GenerarComboEss("6","",Nombre,"",Campo1,"","","","","","",SQL,Conexion,"",Campo2,"")  
end function

function GenerarCombo2(NombreCampo,ValorCampo,SQL,Conexion,Indice,Descripcion,ID)  
GenerarCombo2 = GenerarComboEss("7","",NombreCampo,ValorCampo,"","","","","","","",SQL,Conexion,Indice,Descripcion,ID)  
end function

function GenerarCombo3(ASP,Array,NombreCampo,ValorCampo)  
GenerarCombo3 = GenerarComboEss("8",ASP,NombreCampo,ValorCampo,"","","","",Array,"","","","","","","")  
end function

function GenerarCombo333(Array,NombreCampo,ValorCampo,direc,imagen)
GenerarCombo333 = GenerarComboEss("9","",NombreCampo,ValorCampo,"","",direc,imagen,Array,"","","","","","","")  
end function

function GenerarCombo33(Array,NombreCampo,ValorCampo)  
GenerarCombo33 = GenerarComboEss("a","",NombreCampo,ValorCampo,"","","","",Array,"","","","","","","")  
end function

function comboSiNO(form,valor,nombre)
comboSiNO = GenerarComboEss("b",form,nombre,valor,"","","","","","","","","","","","")  
end function

function comboSiNoEx(form,valor,nombre,s,n)
comboSiNoEx = GenerarComboEss("c",form,nombre,valor,"","","","","",s,n,"","","","","")  
end function

function comboSiNOmin(nombre)
comboSiNOmin = GenerarComboEss("d","",nombre,"","","","","","","","","","","","","")  
end function


function GenerarComboEss(Btipo,Bform,BNCampo,BVCampo,BNCampo2,BVCampo2,Bdir,Bima,BArray,Bs,Bn,BSQL,BCon,BInd,BDesc,BOpc)  


  dim sID
  dim bfnom 
  dim bfval
  dim bfval2

  bfnom = BNCampo 
  bfnom2 = BNCampo2 
  bfval = BVCampo
  bfval2 = BVCampo2
   

  if isnull(BVCampo) then bfval = ""
  if isnull(BVCampo2) then bfval2 = ""

'function GenerarCombo(ASPx ,NombreCampox ,ValorCampo,SQL,Conexion,Indice,Descripcion,ID) 
'function GenerarComboSP(ASP,NombreCampo,ValorCampo,SQL,Conexion,Indice,Descripcion,ID) 
'function GenerarLista(SQL,Conexion,Campo)
'function GenerarCombox(SQL,Conexion,Campo,Nombre)
'function GenerarCombox2(SQL,Conexion,Campo1,Campo2,Nombre)
'function GenerarCombox3(SQL,Conexion,Campo1,Campo2,Nombre)
'function GenerarCombo2(NombreCampo,ValorCampo,SQL,Conexion,Indice,Descripcion,ID)  
'function GenerarCombo3(ASP,Array,NombreCampo,ValorCampo)  
'function GenerarCombo333(Array,NombreCampo,ValorCampo,direc,imagen)
'function GenerarCombo33(Array,NombreCampo,ValorCampo)  
'function comboSiNO(form,valor,nombre)
'function comboSiNoEx(form,valor,nombre,s,n)
'function comboSiNOmin(nombre)

select case Btipo

case "1"



  Array1 = GenerarArray(BSQL,BCon,BInd)
  Array2 = GenerarArray(BSQL,BCon,BDesc)

  if BOpc = "no" or BOpc = "no0" then
    sID = ""
  else
  if BOpc = "sb" or BOpc = "sb0" then
    sID = " onchange=" & chr(34)  & Bform & ".submit(); return true;" & chr(34) & " "
  else
    sID = ValidarActualizacionCombo(Bform,bfnom)
  end if
  end if
  if BOpc = "no0" or BOpc = "sb0" then
     scx = "<option value=0>&nbsp;</option>"
  else 
     scx = "<option value=nsnc>&nbsp;</option>"
  end if

  For nx = 1 To Ubound(Array1)

    if Array1(nx) <> "" then

      If CStr(Array1(nx)) = CStr(bfval) Then
        scx = scx & "<option value=" & convertircar(Array1(nx)," ","_") & " selected>" & Array2(nx) & "</option>"
      Else
        scx = scx & "<option value=" & convertircar(Array1(nx)," ","_") & ">" & Array2(nx) & "</option>"
      End If
    end if
  Next




case "2"

  Array1 = GenerarArray(BSQL,BCon,BInd)
  Array2 = GenerarArray(BSQL,BCon,BDesc)

  sID = ValidarActualizacionCombo(Bform,bfnom)

  scx = "<option value=nsnc>&nbsp;</option>"

  For n = 1 To Ubound(Array1)
    if Array1(n) <> "" then
      If CStr(Array1(n)) = CStr(bfval) Then
        scx = scx & "<option value=" & convertircar(Array1(n)," ","_") & " selected>" & Array2(n) & "</option>"
      Else
        scx = scx & "<option value=" & convertircar(Array1(n)," ","_") & ">" & Array2(n) & "</option>"
      End If
    end if
  Next

case "3"

  Arrayx = GenerarArray(BSQL,BCon,bfnom)

  sID = " align=center "

  bfnom = "lista"

  if ubound(arrayx) > 0 then
     For n = 1 To Ubound(Arrayx)
         scx = "<option value=" & convertircar(Arrayx(n)," ","_") & ">" & Arrayx(n) & "</option>"
     Next
  end if


case "4"

  Arrayx = GenerarArray(BSQL,BCon,bfnom2)

  sID = ""

  For n = 1 To Ubound(Arrayx)
    scx = "<option value=" & chr(34) & convertircar(Arrayx(n)," ","_")
    scx = scx & chr(34) & ">" & Arrayx(n) & "</option>"
  Next

case "5"

  Arrayx1 = GenerarArray(SQL,Conexion,bfnom2)
  Arrayx2 = GenerarArray(SQL,Conexion,Bdesc)

  sID = ""

  scx = " "

  For n = 1 To Ubound(Arrayx1)
    scx = scx & "<option value=" & chr(34) & convertircar(Arrayx1(n)," ","_")
    scx = scx & chr(34) & ">" & Arrayx2(n) & "</option>"
  Next

case "6"

  Arrayx1 = GenerarArray(SQL,Conexion,bfnom2)
  Arrayx2 = GenerarArray(SQL,Conexion,Bdesc)

  sID = ""

  scx = "<option value=nsnc>&nbsp;</option>"

  For n = 1 To Ubound(Arrayx1)
    scx = scx & "<option value=" & chr(34) & convertircar(Arrayx1(n)," ","_")
    scx = scx & chr(34) & ">" & Arrayx2(n) & "</option>"
  Next

case "7"

  Array1 = GenerarArray(BSQL,BCon,BInd)
  Array2 = GenerarArray(BSQL,BCon,BDesc)

  sID = ""

  scx = "<option value=nsnc>&nbsp;</option>"

  For n = 1 To Ubound(Array1)
    if Array1(n) <> "" then
      If CStr(Array1(n)) = CStr(bfval) Then
        scx = scx & "<option value=" & Array1(n) & " selected>" & Array2(n) & "</option>"
      Else
        scx = scx & "<option value=" & Array1(n) & ">" & Array2(n) & "</option>"
      End If
    end if
  Next

case "8"

  sID = ValidarActualizacionCombo(Bform,bfnom)

  scx = "<option value=nsnc>&nbsp;</option>"

  For n = 1 To Ubound(BArray)
    if BArray(n) <> "" then
      If UCase(CStr(BArray(n))) = UCase(CStr(bfval)) Then
        scx = scx & "<option value=" & BArray(n) & " selected>" & BArray(n) & "</option>"
      Else
        scx = scx & "<option value=" & BArray(n) & ">" & BArray(n) & "</option>"
      End If
    end if
  Next

case "9"

  sID = "document." & Bima & ".src='Iconos/' + this.value ;"

  scx = "<option value=nsnc>&nbsp;</option>"

  For n = 1 To Ubound(BArray)
    if BArray(n) <> "" then
      If UCase(CStr(BArray(n))) = UCase(CStr(bfval)) Then
        scx = scx & "<option value=" & BArray(n) & " selected>" & BArray(n) & "</option>"
      Else
        scx = scx & "<option value=" & BArray(n) & ">" & BArray(n) & "</option>"
      End If
    end if
  Next

case "a"

  sID = ""

  scx = "<option value=nsnc>&nbsp;</option>"

  For n = 1 To Ubound(BArray)
    if BArray(n) <> "" then
      If UCase(CStr(BArray(n))) = UCase(CStr(bfval)) Then
        scx = scx & "<option value=" & BArray(n) & " selected>" & BArray(n) & "</option>"
      Else
        scx = scx & "<option value=" & BArray(n) & ">" & BArray(n) & "</option>"
      End If
    end if
  Next

case "b"

 sID = " onchange=document." & Bform & ".submit() "

if bfval then
   scx = " <option selected value=true>Si</option>"
   scx = scx & "<option value=false>No</option>"
else
   scx = " <option  value=true>Si</option>"
   scx = scx & " <option selected value=false>No</option>"
end if

case "c"

 sID = " onchange=document." & Bform & ".submit() "

if bfval then
   scx = " <option selected value=" & Bs &">Si</option>"
   scx = scx & " <option value="& Bn &">No</option>"
else
   scx = " <option  value=" & Bs & ">Si</option>"
   scx = scx & " <option selected value=" &  Bn & ">No</option>"
end if

case "d"

 sID = ""
 scx =  "<option value=True>Si</option>" & chr(13)
 scx = scx & "<option value=False>No</option>" & chr(13)


end select

  sx = " <font class=at><select " &  sID & " name=" & bfnom & "> "
  sx = sx & scx
  sx = sx & "</select></font>"

  GenerarComboEss = sx


end function

function GenerarArray(SQL,Conexion,Campo)

  set rsArray=crearconsultaEx(SQL,Conexion,1,parametros)

  n = 1
  redim s(0)
  do until rsArray.eof
    redim preserve s(n)
    s(n) = rsArray(campo)
    rsArray.movenext
    n = n + 1
  loop

  call cerrarconsulta(rsArray)

  GenerarArray = s

end function

function vueltaasp(mdir,m1,m2)
dim nn
nn=strreverse(request.servervariables("SCRIPT_NAME"))
nn=strreverse(left(nn,instr(nn,"/")-1)) 
nn= mdir & nn & "?cantregmostrar=" & cstr(m1) & "_CantRegMover=" & cstr(m2)
vueltaasp=nn
end function

function vueltaaspEx(mdir,m1,m2)
dim nn
nn=strreverse(request.servervariables("SCRIPT_NAME"))
nn=strreverse(left(nn,instr(nn,"/")-1)) 
nn= mdir & nn & "?CantRegAMostrar=" & cstr(m1) & "_CantRegAMover=" & cstr(m2)
vueltaaspEx=nn
end function

function vueltaaspEx2(mdir,m1,m2)
dim nn
nn=strreverse(request.servervariables("SCRIPT_NAME"))
nn=strreverse(left(nn,instr(nn,"/")-1)) 
nn= mdir & nn & "?CantRegAMostrar2=" & cstr(m1) & "_CantRegAMover2=" & cstr(m2)
vueltaaspEx2=nn
end function

function vueltaasp2(mdir)
dim nn
nn=strreverse(request.servervariables("SCRIPT_NAME"))
nn=strreverse(left(nn,instr(nn,"/")-1)) 
nn= mdir & nn
vueltaasp2=nn
end function

function botonera(rs,cantregmover,cantregmostrar,posi)
dim l

l=strreverse(request.servervariables("SCRIPT_NAME"))
l=strreverse(left(l,instr(l,"/")-1))

if CantRegMoverx <> 0  then
 if rs.eof then
   strx=strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-cantregmostrarx)) & "> <img border=0 height=17 width=17 title='Anterior' src=Imagenes/bluprev2.gif>  </a>"
 else 
   strx=strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-posi))  &"> <img border=0 height=17 width=17 title='Anterior' src=Imagenes/bluprev2.gif>  </a>"
 end if
end if 


if not rs.eof then
  strx=strx & chr(13) & "<a href="& l & "?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr(cantregmoverx+posi)  & "> <img border=0 height=17 width=17 title='Siguiente' src=Imagenes/blunext2.gif>  </a>"
end if 
botonera=strx
end function

' Botonera con boton Volver

function botoneraBV(rs,cantregmover,cantregmostrar,posi,Volver)
dim l

l=strreverse(request.servervariables("SCRIPT_NAME"))
l=strreverse(left(l,instr(l,"/")-1))

strx = "<table id=botonera width=50% marginwidth=0  cellspacing=1 cellpadding=0><tr>" & chr(13)

if CantRegMoverx <> 0  then
 if rs.eof then
   strx = strx & chr(13) & "<td  height=10 width=2% >"
   strx = strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-cantregmostrarx)) & "> <img border=0 height=17 width=17 title='Anterior' src=Imagenes/bluprev2.gif>  </a>"
   strx = strx & "</td>" & chr(13) 
else 
   strx = strx & chr(13) & "<td  height=10 width=2% >"
   strx = strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-posi))  &"> <img border=0 height=17 width=17 title='Anterior' src=Imagenes/bluprev2.gif>  </a>"
   strx = strx & "</td>" & chr(13)  
 end if
end if 


if not rs.eof then
  strx = strx & chr(13) & "<td  height=10 width=2% align=left >"
  strx=strx & chr(13) & "<a href="& l & "?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr(cantregmoverx+posi)  & "> <img border=0 height=17 width=17 title='Siguiente' src=Imagenes/blunext2.gif>  </a>"
  strx = strx & "</td>" & chr(13) 
end if 

'   strx = strx & chr(13) & "<td  height=10 width=2% >"
   strx = strx & Volver
   strx = strx & "</tr></table>" & chr(13)

botoneraBV=strx
end function

' Botonera con boton cancelar y Agregar
function botoneraAC(rs,cantregmover,cantregmostrar,posi,Agregar,Volver)
dim l
l=strreverse(request.servervariables("SCRIPT_NAME"))
l=strreverse(left(l,instr(l,"/")-1))

strx = "<table class=at id=botonera width=2% marginwidth=0  cellspacing=0 cellpadding=0><tr  >" & chr(13)

if CantRegMoverx <> 0  then
 if rs.eof then
   strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-cantregmostrarx)) & "> <img border=0 title=Registros&nbspAnteriores height=17 width=17 src=Imagenes/bluprev2.gif>  </a>"
   strx = strx & "</td>" & chr(13) 
else 
   strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-posi))  &"> <img border=0 title=Registros&nbspAnteriores height=17 width=17 src=Imagenes/bluprev2.gif>  </a>"
   strx = strx & "</td>" & chr(13)  
 end if
end if 


if not rs.eof then
  strx = strx & chr(13) & "<td  height=10 width=2% align=left >"
  strx=strx & chr(13) & "<a href="& l & "?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr(cantregmoverx+posi)  & "> <img border=0 title=Registros&nbspSiguientes height=17 width=17 src=Imagenes/blunext2.gif>  </a>"
  strx = strx & "</td>" & chr(13) 
end if 
pepe=0
if agregar <> "no" then
 '  strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & Agregar
 '  strx = strx & "</td>" & chr(13)
end if
if Volver <> "no" then
'   strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & Volver
'   strx = strx & "</td>" & chr(13)
end if
   strx = strx & "</tr></table>" & chr(13)

botoneraAC=strx
end function

function botoneraACB(rs,cantregmover,cantregmostrar,posi,Agregar,Volver,Buscar)
dim l
l=strreverse(request.servervariables("SCRIPT_NAME"))
l=strreverse(left(l,instr(l,"/")-1))

strx = "<table class=at id=botonera width=2% marginwidth=0  cellspacing=0 cellpadding=0><tr  >" & chr(13)

if CantRegMoverx <> 0  then
 if rs.eof then
   strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-cantregmostrarx))& "&kk=" & convertircar(now," ","-") & rnd & "> <img border=0 title=Registros&nbspAnteriores height=17 width=17 src=Imagenes/bluprev2.gif>  </a>"
   strx = strx & "</td>" & chr(13) 
else 
   strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-posi)) & "&kk=" & convertircar(now," ","-") & rnd &"> <img border=0 title=Registros&nbspAnteriores height=17 width=17 src=Imagenes/bluprev2.gif>  </a>"
   strx = strx & "</td>" & chr(13)  
 end if
botonbuscar = "si"
end if 


if not rs.eof then
  strx = strx & chr(13) & "<td  height=10 width=2% align=left >"
  strx=strx & chr(13) & "<a href="& l & "?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr(cantregmoverx+posi) & "&kk=" & convertircar(now," ","-") & rnd & "> <img border=0 title=Registros&nbspSiguientes height=17 width=17 src=Imagenes/blunext2.gif>  </a>"
  strx = strx & "</td>" & chr(13) 
botonbuscar = "si"
end if 

if agregar <> "no" then
 '  strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & Agregar
 '  strx = strx & "</td>" & chr(13)
end if
if Volver <> "no" then
'   strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & Volver
'   strx = strx & "</td>" & chr(13)
end if


if  botonbuscar = "si" then
'   strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & Buscar
'   strx = strx & "</td>" & chr(13)
end if

   strx = strx & "</tr></table>" & chr(13)

botoneraACB=strx
end function

function botoneraXXX(rs,cantregmover,cantregmostrar,posi,Agregar,Volver,Buscar,critico)
dim l

l=strreverse(request.servervariables("SCRIPT_NAME"))
l=strreverse(left(l,instr(l,"/")-1))

strx = "<table class=at id=botonera width=2% marginwidth=0  cellspacing=0 cellpadding=0><tr  >" & chr(13)

if CantRegMoverx <> 0  then
 if rs.eof then
   strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-cantregmostrarx)) & "> <img border=0 title=Registros&nbsp;Anteriores height=17 width=17 src=Imagenes/bluprev2.gif>  </a>"
   strx = strx & "</td>" & chr(13) 
else 
   strx = strx & chr(13) & "<td  height=10 width=2% align=left>"
   strx = strx & chr(13) & "<a href="& l &"?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr((cantregmoverx-posi))  &"> <img border=0 title=Registros&nbsp;Anteriores height=17 width=17 src=Imagenes/bluprev2.gif>  </a>"
   strx = strx & "</td>" & chr(13)  
 end if
botonbuscar = "si"
end if 


if not rs.eof then
  strx = strx & chr(13) & "<td  height=10 width=2% align=left >"
  strx=strx & chr(13) & "<a href="& l & "?cantregmostrar=" & cstr(CantRegMostrarx) & "&CantRegMover=" & cstr(cantregmoverx+posi)  & "> <img border=0 title=Registros&nbsp;Siguientes height=17 width=17 src=Imagenes/blunext2.gif>  </a>"
  strx = strx & "</td>" & chr(13) 
botonbuscar = "si"
end if 

if agregar <> "no" then
   strx = strx & Agregar
end if
if Volver <> "no" then
   strx = strx & Volver
end if


if  botonbuscar = "si" then
   strx = strx & Buscar
end if

if  critico <> "no" then
   strx = strx & critico
end if

   strx = strx & "</tr></table>" & chr(13)

botoneraXXX=strx
end function


function BotoneraEx(rs,CantRegAMover,CantRegAMostrar,Codigo,Tipo)

  l=strreverse(request.servervariables("SCRIPT_NAME"))
  l=strreverse(left(l,instr(l,"/")-1))

  if CantRegAMover <> 0  then
    if rs.eof then
      strx=strx & chr(13) & "<a href="& l & "?CantRegAMover="
      strx=strx & cstr(CantRegAMover-CantRegAMostrar) & "&CantRegAMostrar="
      strx=strx & CantRegAMostrar & "&Codobjeto=" & Codigo & "&Tipobjeto=" & Tipo & ">"
      strx=strx & "<img border=0 height=17 width=17 title='Anterior' src=Imagenes/bluprev2.gif></a>"
    else 
      strx=strx & chr(13) & "<a href="& l & "?CantRegAMover="
      strx=strx & cstr(CantRegAMover-CantRegAMostrar) & "&CantRegAMostrar="
      strx=strx & CantRegAMostrar & "&Codobjeto=" & Codigo & "&Tipobjeto=" & Tipo & ">"
      strx=strx & "<img border=0 height=17 width=17 title='Anterior' src=Imagenes/bluprev2.gif></a>"
    end if
  end if 

  if not rs.eof then
    strx=strx & chr(13) & "<a href="& l & "?CantRegAMover="
    strx=strx & cstr(CantRegAMover+CantRegAMostrar) & "&CantRegAMostrar="
    strx=strx & CantRegAMostrar & "&Codobjeto=" & Codigo & "&Tipobjeto=" & Tipo & ">"
    strx=strx & "<img border=0 height=17 width=17 title='Siguiente' src=Imagenes/blunext2.gif></a>"
  end if 

  BotoneraEx=strx

end function

function BotoneraEx2(rs2,CantRegAMover2,CantRegAMostrar2,Codigo,Tipo)

  l=strreverse(request.servervariables("SCRIPT_NAME"))
  l=strreverse(left(l,instr(l,"/")-1))

  if CantRegAMover2 <> 0  then
    if rs.eof then
      strx=strx & chr(13) & "<a href="& l & "?CantRegAMover2="
      strx=strx & cstr(CantRegAMover2-CantRegAMostrar2) & "&CantRegAMostrar2="
      strx=strx & CantRegAMostrar2 & "&Codobjeto=" & Codigo & "&Tipobjeto=" & Tipo & ">"
      strx=strx & "<img border=0 height=17 width=17 title='Anterior' src=Imagenes/bluprev2.gif></a>"
    else 
      strx=strx & chr(13) & "<a href="& l & "?CantRegAMover2="
      strx=strx & cstr(CantRegAMover2-CantRegAMostrar2) & "&CantRegAMostrar2="
      strx=strx & CantRegAMostrar2 & "&Codobjeto=" & Codigo & "&Tipobjeto=" & Tipo & ">"
      strx=strx & "<img border=0 height=17 width=17 title='Anterior' src=Imagenes/bluprev2.gif></a>"
    end if
  end if 

  if not rs.eof then
    strx=strx & chr(13) & "<a href="& l & "?CantRegAMover2="
    strx=strx & cstr(CantRegAMover2+CantRegAMostrar2) & "&CantRegAMostrar2="
    strx=strx & CantRegAMostrar2 & "&Codobjeto=" & Codigo & "&Tipobjeto=" & Tipo & ">"
    strx=strx & "<img border=0 height=17 width=17 title='Siguiente' src=Imagenes/blunext2.gif></a>"
  end if 

  BotoneraEx2=strx

end function

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

function CheckUsuario()

 if Session("Usuario") = "" Then
  sHTML = "<html><Title>COR EPlanning</title><style>"
  sHTML = sHTML & "a{font-family:tahoma;color:#990000;font-size:7pt;text-decoration:none;}"
  sHTML = sHTML & "a.link{font-family:tahoma;color:#990000;font-size:7pt;text-decoration:none;}"
  sHTML = sHTML & "input{font-family:tahoma;color:#000077;font-size:7pt;background-color:#f7eedd;}"
  sHTML = sHTML & "input.bt{font-family:tahoma;color:#eeeeee;font-size:7pt;background-color:#336699;}"
  sHTML = sHTML & "select{font-family:tahoma;color:#000077;font-size:7pt;background-color:#f7eedd;}"
  sHTML = sHTML & "tr.st{background-color:#336699;}"
  sHTML = sHTML & "td.st{font-family:verdana;color:#dddddd;font-size:6pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"
  sHTML = sHTML & "td.mt{font-family:tahoma;color:#000077;font-size:7pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:thin;background-color:#eeeeee;}"
  sHTML = sHTML & "td.tt{font-family:tahoma;color:#dddddd;font-size:7pt;text-transform:uppercase;font-weight:600;}"
  sHTML = sHTML & "td.tx{font-family:arial narrow;color:#000077;font-size:7pt;text-transform:uppercase;border-style:outset;border-width:thin;background-color:#aabbdd;}"
  sHTML = sHTML & "td.cc{font-family:tahoma;color:#000077;font-size:7pt;background-color:#f7eedd;border-style:ridge;border-width:thin;text-indent:10;}"
  sHTML = sHTML & "td.ut{font-family:tahoma;color:#000077;font-size:7pt;font-weight:600;}"
  sHTML = sHTML & "td.lt{font-family:verdana;color:#550000;font-size:8pt;font-weight:600;text-indent:10;}"
  sHTML = sHTML & "</style>" & chr(13)
  sHTML = sHTML & "<body bgcolor=#f7eedd >" & chr(13)
  sHTML = sHTML & "<table  border=0 width=50% marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top><td align=left width=4% >"
  sHTML = sHTML & "<img src=Imagenes/Corlogox1.jpg border=0 width=22 height=22></td>"
  sHTML = sHTML & "<td class=ut align=left width=55% >COR EPlanning</td></tr></table>" &chr(13)
  sHTML = sHTML & "<br><table  border=0 width=100% marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top><td class=lt align=left width=60% >Logon Check</td>" &chr(13)
  fecha = fechador()
  sHTML = sHTML & "<td class=ut align=right >" & fecha & "</td></tr></table><hr color=#000088 size=1 width=100% align=left >"
  sHTML = sHTML & "<table  border=0 width=100% marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top><td align=left class=cc width=4% >La Aplicacion no esta disponible sin Logon,"
  sHTML = sHTML & " el login fue invalido o expiro por time-out</td></tr>"
sHTML = sHTML & "<tr align=top><td align=center  width=4%  ><br><br>"
sHTML = sHTML & "<form name=logoff action=inicio.asp?Opcion=I target=Master >"
sHTML = sHTML & "<input class=bt type=submit name=boton value=Logon size=20 >"
sHTML = sHTML & "</form></td></tr>"
  sHTML = sHTML & "</table></body></html>"
else
    sHTML = "Y"
end if

CheckUsuario=sHTML
end function

function fechador()
fec = Date()
ndia = cdia(WeekDay(fec))
nmes = cmes(Month(fec))
nowstr = "  " & ndia & ", " &  day(fec) &  " de " & nmes &  " de " & Year(fec)

fechador=nowstr
end function

function cmes(mes)

xmes = Array ("nulo","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")

cmes= xmes(mes)
end function

function cdia(dia)
dim xdia
xdia = Array ("nulo","Domingo","Lunes","Martes","Miercoles","Jueves","Viernes","Sabado")
cdia= xdia(dia)
end function

function FormStyler()
dim sHTML 
dim wbrow
dim wfont
dim wspacing


wbrow = explorador()
if wbrow = "MSIE" then
   wfont = 7
   wspacing = 1
else
   wfont = 8
   wspacing = 0
end if

sHTML = sHTML & "<meta HTTP-EQUIV=" & chr(34) & "expires" & chr(34) & " CONTENT=" & chr(34) & "0" & chr(34) & " >"
sHTML = sHTML & "<style>"

sHTML = sHTML & "a.link{color:#ffffff;text-decoration:none;}"
sHTML = sHTML & "a{font-size:" & wfont & "pt;font-family:tahoma; color:#ffffff; text-transform: capitalize; "
sHTML = sHTML & "font-weight:500;text-decoration:none;}"
sHTML = sHTML & "a.nn {font-size:" & wfont & "pt;font-family:tahoma; color:#ffffff; text-transform: capitalize; "
sHTML = sHTML & "font-weight:500;text-decoration:none;}"
shtml = shtml & "input.at{font-family:tahoma;color:#000055;font-size:" & wfont & "pt;background-color:#ffffff;}" & chr(13)
sHTML = sHTML & "input{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
sHTML = sHTML & "input.dt{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
sHTML = sHTML & "input.bt{font-family:tahoma;color:#eeeeee;font-size:" & wfont & "pt;background-color:#336699;}"
sHTML = sHTML & "input.it{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#cccccc;}"
sHTML = sHTML & "select{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
sHTML = sHTML & "select.at{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
sHTML = sHTML & "span.at{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
sHTML = sHTML & "span.mt{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1pt;background-color:#eeeeee;width:80%;}"
sHTML = sHTML & "font.at{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;}"
sHTML = sHTML & "option{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
sHTML = sHTML & "tr.st{background-color:#336699;}"
sHTML = sHTML & "table.at{background-color:#999999;}"
sHTML = sHTML & "td.aa{text-transform: capitalize;}"
shtml = shtml & "td.ee{font-size:" & wfont & "pt;font-family:tahoma;  border-width:1pt; border-color:#4466bb;" & chr(13)
shtml = shtml & "color:#770000;font-weight:500;}" & chr(13)
sHTML = sHTML & "td.at{font-family:tahoma;color:#eeeeee;font-size:" & wfont & "pt;font-weight:500;}"
sHTML = sHTML & "td.rr{font-size:" & wfont & "pt;font-family:tahoma;  border-width:1pt; border-color:#4466bb;"
sHTML = sHTML & "color:#ffffff;font-weight:500;}"
sHTML = sHTML & "td.st{font-family:verdana;color:#dddddd;font-size:" & wfont & "pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"
sHTML = sHTML & "td.mt{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1pt;background-color:#eeeeee;}"
sHTML = sHTML & "td.tt{font-family:tahoma;color:#dddddd;font-size:" & wfont & "pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"
sHTML = sHTML & "td.tx{font-family:arial narrow;color:#000077;font-size:" & wfont & "pt;text-transform:uppercase;border-style:outset;border-width:1pt;background-color:#aabbdd;}"
sHTML = sHTML & "td.cc{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;border-style:ridge;border-width:1pt;}"
sHTML = sHTML & "td.ut{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;font-weight:500;}"
sHTML = sHTML & "td.lt{font-family:verdana;color:#550000;font-size:8pt;font-weight:600;text-indent:10;}"
sHTML = sHTML & "td.vt{font-family:tahoma;color:#000055;font-size:" & wfont & "pt;font-weight:600;}"
sHTML = sHTML & "td.wt{font-family:tahoma;color:#550000;font-size:" & wfont & "pt;font-weight:600;}"
sHTML = sHTML & "td.qt{font-family:tahoma;color:#550000;font-size:8pt;font-weight:600;}"
sHTML = sHTML & "div {page-break-after:always;}"
sHTML = sHTML & "</style>"

FormStyler = sHTML

end function

function MenuHeader(scolor,simagen)


sHTML = "<html><head><basefont size=1 color=#000077 face=tahoma>" & FormStyler() & "</head>"

sHTML = sHTML & "<body BGCOLOR=" & scolor & " alink=#ffffff vlink=#ffffff link=#ffffff " & simagen & "  >"

MenuHeader = sHTML

end function



function RepHeader(sTitulo)
  sHTML = "<html><Title>COR EPlanning - " & sTitulo & "</title><style>"
  sHTML = sHTML & "a{font-family:tahoma;color:#990000;font-size:7pt;text-decoration:none;}"
  sHTML = sHTML & "a.link{font-family:tahoma;color:#990000;font-size:7pt;text-decoration:none;}"
  sHTML = sHTML & "input{font-family:tahoma;color:#000077;font-size:7pt;background-color:#f7eedd;}"
  sHTML = sHTML & "input.bt{font-family:tahoma;color:#eeeeee;font-size:7pt;background-color:#336699;}"
  sHTML = sHTML & "select{font-family:tahoma;color:#000077;font-size:7pt;background-color:#f7eedd;}"
  sHTML = sHTML & "tr.st{background-color:#336699;}"
  sHTML = sHTML & "td.st{font-family:verdana;color:#dddddd;font-size:6pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"
  sHTML = sHTML & "td.mt{font-family:tahoma;color:#000077;font-size:7pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:thin;background-color:#eeeeee;}"
  sHTML = sHTML & "td.tt{font-family:tahoma;color:#dddddd;font-size:7pt;text-transform:uppercase;font-weight:600;}"
  sHTML = sHTML & "td.tx{font-family:arial narrow;color:#000077;font-size:7pt;text-transform:uppercase;border-style:outset;border-width:thin;background-color:#aabbdd;}"
  sHTML = sHTML & "td.cc{font-family:tahoma;color:#000077;font-size:7pt;background-color:#f7eedd;border-style:ridge;border-width:thin;}"
  sHTML = sHTML & "td.ut{font-family:tahoma;color:#000055;font-size:7pt;font-weight:500;}"
  sHTML = sHTML & "td.vt{font-family:tahoma;color:#000055;font-size:7pt;font-weight:600;}"
  sHTML = sHTML & "td.lt{font-family:tahoma;color:#000055;font-size:8pt;font-weight:600;}"
  sHTML = sHTML & "</style>" & chr(13)
  sHTML = sHTML & "<body bgcolor=#ffffff >" & chr(13)
  sHTML = sHTML & "<table  border=0 width=600 marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top><td align=left width=4% >"
  sHTML = sHTML & "<img src=" & Session("CliLogo")  & " border=0 width=22 height=22></td>"
  sHTML = sHTML & "<td class=ut align=left width=55% >" & Session("Cliente")  & "</td></tr></table>" &chr(13)
  sHTML = sHTML & "<br><table  border=0 width=600 marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top><td class=lt align=left width=60% >" & sTitulo & "</td>" &chr(13)
  fecha = fechador()
  sHTML = sHTML & "<td class=ut align=right >" & fecha & "</td></tr></table><hr color=#000088 size=1 width=600 align=left >"

  RepHeader = sHTML
end function

function Repbody(sTitulo)
  sHTML = "<html><Title>COR EPlanning - " & sTitulo & "</title>"
  sHTML = sHTML & FormStyler()
  sHTML = sHTML & "<body bgcolor=#ffffff >" & chr(13)
  Repbody = sHTML
end function

function RepHeader2(sTitulo)

  sHTML = "<table  border=0 width=600 marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top><td align=left width=4% >"
  sHTML = sHTML & "<img src=" & Session("CliLogo")  & " border=0 width=22 height=22></td>"
  sHTML = sHTML & "<td class=ut align=left width=55% >" & Session("Cliente")  & "</td></tr></table>" &chr(13)
  sHTML = sHTML & "<br><table  border=0 width=600 marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top><td class=qt align=left width=60% >" & sTitulo & "</td>" &chr(13)
  fecha = fechador()
  sHTML = sHTML & "<td class=ut align=right >" & fecha & "</td></tr></table><hr color=#000088 size=1 width=600 align=left >"

  RepHeader2 = sHTML
end function

function RepHeaderL(sTitulo)

  sHTML = "<table  border=0 width=900 marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top><td align=left width=4% >"
  sHTML = sHTML & "<img src=" & Session("CliLogo")  & " border=0 width=22 height=22></td>"
  sHTML = sHTML & "<td class=ut align=left width=55% >" & Session("Cliente")  & "</td></tr></table>" &chr(13)
  sHTML = sHTML & "<br><table  border=0 width=900 marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top><td class=qt align=left width=60% >" & sTitulo & "</td>" &chr(13)
  fecha = fechador()
  sHTML = sHTML & "<td class=ut align=right >" & fecha & "</td></tr></table><hr color=#000088 size=1 width=900 align=left >"

  RepHeaderL = sHTML
end function


function FormHeader(sTitulo)

dim sHTML 
dim wbrow
dim wfont
dim wspacing


wbrow = explorador()
if wbrow = "MSIE" then
  wbody = "bgcolor=#cccccc"
else
  wbody = "bgcolor=#ccbbaa"
end if

  sHTML = "<html><Title>Normas y Procedimientos</title>" & FormStyler()
  sHTML = sHTML & "<body " & wbody & " >" & chr(13)
  sHTML = sHTML & "<table align=center border=0 width=100% marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top>"
  sHTML = sHTML & "<td class=mt align=center width=35% >"
  sHTML = sHTML & sTitulo 
  sHTML = sHTML & "</td>" &chr(13)
  sHTML = sHTML & "<td class=ut align=right >Usuario: " & Session("SegUsuario") & "</td></tr></table>"
if wbrow = "MSIE" then
  sHTML = sHTML & "<hr color=#000088 size=1>"
else
  sHTML = sHTML & ""
end if

  FormHeader = sHTML
end function

function FormHeaderx(sTitulo)

dim sHTML 
dim wbrow
dim wfont
dim wspacing


wbrow = explorador()
if wbrow = "MSIE" then
  wbody = "bgcolor=#cccccc"
else
  wbody = "bgcolor=#ccbbaa"
end if

  sHTML = "<html><Title>Normas y Procedimientos</title>" & FormStyler()
  sHTML = sHTML & "<body " & wbody & " onload=alert('Aca');  >" & chr(13)
  sHTML = sHTML & "<table align=center border=0 width=100% marginwidth=1 cellspacing=1 cellpadding=0>"
  sHTML = sHTML & "<tr align=top>"
  sHTML = sHTML & "<td class=mt align=center width=35% >"
  sHTML = sHTML & sTitulo 
  sHTML = sHTML & "</td>" &chr(13)
  sHTML = sHTML & "<td class=ut align=right >Usuario: " & Session("SegUsuario") & "</td></tr></table>"
if wbrow = "MSIE" then
  sHTML = sHTML & "<hr color=#000088 size=1>"
else
  sHTML = sHTML & ""
end if

  FormHeaderx = sHTML
end function

function comboSiNO(form,valor,nombre)

if valor then
str="<select class=at onchange=document." & form & ".submit() name=" & nombre & "> <option selected value=true>Si</option>" &_
     "<option value=false>No</option></select>"
else
str="<select class=at onchange=document." & form & ".submit() name=" &nombre &"> <option  value=true>Si</option>" &_
    " <option selected value=false>No</option></select>"
end if

combosino= "<font class=at> " & str & " </font>"
end function

function comboSiNoEx(form,valor,nombre,s,n)

if valor then
str="<select class=at onchange=document." & form & ".submit() name=" & nombre & "> <option selected value=" & s &">Si</option>" &_
     "<option value="& n &">No</option></select>"
else
str="<select class=at onchange=document." & form & ".submit() name=" &nombre &"> <option  value=" & s & ">Si</option>" &_
    " <option selected value=" &  n & ">No</option></select>"
end if

combosinoEx="<font class=at> " & str & " </font>"
end function

function queES(valor)
 select case vartype(valor)
   case 8
     st="'"
  case else
     st=""
 end select

quees=st

end function

function queESX(valor)
 select case valor
   case 8,135,202
     st="'"
   case else
     st=""
 end select

queesx=st

end function

function ValidarTipo(campo,valor)

  ValidarTipo = True

  tipo = vartype(campo)

  select case tipo
    case 2,3,4,5
      if not isnumeric(valor) then ValidarTipo = False
    case 7
      if not isdate(valor) then ValidarTipo = False
    case 11
      if not isnumeric(valor) then ValidarTipo = False
      if valor <> "True" And valor <> "False" then ValidarTipo = False
  end select
 
end function

function GenerarComboe(ASP,Campo,Array)  

  sHTML = sHTML & "<select onchange=" & ASP & ".submit() name=" & Campo.name & ">"
  sHTML = sHTML & "<option value=nsnc>&nbsp"
  For n = 0 To Ubound(Array)
    If Array(n) = Campo Then
      sHTML = sHTML & "<option value=" & Array(n) & " selected>" & Array(n)
    Else
      sHTML = sHTML & "<option value=" & Array(n) & ">" & Array(n)
    End If
  Next
  sHTML = sHTML & "</select><td></td>"

  GenerarComboe = sHTML

end function

function GenerarCombo3(ASP,Array,NombreCampo,ValorCampo)  

  if isnull(ValorCampo) then ValorCampo = ""
  s = s & "<select name=" & NombreCampo & " onchange= " & ASP & ".submit()>"
  s = s & "<option value=nsnc>&nbsp"
  For n = 1 To Ubound(Array)
    if Array(n) <> "" then
      If UCase(CStr(Array(n))) = UCase(CStr(ValorCampo)) Then
        s = s & "<option value=" & Array(n) & " selected>" & Array(n)
      Else
        s = s & "<option value=" & Array(n) & ">" & Array(n)
      End If
    end if
  Next
  s = s & "</select>"

  GenerarCombo3 = s

end function

function GetFiles(Directorio)

  Dim fso,Folder,Files,F, s, n

  Set fso = CreateObject("Scripting.FileSystemObject")
  sPath = server.mappath("1")
  sPath = left(sPath,len(sPath)-1) & Directorio 

  Set Folder = fso.getfolder(sPath)
  Set Files = Folder.Files
  
  redim s(0)
  n = 1
  For Each F in Files
    redim preserve s(n)
    s(n) = F.name
    n = n + 1
  Next
  GetFiles = s

end function

function ObtenerNOM(Vector,VectorD,rs)  
  for i=0 to Ubound(Vector)
    if rs = Vector(i) then 
	ObtenerNOM = VectorD(i)
	Exit for
    End If
  Next
end function

function ComboBuscar(strsql1,strsql2,CantRegMostrarx,CantRegMoverx,vuelta)
'response.write strsql1 & chr(13)
' cargo los vectores para que coincidan con el orden en la pagina
  set rx=crearconsultaEx(StrSql1,cn,1,parametros)
  i = 0
  j = 1

  Do until rx.eof
     for YY = 1 to CantRegMostrarx
       if rx.eof then exit for end if   
         Redim Preserve Perfiles(j)
         Redim Preserve Perfilesd(j)
         Perfiles(j) = i
         Perfilesd(j) = rx(0)
         rx.movenext   
         j=j+1
     next  
     i=i + CantRegMostrarx
  Loop
  call cerrarconsulta(rx)

' muestro el contenido de buscar ord. alfab. pero con los indices del orden de la pagina

  set rx=crearconsultaEx(StrSql2,cn,1,parametros) ' consulta ordenada alfabeticamente
  buscar = "<form name=formbuscar action=xxxxx " 
  buscar = buscar & " method=post>" &chr(13)  
  buscar = buscar & "<td height=20 align=left>" & chr(13)
  buscar = buscar & "<input class=bt type=button size=10% value='Buscar' ></td>" & chr(13)
  buscar = buscar & "<td><font class=at><select class=at onchange=" & chr(34) & "document.formbuscar.action = '" & vuelta & "?CantRegMostrar=" & Cantregmostrarx & "&CantRegMover=' + document.formbuscar.buscarlis[this.selectedIndex].value; document.formbuscar.submit();" & chr(34) & " name=buscarlis >"
  buscar = buscar & "<option value=" & CantRegMoverx & " selected>&nbsp;" 
  Do until rx.eof
     for YY = 1 to ubound(Perfiles)
       If Perfilesd(yy) = rx(0) then   
          buscar = buscar & "<option value=" & Perfiles(yy) & " >" & Perfilesd(yy)
          Perfilesd(yy) = "***" 
          exit for
       End If
     next
     rx.movenext   
  Loop
  buscar = buscar & "</select></font></td></form>"
  call cerrarconsulta(rx)

  Combobuscar = buscar
end function

function combofiltrar (CantRegMostrarx,CantRegMoverx,vuelta,filtro)
  filtrar = "<form name=formfiltrar action=xxxxx " 
  filtrar = filtrar & " method=post>" &chr(13)  
  filtrar = filtrar & "<td height=20 align=left>" & chr(13)
  filtrar = filtrar & "<input class=bt type=button size=10% value='Mostrar' ></td>" & chr(13)
  filtrar = filtrar & "<td><font class=at><select class=at onchange=" & chr(34) & "document.formfiltrar.action = '" & convertircar(vuelta,"_","&") & "&Filtrar=' + document.formfiltrar.filtrarlis[this.selectedIndex].value; document.formfiltrar.submit();" & chr(34) & " name=filtrarlis >"
  filtrar = filtrar & "<option value=no>TODOS" 
if filtro = "cri" then
  filtrar = filtrar & "<option value=cri selected>CRITICOS"
else
  filtrar = filtrar & "<option value=cri>CRITICOS"
end if
if filtro = "nocri" then
  filtrar = filtrar & "<option value=nocri selected>NO CRITICOS"
else
  filtrar = filtrar & "<option value=nocri>NO CRITICOS"
end if

  filtrar = filtrar & "</select></font></td></form>"

  Combofiltrar = filtrar
end function

function ValClaveRep(Clave) ' testea si un valor en un imput esta contenido en un vector

repetido =  "<script language=JavaScript >"
repetido = repetido & " function vectord(Clave) "
repetido = repetido & " { "
repetido = repetido & " var vectclave = new Array("
for j=0 to ubound(clave)
    repetido = repetido & chr(34) & Clave(j) & chr(34) & ","
next
repetido =  mid(repetido,1,len(repetido)-1) 
repetido = repetido & "); "
repetido = repetido & " return vectclave[Clave]; } "
repetido = repetido & " "
repetido = repetido & " function valida(dato) "
repetido = repetido & " { sale=0; "
repetido = repetido & " for(i=0; i < " & Ubound(clave) + 1 & "; i++) " 
repetido = repetido & "{ if ( vectord(i) == dato.value) { "
repetido = repetido & " alert ('El Codigo ' + dato.value + ' ya "
repetido = repetido & "fue utilizado.'); dato.value = dato.defaultValue; dato.focus(); "
repetido = repetido & " sale=1; i=" & ubound(Clave) & " + 2; } "
repetido = repetido & "}"
repetido = repetido & " if (sale==1) {return false;} "
repetido = repetido & " } "
repetido = repetido & "</script>"
ValClaveRep = repetido
end function

function VectorClave(str,rspos)
   set cn=conectar(strconn2)
   set rs=crearconsultaEx(Str,cn,1,parametros)
   n=0
 
   If rs.eof then 
      Redim Preserve Clave(0)
      Clave(0) = ""
   End If
   Do until rs.Eof 
      Redim Preserve Clave(n)
      Clave(n) = rs(rspos)
      rs.Movenext
      n = n + 1
   Loop
   call cerrarconsulta(rs)
   VectorClave = clave
end function

function nombre(str,valor)
  set rnom=crearconsultaEx(Str,cn,1,parametros)
  if not rnom.eof then 
     nombre = rnom(0)
  else
     nombre= ""
  end if

  call cerrarconsulta(rnom) 
end function

%>
