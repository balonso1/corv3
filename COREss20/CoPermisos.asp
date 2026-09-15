<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

function genlinea(indi,stipo,stit,stit2,sdesc,s1,s2,s3,s4,s5,sp1,sp2)
   if indi > 40 or indi = 0 then
      if indi > 0 then
         response.write ofv.GenFTabla()  
         response.write ofv.GenFcelda() 
         response.write ofv.GenFRow() 
         response.write ofv.GenFTabla()  
         if ofv.explorador = "MSIE" then response.write "<DIV>&nbsp;</DIV>"
         response.flush  
      end if  

      response.write ofv.GenTabla("","aa  border=0","","","","0","0","0","0","","","","no")
      response.write ofv.GenRow("","","","","","","","","no")
      response.write ofv.gencelda("","","","","100%","","","","no")
      response.write ofv.RepHeader2(stit,"600")
      response.write "<br>"
      response.write ofv.GenTabla("","aa  border=0","","600","","0","0","1","0","","","","no")

    if stit2 <> "" then
       Celda = ofv.gencelda("","vt","","50% colspan=10","","","",stit2,"si")
       response.write  ofv.GenRow("","","","","","","",Celda,"si")
    end if

    Celda = ofv.gencelda("","vt","","50% colspan=5","","#99aadd","",Ucase("Item"),"si")
    Celda = Celda & ofv.gencelda("","vt","","2%","","#99aadd","",Ucase("Activo"),"si")
    Celda = Celda & ofv.gencelda("","vt","","2%","","#99aadd","",Ucase("a"),"si")
    Celda = Celda & ofv.gencelda("","vt","","2%","","#99aadd","",Ucase("m"),"si")
    Celda = Celda & ofv.gencelda("","vt","","2%","","#99aadd","",Ucase("e"),"si")
    Celda = Celda & ofv.gencelda("","vt","","2%","","#99aadd","",Ucase("c"),"si")
    response.write ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ""
    for rr=1 to 10 
        if rr = 5 then
           Celda = Celda & ofv.gencelda("","ut","","50%","","","","&nbsp;","si")
        else
           Celda = Celda & ofv.gencelda("","ut","","1%","","","","&nbsp;","si")
        end if
    next
    response.write ofv.GenRow("","","","","","","",Celda,"si")
    indi = 0
  end if
    if stipo = "T" then
          celdas = ofv.GenCelda("","ut colspan=" & sp1,"","","","","",sdesc,"si")
    else

    if sp1 > 0 then
       if sp1 = 1 then
          celdas = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
       else
          celdas = ofv.GenCelda("","ut colspan=" & sp1,"","","","","","&nbsp;","si")
       end if
    end if 

    if sp2 > 0 then
       if sp2 = 1 then
          celdas = celdas & ofv.GenCelda("","ut","","","","","",sdesc,"si")
       else
          celdas = celdas & ofv.GenCelda("","ut colspan=" & sp2,"","","","","",sdesc,"si")
       end if
    end if 

    if stipo = "P" then
       celdas = celdas & ofv.GenCelda("","ut","","","","","","&nbsp;","si")
       celdas = celdas & ofv.GenCelda("","ut","","","","","",s2,"si")
       celdas = celdas & ofv.GenCelda("","ut","","","","","",s3,"si")
       celdas = celdas & ofv.GenCelda("","ut","","","","","",s4,"si")
       celdas = celdas & ofv.GenCelda("","ut","","","","","",s5,"si")
    else
       celdas = celdas & ofv.GenCelda("","ut colspan=5","","","","","",s1,"si")
    end if
    end if
    response.write  ofv.GenRow("","","","","","","",Celdas,"si")
    indi = indi + 1


end function

ok=ofv.CheckUsuario()
response.buffer = true
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
'  response.expires=0

  ID = request.querystring("ID")
  t = request.querystring("t")

  set cn = ofv.conectar(ofv.strconn0)
  set cx = ofv.conectar(ofv.strconn1)

  Strsql = "Select * From CorAplic where enable<>0 Order By DescApli"
  set rs = ofv.crearconsultaEx(StrSql,cx,1,parametros)
  redim preserve aApli(0)
  redim preserve aAplicd(0)
  ix = 0
  do until rs.eof
    redim preserve aApli(ix)
    redim preserve aAplicd(ix)
    aApli(ix) = ofv.convertircar(rs(2),"*"," ") & rs(4)
    aAplicd(ix) = rs(1)
    ix = ix + 1
    rs.movenext

  loop
  call ofv.cerrarconsulta(rs)

  if t = "G" then
    tabla = "GRUPO"
  else
    tabla = "USUARIO"
  end if
  xdescri = tabla

  Strsql = "Select * From " & tabla & "s where ID = '" & ID & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     xdescri = tabla & ": " & ID & " - " & rs("descripcion") 
  end if 
  if t <> "G" then
    tabla = "GRUPO"
    xdescri2 = ""
    Strsql = "Select * From " & tabla & "s where ID = '" & rs("Grupo") & "' "
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then
       xdescri2 = "Integrante del Grupo: " & rs("ID") & " - " & rs("descripcion") 
    end if 
  end if
  call ofv.cerrarconsulta(rs)
  
  sfsn = "<font face=tahoma,arial size=1 color=#dddddd><b>"
  sfnn = "<font face=tahoma,arial size=1 color=#dddddd><b>"

  sfsg = "<font face=tahoma,arial size=1 color=#900000><b>"
  sfng = "<font face=tahoma,arial size=1 color=#900000><b>"

  sfsu = "<font face=tahoma,arial size=1 color=#336699><b>"
  sfnu = "<font face=tahoma,arial size=1 color=#336699><b>"

  sfs = ""
  sfn = ""


  sff = "</b></font>"

  response.write ofv.Repbody(Ucase("permisos del ") & XDESCRI)
    indi = 0
    xtit = Ucase("permisos del ") & XDESCRI
    if xdescri2 <> "" then
       xtit2 =  Ucase(xdescri2)
    end if

    call genlinea(indi,"M",xtit,xtit2,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","&nbsp;","&nbsp;","0","5")

'xxxxap


  strsql = "SELECT * From Atrm Where tipo = 'A' order by atributo"
  set rs0 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  xgrupo = " "
  xgrp = "S"
  xusr = "S" 

  if t <> "G" then
     strsql = "SELECT grupo From usuarios Where ID = '" & ID & "' "
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if not rsu.eof then xgrupo = rsu(0) 
     call ofv.cerrarconsulta(rsu)

     strsql = "SELECT * From AtrmXusuario Where usuario = '" & ID & "' "
     strsql = strsql & "And Tipo = 'A' Order By atributo"
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsu.eof then xusr = "N"
  else

      xgrupo = ID
      xusr = "N" 

  end if

     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And Tipo = 'A' Order By atributo"
     set rsg = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsg.eof then xgrp = "N"

  ix = 0
  ixx = 0
  do until rs0.eof
     sfs = sfsn
     sfn = sfnn
     xestado = 1
     xestadoD = 1
     xdef = "D"
     xcolor = ""
     for ix=0 to  Ubound(aApli)
         if rs0("Aplicacion") = aAplicd(ix) then
            ixx = ix
            exit for
         end if  
     next   

    if rs0("Aplicacion") = aAplicd(ixx) then
      ix = ixx
      sAplicacion = "Aplicacion: " & UCase(aApli(ix))

    if xgrp <> "N" then
       if rsg("atributo") = rs0("atributo") then
          if not rsg("grpestado") then 
             xestado = 0  
             xestadoD = 0  
             xdef = "G"
             sfs = sfsg
             sfn = sfng
          end if
          rsg.movenext
          if rsg.eof then xgrp = "N"
        end if
     end if

    if xusr <> "N" then
       if rsu("atributo") = rs0("atributo") then
             xestado = abs(cint(rsu("usrestado")))
             xcolor = "#6699cc"
             sfs = sfsu
             sfn = sfnu
          rsU.movenext
          if rsU.eof then xusr = "N"
        end if
     end if

      if xestado = 1 then
        sInput = sfs & "SI" & sff
      else
        sInput = sfn & "NO" & sff
      end if

    call genlinea(indi,"M",xtit,xtit2,sAplicacion,sInput,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","0","5")

      if xestado = 1 then


' xxxxxmn

  set cn2 = ofv.conectar(ofv.strconn0)
  set cx2 = ofv.conectar(ofv.strconn1)


  strsql = "SELECT * From Atrm Where tipo = 'S'  And aplicacion = '" & rs0("Aplicacion") & "'  order by alias"
  set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)


  xgrp2 = "S"
  xusr2 = "S" 

  if t <> "G" then
     strsql = "SELECT * From AtrmXusuario Where usuario = '" & ID & "' "
     strsql = strsql & "And Tipo = 'S'  And aplicacion = '" & rs0("Aplicacion") & "'  Order By Alias"
     set rsu2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros) 
     if rsu2.eof then xusr2 = "N"
  else

      xgrupo = ID
      xusr2 = "N" 

  end if

     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And Tipo = 'S'  And aplicacion = '" & rs0("Aplicacion") & "' Order By Alias"
     set rsg2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros) 
     if rsg2.eof then xgrp2 = "N"

  do until rs2.eof
     sfs = sfsn
     sfn = sfnn
     xestado = 1
     xestadoD = 1
     xdef = "D"
     xcolor = ""
      xdescri = ""
      xdesatr = " <font color=#993333 face=tahoma size=1>(" & rs2("Alias") & ")</font>"
      strsql = "select menu from menux where menuatr = '" & trim(rs2("Alias")) & "' "
      set rsm = ofv.crearconsultaEx(StrSql,cx2,1,parametros)
      if not rsm.eof then
         xdescri = Ucase(rsm(0)) & " - "
      end if 
      call ofv.cerrarconsulta(rsm)

    if xgrp2 <> "N" then
       if rsg2("Alias") = rs2("Alias") then
          if not rsg2("grpestado") then 
             xestado = 0  
             xestadoD = 0  
             xdef = "G"
             xcolor = "#993333"
             sfs = sfsg
             sfn = sfng
          end if
          rsg2.movenext
          if rsg2.eof then xgrp2 = "N"
        end if
     end if

    if xusr2 <> "N" then
       if rsu2("Alias") = rs2("Alias") then
             xestado = abs(cint(rsu2("usrestado")))
             xcolor = "#6699cc"
             sfs = sfsu
             sfn = sfnu
          rsU2.movenext
          if rsU2.eof then xusr2 = "N"
        end if
     end if

      if  xestado = 1 then
        sInput2 = sfs & "SI" & sff
      else
        sInput2 = sfn & "NO" & sff
      end if

    call genlinea(indi,"M",xtit,xtit2,"Solapa: " & xdescri & xdesatr,sInput2,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","1","4")
      if  xestado = 1 then

' xxxxxsmn
  set cn3 = ofv.conectar(ofv.strconn0)
  set cx3 = ofv.conectar(ofv.strconn1)


  strsql = "SELECT * From Atrm Where tipo = 'M'  And aplicacion = '"
  strsql = strsql & rs0("Aplicacion") & "' and depen = '" & trim(rs2("Alias")) & "'  order by atributo"
  set rs3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)


  xgrp3 = "S"
  xusr3 = "S" 

  if t <> "G" then
     strsql = "SELECT * From AtrmXusuario Where usuario = '" & ID & "' "
     strsql = strsql & "And Tipo = 'M'  and depen = '" & trim(rs2("Alias")) & "'  And aplicacion = '" & rs0("Aplicacion")  & "'  Order By atributo"
     set rsu3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros) 
     if rsu3.eof then xusr3 = "N"
  else

      xgrupo = ID
      xusr3 = "N" 

  end if

     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And Tipo = 'M'  and depen = '" & trim(rs2("Alias")) & "'  And aplicacion = '" & rs0("Aplicacion")  & "' Order By atributo"
     set rsg3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros) 
     if rsg3.eof then xgrp3 = "N"


  do until rs3.eof

     sfs = sfsn
     sfn = sfnn
     xestado = 1
     xestadoD = 1
     xdef = "D"
     xcolor = ""
      xdescri = ""
      xdesatr = " <font color=#993333 face=tahoma size=1>(" & rs3("Alias") & ")</font>"
      strsql = "select menu from menux where menuatr = '" & trim(rs3("Alias")) & "' "
      set rsm = ofv.crearconsultaEx(StrSql,cx3,1,parametros)
      if not rsm.eof then
         xdescri = Ucase(rsm(0)) & " - "
      end if 
      call ofv.cerrarconsulta(rsm)

    if xgrp3 <> "N" then
       if rsg3("atributo") = rs3("atributo") then
          if not rsg3("grpestado") then 
             xestado = 0  
             xestadoD = 0  
             xdef = "G"
             xcolor = "#993333"
             sfs = sfsg
             sfn = sfng
          end if
          rsg3.movenext
          if rsg3.eof then xgrp3 = "N"
        end if
     end if

    if xusr3 <> "N" then
       if rsu3("atributo") = rs3("atributo") then
             xestado = abs(cint(rsu3("usrestado")))
             xcolor = "#6699cc"
             sfs = sfsu
             sfn = sfnu
          rsU3.movenext
          if rsU3.eof then xusr3 = "N"
        end if
     end if



      if  xestado = 1 then
        sInput3 = sfs & "SI" & sff
      else
        sInput3 = sfn & "NO" & sff
      end if

    call genlinea(indi,"M",xtit,xtit2,"Menu: " & xdescri & xdesatr,sInput3,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","2","3")

      if  xestado = 1 then
' xxxxxsmn2
  set cn4 = ofv.conectar(ofv.strconn0)
  set cx4 = ofv.conectar(ofv.strconn1)

  strsql = "SELECT * From Atrm Where tipo = 'M'  And aplicacion = '"
  strsql = strsql & rs0("Aplicacion") & "' and depen = '" & trim(rs3("Alias")) & "'  order by atributo"
  set rs4 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)


  xgrp4 = "S"
  xusr4 = "S" 

  if t <> "G" then
     strsql = "SELECT * From AtrmXusuario Where usuario = '" & ID & "' "
     strsql = strsql & "And Tipo = 'M'  and depen = '" & trim(rs3("Alias")) & "'  And aplicacion = '" & rs0("Aplicacion")  & "'  Order By atributo"
     set rsu4 = ofv.crearconsultaEx(StrSql,cn4,1,parametros) 
     if rsu4.eof then xusr4 = "N"
  else

      xgrupo = ID
      xusr4 = "N" 

  end if

     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And Tipo = 'M'  and depen = '" & trim(rs3("Alias")) & "'  And aplicacion = '" & rs0("Aplicacion")  & "' Order By atributo"
     set rsg4 = ofv.crearconsultaEx(StrSql,cn4,1,parametros) 
     if rsg4.eof then xgrp4 = "N"


  do until rs4.eof
     sfs = sfsn
     sfn = sfnn
     xestado = 1
     xestadoD = 1
     xdef = "D"
     xcolor = ""
      xdescri = ""
      xdesatr = " <font color=#993333 face=tahoma size=1>(" & rs4("Alias") & ")</font>"
      strsql = "select menu from menux where menuatr = '" & trim(rs4("Alias")) & "' "
      set rsm = ofv.crearconsultaEx(StrSql,cx4,1,parametros)
      if not rsm.eof then
         xdescri = Ucase(rsm(0)) & " - "
      end if 
      call ofv.cerrarconsulta(rsm)

    if xgrp4 <> "N" then
       if rsg4("atributo") = rs4("atributo") then
          if not rsg4("grpestado") then 
             xestado = 0  
             xestadoD = 0  
             xdef = "G"
             xcolor = "#993333"
             sfs = sfsg
             sfn = sfng
          end if
          rsg4.movenext
          if rsg4.eof then xgrp4 = "N"
        end if
     end if

    if xusr4 <> "N" then
       if rsu4("atributo") = rs4("atributo") then
             xestado = abs(cint(rsu4("usrestado")))
             xcolor = "#6699cc"
             sfs = sfsu
             sfn = sfnu
          rsU4.movenext
          if rsU4.eof then xusr4 = "N"
        end if
     end if



      if  xestado = 1 then
        sInput4 = sfs & "SI" & sff
      else
        sInput4 = sfn & "NO" & sff
      end if


    call genlinea(indi,"M",xtit,xtit2,"Submenu: " & xdescri & xdesatr,sInput4,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","3","2")
      if  xestado = 1 then
' xxxxxsmn3
  set cn5 = ofv.conectar(ofv.strconn0)
  set cx5 = ofv.conectar(ofv.strconn1)

  strsql = "SELECT * From Atrm Where (tipo = 'F' or tipo = 'N') And aplicacion = '"
  strsql = strsql & rs0("Aplicacion") & "' and depen = '" & trim(rs4("Alias")) & "'  order by atributo"
  set rs5 = ofv.crearconsultaEx(StrSql,cn5,1,parametros)


  xgrp5 = "S"
  xusr5 = "S" 

  if t <> "G" then
     strsql = "SELECT * From AtrmXusuario Where usuario = '" & ID & "' "
     strsql = strsql & "And (tipo = 'F' or tipo = 'N')  and depen = '" & trim(rs4("Alias")) & "' And aplicacion = '" & rs0("Aplicacion") & "'  Order By atributo"
     set rsu5 = ofv.crearconsultaEx(StrSql,cn5,1,parametros) 
     if rsu5.eof then xusr5 = "N"
  else

      xgrupo = ID
      xusr5 = "N" 

  end if

     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And (tipo = 'F' or tipo = 'N')  and depen = '" & trim(rs4("Alias")) & "' And aplicacion = '" & rs0("Aplicacion") & "' Order By atributo"
     set rsg5 = ofv.crearconsultaEx(StrSql,cn5,1,parametros) 
     if rsg5.eof then xgrp5 = "N"



  redim aAtributos(3)
  aAtributos(0) = "AGREGAR"
  aAtributos(1) = "MODIFICAR"
  aAtributos(2) = "ELIMINAR"
  aAtributos(3) = "CONSULTAR"

  redim aAtrm(3)
  aAtrm(0) = 1
  aAtrm(1) = 1
  aAtrm(2) = 1
  aAtrm(3) = 1

  redim aAtrmD(3)
  aAtrmD(0) = 1
  aAtrmD(1) = 1
  aAtrmD(2) = 1
  aAtrmD(3) = 1


  do until rs5.eof
     sfs = sfsn
     sfn = sfnn
    if ucase(rs5("tipo")) = "F" then
       xdescri = "Permisos de Pantalla"
    else
       xdescri = "Permisos del Nodo " & rs5("alias") 
    end if

     xdef = "D"
     xcolor = ""

     aAtrm(0) = 1
     aAtrm(1) = 1
     aAtrm(2) = 1
     aAtrm(3) = 1

     aAtrmD(0) = 1
     aAtrmD(1) = 1
     aAtrmD(2) = 1
     aAtrmD(3) = 1



    if xgrp5 <> "N" then
       if rsg5("atributo") = rs5("atributo") then
          if not rsg5("grpagregar") or not rsg5("grpmodificar") or not rsg5("grpeliminar") or not rsg5("grpconsultar")  then 
             aAtrm(0) = abs(cint(rsg5("grpagregar")))
             aAtrm(1) = abs(cint(rsg5("grpmodificar")))
             aAtrm(2) = abs(cint(rsg5("grpeliminar")))
             aAtrm(3) = abs(cint(rsg5("grpconsultar")))

             aAtrmD(0) = abs(cint(rsg5("grpagregar")))
             aAtrmD(1) = abs(cint(rsg5("grpmodificar")))
             aAtrmD(2) = abs(cint(rsg5("grpeliminar")))
             aAtrmD(3) = abs(cint(rsg5("grpconsultar"))) 

             xdef = "G"
             xcolor = "#993333"
             sfs = sfsg
             sfn = sfng
          end if
          rsg5.movenext
          if rsg5.eof then xgrp5 = "N"
        end if
     end if

    if xusr5 <> "N" then
       if rsu5("atributo") = rs5("atributo") then
             aAtrm(0) = abs(cint(rsu5("usragregar")))
             aAtrm(1) = abs(cint(rsu5("usrmodificar")))
             aAtrm(2) = abs(cint(rsu5("usreliminar")))
             aAtrm(3) = abs(cint(rsu5("usrconsultar")))
             xcolor = "#6699cc"
             sfs = sfsu
             sfn = sfnu
          rsU5.movenext
          if rsU5.eof then xusr5 = "N"
        end if
     end if

      if aAtrm(0) then
        sInput4 = sfs & "SI" & sff
      else
        sInput4 = sfn & "NO" & sff
      end if
     si1 = sInput4

      if aAtrm(1)  then
        sInput4 = sfs & "SI" & sff
      else
        sInput4 = sfn & "NO" & sff
      end if
     si2 = sInput4

      if aAtrm(2)  then
        sInput4 = sfs & "SI" & sff
      else
        sInput4 = sfn & "NO" & sff
      end if
     si3 = sInput4

      if aAtrm(3)  then
        sInput4 = sfs & "SI" & sff
      else
        sInput4 = sfn & "NO" & sff
      end if
     si4 = sInput4


    call genlinea(indi,"P",xtit,xtit2,xdescri,"&nbsp;",si1,si2,si3,si4,"4","1")

    rs5.movenext
  loop
  call ofv.cerrarconsulta(rs5)
  call ofv.cerrarconn(cx5)
  call ofv.cerrarconn(cn5)

' xxxxxsmn3
  end if
    rs4.movenext
  loop
  call ofv.cerrarconsulta(rs4)
  call ofv.cerrarconn(cx4)
  call ofv.cerrarconn(cn4)

  set cn4 = ofv.conectar(ofv.strconn0)
  set cx4 = ofv.conectar(ofv.strconn1)

  strsql = "SELECT * From Atrm Where (tipo = 'F' or tipo = 'N') And aplicacion = '"
  strsql = strsql & rs0("Aplicacion") & "' and depen = '" & trim(rs3("Alias")) & "'  order by atributo"
  set rs4 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)


  xgrp4 = "S"
  xusr4 = "S" 

  if t <> "G" then
     strsql = "SELECT * From AtrmXusuario Where usuario = '" & ID & "' "
     strsql = strsql & "And (tipo = 'F' or tipo = 'N')  and depen = '" & trim(rs3("Alias")) & "' And aplicacion = '" & rs0("Aplicacion") & "'  Order By atributo"
     set rsu4 = ofv.crearconsultaEx(StrSql,cn4,1,parametros) 
     if rsu4.eof then xusr4 = "N"
  else

      xgrupo = ID
      xusr4 = "N" 

  end if

     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And (tipo = 'F' or tipo = 'N')  and depen = '" & trim(rs3("Alias")) & "' And aplicacion = '" & rs0("Aplicacion") & "' Order By atributo"
     set rsg4 = ofv.crearconsultaEx(StrSql,cn4,1,parametros) 
     if rsg4.eof then xgrp4 = "N"



  redim aAtributos(3)
  aAtributos(0) = "AGREGAR"
  aAtributos(1) = "MODIFICAR"
  aAtributos(2) = "ELIMINAR"
  aAtributos(3) = "CONSULTAR"

  redim aAtrm(3)
  aAtrm(0) = 1
  aAtrm(1) = 1
  aAtrm(2) = 1
  aAtrm(3) = 1

  redim aAtrmD(3)
  aAtrmD(0) = 1
  aAtrmD(1) = 1
  aAtrmD(2) = 1
  aAtrmD(3) = 1


  do until rs4.eof
     sfs = sfsn
     sfn = sfnn
    if ucase(rs4("tipo")) = "F" then
       xdescri = "Permisos de Pantalla"
    else
       xdescri = "Permisos del Nodo " & rs4("alias") 
    end if

     xdef = "D"
     xcolor = ""

     aAtrm(0) = 1
     aAtrm(1) = 1
     aAtrm(2) = 1
     aAtrm(3) = 1

     aAtrmD(0) = 1
     aAtrmD(1) = 1
     aAtrmD(2) = 1
     aAtrmD(3) = 1



    if xgrp4 <> "N" then
       if rsg4("atributo") = rs4("atributo") then
          if not rsg4("grpagregar") or not rsg4("grpmodificar") or not rsg4("grpeliminar") or not rsg4("grpconsultar")  then 
             aAtrm(0) = abs(cint(rsg4("grpagregar")))
             aAtrm(1) = abs(cint(rsg4("grpmodificar")))
             aAtrm(2) = abs(cint(rsg4("grpeliminar")))
             aAtrm(3) = abs(cint(rsg4("grpconsultar")))

             aAtrmD(0) = abs(cint(rsg4("grpagregar")))
             aAtrmD(1) = abs(cint(rsg4("grpmodificar")))
             aAtrmD(2) = abs(cint(rsg4("grpeliminar")))
             aAtrmD(3) = abs(cint(rsg4("grpconsultar"))) 

             xdef = "G"
             xcolor = "#993333"
             sfs = sfsg
             sfn = sfng
          end if
          rsg4.movenext
          if rsg4.eof then xgrp4 = "N"
        end if
     end if

    if xusr4 <> "N" then
       if rsu4("atributo") = rs4("atributo") then
             aAtrm(0) = abs(cint(rsu4("usragregar")))
             aAtrm(1) = abs(cint(rsu4("usrmodificar")))
             aAtrm(2) = abs(cint(rsu4("usreliminar")))
             aAtrm(3) = abs(cint(rsu4("usrconsultar")))
             xcolor = "#6699cc"
             sfs = sfsu
             sfn = sfnu
          rsU4.movenext
          if rsU4.eof then xusr4 = "N"
        end if
     end if

      if aAtrm(0) then
        sInput4 = sfs & "SI" & sff
      else
        sInput4 = sfn & "NO" & sff
      end if
     si1 = sInput4

      if aAtrm(1)  then
        sInput4 = sfs & "SI" & sff
      else
        sInput4 = sfn & "NO" & sff
      end if
     si2 = sInput4

      if aAtrm(2)  then
        sInput4 = sfs & "SI" & sff
      else
        sInput4 = sfn & "NO" & sff
      end if
     si3 = sInput4

      if aAtrm(3)  then
        sInput4 = sfs & "SI" & sff
      else
        sInput4 = sfn & "NO" & sff
      end if
     si4 = sInput4

    call genlinea(indi,"P",xtit,xtit2,xdescri,"&nbsp;",si1,si2,si3,si4,"3","2")
    rs4.movenext
  loop
  call ofv.cerrarconsulta(rs4)
  call ofv.cerrarconn(cx4)
  call ofv.cerrarconn(cn4)
' xxxxxsmn2
  end if
    rs3.movenext
  loop
  call ofv.cerrarconsulta(rs3)
  call ofv.cerrarconn(cx3)
  call ofv.cerrarconn(cn3)
' xxxxxsmn
    end if
    rs2.movenext
  loop
  call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cx2)
  call ofv.cerrarconn(cn2)

' xxxxxmn
   end if
        sinput = "<hr size=1 color=#5588bb widht=100% >"

    call genlinea(indi,"T",xtit,xtit2,sinput,"&nbsp;",si1,si2,si3,si4,"10","0")
   end if
   rs0.movenext

  loop

'xxxxxxxap



'  else
'  response.write  "<font face=tahoma,arial size=2 color=#336699>"   
'  response.write  "<p align=center>No existen permisos habilitados</p>"
'  response.write  "</font>"
'  end if

  response.write ofv.GenFTabla()  
  response.write ofv.GenFcelda() 
  response.write ofv.GenFRow() 
  response.write ofv.GenFTabla()  
  response.flush  

  call ofv.cerrarconsulta(rs0)
  call ofv.cerrarconn(cx)
  call ofv.cerrarconn(cn)

  response.write  "</center></body></html>"
end if

response.write sHTML
response.flush 
%>