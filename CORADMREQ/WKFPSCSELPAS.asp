<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = "../WKFPASOSPROP.ASP"

  set cn = ofv.conectar(ofv.strconn5)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")
 WKFPSCTPC = request.querystring("WKFPSCTPC")

 if WKFPSCTPC = "" then
    WKFPSCTPC = session("WKFPSCTPC")
 else
    session("WKFPSCTPC") = WKFPSCTPC
 end if 

    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenF = "<B><FONT FACE=webDINGS COLOR=#993333 SIZE=5>q</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
    sImagenSI = "<B><FONT FACE=WINGDINGS COLOR=#0000aa SIZE=4>&#0236;</FONT></B>"
    sImagenSF = "<B><FONT FACE=WINGDINGS COLOR=#000000 SIZE=4>&#0237;</FONT></B>"
    sImagenSA = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=5>&#0252;</FONT></B>"
    sImagenSR = "<B><FONT FACE=WINGDINGS COLOR=#aa0000 SIZE=5>&#0251;</FONT></B>"
    sImagenPA = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=4>&#0231;</FONT></B>"
    sImagenPS = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=4>&#0232;</FONT></B>"


IF SOPCION = "SI" THEN

   NEWPASO = GENWKFPASONEW()

   STRSQL = "UPDATE WKFPSCTPC SET PASOSIG = '" & NEWPASO & "' "
   STRSQL = STRSQL & " WHERE ID = " & session("WKFPSCTPC")
   call ofv.crearconsultaEx(StrSql,cn,1,volver)

  StrSql = "Select * From WKFPSCTPC WHERE ID = " & session("WKFPSCTPC")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then


     IF RS(3) = "INI" OR RS(3) = "APR" THEN
        STRSQL = "UPDATE WKFPASOS SET APRPSC = '" & NEWPASO & "' "
        STRSQL = STRSQL & " WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)
     ELSE
        STRSQL = "UPDATE WKFPASOS SET RCHPSC = '" & NEWPASO & "' "
        STRSQL = STRSQL & " WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)  
     END IF


  END IF

  RESPONSE.REDIRECT "WKFPASOSPROP.ASP"

ELSE


  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & session("WKF") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
'     IF RS(3) THEN
'        bAgregar = FALSE
'        bModificar = FALSE
'        bSuprimir = FALSE
'     END IF


    sInput = "CIRCUITO:"
    celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

    sInput = RS(1) & " - " & RS(2)
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


  END IF



    celdas = ""





  StrSql = "Select * From WKFPASOS WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then



    sInput = "PASO:"
    celdas = ofv.GenCelda("","vt","","10%","","","",sInput,"si")

    sInput = "<big>" & RS(3) & ". " & RS(4) & "</big>"
    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

        sinputK = "<a class=nt style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href=WKFPASOSPROP.asp"
        sinputK = sinputK  & "  TITLE='Volver&nbsp;al&nbsp;Paso'  ><b>" & sImagenF & "&nbsp;<b><big>Volver&nbsp;al&nbsp;Paso</big></b></b></a>" 


    celdas = celdas & ofv.GenCelda("","vt","CENTER","10%","","","",sInputK,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


 END IF




  StrSql = "Select * From WKFPSCTPC WHERE ID = " & session("WKFPSCTPC")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then


     IF RS(3) = "INI" THEN sInput = "INICIAR"     
     IF RS(3) = "APR" THEN sInput = "APROBAR"     
     IF RS(3) = "RCH" THEN sInput = "RECHAZAR"     

    sInput = "DERIVAR&nbsp;AL&nbsp;" & SINPUT & "&nbsp;A:&nbsp;&nbsp;"
    celdas = ofv.GenCelda("","vt","","15%","","","",sInput,"si")


   XPASO = RS(6) & ""
   if XPASO <> "" then 
     sInput = GENWKFPASO(RS(6))
   ELSE
    sInput = "SELECCIONAR"
   END IF

    sInput = sImagenPS & "&nbsp;" & SINPUT	

    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"





''' ANTERIOR

        sinputK = "<a class=nt style='text-decoration:none;color:#993333;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='WKFPSCSELPAS.asp?OPCION=SI' "
        sinputK = sinputK  & "  TITLE='Crear nuevo paso'  ><b>" & sImagenPS & "&nbsp;&nbsp;<big>" & sImagenPC & "&nbsp;CREAR NUEVO PASO</big></b></a>" 

        STABC = ofv.GenCelda("","ut  ","","","","","",sInputK,"si")
        STAB = ofv.GenRow("","","","","","","",STABC,"si")
      strsql = "Select * From WKFPASOS WHERE WKFCOD = '" & session("WKF") & "' AND WKFPASCOD <> '" & RS(2) & "' AND WKFPASCOD <> '" & RS(6) & "'  Order By ORDEN, DESCRIPCION "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF not RSNIV.EOF THEN

      DO WHILE NOT RSNIV.EOF


          xpaso = GENWKFPASO(RSNIV(2))

        sinputK = "<a class=nt style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='javascript:void(0);' onclick=' document.all." & rs(6).name & ".value=" & chr(34) & RSNIV(2) & chr(34) & ";submit();return false;' "
        sinputK = sinputK  & "  TITLE='Derivar a " & xpaso & "'  ><b>" & sImagenPS & "&nbsp;&nbsp;<big>" & xpaso & "</big></b></a>" 


        
        STABC = ofv.GenCelda("","ut  ","","","","","",sInputK,"si")
        STAB = STAB & ofv.GenRow("","","","","","","",STABC,"si")

         RSNIV.MOVENEXT
      LOOP
        STAB = ofv.GenTabla("","","","80%","","0","","0","0","","",STAB,"si")
        sInput = "<input id=" & rs(6).name & " name=" & rs(6).name & " type=hidden value=" & rs(6) & " >"


        sAccion = "asp/ActualizarWK.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPSCTPC"
        STAB = ofv.GenForm("fder","",sAccion,"","post","",STAB & sinput,"si")

        END IF
      call ofv.cerrarconsulta(rsniv)


''' ANTERIOR
    celdas = ofv.GenCelda("","vt STYLE='BORDER-bottom-STYLE:solid;BORDER-bottom-WIDTH:2;BORDER-bottom-color:#993333;' colspan=2 ","CENTER","","","","","&nbsp;","si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")

    celdas = ofv.GenCelda("","vt bgcolor=#ffffff ","CENTER","15%","","","","<big>SELECCIONAR</big>","si")
    celdas = celdas & ofv.GenCelda("","ut  ","CENTER","85%","","","",STAB,"si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")
    STAB = ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")

    celdas = ofv.GenCelda("","","CENTER","","","","",STAB,"si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")
 
    sHTML = sHTML  & "<CENTER>"
'' & ofv.GenTabla("","AAA STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:2;' ","","80%","","0","","0","0","","",filas,"si") & "<BR>"


  sHTML = sHTML  & ofv.GenTabla("","aaa style='filter:shadow(color=#cccccc,direction=125);' ","","100%","","0","","0","0","","",filas,"si") & "<BR>"



END IF





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




  'Botonera
  Agregar = ""


  Buscar = ""






  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"

 END IF

end if

response.write sHTML

FUNCTION GENWKFPASONEW()

    WKFCOD = GTWKFPSC()

  StrSql = "Select MAX(ORDEN) From WKFPASOS WHERE WKFCOD = '" & session("WKF") & "' and orden < 999 "
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then 
     NORDEN = RS3(0) + 10
  ELSE

     NORDEN = 20
  END IF

    strsql = "insert into WKFPASOS (WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,APR,RCH,TIPO) values ('" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','NUEVO PASO WKF','1','1','U')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  GENWKFPASONEW = WKFCOD

END FUNCTION

FUNCTION GENWKFPASO(PASO)

  GENWKFPASO = "SIN ASOCIAR"

  StrSql = "Select * From WKFPASOS WHERE WKFPASCOD = '" & PASO & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs2.eof then
     sinput = RS2(3) & ". " & RS2(4)
     GENWKFPASO = sinput
  END IF
  call ofv.cerrarconsulta(rs2)

END FUNCTION

function pp
    sInput = "SELECCIONAR"
    celdas = ofv.GenCelda("","vt","CENTER","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")
    formularioA = filas


    if bModificar then
      strsql = "Select * From WKFPASOS WHERE WKFCOD = '" & session("WKF") & "'  Order By ORDEN, DESCRIPCION "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF RSNIV.EOF THEN
           sInput = "NO HAY PASOS PARA ASIGNAR<BR>SELECCIONE ""CREAR NUEVO PASO"" "

        ELSE
        sInput = "<font class=at><select SIZE=5 name=" & rs(6).name & " onchange=submit()>"
      DO WHILE NOT RSNIV.EOF
          sInput = sInput & "<option value=" & RSNIV(2) & " "
          if rs(6) = RSNIV(2) then sInput = sInput & "selected"
          sInput = sInput & " >" & GENWKFPASO(RSNIV(2))
         RSNIV.MOVENEXT
      LOOP
        sInput = sInput & "</select></font>"
        END IF
      call ofv.cerrarconsulta(rsniv)
    else
      SNIV = "SIN ASOCIAR" 
      strsql = "Select * From WKFPASOS WHERE WKFPASCOD = '" & RS(6) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if
    celdas = ofv.GenCelda("","vt","CENTER","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")
    if bModificar then
      sAccion = "asp/ActualizarWK.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPSCTPC"
      formularioC = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formularioC = ofv.GenForm("","","","","post","",filas,"si")
    end if
end function

%>