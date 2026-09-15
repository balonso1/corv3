<%@LCID = 11274%> 
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else


  response.expires=0

  CODFAM = request.querystring("CODFAM")

  IF CODFAM <> "" THEN
     SESSION("CODFAM") = CODFAM
  ELSE
     CODFAM = SESSION("CODFAM")
  END IF 



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sopcion = request.querystring("opcion")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  VOPT = request.querystring("VOPT")

  IF VOPT = "SI" THEN

     VKEY = request.querystring("VKEY")
     DATO = request.FORM("xsi")

     strsql = "UPDATE SYSDEFDSKTOPDET SET PRTY1 = '" & DATO & "' WHERE OBJID = " & VKEY
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    ''' RESPONSE.WRITE STRSQL & "<BR>"

  END IF

    sImagen4 = "<B><FONT FACE=WingDINGS COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"

  sHTML = ofv.FormHeader(session("FormN"))

 '' sHTML = replace(ofv.FormHeader(session("FormN")),"<body ","<body onload='uno();' ")  & chr(13)
  shtml = shtml & "<script  language=javascript>" & chr(13)
  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function uno() {  " & chr(13)
  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - 30; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)


  set ofr = server.createobject("corscr.Corrender")
  ofr.verfonts

  sxopc1A = ""
  sxopc1C = ""
  redim preserve aFontName(0)
  aFontName(0) = ""
  for ii=0 to ofr.wscreen.fontcount -1
      redim preserve aFontName(ii + 1)
      aFontName(ii + 1) = UCASE(ofr.wscreen.fonts(ii))
      sxopc1A = sxopc1A & sxopc1C & ucase(ofr.wscreen.fonts(ii))
      sxopc1C = chr(9) 
  next

'Font Size
  dim aFontSize(13)
  aFontSize(0) = ""
  aFontSize(1) = "7"
  aFontSize(2) = "8"
  aFontSize(3) = "9"
  aFontSize(4) = "10"
  aFontSize(5) = "11"
  aFontSize(6) = "12"
  aFontSize(7) = "14"
  aFontSize(8) = "16"
  aFontSize(9) = "18"
  aFontSize(10) = "20"
  aFontSize(11) = "22"
  aFontSize(12) = "24"
  aFontSize(13) = "36"

  sxopc2A = ""
  sxopc2C = ""
for ii=1 to ubound(aFontSize)
  sxopc2A = sxopc2A & sxopc2C & aFontSize(ii) 
  sxopc2C = chr(9) 
next


  'Encabezados
  Formulario = ofv.Encabezados("ITEM","SET","MUESTRA")

  SIMAGE1 = "<B><FONT FACE=WINGDINGS SIZE=2 COLOR=#005500>&#0232;</FONT></B>"
  SIMAGE2 = "<B><FONT FACE=WINGDINGS SIZE=1 COLOR=#AA0000>l</FONT></B>"

  XSUBFAM    = "NO"
  XSUBFAMTXT = ""
  XSUBFAMCNT = 0


  Strsql = ""

  Strsql = Strsql & " SELECT A.OBJID AS OBJID, A.PRTCOD AS PRTCOD, A.PRTY1 AS PRTY1, B.PRTY1 AS MS1, "
  Strsql = Strsql & " B.PRTY2 AS MS2, B.PRTY3 AS MS3, B.PRTY4 AS MS4, B.PRTY5 AS MS5, B.PRTY6 AS MS6 "
  Strsql = Strsql & " FROM SYSDEFDSKTOPDET A "
  Strsql = Strsql & " INNER JOIN SYSDEFDSKTOPDET B "
  Strsql = Strsql & "       ON A.PRTCOD = B.PRTCOD AND B.CODSET = 0 "
  Strsql = Strsql & " WHERE A.CODSET = " & SESSION("VSET") & " AND B.PRTY1 = '" & CODFAM & "' and b.prty5 = 'SI' "
  Strsql = Strsql & " ORDER BY B.PRTY6, A.OBJID "

''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    IF RS(8) <> XSUBFAM THEN
       XSUBFAM = RS(8)
       XSUBFAMTXT = XSUBFAM
       XSUBFAMCNT = XSUBFAMCNT + 1

       IF XSUBFAMTXT = "" THEN XSUBFAMTXT = "PARAMETROS GENERALES"

       Celda = ofv.GenCelda("","VT COLSPAN=4 STYLE='BORDER-BOTTOM-STYLE:RIDGE;BORDER-BOTTOM-WIDTH:2;BORDER-BOTTOM-COLOR:#F7AA22;' VALIGN=BOTTOM  ","","","","","",XSUBFAMTXT,"si")

       Fila = ofv.genrow("","","","","35","","",Celda,"si")
 
       Formulario = Formulario & Fila

    END IF

      Celda = ofv.GenCelda("","","","1","","","","&nbsp;","si")


   'Descripcion

    Celda = Celda & ofv.gencelda("","VT","","50%","","","",SIMAGE2 & "&nbsp;" & rs(5),"si")

   'SET


    VFORMAT = RS(6)

''  CantRegMoverx = cdbl(request.querystring("CantRegMover"))


        SELECT CASE VFORMAT
               CASE "COLOR" 



                     sInput  = "<A CLASS=NT HREF='CHGCOLOR.ASP?VCOLOR=" & REPLACE(RS(2),"#","") 
                     sInput  = sInput & "&VVUELT=DEFESCRIDETEXTDET.ASP?CantRegMover=" & CantRegMoverx
                     sInput  = sInput & "_VKEY=" & RS(0) & "_VOPT=SI"
                     sInput  = sInput & "' TITLE='MODIFICAR COLOR DE FONDO O BASE' >" & RS(2) & "</A>"

                     sInput2 = "<DIV STYLE='BORDER-STYLE:OUTSET;BORDER-WIDTH:1;BACKGROUND-COLOR:" & RS(2) & "' >&nbsp;</DIV>"

               CASE "COLORFT" 

                     sInput  = "<A CLASS=NT HREF='CHGCOLOR.ASP?VCOLOR=" & REPLACE(RS(2),"#","") 
                     sInput  = sInput & "&VVUELT=DEFESCRIDETEXTDET.ASP?CantRegMover=" & CantRegMoverx
                     sInput  = sInput & "_VKEY=" & RS(0) & "_VOPT=SI"
                     sInput  = sInput & "' TITLE='MODIFICAR COLOR DE FONT' >" & RS(2) & "</A>"

                     sInput2 = "<DIV STYLE='BORDER-STYLE:OUTSET;BORDER-WIDTH:1;BACKGROUND-COLOR:" & RS(2) & "' >&nbsp;</DIV>"

               CASE "SIZE" 

                     SINPUT = "<font class=at><select name=" & RS(2).NAME
                     SINPUT = SINPUT & " onchange=document." & sForm & ".submit();>"
                     For n = 0 To Ubound(aFontSize)
                         sSelect = ""
                         If aFontSize(n) = RS(2) Then sSelect = " selected"
                         SINPUT = SINPUT & "<option value='" & aFontSize(n) & "' " & sSelect & ">"
                         SINPUT = SINPUT & aFontSize(n) & "</option>"
                     Next
                     SINPUT = SINPUT & "</select></font>"

                     SINPUT = ofv.gencomboFX(sForm,rs(2).name,rs(2),sxopc2A,sxopc2A,"sb0")

                     sInput2 = "&nbsp;"

               CASE "FONT" 

                     SINPUT = "<font class=at><select name=" & RS(2).NAME
                     SINPUT = SINPUT & " onchange=document." & sForm & ".submit();>"
                     For n = 0 To Ubound(aFontName)
                         sSelect = ""
                         If aFontName(n) = RS(2) Then sSelect = " selected"
                         SINPUT = SINPUT & "<option value='" & aFontName(n) & "' " & sSelect & ">"
                         SINPUT = SINPUT & aFontName(n) & "</option>"
                     Next
                     SINPUT = SINPUT & "</select></font>"


                     SINPUT = ofv.gencomboFX(sForm,rs(2).name,rs(2),sxopc1A,sxopc1A,"sb")

                     sInput2 = "<B><FONT FACE='" & RS(2) & "' SIZE=3 COLOR=#000080 >Test</FONT></B>"

               CASE "IMAGE" 

                     sInput = ofv.GenerarCombo3(sForm,ofv.Getfiles("../COREss20/CorFondos"),RS(2).NAME,RS(2))

                     sinput = replace(sinput,"value=nsnc","value=NO")

                     if rs(2) = "NO" then 
                        sInput2 = "&nbsp;"
                     else
                        sInput2 = "<img  src=../COREss20/CorFondos/" & rs(2) & ">"
                     end if

               CASE "BOLD" 

                     sInput = "<font class=at><select onchange=document." & sForm & ".submit(); "
                     sInput = sInput & " name=" & rs(2).name & " ><option value=1>SI  "
                     if  rs(2) = 1 then
                         sInput = sInput & "<option value=0>NO  "
                     else
                         sInput = sInput & "<option value=0 selected>NO  "
                     end if
                     sInput = sInput & "</select></font>"

                   
                     SINPUT = ofv.combosino(SForm,rs(2),rs(2).name)

                     sInput2 = "&nbsp;"

               CASE "ITAL" 

                     sInput = "<font class=at><select onchange=document." & sForm & ".submit(); "
                     sInput = sInput & " name=" & rs(2).name & " ><option value=1>SI  "
                     if  rs(2) = 1 then
                         sInput = sInput & "<option value=0>NO  "
                     else
                         sInput = sInput & "<option value=0 selected>NO  "
                     end if
                     sInput = sInput & "</select></font>"

                     SINPUT = ofv.combosino(SForm,rs(2),rs(2).name)

                     sInput2 = "&nbsp;"

               CASE "SUBR" 

                     sInput = "<font class=at><select onchange=document." & sForm & ".submit(); "
                     sInput = sInput & " name=" & rs(2).name & " ><option value=1>SI  "
                     if  rs(2) = 1 then
                         sInput = sInput & "<option value=0>NO  "
                     else
                         sInput = sInput & "<option value=0 selected>NO  "
                     end if
                     sInput = sInput & "</select></font>"

                     SINPUT = ofv.combosino(SForm,rs(2),rs(2).name)

                     sInput2 = "&nbsp;"

        END SELECT 




    Celda = Celda & ofv.gencelda("","XX","","30%","","","",SINPUT,"si")



    Celda = Celda & ofv.gencelda("","XX ","","5%","","","",SINPUT2,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
     SACCION = "asp/actualizard.asp?fname=" 
     SACCION = SACCION & rs(0) & "&tabla=SYSDEFDSKTOPDET&volver=../DEFESCRIDETEXTDET.asp"
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop


  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa  ","","90%","","0","","0","0","","",filas,"si")

  'Botonera
   Agregar = ""
   Buscar = ""
   Retornar = ""

     sAccion = "DEFESCRIDETEXT.asp"
     retornar = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

Response.Write sHTML

%>