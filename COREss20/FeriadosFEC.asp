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

  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  cniv = request.querystring("Cniv")
  cnom = request.querystring("Cnom")
  cx = request.querystring("Cx")
  cA = request.querystring("CA")

  if cniv = "" then 
     cniv = session("ENC7")
     cnom = session("ENC8")
     cx = session("ENC9")
     cA = session("ENCA")
  else
     session("ENC7") = cniv
     session("ENC8") = cnom
     cx = cdbl(cx)
     session("ENC9") = cx
     session("ENCA") = cA
  end if


  sHTML = ofv.FormHeader("Fijar Feriado") & "<CENTER><BR>"

  'Encabezados
  Formulario = ""


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn0)

  Strsql = "Select * From feriados where id = " & cniv
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 
  posi = 0
    sForm = "form" & posi
       sclase = "CC" 

     Celda = ofv.GenCelda("","TT","","10%","","","","fECHA:","si")

    'codigo
    'relevamiento
    if clng(rs(1)) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(rs(1)) + clng(cdate("01-01-1900")) )
      ' sXinput=sInput1
    swd = datepart("w",sInput1)

    select case swd
           case 1

                sinput1 = "Domingo, " & sinput1

           case 2
                sinput1 = "Lunes, " & sinput1

           case 3
                sinput1 = "Martes, " & sinput1

           case 4
                sinput1 = "Miercoles, " & sinput1

           case 5
                sinput1 = "Jueves, " & sinput1

           case 6
                sinput1 = "Viernes, " & sinput1

           case else
                sinput1 = "Sabado, " & sinput1

    end select
    end if



     Celda = Celda & ofv.gencelda("","cc colspan=3 ","","40%","","","",sInput1 ,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","35%","","0","","0","0","","",Fila,"si") & "<br>"

    'Descripcion
    CELDA = ""

    sanio = datepart("yyyy",date())
    smes = datepart("m",date())

    SMESC = SMES

    if not isnull(rs(1)) then

       if clng(rs(1)) > 0 then

          smesc = cdate(clng(rs(1)) + clng(cdate("01-01-1900")) )
          smesc = datepart("m",smesc)

       end if 


    end if

    IF CA <> "" THEN 
       smesc = request.form("smesc")
 '      response.write "smesc " & smesc & " ca " & ca
       if isnull(smesc) then  SMESC = ca
       if len(smesc) = 0 then  SMESC = ca
       ca =  SMESC
      
    else
       ca = smesc

    end if  

  '     response.write "smesc " & smesc & " ca " & ca

       session("ENCA") = smesc

    if cint(smesc) = cint(smes) then
       sanioc = sanio
    elseif cint(smesc) < cint(smes) then
       sanioc = sanio + 1
    else
       sanioc = sanio
    end if

    
      Celda = ofv.GenCelda("","TT","","10%","","","","mes:","si")

    'codigo

     Strsql = "Select * From meses Order By mes"
     sInput = ofv.GenerarCombo("document." & sForm & "A","smesc",smesc,Strsql,cn,1,2,"sb")
     Celda = Celda & ofv.gencelda("",sclase,"","40%","","","",sInput,"si")

     Celda = Celda & ofv.GenCelda("","TT","","10%","","","","año:","si")
     
     sinput = sanioc 
     Celda = Celda & ofv.gencelda("",sclase,"","40%","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "FeriadosFEC.asp?cniv=" & cniv & "&cnom=" &  cnom & "&cx=" & cx & "&ca=" & smesc
    end if
    Formulario = ofv.genform(sForm & "A","",sAccion,"","post","",Fila,"si")
    
    

    sHTML = sHTML & ofv.GenTabla("","","","35%","","0","","0","0","","",Formulario,"si") & "<BR>"


    dim afech(33,1)
    dim afechs(42,1)

    CELDA = ""

    sfecds = "01-" & cstr(smesc) & "-" & cstr(sanioc)

    'sfecd = clng(cdate(sfecds))
    sfecd = clng(dateserial(sanioc,smesc,1))

    'swd = datepart("w",date(sfecds))
     swd = datepart("w",dateserial(sanioc,smesc,1))
     
    sfechs = "28-" & cstr(smesc) & "-" & cstr(sanioc)

    'sfech = clng(date(sfechs))
     sfech = clng(dateserial(sanioc,smesc,28))
 '   response.write " fd " & sfecd & " fh " & sfech

    yy = 0

    for ii = 1 to 3
        sfecha = sfech + ii
        sfecha = cdate(clng(sfecha))
        smess = datepart("m",sfecha)
        if cint(smesc) <> cint(smess) then
           exit for
        else
           yy = ii
        end if

    next  

    sfech = sfech + yy

 '  response.write " wd " & swd & " fd " & sfecd & " fh " & sfech
    srest = clng(cdate("01-01-1900"))
   ' response.write  sfech &" -"&sfecd 
    topefech = sfech - sfecd 

    yy = 0



    wswd = swd

    for ii = 0 to topefech
        
        afech(ii+1,0) = wswd
        wswd = wswd + 1
        if wswd > 7 then wswd = 1        

        afech(ii+1,1) = sfecd + ii


    next


    for ii = 1 to (cint(afech(1,0)) - 1)
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    yy = 1

 '   RESPONSE.WRITE "II " & II & " TOPE " & (ii + topefecH ) & " TOPEFECH " & topefecH
    for ii = ii to (ii + topefecH )
        afechs(ii,0) = yy
        afechs(ii,1) = afech(yy,1)
        yy = yy + 1
    next

'    RESPONSE.WRITE "II " & II

    for ii = ii to 42
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    CELDA = ""
    Celda = Celda & ofv.GenCelda("","tt style='border-style:outset;border-width:2;'","CENTER","10%","","","","D","si")
    Celda = Celda & ofv.GenCelda("","tt style='border-style:outset;border-width:2;'","CENTER","10%","","","","L","si")
    Celda = Celda & ofv.GenCelda("","tt style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","tt style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","tt style='border-style:outset;border-width:2;'","CENTER","10%","","","","J","si")
    Celda = Celda & ofv.GenCelda("","tt style='border-style:outset;border-width:2;'","CENTER","10%","","","","V","si")
    Celda = Celda & ofv.GenCelda("","tt style='border-style:outset;border-width:2;'","CENTER","10%","","","","S","si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    CELDA = ""
    yy = 0

    for ii = 1 to 43

        if yy = 7 then
           yy = 1
           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
           CELDA = ""
        else
           yy = yy + 1 
        end if

        if ii > 42 then exit for  

        if yy = 1 or yy = 7 then
           sclase = "tt3 style='border-style:outset;border-width:2;'" 
        else
           sclase = "TT style='border-style:outset;border-width:2;'" 
        end if 

        if afechs(ii,0) = "NO" then
           sinput = "&nbsp;"
           sclase = "CC " 
           Celda = Celda & ofv.GenCelda("","cc","CENTER","10%","","","",sinput,"si")
        else
           sinput1 = afechs(ii,0)

           sinput1 = cstr(sinput1)

           if len(sinput1) = 1 then sinput1 = "0" & sinput1 

           sinput1 = " " & sinput1 & " " 

'           sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
'           sAccion = sAccion & "&tabla=feriados&volver=" & vuelta

            sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=feriados"

           t1 = ofv.GenerarInput("bvo1",sInput1,"submit","","10",sclase,"")
           t1 = t1 & ofv.GenerarInput(rs(1).name,(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","cc","center","10%","","","",t1,"si")
           t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si") 
           Celda = Celda & t1



 


        end if


    next


   sHTML = sHTML & ofv.GenTabla("","","","35%","","0","","0","0","","",FILA,"si") & "<br>" & chr(13)

 
  END IF


  svuelve = "feriados.asp?CantRegMover" & CX

  Retornar = ofv.GenerarInput("bvo1","Volver","submit","","10","bt","")
  Retornar = ofv.GenCelda("","","LEFT","","","","",Retornar,"si")
  Retornar = ofv.GenForm("Retornar","",sVuelve,"","post","",Retornar,"si") 


  Fila = ofv.genrow("","","","","","","",Retornar,"si")

  sHTML = sHTML & ofv.GenTabla("","","","35%","","0","","0","0","","",Fila,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</CENTER></body></html>"
end if

Response.Write sHTML

%>