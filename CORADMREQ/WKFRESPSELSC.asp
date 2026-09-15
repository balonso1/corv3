<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0



  sHTML = ofv.MENUHeader("","")

  CantRegAMover = cdbl(request.querystring("CantRegAMover"))
  CantRegMover = cdbl(request.querystring("CantRegMover"))
  if CantRegMover > 0 then CantRegAMover = CantRegMover
  CantRegAMostrar = 10
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta2 = ofv.vueltaaspEx("",CantRegAMostrar,CantRegAMover)
  OpcAgregar = request.querystring("OpcAgregar")



  ofv.ObtenerAtributos request.querystring("seg"),""
  Session("Form") = request.querystring("seg")

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  set cnX = ofv.conectar(ofv.strconn5)

  StrSql = "Select * From WKFresp where id = " & session("WKFRESP")
  set rsX = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
  if not rsX.eof then


  SELECT CASE RSX(2)



         CASE "GRP"

'''''' GRUPOS ''''''''''''''''''''''  

  set cn = ofv.conectar(ofv.strconn0)

    if CantRegAMover > 0 then
       NPagina =  Cint((CantRegAMover / CantRegAMostrar)) + 1
    else
       NPagina = 1
    end if   

  XIN = ""

  StrSql = "SELECT DISTINCT grupo FROM USUARIOS"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     DO WHILE NOT RS.EOF
        XIN = XIN & "'" & RS(0) & "',"
        RS.MOVENEXT
     LOOP
     IF LEN(XIN) > 0 THEN XIN = LEFT(XIN,LEN(XIN) - 1)
  END IF


    celdas = ofv.GenCelda("","tt","","","","","","Lista de GRUPOS","si")
    celdas = celdas & ofv.GenCelda("","tt","right","","","","","Pagina " & NPagina,"si")
    filas =  ofv.GenRow("","","","","10","","",celdas,"si")


  StrSql = "Select ID,Descripcion From GRUPOS where ID IN (" & XIN & ") ORDER BY DESCRIPCION "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegAMover)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegAMostrar)
    nID = rs(0)

       sHref = "WKFRESPSELHC.asp?ACT=SI&DATO=" & rs(0)
       sLink = ofv.GenLink("","cc",sHref,"TOPE","","","<b><font size=2>" & rs(1) & "<font></b>","si")
       celdas = ofv.GenCelda("","vt","","80%","","","",sLink,"si")

       sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & rs(0) & "</b>","si")
       celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sLink,"si")

    filas =  filas & ofv.GenRow("","","","","10","","",celdas,"si")

  posi = posi + 1
  rs.movenext
  loop

    sHTML = sHTML & "<BR>"
    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","100%","","0","","-1","-1","","",filas,"si")




    Buscar = ""
      restval = ofv.GenCelda("","","","","","","","&nbsp;","si")
      sHTML = sHTML & ofv.botoneraESSX("a",rs,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no","no",buscar,"")



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



'''''' FIN GRUPOS ''''''''''''''''''''''  

         CASE "SEC"

'''''' SECTORES ''''''''''''''''''''''  

  set cn = ofv.conectar(ofv.strconn0)

    if CantRegAMover > 0 then
       NPagina =  Cint((CantRegAMover / CantRegAMostrar)) + 1
    else
       NPagina = 1
    end if   

  XIN = ""

  StrSql = "SELECT DISTINCT SECTOR FROM USUARIOS WHERE SECTOR <> '' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     DO WHILE NOT RS.EOF
        XIN = XIN & "'" & RS(0) & "',"
        RS.MOVENEXT
     LOOP
     IF LEN(XIN) > 0 THEN XIN = LEFT(XIN,LEN(XIN) - 1)
  END IF

  set cn = ofv.conectar(ofv.strconn2)

    celdas = ofv.GenCelda("","tt","","","","","","Lista de SECTORES","si")
    celdas = celdas & ofv.GenCelda("","tt","right","","","","","Pagina " & NPagina,"si")
    filas =  ofv.GenRow("","","","","10","","",celdas,"si")


  StrSql = "Select ECO,Descri From ECO where ECO IN (" & XIN & ") ORDER BY DESCRI "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegAMover)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegAMostrar)
    nID = rs(0)

       sHref = "WKFRESPSELHC.asp?ACT=SI&DATO=" & rs(0)
       sLink = ofv.GenLink("","cc",sHref,"TOPE","","","<b><font size=2>" & rs(1) & "<font></b>","si")
       celdas = ofv.GenCelda("","vt","","80%","","","",sLink,"si")

       sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & rs(0) & "</b>","si")
       celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sLink,"si")

    filas =  filas & ofv.GenRow("","","","","10","","",celdas,"si")

  posi = posi + 1
  rs.movenext
  loop

    sHTML = sHTML & "<BR>"
    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","100%","","0","","-1","-1","","",filas,"si")




    Buscar = ""
      restval = ofv.GenCelda("","","","","","","","&nbsp;","si")
      sHTML = sHTML & ofv.botoneraESSX("a",rs,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no","no",buscar,"")



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'''''' FIN SECTORES ''''''''''''''''''''''  


         CASE "USR"

'''''' USUARIOS ''''''''''''''''''''''  

  set cn = ofv.conectar(ofv.strconn0)

  Sopcion = request.querystring("Sopcion")

  IF Session("Sletra") = "" THEN Session("Sletra") = "A"

  if Sopcion = "A" then
      SBuscar = request.querystring("SLetra")
      Session("Sletra") = SBuscar
  elseif Sopcion = "I" then
      SBuscar = request.form("alfa")
      Session("Sletra") = SBuscar
  else
      SBuscar = Session("Sletra")
  end if

  if sbuscar <> "" then sbuscar = replace(sbuscar,"'","")    

    if CantRegAMover > 0 then
       NPagina =  Cint((CantRegAMover / CantRegAMostrar)) + 1
    else
       NPagina = 1
    end if   

  redim nivel(27)
  nivel(0) = "A"
  nivel(1) = "B"
  nivel(2) = "C"
  nivel(3) = "D"
  nivel(4) = "E"
  nivel(5) = "F"
  nivel(6) = "G"
  nivel(7) = "H"
  nivel(8) = "I"
  nivel(9) = "J"
  nivel(10) = "K"
  nivel(11) = "L"
  nivel(12) = "M"
  nivel(13) = "N"
  nivel(14) = "Ñ"
  nivel(15) = "O"
  nivel(16) = "P"
  nivel(17) = "Q"
  nivel(18) = "R"
  nivel(19) = "S"
  nivel(20) = "T"
  nivel(21) = "U"
  nivel(22) = "V"
  nivel(23) = "W"
  nivel(24) = "X"
  nivel(25) = "Y"
  nivel(26) = "Z"



    sInput = "<font class=at><select  name=alfa >"
    For i=0 to 26
      if nivel(i) = Ucase(SBuscar) then
        sInput = sInput & "<option value=" & nivel(i) & " selected>" & nivel(i)
      else
        sInput = sInput & "<option value=" & nivel(i) & ">" & nivel(i) 
      end if
    next
    sInput = sInput & "</select></font>"
    Celda =  ofv.gencelda("","","center","","","","",sInput,"si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sInput = "<input  type=SUBMIT  size=10%  value='Buscar' class=bt >"
    Celda =  ofv.gencelda("","","center","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "WKFRESPSELSC.asp?Sopcion=I"
    Formulario =  ofv.genform("tab1","",sAccion,"","post","",Fila,"si")


    sInput = ofv.GenerarInput("alfa",sbuscar,"agregar","","70","dt","")
    Celda =  ofv.gencelda("","","center","","","","",sInput,"si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sInput = "<input  type=SUBMIT  size=10%  value='Buscar' class=bt >"
    Celda =  ofv.gencelda("","","center","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "WKFRESPSELSC.asp?Sopcion=I"
    Formulario2 =  ofv.genform("tab2","",sAccion,"","post","",Fila,"si")


    TABLA1 = ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","100%","","0","","0","0","","",Formulario,"si")
    TABLA2 = ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","100%","","0","","0","0","","",Formulario2,"si")

    celdas = ofv.GenCelda("","cc","center","40%","","","",TABLA1,"si")
    celdas = celdas & ofv.GenCelda("","cc","center","60%","","","",TABLA2,"si")
    filas =  ofv.GenRow("","","","","10","","",celdas,"si")

 '   sHTML = sHTML & "<center>"

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","100%","","0","","0","0","","",filas,"si")
     

    celdas = ofv.GenCelda("","tt","","","","","","Lista de USUARIOS","si")
    celdas = celdas & ofv.GenCelda("","tt","right","","","","","Pagina " & NPagina,"si")
    filas =  ofv.GenRow("","","","","10","","",celdas,"si")


  StrSql = "Select ID,Descripcion From USUARIOS where Descripcion like '" & SBuscar & "%' Order By Descripcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegAMover)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegAMostrar)
    nID = rs(0)

       sHref = "WKFRESPSELHC.asp?ACT=SI&DATO=" & rs(0)
       sLink = ofv.GenLink("","cc",sHref,"TOPE","","","<b><font size=2>" & rs(1) & "<font></b>","si")
       celdas = ofv.GenCelda("","vt","","80%","","","",sLink,"si")

       sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & rs(0) & "</b>","si")
       celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sLink,"si")

    filas =  filas & ofv.GenRow("","","","","10","","",celdas,"si")

  posi = posi + 1
  rs.movenext
  loop

    sHTML = sHTML & "<BR>"
    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","100%","","0","","-1","-1","","",filas,"si")




    Buscar = ""
      restval = ofv.GenCelda("","","","","","","","&nbsp;","si")
      sHTML = sHTML & ofv.botoneraESSX("a",rs,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no","no",buscar,"")



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'''''' FIN USUARIOS ''''''''''''''''''''''  

  END SELECT

  END IF
  call ofv.cerrarconsulta(rsX)
  call ofv.cerrarconn(cnX)



  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>