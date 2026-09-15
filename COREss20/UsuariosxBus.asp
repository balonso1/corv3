<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0



  sHTML = ofv.FormHeader(session("FormN"))

  CantRegAMover   = cdbl(request.querystring("CantRegAMover"))
  CantRegMover    = cdbl(request.querystring("CantRegMover"))
  if CantRegMover > 0 then CantRegAMover = CantRegMover
  CantRegAMostrar = 8
  vuelta          = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta2         = ofv.vueltaaspEx("",CantRegAMostrar,CantRegAMover)

  OpcBuscar      = request.querystring("OpcBuscar")

  if opcbuscar <> "" then
     session("OPCBUS") =  opcbuscar
  ELSE
     OpcBuscar         = session("OPCBUS")
  END IF


  set cn  = ofv.conectar(ofv.strconn0)
  set cn2 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos request.querystring("seg"),""
  Session("Form") = request.querystring("seg")

  bAgregar   = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir  = ofv.AEliminar
  bConsultar = ofv.AConsultar



  Sopcion = request.querystring("Sopcion")
  Scodobj = request.querystring("codobjeto")
  Stipo   = request.querystring("tipobjeto")

  if sopcion = "" then sopcion = scodobj

  if stipo <> "V" then 

  if Sopcion = "A" then
      SBuscar = request.querystring("SLetra")
      Session("Sletra") = SBuscar
  elseif Sopcion = "I" then
      SBuscar = request.form("alfa")
      Session("Sletra") = SBuscar
  elseif Sopcion = "S" then
      SBuscar = request.form("sector")
      Session("Sletra") = SBuscar
  else
      SBuscar = Session("Sletra")
  end if

  else
      SBuscar = Session("Sletra")
  end if

  if sbuscar <> "" then
     sbuscar = replace(sbuscar,"'","")    
''  else
''     sbuscar = "TODOS"   
  end if

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


''' por alfabeto '''''''''''''''''''''''''''''''''''''''''''''''
    sInput = "seleccion alfabetica"
    Celda =  ofv.gencelda("","ttt","center","","","","",sInput,"si")
    Fila = ofv.genrow("","","","","20","","",Celda,"si")

     SCOMSEP = ""
     scomK   = ""
 
     For i=0 to 26
         scomK = scomK & SCOMSEP & nivel(i)
         SCOMSEP = chr(9)
     next

     scomD = scomK

    sInput = ofv.gencomboFX("tab1","alfa",Ucase(SBuscar),scomK,scomD,"")


    Celda =  ofv.gencelda("","aaa bgcolor=#f9f9f9 ","center","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sInput = "<input  type=SUBMIT  size=10%  value='Buscar' class=bt >"
    Celda =  ofv.gencelda("","aaa bgcolor=#f9f9f9 ","center","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = "UsuariosxBus.asp?Sopcion=I&OpcBuscar=" & OpcBuscar
    Formulario =  ofv.genform("tab1","",sAccion,"","post","",Fila,"si")

''' por texto '''''''''''''''''''''''''''''''''''''''''''''''
    sInput = "seleccion por texto"
    Celda =  ofv.gencelda("","ttt","center","","","","",sInput,"si")
    Fila = ofv.genrow("","","","","20","","",Celda,"si")

    sInput = ofv.GenerarInput("alfa",sbuscar,"agregar","","70","dt","")
    Celda =  ofv.gencelda("","aaa bgcolor=#f9f9f9 ","center","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sInput = "<input  type=SUBMIT  size=10%  value='Buscar' class=bt >"
    Celda =  ofv.gencelda("","aaa bgcolor=#f9f9f9 ","center","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "UsuariosxBus.asp?Sopcion=I&OpcBuscar=" & OpcBuscar
    Formulario2 =  ofv.genform("tab2","",sAccion,"","post","",Fila,"si")

''' por sector '''''''''''''''''''''''''''''''''''''''''''''''
    sInput = "seleccion por sector"
    Celda =  ofv.gencelda("","ttt","center","","","","",sInput,"si")
    Fila = ofv.genrow("","","","","20","","",Celda,"si")

    StrSql = "Select * from eco Order By descri"
    sInput = ofv.GenerarCombo("tab3","sector",cstr(sbuscar),strsql,cn2,0,1,"")


    Celda =  ofv.gencelda("","aaa bgcolor=#f9f9f9 ","center","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sInput = "<input  type=SUBMIT  size=10%  value='Buscar' class=bt >"
    Celda =  ofv.gencelda("","aaa bgcolor=#f9f9f9 ","center","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "UsuariosxBus.asp?Sopcion=S&OpcBuscar=" & OpcBuscar
    Formulario3 =  ofv.genform("tab3","",sAccion,"","post","",Fila,"si")



    TABLA1 = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#DDDDDD;' ","","100%","","0","","-1","-1","","",Formulario,"si")
    TABLA2 = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#DDDDDD;' ","","100%","","0","","-1","-1","","",Formulario2,"si")
    TABLA3 = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#DDDDDD;' ","","100%","","0","","-1","-1","","",Formulario3,"si")

    celdas = ofv.GenCelda("","ut","center","40%","","","",TABLA1,"si")
    celdas = celdas & ofv.GenCelda("","ut","center","60%","","","",TABLA2,"si")
    filas =  ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut COLSPAN=2","center","","","","",TABLA3,"si")
    filas =  FILAS & ofv.GenRow("","","","","10","","",celdas,"si")

    sHTML = sHTML & "<center>"

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:2;border-color:" & session("CMFMENU") & ";' ","","80%","","0","","0","0","","",filas,"si")
     

    sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=2>&#0128;</FONT></B>"

    celdas = ofv.GenCelda("","ttR","","","","","","Lista de usuarios","si")
    celdas = celdas & ofv.GenCelda("","ttR","right","","","","","Pagina " & NPagina,"si")
    filas =  ofv.GenRow("","","","","20","","",celdas,"si")

  if Sopcion = "S" then
     StrSql = "Select id,Descripcion From Usuarios where sector = '" & SBuscar & "' Order By descripcion"
  else
     StrSql = "Select id,Descripcion From Usuarios where descripcion like '" & SBuscar & "%' Order By descripcion"
  end if
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegAMover)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegAMostrar)
    nID = rs(0)

    if  OpcBuscar = "U" then 
       sHref = "UsuariosX.asp?Nusuario=" & rs(0)
    else
       sHref = "habilit.asp?Nusuario=" & rs(0)
    end if

   ''    sHref = "UsuariosX.asp?Nusuario=" & rs(0)
       sLink = ofv.GenLink("","nt",sHref,"","","",sImagen4si & "&nbsp;" & rs(1),"si")
       celdas = ofv.GenCelda("","UT","","80%","","","",sLink,"si")

   ''    sHref = "UsuariosX.asp?Nusuario=" & rs(0)
       sLink = ofv.GenLink("","nt",sHref,"","","",rs(0),"si")
       celdas = celdas & ofv.GenCelda("","UT","","20%","","","",sLink,"si")

    filas =  filas & ofv.GenRow("",grcol,"","","10","","",celdas,"si")

    if  grcol = "aaa bgcolor=#f9f9f9 " then
        grcol = "aaa bgcolor=#ffffff "
    else
        grcol = "aaa bgcolor=#f9f9f9 "
    end if 

  posi = posi + 1
  rs.movenext
  loop

    sHTML = sHTML & "<BR>"
    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:2;border-color:" & session("CMFMENU") & ";' ","","80%","","0","","-1","-1","","",filas,"si")


    sHTML = sHTML & "</center>"


    Buscar = ""
      restval = ofv.GenCelda("","","","","","","","&nbsp;","si")
      sHTML = sHTML & ofv.botoneraESSX("a",rs,CantRegAMover,CantRegAMostrar,"","","",sopcion,"V","no","no",buscar,"")







 

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>