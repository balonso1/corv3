<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

 ''' RESPONSE.WRITE REQUEST.QUERYSTRING
  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))


  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


 on error resume next
 ACT = request.querystring("ACT")
 DATO = request.querystring("DATO")

 if ACT = "SI" then

    strsql = "update WKFRESP set CODIGO = '" & DATO & "' WHERE ID = " & session("WKFRESP")
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

 end if


  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & session("WKF") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     IF RS(3) THEN
        bAgregar = FALSE
        bModificar = FALSE
        bSuprimir = FALSE
     END IF

    sInput = "CIRCUITO:"
    celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

    sInput = RS(1) & " - " & RS(2)
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"

  END IF


    celdas = ""



  'Encabezados
   filas = ofv.Encabezados("TIPO","RESPONSABLE")


  StrSql = "Select * From WKFresp where id = " & session("WKFRESP")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then


    celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")



        if rs(2) = "USR" then sInput = "USUARIO"
        if rs(2) = "GRP" then sInput = "GRUPO"
        if rs(2) = "SEC" then sInput = "SECTOR"
        sInput = ofv.GenerarInput("",sInput,"readonly","","20","dt","")


       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    XRESPO = RS(3) & ""

    

    if XRESPO = "" then
       sInput = "--- NO ASIGNADO ---------------------"
    else
       sInput = "S/D"
        if rs(2) = "USR" then
           strsql = "Select * From USUARIOS where ID = '" & rs(3) & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
        end if
        if rs(2) = "GRP" then
           strsql = "Select * From GRUPOS where ID = '" & rs(3) & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
        end if
        if rs(2) = "SEC" then
           strsql = "Select * From ECO where ECO = " & rs(3)
           set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
        end if

    end if




       sInput = ofv.GenerarInput("",sInput,"readonly","","70","dt","")
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")




    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

 
  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")


  'Botonera
  Agregar = ""


  Buscar = ""
  
  sinput = "<a CLASS=TT href=WKFresp.asp"
  sinput = sinput  & " TARGET=_parent ><b>Volver&nbsp;a&nbsp;CIRCUITO</b></a>" 
  Buscar = ofv.GenCelda("","TT","","","","","",sInput,"si")

       celdas = Buscar

       celdas = celdas & ofv.GenCelda("","","","90%","","","","&nbsp;","si")


    filas = ofv.GenRow("","","","","","","",celdas,"si")

 
  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")
  


  END IF

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENWKFRESP()



    strsql = "insert into WKFRESP (WKFCOD,TIPO) values ('" & WKF & "','USR')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>