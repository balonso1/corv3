<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  set cn = ofv.conectar(ofv.strconn4)

  sHTML = ofv.FormHeader("CASOS DE CONTINGENCIA")

  'Encabezados
  Formulario = ofv.Encabezados("Cod","Descripcion","Asociacion")

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Strsql = "Select * from Tipocont"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  n=0
  Redim Preserve Clave(n)
  If rs.eof then Clave(0) = ""
  Do until rs.Eof 
    Redim Preserve Clave(n)
    Clave(n) = rs(0)
    rs.Movenext
    n = n + 1
  Loop
  call ofv.cerrarconsulta(rs)

  sHTML = sHTML & ofv.ValClaveRep(Clave)

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  Strsql = "Select * From Asociacion Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  n = 0
  If rs.Eof then 
    sHTML = sHTML & "<p align=center><font face=tahoma size=1 color=#990000><b>"
    sHTML = sHTML & "NO PUEDE CARGAR * CONTINGENCIAS * HASTA QUE NO HALLA CARGADO POR LO MENOS "
    sHTML = sHTML & "UNA ASOCIACION</b><br><br></FONT></P>"
  Else  
    Strsql = "Select * from Tipocont"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    if rs.eof then
      sHTML = sHTML & "<p align=center><font face=tahoma size=1 color=#990000><b>"
      sHTML = sHTML & "NO HAY ELEMENTOS ASOCIADOS</b><br><br></FONT></P>"
    else 
      rs.move(CantRegMoverx)
      posI=0
      Do While Not rs.EOF and posi < CantRegMostrarx
        sForm= "form" & posi

        'Eliminar
        if bSuprimir then
          sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
          sHref = sHref & "&Tabla=Tipocont&volver=" & vuelta & "&Eliminar=NSNC"
          Celda = ofv.BotonEliminar(rs(1),sHref)
        else
          Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
        end if

        'Codigo
        sInput = ofv.GenerarInput(rs(0).name,rs(0),"text",sForm,"5","dt MAXLENGTH=5","SIR")
        Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

        'Descripcion
        sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"70","dt MAXLENGTH=50","")
        Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

        'Asociacion
        if bModificar then
          Strsql = "Select * From Asociacion Order By descri"
          sID = "document." & sForm
          sInput = ofv.GenerarCombo(sID,rs(2).name,rs(2),Strsql,cn,0,1,"")
        else
          sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"80","dt MAXLENGTH=50","")
        end if
        Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

        Celda = Celda & ofv.gencelda("","","","","","","","","si")
        Fila = ofv.genrow("","","","","","","",Celda,"si")

        sAccion = ""
        if bModificar then
          sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
          sAccion = sAccion & "&tabla=Tipocont&volver=" & vuelta
        end if
        Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

        rs.Movenext
        posi=posi+1
      Loop
      sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
    End If

    'Botonera
    Agregar = ""
    if bAgregar then
      sAccion = "AgregarContingencias.asp?Regisvuelta=" & CantRegMoverx
      sAccion = sAccion & "&Opcion=si&tabla=Tipocont"
      Agregar = ofv.BotonAgregar(sAccion)
    end if

    str1 = "Select descri From Tipocont "    ' orden pagina
    str2 = "Select descri From Tipocont Order By descri "  ' orden alfabetico
    Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

    sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

    call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)
  End If

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>