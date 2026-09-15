<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0



  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  EMP = request.querystring("EMP")

  IF EMP <> "" then
     session("EMP") = EMP
  else
     EMP = session("EMP")
  end if

  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))


  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From SYSDEFDSKTOP where OBJUN = " & EMP
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)  
  IF RS.EOF THEN

        strsql = "Insert into SYSDEFDSKTOP (OBJUN,DEFSET,DSCSET,USRGRP,TIPO) "
        strsql = STRSQL & " values (" & EMP & ",'SI','ESCRITORIO DEFAULT','NO','NO')"
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        Strsql = "Select * From SYSDEFDSKTOP where OBJUN = " & EMP
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)  
        IF NOT RS.EOF THEN

           strsql = "UPDATE Clienteplan SET CODSET = " & RS(0) & " where ID = " & EMP
           call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        END IF

  END IF
     call ofv.cerrarconsulta(rs)


  Strsql = "Select * From Clienteplan where ID = " & EMP
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)  



  sForm = "form" & rs(0)
''''''''''''''''''''''''''

    sInput = "DATOS DE LA EMPRESA"
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Nombre
     Celda = ofv.gencelda("","tt","","20%","","","","Nombre de la entidad","si")

     sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"50","dt","")
     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Perfil
     Celda = ofv.gencelda("","tt","","20%","","","","Con logo?","si")

     sInput = ofv.combosinoEx(sform,(rs(2)="S"),rs(2).name,"S","N")
     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Perfil
     Celda = ofv.gencelda("","tt","","20%","","","","Archivo Imagen","si")

     sImagenValor = rs(5)
     sImagenCampo = rs(5).name

     sInput = ofv.GenerarCombo3(sform,ofv.Getfiles("../COREss20/Empresa"),sImagenCampo,sImagenValor)
     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Perfil
     Celda = ofv.gencelda("","tt","","20%","","","","Logo de la entidad","si")

     sInput = "<img name='" & rs(3).name & "' src='../COREss20/Empresa/" & rs(5) & "' >"
     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Perfil
     Celda = ofv.gencelda("","tt","","20%","","","","Archivo Adicional","si")

     sImagenValor = rs(11) & ""
     sImagenCampo = rs(11).name

     sInput = ofv.GenerarCombo3(sform,ofv.Getfiles("../COREss20/Empresa"),sImagenCampo,sImagenValor)
     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    if simagenvalor <> "" then 

    'Perfil
     Celda = ofv.gencelda("","tt","","20%","","","","Logo Adicional","si")

     sInput = "<img name='" & rs(3).name & "adc' src='../COREss20/Empresa/" & rs(11) & "' >"
     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    end if



    'Perfil
     Celda = ofv.gencelda("","tt","","20%","","","","Tipo","si")

     xtipo = "UNIDAD DE NEGOCIOS"
     if not isnull(rs(6)) then
        if rs(6) = "1" then xtipo = "GRUPO"
     end if

     sInput = xtipo
     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Perfil
     Celda = ofv.gencelda("","tt","","20%","","","","Configuracion de Escritorio","si")

     if bModificar then
        Strsql = "Select * from SYSDEFDSKTOP WHERE OBJUN =  " & EMP & " ORDER BY OBJID "
        sInput = ofv.GenerarCombo(sform,rs(10).name,rs(10).value,Strsql,cn,0,3,"SB")
     else
        sInput = ofv.GenerarInput("",rs(10),"readonly","","10","dt","")
     end if

     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


'''''''''''''''''''''''''



  IF SESSION("EMPMUN") = "S" AND RS(0) = 1 THEN

    sInput = "<br><br>CONFIGURACION MULTIEMPRESA"
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Nombre
     Celda = ofv.gencelda("","tt","","20%","","","","Mostrar Logo Y Nombre","si")

    if bModificar then

       scomK = "1" & chr(9) & "2"
       scomD = "DEL GRUPO" & chr(9) & "DE LA EMPRESA DEL USUARIO"
       sInput = ofv.gencomboFX(sForm,rs(8).name,rs(8),scomK,scomD," ")

    else
      if rs(8) = 1 then
         xestado = "DEL GRUPO"
      else
         xestado = "DE LA EMPRESA DEL USUARIO"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
    end if

     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Nombre
     Celda = ofv.gencelda("","tt","","20%","","","","(Consultas)&nbsp;Ver&nbsp;Documentos&nbsp;en&nbsp;","si")

      if bModificar then

         scomK = "1" & chr(9) & "2"
         scomD = "DATOS DEL GRUPO" & chr(9) & "TODAS LAS CONSULTAS"
         sInput = ofv.gencomboFX(sForm,rs(9).name,rs(9),scomK,scomD," ")

    else
      if rs(9) = 1 then
         xestado = "DATOS DEL GRUPO"
      else
         xestado = "TODAS LAS CONSULTAS"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
    end if


     Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


  END IF

    sAccion = ""
    if bModificar then
       SACCION = "asp/actualizard.asp?fname=" 
       SACCION = SACCION & rs(0) & "&tabla=Clienteplan&volver=../DatosDelClientePROP.asp"
    end if
    Formulario = ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & ofv.GenTabla("","aa  ","","80%","","0","","0","0","","",Formulario,"si")




  'Botonera
   Agregar  = ""
   Buscar   = ""
   Retornar = ""

  IF SESSION("EMPMUN") = "S"  THEN

     sAccion = "DatosdelCliente.asp"
     retornar = ofv.BotonVolver(sAccion)

  end if

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</body></html>"

end if

Response.Write sHTML

%>