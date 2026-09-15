<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

      Response.expires = 0


      fondo = " bgcolor=" & Session("CMFMENU")


      sHTML = "<html><body " & fondo & "  >"


        if  session("SMPLMENU") <> "S" then

      xsimagenf  = "<font FACE=WingDINGS COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;" 

        end if

      Menu = Request.querystring("Menu")
      If  Menu <> "I" Then
          Submenu  = Request.querystring("ID")
          Menu     = OFV.convertircar(Menu, "_", " ")
          sMenu    = "Menu " & Menu
          Session("cormenu") = Menu
  
          Fila = ""

        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;"
        else     
            styl = XspEx("#888888","#cccccc",0) & ";"
        end if

          xlfun = " onclick=""despmenu(); "" "

        if  session("SMPLMENU") <> "S" then


          sImagenx3  = "<B><FONT FACE=WebDINGS COLOR='#ffffff' SIZE=1>2</FONT></B>"

          sImagenx3  =   "&nbsp;" & sImagenx3 & "<font color='#ffffff' >" & "&nbsp;" & Session("cormenu") & "</font>"

          Celda =  OFV.gencelda("", "at STYLE='" & styl & ";border-style:solid;border-color:#ffffff;border-width:1;padding-left:3px;' title='Ocultar solapas.' " & xlfun, "", "", "", "", "", sImagenx3, "si")

          Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")

         end if


          Set cx = OFV.conectar(OFV.strconn0)


          Set cn = OFV.conectar(OFV.strconn1)
            xsel = 1
          StrSql = "Select * From Menux Where mdepen='" & Submenu & "' and mcodapli = '" & Session("CorApli") & "' "

          Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
            If Not rs.EOF Then
               If rs.fields.Count > 6 Then
                  xsel = 6
                  Call OFV.cerrarconsulta(rs)
                  StrSql = "Select * From Menux Where mdepen='" & Submenu & "' and mcodapli = '" & Session("CorApli") & "' order by Morden "

                  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
               Else
                  xsel = 1
               End If
            End If
          Do Until rs.EOF



              atrok = "N"
              StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'M' "
              StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
              StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & Session("CorApli") & "'"

              Set rs1 = OFV.crearconsultaEx(StrSql, cx, 1, "")
              If rs1.EOF Then
                 atrok = "O"
              ElseIf rs1(1) Then
                     atrok = "S"
                  Else
                     atrok = "N"
              End If
              Call OFV.cerrarconsulta(rs1)

              xcount  = 0

              If atrok = "O" Then

                 StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'M' "
                 StrSql = StrSql & " And alias = '" & rs(xsel) & "' And  "

                 if  Session("DELLOG") = "S" then

                     XMFTGRP = split(session("SegGrpLst"),chr(9))
                     xcount  = ubound(XMFTGRP) + 1
                     XMFSEP  = "" 
                     StrSql = StrSql & " grupo in ("

                     for XMFI = 0 to ubound(XMFTGRP)
                         StrSql  = StrSql  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                         XMFSEP  = "," 
                     next

                     StrSql = StrSql & ") "

                 else
                     StrSql = StrSql & " grupo = '" & Session("seggrp") & "' "
                 end if 


                 StrSql = StrSql & " And Aplicacion = '" & Session("CorApli") & "'"
                 Set rs2 = OFV.crearconsultaEx(StrSql, cx, 1, "")
                 If  rs2.EOF Then
                        atrok = "O"
                 Else
                     if  Session("DELLOG") = "S" then
                         XMFSUM = 0 
                         XMFSUMN = 0 
                         do until rs2.eof
                            XMFSUM = XMFSUM + 1  
                            If rs2(1) Then
                               atrok = "S"
                               exit do
                            Else
                               atrok   = "N"
                               XMFSUMN = XMFSUMN + 1
                            End If
                         rs2.movenext
                         loop 
                       
                         if  atrok = "N" then 
                             if  xcount >  XMFSUMN then atrok = "S"
                         end if 

                     else 
                         If rs2(1) Then
                            atrok = "S"
                         Else
                            atrok = "N"
                         End If
                     end if 
                 End If
                 Call OFV.cerrarconsulta(rs2)

              End If


              If atrok = "O" Or atrok = "S" Then


                          sInput = Trim(rs(xsel))
                          sHref = Trim(rs(2))
                              sTarget = "info"
                              sPath = Session("CorPath")
                              If rs(3) <> "" Then sTarget = rs(3)
                              spt = InStr(rs(2), ".asp")
                              xesm = "N"
                              If spt > 1 Then
                                  If Left(rs(2), spt - 1) = "Submenus" Then
                                     sPath = "../coress20/"
                                     xesm = "S"
                                  End If
                              End If
                              sseg = InStr(sHref, "?")
                              If sseg = 0 Then
                                  sHref = sHref & "?"
                              Else
                                  sHref = sHref & "&"
                              End If
                              sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rs(1)), " ", "_")

                           if  session("SMPLMENU") <> "S" then

                              If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then

                                 sInput  = UCase(rs(1))
                                 xsclase = "nt STYLE='border-bottom-style:solid;border-bottom-color:#dddddd;border-bottom-width:1;padding-left:10px;' "

                              else

                                 sInput  = UCase(rs(1))
                                 xsclase = "nt STYLE='border-bottom-style:solid;border-bottom-color:#dddddd;border-bottom-width:1;padding-left:10px;' "

                              End If


                              sLink = xsimagenf & OFV.GenLink("", "xts", sPath & sHref, sTarget, "", "", sInput, "si")
                              Celda = OFV.gencelda("", xsclase, "", "", "", "", "", sLink, "si")
                              Fila = Fila & OFV.genrow("", "", "", "", "20", "", "", Celda, "si")

                           end if


                  End If

                          


              rs.movenext
          Loop
          Call OFV.cerrarconsulta(rs)
          Call OFV.cerrarconn(cn)
          Call OFV.cerrarconn(cx)

      if  session("SMPLMENU") <> "S" then


        If ofv.ModTitCls = "1" Then

           sHTML = sHTML & OFV.GenTabla("DBSL", "aaa ", "", "100%", "", "0", "", "-2", "-2", "", "", Fila, "si")

        else

           sHTML = sHTML & OFV.GenTabla("DBSL", "", "", "100%", "", "0", "", "-2", "-2", "", "", Fila, "si")

        end if

      end if 




      Else


        if  session("SMPLMENU") <> "S" then

          sHTML = sHTML & "<center><br><br>"


         
          sHTML = sHTML & "</center>"

        end if

      End If

      Fila = ""

      sInput = OFV.GenerarInput("boton", "", "hidden", "", "1", "it", "")
      Formulario = OFV.GenForm("nulo", " aa style='display:none;' ", "nulo.asp", "info", "post", "", sInput, "si")
      Formulario = Formulario & "<script language='javascript' > document.nulo.submit(); </script>"  

      sHTML = sHTML & Formulario


     shtml = shtml  & "</body></html>"



Response.Write sHTML

%>
