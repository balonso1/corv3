<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()


     Response.expires = 0

        
              scp = ""
       
        fondo = " bgcolor=" & Session("CMFMENU")


      sHTML = "<html>"

        sHTML = sHTML & "<body width=400 " & fondo
        sHTML = sHTML & " >"

      if session("SMPLMENU") <> "xS" then




        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;"
        else     
            styl = XspEx("#888888","#cccccc",0) & ";"
        end if


        ctablas = 0

        ctabcol = 99



      sOpcion = Trim(Request.querystring("Opcion"))
        Celda = "N"
      If sOpcion = "I" Then
          sOpcion = ""

        If ofv.ModTitCls = "1"  Then

           sHTML = sHTML & OFV.GenTabla("DBML" & ctablas, "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "90%", "", "0", "", "-2", "2", "", "", "", "no")

        else

           sHTML = sHTML & OFV.GenTabla("DBML" & ctablas, "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "90%", "", "0", "", "2", "2", "", "", "", "no")

        end if



         sHTML = sHTML & OFV.genrow("", "", "", "", "", "", "", "", "no")

          sHTML = sHTML & "<td  ><p align=center><font face=" & Session("CMFMEFTFC") & " size=1 color=" & Session("CMFMNFT") & " ><B>Bienvenido a COR ESolution Suite &reg - por favor ingrese su Usuario y Password</b></font></p></td>"
          Session("CorApli") = ""
          Session("CorPath") = ""

         sHTML = sHTML & OFV.genfrow()

      Else
          Set cx = OFV.conectar(OFV.strconn0)

          Celda = "N"
          Session("CorApli") = sOpcion
          sApli = sOpcion
          If sOpcion = "CORESS20" And Session("CorPath") = "" Then
              sOpcion = ""
          Else
              sOpcion = "?Opcion=I"
          End If

          Set cn = OFV.conectar(OFV.strconn1)

            Session("Coraplname") = Session("CorApli")

          If sApli = "xxCORESS20" Then
              Session("CorPath") = ""
          Else
              StrSql = "Select DirApli,titapli From CorAplic where CorAplic.CodApli = '" & sApli & "'"
              Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
              If Not rs.EOF Then
                  Session("CorPath") = "../" & Trim(rs(0)) & "/"
                  Session("CorTitApli") = ucase(rs(1))

              Else
                  Session("CorPath") = ""
                  Session("CorTitApli") = " "

              End If
              Call OFV.cerrarconsulta(rs)
          End If


    '    if  session("SMPLMENU") <> "xS" then

        if  session("SMPLMENU") <> "S" then

          sImagenx3  = "<B><FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1>&#0204;</FONT></B>"

          sImagenx3  =   "&nbsp;" & sImagenx3 & "<font color=" & session("CMFMNBOT") & ">" & "&nbsp;" & Session("CorTitApli") & "</font>"

          celdaxx = OFV.gencelda("", "vt STYLE='" & XspSShw("#aaaaaa","125") & ";' ", "", "", "", "", "", sImagenx3, "si")

        end if

            xsel = 1
            
' xxxxxxxxxxxxxx

               atrok = "N"
              StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'A' "
              StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & sApli & "'"
              Set rs1 = OFV.crearconsultaEx(StrSql, cx, 1, "")
              If rs1.EOF Then
                 atrok = "O"
              ElseIf rs1(1) Then
                     atrok = "S"
                  Else
                     atrok = "N"
              End If
              Call OFV.cerrarconsulta(rs1)

              If atrok = "O" Then

''''''''''''''''''''''''''''''''''


                 StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'A' "
                 StrSql = StrSql & " And  "

                 if  Session("DELLOG") = "S" then

                     XMFTGRP = split(session("SegGrpLst"),chr(9))

                     xcount = ubound(XMFTGRP) + 1
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


                 StrSql = StrSql & " And Aplicacion = '" & sApli & "'"
                 Set rs1 = OFV.crearconsultaEx(StrSql, cN, 1, "")
                 If  rs1.EOF Then
                        atrok = "O"
                 Else
                     if  Session("DELLOG") = "S" then
                         XMFSUM = 0 
                         XMFSUMN = 0 
                         do until rs1.eof
                            XMFSUM = XMFSUM + 1  
                            If rs1(1) Then
                               atrok = "S"
                               exit do
                            Else
                               atrok   = "N"
                               XMFSUMN = XMFSUMN + 1
                            End If
                         rs1.movenext
                         loop 
                       
                         if  atrok = "N" then 
                             if  xcount >  XMFSUMN then atrok = "S"
                         end if 

                     else 
                         If rs1(1) Then
                            atrok = "S"
                         Else
                            atrok = "N"
                         End If
                     end if 
                 End If
                 Call OFV.cerrarconsulta(rs1)


''''''''''''''''''''''''''''''''''





              End If

              If atrok = "O" Or atrok = "S" Then


' xxxxxxxxxxxxx


          StrSql = "Select Menux.* From Menux, CorAplic "
          StrSql = StrSql & " where CorAplic.codapli = Menux.mcodapli and "
          StrSql = StrSql & "CorAplic.CodApli = '" & sApli & "' AND Menux.mdepen = '0'"
          Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
            If Not rs.EOF Then
               If rs.fields.Count > 6 Then
                  xsel = 6
                  Call OFV.cerrarconsulta(rs)

          StrSql = "Select Menux.* From Menux, CorAplic "
          StrSql = StrSql & " where CorAplic.codapli = Menux.mcodapli and "
          StrSql = StrSql & "CorAplic.CodApli = '" & sApli & "' AND Menux.mdepen = '0' order by Menux.Morden"
                  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
               Else
                  xsel = 1
               End If
            End If
          Fila = ""
          cnm = 0
          Do Until rs.EOF
              atrok = "N"
              StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'S' "
              StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
              StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & sApli & "'"

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

                 StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'S' "
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


                 StrSql = StrSql & " And Aplicacion = '" & sApli & "'"
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
                            If Celda = "N" Then
                               Celda = ""
                               sInput = OFV.GenerarInput("boton2", "menu", "hidden", "", "10", "", "")
                               Formulario = OFV.GenForm("initaplx", "aaa STYLE='display:none;' ", "../coress20/Submenus.asp?ID=" & rs(6) & "&Menu=" & Replace(UCase(rs(1)), " ", "_"), "Master", "post", "", sInput, "si")
                               scp = "<script language='javascript' > document.initaplx.submit(); </script>"
                            End If


               ctabcol = ctabcol + 1

               if ctabcol > 4 then

                  if ctabcol < 99 then

                   if  session("SMPLMENU") <> "S" then

                     xsclase = "nt style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:webdings;color:#cc0000;cursor:hand;' "


                     xlfun = " onclick=""document.all.DBML" & (ctablas + 1) & ".style.display='block';document.all.DBML" & (ctablas) & ".style.display='none'; "" "


                     Celda = Celda & OFV.gencelda("BAT" & CTABLAS & "S*", xsclase, "center title='siguiente' " & xlfun, "1", "", "", "", "<b>4</b>", "si")

                     
                     sHTML = sHTML & OFV.genrow("", "", "", "", "20", "", "", celda, "si")

                     sHTML = sHTML & OFV.GenFTabla()

                   end if

                     Celda = ""


                  end if

                  ctabcol = 1

                  ctablas = ctablas + 1


                  if ctablas > 1 then

                    if  session("SMPLMENU") <> "S" then

                     styl2 = "display:none;"

                     xsclase = "nt style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:webdings;color:#cc0000;cursor:hand;' "


                     xlfun = " onclick=""document.all.DBML" & (ctablas - 1) & ".style.display='block';document.all.DBML" & (ctablas) & ".style.display='none'; "" "


                     Celda = OFV.gencelda("BAT" & CTABLAS & "A*", xsclase, "center title='anterior' " & xlfun, "1", "", "", "", "<b>3</b>", "si")

                    end if

                  else

                     styl2 = "display:block;"


                  end if

        if  session("SMPLMENU") <> "S" then


        If ofv.ModTitCls = "1"  Then

           sHTML = sHTML & OFV.GenTabla("DBML" & ctablas, "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & styl2 & "' ", "", "100%", "", "0", "", "-2", "1", "", "", "", "no")

        else

           sHTML = sHTML & OFV.GenTabla("DBML" & ctablas, "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & styl2 & "' ", "", "100%", "", "0", "", "-2", "1", "", "", "", "no")

        end if
 
         end if


               end if 

               if  session("SMPLMENU") <> "S" then


                    If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
                       xGenConti = UCase(rs(1))
                       sInput    = xGenConti
                       xsclase   = "nt style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;' "

                    else

                       sInput    = UCase(rs(1))
                       xsclase   = "nt style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;' "
                    End If



              

                          sHref = "../coress20/Submenus.asp?ID=" & rs(6) & "&Menu=" & Replace(UCase(rs(1)), " ", "_")
                          sLink = OFV.GenLink("", "aaa ", sHref, "Master", "", "", sInput, "si")
                          Celda = Celda & OFV.gencelda("", xsclase, "", "", "", "", "", sLink, "si")

                  end if
     
                   End If






                   rs.movenext
                Loop
                Call OFV.cerrarconsulta(rs)
          
          Else
          
                 Celda = "<B>Seleccione la Aplicacion al pie de la pantalla</b>"
                 Celda = OFV.gencelda("", "ts  ", "center", "", "", "", "", Celda, "si")

          End If
          
          
          Call OFV.cerrarconn(cn)
            Call OFV.cerrarconn(cx)

         end if


    '  End If

        if  session("SMPLMENU") <> "S" then
                  
           sHTML = sHTML & OFV.genrow("", "", "", "", "20", "", "", celda, "si")

           sHTML = sHTML & OFV.GenFTabla()

        end if 


      else

           sInput = OFV.GenerarInput("boton2", "menu", "hidden", "", "10", "", "")
           Formulario = OFV.GenForm("initaplx", "aaa STYLE='display:none;' ", "../coress20/Submenus.asp?ID=0000&Menu=NO", "Master", "post", "", sInput, "si")
           scp = "<script language='javascript' > document.initaplx.submit(); </script>"



      end if  

         


      sHTML = sHTML & Formulario & scp & "</body></html>"


Session("Usuario") = Ucase(Session("Usuario"))



Response.Write sHTML

%>