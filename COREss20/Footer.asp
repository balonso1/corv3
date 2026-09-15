<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()




      Response.expires = 0

      sOpcion = Request.querystring("Opcion")

        If Session("USRES") <> "B" Then XFondo = " background='CorFondos/" & Session("CMIMENU") & "' "

      sHTML = "<html>"

      sHTML = sHTML & "<body bgcolor='#dddddd' >"

      Celda = ""

      Formulario = "" 


      If sOpcion = "I" Then

       '   Fila = OFV.genrow("", "", "", "", "", "", "", Celda, "si")
       '   sHTML = sHTML & OFV.GenTabla("", "", "", "100%", "100%", "", "", "-1", "-2", "", "", Fila, "si")

      Else
            XFondo = "N"



          Dim aAplicacion()
            Dim aLinks()
            Dim aLinksImg()
            Dim aLinksLet()
            Dim aLinksTLT()
            Dim aLinksURL()
            Dim aLinksTRG()
            Dim aLinksDSC()

            Set cx = OFV.conectar(OFV.strconn1)
            StrSql = "select * from CorAplic"
            Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
            If Not rx.EOF Then
               If rx.fields.Count > 8 Then
                  Call OFV.cerrarconsulta(rx)
                  StrSql = "select * from CorAplic where enable <> 0 order by ordenapli "
                  Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
                  ind = 0
                  Do Until rx.EOF
                     ReDim Preserve aAplicacion(ind)
                     ReDim Preserve aLinks(ind)
                     ReDim Preserve aLinksImg(ind)
                     ReDim Preserve aLinksLet(ind)
                     ReDim Preserve aLinksTLT(ind)
                     ReDim Preserve aLinksURL(ind)
                     ReDim Preserve aLinksTRG(ind)
                     ReDim Preserve aLinksDSC(ind)
                     
                     aAplicacion(ind) = rx(1)
                     aLinks(ind) = "<a class=bt title='" & rx(12)
                     aLinks(ind) = aLinks(ind) & "' href='../COREss20/" & rx(11)
                     If UCase(rx(11)) = UCase("Menus.asp") Then
                        aLinks(ind) = aLinks(ind) & "?Opcion=" & UCase(rx(1))
                     End If
                     aLinks(ind) = aLinks(ind) & "' target='" & rx(10)
                     aLinks(ind) = aLinks(ind) & "' >"
                     aLinksImg(ind) = "<img border=0 src='../COREss20/Imagenes/" & rx(6) & "' >"
                     apllet = Replace(rx(2), "*", "")
                     aLinksLet(ind) = "<b>" & Replace(apllet, " ", "&nbsp;") & "</b>"
                     
'                     aLinks(ind) = aLinks(ind) & "</a>"

                     aLinksTLT(ind) = rx(12)
                     aLinksURL(ind) = "../COREss20/" & rx(11)
                     If UCase(rx(11)) = UCase("Menus.asp") Then
                        aLinksURL(ind) = aLinksURL(ind) & "?Opcion=" & UCase(rx(1))

                     else 
                        aLinksURL(ind) = rx(11)
                     End If
                     aLinksTRG(ind) = rx(10)
                     aLinksDSC(ind) = ucase(Replace(rx(2), "*", ""))
                     
                     ind = ind + 1
                     rx.movenext
                  Loop
                 
                  
               Else
                  
            ReDim Preserve aAplicacion(4)
          aAplicacion(0) = "CORESS20"
          aAplicacion(1) = "CORCHT20"
          aAplicacion(2) = "COREP20"
          aAplicacion(3) = "COREM20"
          aAplicacion(4) = "COREC20"
        
          ReDim Preserve aLinks(4)
          aLinks(0) = "<a title='COR&nbsp;ESS&nbsp;2.0' href=Menus.asp?Opcion=CORESS20 target=Menus><img border=0 src=Imagenes/iCorESS.jpg></a>"
          aLinks(1) = "<a title='COR&nbsp;Chat&nbsp;2.0' href=../CORChat20/chat.htm target=_blank><img border=0 src=Imagenes/iCorChat2.jpg></a>"
          aLinks(2) = "<a title='COR&nbsp;EPlanning&nbsp;2.0' href=Menus.asp?Opcion=COREP20 target=Menus><img border=0 src=Imagenes/iCorplanning.jpg></a>"
          aLinks(3) = "<a title='COR&nbsp;EManager&nbsp;2.0' href=Menus.asp?Opcion=COREM20 target=Menus><img border=0 src=Imagenes/iCormanager.jpg></a></td>"
          aLinks(4) = "<a title='COR&nbsp;EComponent&nbsp;2.0' href=Menus.asp?Opcion=COREC20 target=Menus><img border=0 src=Imagenes/iCorComponent.jpg></a>"
        
            End If
            End If
            Call OFV.cerrarconsulta(rx)

        if  session("SMPLMENU") <> "S" then
            
            
        If Session("USRES") <> "xB" Then
           sInput = "<a title='Ayuda' href=../COREss20/Ayuda/NormasHelp.asp target=info><img border=0 src='../COREss20/Imagenes/iCorSHelp.jpg' ></a>"
         '  Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a title='About' href=CorAbout.asp target=info><img border=0 src='../COREss20/Imagenes/iCorAbout.jpg' ></a>"
         '  Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a title='Usuario' href=InfoUser.asp target=info><img border=0 src='../COREss20/Imagenes/iCorSInfoUser.jpg' ></a>"
         '  Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
        Else
           sInput = "<a  class=bt title='Ayuda' href=../COREss20/Ayuda/NormasHelp.asp target=info><b>Ayuda</b></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a  class=bt title='About' href=CorAbout.asp target=info><b>Acerca&nbsp;de</b></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a  class=bt title='Usuario' href=InfoUser.asp target=info><b>Usuario</b></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
        
        End If

        end if

          xcount = 0
          scp    = ""  
        
          Set cn = OFV.conectar(OFV.strconn0)

          Fila2 = ""

          For n = 0 To UBound(aAplicacion)

               atrok = "N"
              StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'A' "
              StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & aAplicacion(n) & "'"
              Set rs1 = OFV.crearconsultaEx(StrSql, cn, 1, "")
              If rs1.EOF Then
                 atrok = "O"
              ElseIf rs1(1) Then
                     atrok = "S"
                  Else
                     atrok = "N"
              End If
              Call OFV.cerrarconsulta(rs1)

              If atrok = "O" Then


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


                 StrSql = StrSql & " And Aplicacion = '" & aAplicacion(n) & "'"
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

              End If




              If atrok = "O" Or atrok = "S"  Then
                            If XFondo = "N" Then
                               XFondo = n
                               XFdsc  = aLinksDSC(n)
                           '    sInput = OFV.GenerarInput("boton2", "menu", "hidden", "", "10", "", "")
                           '    Formulario = OFV.GenForm("initaplx", "aa style='display:none;' ", "../COREss20/Menus.asp?Opcion=" & aAplicacion(n), "Menus", "post", "", sInput, "si")
                           '    scp = "<script language='javascript'> document.initaplx.submit(); </script>"
                            End If


                            if  sopcion = aAplicacion(n) then
                                XFondo = n
                                XFdsc  = aLinksDSC(n)
                            end if


                            
                 If Session("USRES") <> "xB" Then
                    sInput = aLinks(n) & aLinksImg(n) & "</a>"
                 Else
                    sInput = aLinks(n) & aLinksLet(n) & "</a>"
                 End If

        if  session("SMPLMENU") <> "S" then



                 sinput1 = "<a  title='" & aLinksTLT(n) & "' href='../COREss20/footer.asp?Opcion="
                 sinput1 = sinput1 & aAplicacion(n) & "' target='Footer' >" & aLinksImg(n) & "</a>"

                 sinput2 = "<a  title='" & aLinksTLT(n) & "' href='../COREss20/footer.asp?Opcion="
                 sinput2 = sinput2 & aAplicacion(n) & "' target='Footer' >" & aLinksDSC(n) & "</a>"

                 xlfun = " onclick=""document.all.DFTRL.style.display='none' "" "
              
                 Celda2 = OFV.gencelda("", "at", "center " & xlfun, "2", "", "", "", sInput1, "si")

        xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;' "

                 Celda2 = Celda2 & OFV.gencelda("", xsclase & xlfun, "", "", "", "", "", sInput2, "si")

                 Fila2 = Fila2 & OFV.genrow("", "", "", "", "20", "", "", Celda2, "si")

           end if


              End If

             Next

        if  session("SMPLMENU") <> "S" then

      xsimagena  = "<b><font FACE=WebDINGS COLOR='#ffffff' SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >s</font></b>" 
      xsimagenb  = "<b><font FACE=WebDINGS COLOR='#ffffff' SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0128;</font></b>" 
      xsimagenc  = "<b><font FACE=WebDINGS COLOR='#ffffff' SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >r</font></b>" 


        xsclase = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;' "

           xlfun = " onclick=""document.all.DFTRL.style.display='block';document.all.DFTRTD1.style.display='none';document.all.DFTRTD2.style.display='block'; "" "
          xlfun2 = " onclick=""document.all.DFTRL.style.display='none';document.all.DFTRTD1.style.display='block';document.all.DFTRTD2.style.display='none'; "" "       

          Celda = OFV.gencelda("", "at style='padding-left:3px;filter: chroma(color=#ffffff);' ", "", "2", "", "", "", "<IMG  valign=middle border=0 src='../coress20/imagenes/logocorP4ico.gif' style='' >", "si")

          Celda = Celda & OFV.gencelda("", xsclase & xlfun, "", "90%", "", "", "", "&nbsp;" & XFdsc, "si")

          xsclase = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:webdings;font-size:10pt;color:#ffffff;cursor:hand;' "
          xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:webdings;font-size:10pt;color:#ffffff;cursor:hand;"


          Celda = Celda & OFV.gencelda("DFTRTD1", xsclase2 & "display:block;' ", "center title='Lista de aplicaciones.' " & xlfun, "", "1", "", "", "2", "si")
          Celda = Celda & OFV.gencelda("DFTRTD2", xsclase2 & "display:none;' ", "center title='Ocultar aplicaciones.' " & xlfun2, "", "1", "", "", "0", "si")

       '   Celda = Celda & OFV.gencelda("", xsclase, "center title='Ayuda.' onclick='mostrarHLP();' STYLE='display:block;' ", "", "1", "", "", xsimagena, "si")
           Celda = Celda & OFV.gencelda("", xsclase, "center title='Usuario.' onclick='mostrarUSR();' STYLE='display:block;' ", "", "1", "", "", xsimagenb, "si")
          Celda = Celda & OFV.gencelda("", xsclase, "center title='Logoff.' onclick='document.logoff.submit();return false;' STYLE='display:block;' ", "", "1", "", "", xsimagenc, "si")


          Fila  = OFV.genrow("", "", "", "", "20", "", "", Celda, "si")

        end if


          If XFondo <> "N" and xfondo <> "" Then

             sInput = OFV.GenerarInput("boton2", "menu", "hidden", "", "10", "", "")
             Formulario = OFV.GenForm("initaplx", "aa style='display:none;' ",aLinksURL(XFondo), "'" & aLinksTRG(XFondo) & "'", "post", "", sInput, "si")
             scp = "<script language='javascript'> document.initaplx.submit(); </script>"
          End If

        if  session("SMPLMENU") <> "S" then

        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;"
        else     
            styl = XspEx("#888888","#cccccc",0) & ";"
        end if

           sHTML = sHTML & OFV.GenTabla("DFTR", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-1", "", "", Fila, "si")

            styl = ""

           sHTML = sHTML & OFV.GenTabla("DFTRL", "aaa style='display:none;border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "1", "", "", Fila2, "si")
         end if

      End If


      sHTML = sHTML & Formulario & scp & "</body></html>"



Response.Write sHTML

%>
