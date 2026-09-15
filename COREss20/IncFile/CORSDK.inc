<%

' ****************************************************************************************
' ** Cor E-Solution 3 - Esquema de Seguridad ---------------------------- Enero, 2006   **
' ****************************************************************************************


''''''''''''

dim xtabla(255)
dim atabla(255)


Public Function getrdm()

Dim d1, d2, d3, d4, d5, d6, d7, i

d1 = Session.SESSIONID
d2 = Len(d1) Mod 2
d3 = Session.SESSIONID
d4 = Len(d3) Mod 2
If d2 > 0 Then
   d3 = (d1 * 10) + CInt(Mid(CStr(d1), Len(d1), 1))
   d4 = Len(d3) Mod 2
End If
d5 = 0
d7 = CInt(Mid(CStr(d3), (Len(d3) - 2), 1)) + CInt(Mid(CStr(d3), (Len(d3) - 1), 1)) + CInt(Mid(CStr(d3), Len(d3), 1))
For i = 0 To 10
    If Len(d7) > 1 Then
       d7 = CInt(Mid(CStr(d7), 1, 1)) + CInt(Mid(CStr(d7), 2, 1))
    Else
       Exit For
    End If
Next

d5 = Len(d7)

getrdm = d7

End Function

Public Function getrdma(crdm)

Dim d1, d2, d3, d4, d5, d6, d7, i, y, x
Dim a6()
Dim a3()

d1 = Session.SESSIONID
d2 = Len(d1) Mod 2
d3 = Session.SESSIONID
d4 = Len(d3) Mod 2
If d2 > 0 Then
   d3 = (d1 * 10) + CInt(Mid(CStr(d1), Len(d1), 1))
   d4 = Len(d3) Mod 2
End If

d6 = d3



ReDim Preserve a6(Len(d6))
ReDim Preserve a3(Len(d6))

For i = 1 To Len(d6)
    a3(i) = CInt(Mid(CStr(d3), i, 1))
    a6(i) = CInt(Mid(CStr(d6), i, 1))
Next



'crdm = cint(crdm)

Select Case crdm
       Case 1

            For i = 1 To Len(d6)
                y = a6(i) + 5
                For x = 0 To 10
                    If Len(y) > 1 Then
                       y = CInt(Mid(CStr(y), 1, 1)) + CInt(Mid(CStr(y), 2, 1))
                    Else
                       Exit For
                    End If
                Next
        
                       
                a6(i) = y
            Next


      
       Case 2
            For i = 1 To (Len(d6) / 2)
                a6(i + (Len(d6) / 2)) = Mid(CStr(d6), i, 1)
                a6(i) = Mid(CStr(d6), i + (Len(d6) / 2), 1)
            Next

       Case 3
            x = (Len(d6) / 2)
            If (x Mod 2) <> 0 Then
                y = x - 1
            Else
                y = x
            End If
            y = y / 2
            For i = 1 To y
                a6(x) = Mid(CStr(d6), i, 1)
                a6(i) = Mid(CStr(d6), x, 1)
                a6(x + (Len(d6) / 2)) = Mid(CStr(d6), i + (Len(d6) / 2), 1)
                a6(i + (Len(d6) / 2)) = Mid(CStr(d3), x + (Len(d6) / 2), 1)
                x = x - 1
            Next

       Case 4
            For i = 1 To (Len(d6) - 1)
                a6(i + 1) = a6(i)
                a6(i) = a3(i + 1)

                a3(i + 1) = a6(i + 1)
                a3(i) = a6(i)
            Next

       Case 5
            For i = 1 To Len(d6)
                y = a6(i) + 3

                For x = 0 To 10
                    If Len(y) > 1 Then
                       y = CInt(Mid(CStr(y), 1, 1)) + CInt(Mid(CStr(y), 2, 1))
                    Else
                       Exit For
                    End If
                Next
        
                       
                a6(i) = y
            Next

       Case 6
            For i = 1 To Len(d6)
                y = a6(i) + 6

                For x = 0 To 10
                    If Len(y) > 1 Then
                       y = CInt(Mid(CStr(y), 1, 1)) + CInt(Mid(CStr(y), 2, 1))
                    Else
                       Exit For
                    End If
                Next
        
                       
                a6(i) = y
            Next


       Case 7
            For i = 1 To Len(d6)
                y = a6(i) + 7
                For x = 0 To 10
                    If Len(y) > 1 Then
                       y = CInt(Mid(CStr(y), 1, 1)) + CInt(Mid(CStr(y), 2, 1))
                    Else
                       Exit For
                    End If
                Next
        
                       
                a6(i) = y
            Next

       Case 8
            For i = 1 To Len(d6)
                y = a6(i) + 8
                For x = 0 To 10
                    If Len(y) > 1 Then
                       y = CInt(Mid(CStr(y), 1, 1)) + CInt(Mid(CStr(y), 2, 1))
                    Else
                       Exit For
                    End If
                Next
        
                       
                a6(i) = y
            Next

       Case 9
            y = Len(d6)
            For i = 1 To (Len(d6) / 2)
                a6(y) = Mid(CStr(d6), i, 1)
                a6(i) = Mid(CStr(d6), y, 1)
                y = y - 1
            Next

       Case Else
            For i = 1 To Len(d6) Step 2
                a6(i + 1) = a6(i)
                a6(i) = a3(i + 1)
                If (i + 2) > Len(d6) Then Exit For
            Next


 

End Select


d6 = ""
For i = 1 To Len(d3)
    d6 = d6 & CStr(a6(i))
Next

getrdma = d6

End Function

Public Function getsid()

    getsid = CInt(Mid(CStr(Session.SESSIONID), Len(Session.SESSIONID), 1)) * 17

End Function

Function genMkt()

Dim i, xi, n, m, drdm, drdma, mtop, si
Dim masc(255)


drdm = getrdm()
xi = 0

For n = 0 To 255
    drdma = getrdma(drdm)
    mtop = Len(drdma)
    For m = 1 To mtop
        masc(xi) = CInt(Mid(CStr(drdma), m, 1))
        xi = xi + 1
        If xi > 255 Then Exit For
    Next
    If xi > 255 Then
       Exit For
    Else
       drdm = drdm + 1
       If drdm > 9 Then
          drdm = 0
       End If
    End If
Next

For i = 0 To 255
    si = (masc(i) * 900) + i + ((masc(i) + 3) * 7)
    xtabla(i) = 9999 - si
Next


End Function

Function genMkCh()

Dim i

For i = 0 To 255
    atabla(i) = Chr(i)
Next

End Function

Function Getcrpmk(sopcion, sdato, sid)


If sopcion = "D" Then
   xdatoi = ""
   datoS = ""
   y = 1
   xy = 0
   For i = 1 To Len(sdato) Step 4
       xdatoi = CInt(Mid(sdato, i, 4))
       mascx = CInt(y * 7) + CInt(y * 31 / 17) + sid
       xdatoi = xdatoi + mascx
       For xy = 0 To 255
           If CInt(xdatoi) = CInt(xtabla(xy)) Then
              datoS = datoS & atabla(xy)
              Exit For
           End If
       Next
       y = y + 1
  Next
  Getcrpmk = datoS
Else
  Getcrpmk = ""
End If


End Function

Function Inicio(sopcion)




      xopcion = sopcion

      vu = OFV.ValidarUsuario()
      If vu = "OK" Then

  Set cn = OFV.conectar(OFV.strconn2)

  StrSql = "Select * From Clienteplan"
  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")

  Session("Cliente") = rs(1)
  Session("CliLogo") = "../COREss20/Empresa/" & rs(5)

  Call OFV.cerrarconsulta(rs)
  Call OFV.cerrarconn(cn)

          sHTML = OFV.MenuHeader("", fondo) & Chr(13)

sHTML = sHTML & "<STYLE>"
sHTML = sHTML & " td.at{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:500;}"
if ofv.modtitcls <> "10" then
   sHTML = sHTML & " input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;" & XspEx("#dddddd",session("APLBGD"),0) & ";}"
else
   sHTML = sHTML & " input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;}"
end if
sHTML = sHTML & "</STYLE>"

sHTML = sHTML & "<script language=javascript >" & Chr(13)
sHTML = sHTML & " var ICript = 0;" & Chr(13)
sHTML = sHTML & "function cript()" & Chr(13)
sHTML = sHTML & "{ " & Chr(13)
sHTML = sHTML & " if (ICript==0) { " & Chr(13)
sHTML = sHTML & " ICript = 1;" & Chr(13)
sHTML = sHTML & "var sid = " & getsid() & "; " & Chr(13)
sHTML = sHTML & "var i = 0;" & Chr(13)
sHTML = sHTML & "var tab2 = new Array("
genMkt()

For i = 0 To 255
    If i > 0 Then sHTML = sHTML & ","
    sHTML = sHTML & xtabla(i)
Next
sHTML = sHTML & ");" & Chr(13)
sHTML = sHTML & "var Y     = 0;" & Chr(13)
sHTML = sHTML & "var i     = 0;" & Chr(13)
sHTML = sHTML & "var dato  = document.all.clave.value;" & Chr(13)
sHTML = sHTML & "var ldato = dato.length;" & Chr(13)
sHTML = sHTML & "var datos = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "var masc  = 0;" & Chr(13)

''sHTML = sHTML & " alert(dato);" & Chr(13)
''sHTML = sHTML & "}" & Chr(13)
''sHTML = sHTML & "}" & Chr(13)
''sHTML = sHTML & "function xabrir() {"

sHTML = sHTML & "if (ldato < 31) " & Chr(13)
sHTML = sHTML & " {for (i=1;i<=ldato;i++)" & Chr(13)
sHTML = sHTML & "   {" & Chr(13)
sHTML = sHTML & "     Y     = dato.charCodeAt(i-1);" & Chr(13)
sHTML = sHTML & "     masc  = Math.round(i * 7) + Math.round(i * 31 / 17) + sid;" & Chr(13)
sHTML = sHTML & "     datos = datos + String(Math.round(tab2[Y]) - masc);" & Chr(13)
sHTML = sHTML & "   } " & Chr(13)
sHTML = sHTML & "  document.all.clave.value = datos;" & Chr(13)
sHTML = sHTML & "  document.all.Inicio.submit();" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & " else " & Chr(13)
sHTML = sHTML & " {" & Chr(13)
sHTML = sHTML & "  document.all.clave.value = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "  alert('La password no puede exceder de 30 caracteres');" & Chr(13)
sHTML = sHTML & "  ICript = 0;" & Chr(13)
sHTML = sHTML & "  }" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & "}" & Chr(13)


sHTML = sHTML & "function abrir()"
sHTML = sHTML & "{"
sHTML = sHTML & "document.all.u20.style.display='none';"
sHTML = sHTML & "document.all.u21.style.display='none';"
sHTML = sHTML & "document.all.u22.style.display='none';"
sHTML = sHTML & "document.all.u23.style.display='none';"
sHTML = sHTML & "document.all.u24.style.display='none';"

sHTML = sHTML & "document.all.u10.style.display='block';"
sHTML = sHTML & "document.all.u11.style.display='block';"
sHTML = sHTML & "document.all.u12.style.display='block';"
sHTML = sHTML & "document.all.u13.style.display='block';"
sHTML = sHTML & "document.all.u14.style.display='block';"

sHTML = sHTML & "document.all.u11.focus();"
sHTML = sHTML & "}"

sHTML = sHTML & "function cerrar()"
sHTML = sHTML & "{"
sHTML = sHTML & "document.all.u20.style.display='block';"
sHTML = sHTML & "document.all.u21.style.display='block';"
sHTML = sHTML & "document.all.u22.style.display='block';"
sHTML = sHTML & "document.all.u23.style.display='block';"
sHTML = sHTML & "document.all.u24.style.display='block';"

sHTML = sHTML & "document.all.u10.style.display='none';"
sHTML = sHTML & "document.all.u11.style.display='none';"
sHTML = sHTML & "document.all.u12.style.display='none';"
sHTML = sHTML & "document.all.u13.style.display='none';"
sHTML = sHTML & "document.all.u14.style.display='none';"
sHTML = sHTML & "}"

sHTML = sHTML & "</script>" & Chr(13)




    Set cn = OFV.conectar(OFV.strconn1)
    StrSql = "select * from CorAplic where Codapli= '" & Session("TOPAPL") & "'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then
       AplMain = Replace(rs(2), "*", " ") & " " & rs(4)
    Else
       AplMain = "Cor E-Solution Suite V3"
    End If

    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)
 
    Session("CAplMain") = AplMain



sHTML = sHTML & "<center>"

 if  SESSION("FULLGUI") = "S" and SESSION("RDCGUID") <> "S" THEN

'''''''''''''''''''''

sHTML = sHTML & "  <div id='LGM0' style='z-index:9;position:absolute;top:242;left:460;' > "
sHTML = sHTML & " <table style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#eeeeee;' > "
sHTML = sHTML & " <td style='line-height:25pt;font-size:24pt;font-family:""verdana,tahoma,arial"";color:#888888;' >"
sHTML = sHTML & " <b>&nbsp;COR&nbsp;&nbsp;E-Solution&nbsp;Suite&nbsp;V34&nbsp;</b></td></table></div> "
sHTML = sHTML & "  <div id='LGM0a' style='z-index:9;position:absolute;top:240;left:460;' > "
sHTML = sHTML & " <table style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#ff0000;filter:shadow(color=#CCCCCC,direction=100);' > "
sHTML = sHTML & " <td style='line-height:25pt;font-size:24pt;font-family:""verdana,tahoma,arial"";color:#FFFFFF;' ><b>&nbsp;COR&nbsp;&nbsp;E-Solution&nbsp;Suite&nbsp;V34&nbsp;</b></td> "
sHTML = sHTML & " </table></div> "



sHTML = sHTML & "   <DIV id='THMX'  style='z-index:0;position:absolute;top:200;left:160;width:820px;height:230px;background-color:#6699cc;filter: alpha(opacity=40,finishOpacity=100,STYLE=1);' > "
sHTML = sHTML & " <P >&nbsp;</P> "
sHTML = sHTML & "   </DIV> "

sHTML = sHTML & "  <div id='LGM0a' style='z-index:9;position:absolute;top:480;left:700;' > "
sHTML = sHTML & " <table  > "
sHTML = sHTML & " <td style='line-height:16pt;font-size:14pt;font-family:""verdana,tahoma,arial"";color:#FF0000;' ><b>&nbsp;COR&nbsp;&nbsp;CONSULTING&nbsp;</b></td> "
sHTML = sHTML & " <td style='line-height:16pt;font-size:14pt;font-family:""verdana,tahoma,arial"";color:#FF0000;' ><img src='../CORESS20/IMAGENES/logocorP3sm.jpg' WIDTH='70%' ></td> "
sHTML = sHTML & " </table></div> "



sHTML = sHTML & "      <div id='xx14' style='z-index:3;position:absolute;top:260;left:200;filter:shadow(color=#dddddd,direction=100);' > "
sHTML = sHTML & " <img src='../CORESS20/IMAGENES/logocorP4sm1.gif' ></div> "






'''''''''''''''''''''''''''''''

 Whtml1 = ""






sHTML = sHTML & Whtml1
''''''''''''''''''''''''''''''
else
''''''''''''''''''''''''''''''''''''''''''''

sHTML = sHTML & "  <div id='LGM0' style='z-index:9;position:absolute;top:242;left:460;' > "
sHTML = sHTML & " <table style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#eeeeee;' > "
sHTML = sHTML & " <td style='line-height:25pt;font-size:24pt;font-family:""verdana,tahoma,arial"";color:#888888;' >"
sHTML = sHTML & " <b>&nbsp;COR&nbsp;&nbsp;E-Solution&nbsp;Suite&nbsp;V34&nbsp;</b></td></table></div> "
sHTML = sHTML & "  <div id='LGM0a' style='z-index:9;position:absolute;top:240;left:460;' > "
sHTML = sHTML & " <table style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#ff0000;filter:shadow(color=#CCCCCC,direction=100);' > "
sHTML = sHTML & " <td style='line-height:25pt;font-size:24pt;font-family:""verdana,tahoma,arial"";color:#FFFFFF;' ><b>&nbsp;COR&nbsp;&nbsp;E-Solution&nbsp;Suite&nbsp;V34&nbsp;</b></td> "
sHTML = sHTML & " </table></div> "



sHTML = sHTML & "   <DIV id='THMX'  style='z-index:0;position:absolute;top:200;left:160;width:820px;height:230px;background-color:#6699cc;filter: alpha(opacity=40,finishOpacity=100,STYLE=1);' > "
sHTML = sHTML & " <P >&nbsp;</P> "
sHTML = sHTML & "   </DIV> "

sHTML = sHTML & "  <div id='LGM0a' style='z-index:9;position:absolute;top:480;left:700;' > "
sHTML = sHTML & " <table  > "
sHTML = sHTML & " <td style='line-height:16pt;font-size:14pt;font-family:""verdana,tahoma,arial"";color:#FF0000;' ><b>&nbsp;COR&nbsp;&nbsp;CONSULTING&nbsp;</b></td> "
sHTML = sHTML & " <td style='line-height:16pt;font-size:14pt;font-family:""verdana,tahoma,arial"";color:#FF0000;' ><img src='../CORESS20/IMAGENES/logocorP3sm.jpg' WIDTH='70%' ></td> "
sHTML = sHTML & " </table></div> "



sHTML = sHTML & "      <div id='xx14' style='z-index:3;position:absolute;top:260;left:200;filter:shadow(color=#dddddd,direction=100);' > "
sHTML = sHTML & " <img src='../CORESS20/IMAGENES/logocorP4sm1.gif' ></div> "



 Whtml1 = ""



sHTML = sHTML & Whtml1
'''''''''''''''''''''''''''''''''''''''''''
end if

sHTML = sHTML & OFV.GenForm("Inicio", "", "Inicio.asp?Opcion=V", "", "post", "", "", "no")

sHTML = sHTML & "<DIV  CLASS='aFilter' ID='oDiv' style='DISPLAY:NONE;z-index:80;position:absolute;top:300;left:540;width:150px;height:200px;' >"

          xwsize = 45
          If OFV.explorador = "MSIE" Then xwsize = 42



          Fila = ""

          UserNt = UCase(Request.ServerVariables("LOGON_USER"))
          UserNt = Mid(UserNt, InStrRev(UserNt, "\") + 1, Len(UserNt))

          DOMAINNT = UCase(Request.ServerVariables("LOGON_USER"))
          DOMAINNT = Left(DOMAINNT, InStr(DOMAINNT, "\")) & "\"
          DOMAINNT = Replace(DOMAINNT, "\", "")
           
          If Session("Modal") = "MSNTT" Then
           
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("1","1",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", "<B>USUARIO NT:</B>", "si")
              Fila = Fila & OFV.GenRow("u21", "aaa style='display:block;cursor:hand;'", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("usuarNt readonly style='border-style:solid;border-width:1;BACKGROUND-COLOR:#f9f9f9;" & XspDShw("2","2",session("CMFPNTIT")) & ";' ", UserNt, "agregar", "", "40", "at", "no")
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "center ", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u22", "aaa style='display:block;cursor:hand;'", "", "", "", "", "", Celda, "si")

              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("1","1",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", "<B>DOMINIO NT:</B>", "si")
              Fila = Fila & OFV.GenRow("u23", "aaa style='display:block;cursor:hand;'", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("DOMAINNt readonly style='border-style:solid;border-width:1;BACKGROUND-COLOR:#f9f9f9;" & XspDShw("2","2",session("CMFPNTIT")) & ";' ", DOMAINNT, "agregar", "", "40", "at", "no")
              Celda = OFV.gencelda("", "at style='color:#336699;" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "center ", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u24", "aaa style='display:block;cursor:hand;'", "", "", "", "", "", Celda, "si")
          
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("1","1",session("CMFPNTIT")) & ";' onclick='abrir();' ", "center", "80%", "20", "", "", "<B><FONT face=wingdings size=2>1</font></B>", "si")
              Fila = Fila & OFV.GenRow("u20", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
            
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("1","1",session("CMFPNTIT")) & ";' onclick='cerrar();' ", "center", "80%", "20", "", "", "<B><FONT face=wingdings size=2>0</font></B>", "si")
              Fila = Fila & OFV.GenRow("u10", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
          
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("1","1",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", "<B>USUARIO:</B>", "si")
              Fila = Fila & OFV.GenRow("u11", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("usuar", "", "agregar", "", "40", "dt", "no")
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u12", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("1","1",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", "<B>PASSWORD:</B>", "si")
              Fila = Fila & OFV.GenRow("u13", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("clave", "", "password", "", "40", "dt", "no")
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u14", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
          
          
          Else
          
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("1","1",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", "<B>USUARIO:</B>", "si")
              Fila = Fila & OFV.GenRow("u11", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("usuar", "", "agregar", "", "40", "dt", "no")
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u12", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("1","1",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", "<B>PASSWORD:</B>", "si")
              Fila = Fila & OFV.GenRow("u13", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("clave", "", "password", "", "40", "dt", "no")
              Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "center", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
          
          End If
          

 '' Formulario = OFV.GenForm("Inicio", "", "Inicio.asp?Opcion=V", "", "post", "", Fila, "si")

  Formulario = Fila

  sHTML = sHTML & OFV.GenTabla("", "", "", "30% ", "", "0", "", "0", "0", "", "", Formulario, "si")

  sHTML = sHTML & "</DIV> "

  xsiniciar = ofv.GenBoton(20, "bini", "tt", "N", " onclick='cript();return false;' ", "N", "N", "S", "", "")



sHTML = sHTML & "<DIV  ID='oBDiv' style='DISPLAY:NONE;z-index:82;position:absolute;top:340;left:850;' >"
sHTML = sHTML & xsiniciar
sHTML = sHTML & "</DIV>"
sHTML = sHTML & "</center>"

  sHTML = sHTML & OFV.GenFForm()

          serrordes = ""
          If xopcion <> "" Then
             Select Case xopcion
                    Case "1"
                         serrordes = "SKR001 - Usuario invalido: Usuario nulo o no existente "
                    Case "11"
                         serrordes = "SKR011 - Dominio no habilitado para la aplicacion "
                    Case "2"
                         serrordes = "SKR002 - Password invalida: Password nula o invalida "
                    Case "3"
                         serrordes = "SKR003 - Usuario bloqueado: Usuario bloqueado por "
                         serrordes = serrordes & " superar la cantidad de intentos fallidos "
                    Case "6"
                         serrordes = "SKR006 - Usuario bloqueado: Usuario bloqueado por "
                         serrordes = serrordes & " el Administrador "
                    Case "7"
                         serrordes = "SKR007 - Acceso Restringido: Usuario no habilitado "
                         serrordes = serrordes & " para acceder Sabados y Domingos "
                    Case "8"
                         serrordes = "SKR008 - Acceso Restringido: Usuario no habilitado "
                         serrordes = serrordes & " para acceder Sabados, Domingos y Feriados"
                    Case "9"
                         serrordes = "SKR009 - Acceso Restringido: Usuario no puede acceder "
                         serrordes = serrordes & " fuera del horario autorizado "
                         
                    Case Else
                         serrordes = "SKR901 - Password incorrecta: Error imprevisto en "
                         serrordes = serrordes & " Log-In "

             End Select




 
          End If

sHTML = sHTML & "<DIV  ID='oEDiv' style='DISPLAY:NONE;z-index:85;position:absolute;top:450;left:450;' >"
sHTML = sHTML & "<p  ><font face=verdana color=" & session("CMFPNTIT") & " size=2 ><b>" & serrordes & "</font></p>"
sHTML = sHTML & "</DIV>"

sHTML = sHTML & chr(13)
sHTML = sHTML & "<script language=javascript >" & chr(13)
sHTML = sHTML & " var ST=0;               "  & chr(13)
sHTML = sHTML & " var STW1=65;             "  & chr(13)
sHTML = sHTML & " var STH1=22;             "  & chr(13)
sHTML = sHTML & " var STT1=20;             "  & chr(13)
sHTML = sHTML & " var STT2=19;             "  & chr(13)
sHTML = sHTML & " var STT3=29;             "  & chr(13)
sHTML = sHTML & " var STT4=32;             "  & chr(13)
sHTML = sHTML & " var STW=65;             "  & chr(13)
sHTML = sHTML & " var STH=22;             "  & chr(13)
sHTML = sHTML & " var ST1=65;             "  & chr(13)
sHTML = sHTML & " var ST2=22;             "  & chr(13)
sHTML = sHTML & " var ST3=65;             "  & chr(13)
sHTML = sHTML & " var ST4=22;             "  & chr(13)
sHTML = sHTML & " function FMtmp() {      "  & chr(13)
sHTML = sHTML & "    document.all.oDiv.style.display   = 'block';"  & chr(13) 
sHTML = sHTML & "    document.all.oBDiv.style.display  = 'block';"  & chr(13) 
sHTML = sHTML & "    document.all.oEDiv.style.display  = 'block';"  & chr(13) 
sHTML = sHTML & " }                       "  & chr(13)    
sHTML = sHTML & " FMtmp();"  & chr(13)
sHTML = sHTML & "</script>"  & chr(13)


          sHTML = sHTML & "</BODY></html>"
          Inicio = ""
      Else
          sHTML = vu
      End If

      Inicio = sHTML
    

End Function

Function InicioSesion()



     InicioSesion = ""

      ok = OFV.CheckUsuario()
      If ok <> "Y" Then
          uv = OFV.uv
          If uv = "NO" Then ok = OFV.ValidarUsuario
          sHTML = ok
      Else
          Response.expires = 0
          xscr = "    onload=" & Chr(34)
          xscr = xscr & "document.Menuf.submit();"
 if  SESSION("FULLGUI") = "S" THEN
          xscr = xscr & "FMtmp();"
 end if 
          xscr = xscr & Chr(34)

'''xscr = ""
    
          sHTML = OFV.MenuHeader("", xscr)

sHTML = sHTML & "<STYLE>"
sHTML = sHTML & " td.at{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:500;}"
if ofv.modtitcls <> "10" then
   sHTML = sHTML & " input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;" & XspEx("#dddddd",session("APLBGD"),0) & ";}"
else
   sHTML = sHTML & " input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;}"
end if
sHTML = sHTML & "</STYLE>"


 if  SESSION("FULLGUI") = "S"  and SESSION("RDCGUID") <> "S" THEN

'''''''''''''''''''''


sHTML = sHTML & "  <div id='LGM0' style='z-index:9;position:absolute;top:242;left:460;' > "
sHTML = sHTML & " <table style='width:500px;border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#eeeeee;' > "
sHTML = sHTML & " <td style='line-height:25pt;font-size:24pt;font-family:""verdana,tahoma,arial"";color:#888888;' >"
sHTML = sHTML & " <b>&nbsp;" & REPLACE(Session("CAplMain")," ","&nbsp;") & "&nbsp;</b></td></table></div> "
sHTML = sHTML & "  <div id='LGM0a' style='z-index:9;position:absolute;top:240;left:460;' > "
sHTML = sHTML & " <table style='width:500px;border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#ff0000;filter:shadow(color=#CCCCCC,direction=100);' > "
sHTML = sHTML & " <td style='line-height:25pt;font-size:24pt;font-family:""verdana,tahoma,arial"";color:#FFFFFF;' ><b>&nbsp;" & REPLACE(Session("CAplMain")," ","&nbsp;") & "&nbsp;</b></td> "
sHTML = sHTML & " </table></div> "



sHTML = sHTML & "   <DIV id='THMX'  style='z-index:0;position:absolute;top:200;left:160;width:820px;height:230px;background-color:#6699cc;filter: alpha(opacity=40,finishOpacity=100,STYLE=1);' > "
sHTML = sHTML & " <P >&nbsp;</P> "
sHTML = sHTML & "   </DIV> "

sHTML = sHTML & "  <div id='LGM0a' style='z-index:9;position:absolute;top:480;left:700;' > "
sHTML = sHTML & " <table  > "
sHTML = sHTML & " <td style='line-height:16pt;font-size:14pt;font-family:""verdana,tahoma,arial"";color:#FF0000;' ><b>&nbsp;COR&nbsp;&nbsp;CONSULTING&nbsp;</b></td> "
sHTML = sHTML & " <td style='line-height:16pt;font-size:14pt;font-family:""verdana,tahoma,arial"";color:#FF0000;' ><img src='../CORESS20/IMAGENES/logocorP3sm.jpg' WIDTH='70%' ></td> "
sHTML = sHTML & " </table></div> "



sHTML = sHTML & "      <div id='xx14' style='z-index:3;position:absolute;top:260;left:200;filter:shadow(color=#dddddd,direction=100);' > "
sHTML = sHTML & " <img src='../CORESS20/IMAGENES/logocorPsemitonosm.gif' ></div> "





''''''''''''''''''''''''''''''
else
''''''''''''''''''''''''''''''''''''''''''''

sHTML = sHTML & "  <div id='LGM0' style='z-index:9;position:absolute;top:242;left:460;' > "
sHTML = sHTML & " <table style='width:500px;border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#eeeeee;' > "
sHTML = sHTML & " <td style='line-height:25pt;font-size:24pt;font-family:""verdana,tahoma,arial"";color:#888888;' >"
sHTML = sHTML & " <b>&nbsp;" & REPLACE(Session("CAplMain")," ","&nbsp;") & "&nbsp;</b></td></table></div> "
sHTML = sHTML & "  <div id='LGM0a' style='z-index:9;position:absolute;top:240;left:460;' > "
sHTML = sHTML & " <table style='width:500px;border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#ff0000;filter:shadow(color=#CCCCCC,direction=100);' > "
sHTML = sHTML & " <td style='line-height:25pt;font-size:24pt;font-family:""verdana,tahoma,arial"";color:#FFFFFF;' ><b>&nbsp;" & REPLACE(Session("CAplMain")," ","&nbsp;") & "&nbsp;</b></td> "
sHTML = sHTML & " </table></div> "




sHTML = sHTML & "   <DIV id='THMX'  style='z-index:0;position:absolute;top:200;left:160;width:820px;height:230px;background-color:#6699cc;filter: alpha(opacity=40,finishOpacity=100,STYLE=1);' > "
sHTML = sHTML & " <P >&nbsp;</P> "
sHTML = sHTML & "   </DIV> "

sHTML = sHTML & "  <div id='LGM0a' style='z-index:9;position:absolute;top:480;left:700;' > "
sHTML = sHTML & " <table  > "
sHTML = sHTML & " <td style='line-height:16pt;font-size:14pt;font-family:""verdana,tahoma,arial"";color:#FF0000;' ><b>&nbsp;COR&nbsp;&nbsp;CONSULTING&nbsp;</b></td> "
sHTML = sHTML & " <td style='line-height:16pt;font-size:14pt;font-family:""verdana,tahoma,arial"";color:#FF0000;' ><img src='../CORESS20/IMAGENES/logocorP3sm.jpg' WIDTH='70%' ></td> "
sHTML = sHTML & " </table></div> "



sHTML = sHTML & "      <div id='xx14' style='z-index:3;position:absolute;top:260;left:200;filter:shadow(color=#dddddd,direction=100);' > "
sHTML = sHTML & " <img src='../CORESS20/IMAGENES/logocorPsemitonosm.gif' ></div> "



'''''''''''''''''''''''''''''''''''''''''''
end if

sHTML = sHTML & "<DIV  CLASS='aFilter' ID='oDiv' style='z-index:80;position:absolute;top:320;left:650;width:300px;height:200px;' ><BIG>"





        
          xwsize = 45
          If OFV.explorador = "MSIE" Then xwsize = 42

          Fila = ""
          Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "", "20%", "20", "", "", "<BIG>SESION:", "si")
          Celda = Celda & OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";' ", "", "80%", "20", "", "", "<b><BIG>" & Session.SESSIONID & "</b>", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "", "20%", "20", "", "", "<BIG>USUARIO:", "si")
          Celda = Celda & OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "", "80%", "20", "", "", "<b><BIG>" & Session("usuario") & "</b>", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


          Celda = OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "", "20%", "20", "", "", "<BIG>NOMBRE:", "si")
          Celda = Celda & OFV.gencelda("", "at style='color:" & Session("CMFMEFT") & ";" & XspDShw("2","2",session("CMFPNTIT")) & ";'", "", "80%", "20", "", "", "<b><BIG>" & Session("segusuario") & "</b>", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")






           Fila = Fila & OFV.GenForm("Menuf", "", "Footer.asp?Opcion=A", "Footer", "post", "", "<input type=hidden size=" & OFV.wsize & " name=boton value=menu></input>", "si")


  Set cn = OFV.conectar(OFV.strconn2)

  StrSql = "Select * From Clienteplan"
  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")

  Session("Cliente") = rs(1)
  Session("CliLogo") = "../COREss20/Empresa/" & rs(5)

  Call OFV.cerrarconsulta(rs)
  Call OFV.cerrarconn(cn)

          If Session("USRES") = "B" Then
          sHTML = sHTML & OFV.GenTabla("panel", "aaa ", "", "100%", "", 0, 0, 0, 0, "", "", Fila, "si")
          Else
          sHTML = sHTML & OFV.GenTabla("panel", "aaa ", "", "100%", "", 0, 0, 0, 0, "", "", Fila, "si")
          End If
      End If

sHTML = sHTML & "</DIV>" 


sHTML = sHTML & "<script language=javascript >" & chr(13)
sHTML = sHTML & " var tt=0;          "  & chr(13)
sHTML = sHTML & " function FMtmp() { "  & chr(13)
sHTML = sHTML & " }                  "  & chr(13)      
sHTML = sHTML & "</script>"  & chr(13)


sHTML = sHTML & "</BODY></html>"
    
      InicioSesion = sHTML

End Function



Function LogFinal()



     LogFinal = ""

      ok = OFV.CheckUsuario()
      If ok <> "Y" Then
          uv = OFV.uv
          If uv = "NO" Then ok = OFV.ValidarUsuario
          sHTML = ok
      Else
          Response.expires = 0
          xscr = "    onload=" & Chr(34)
 if  SESSION("FULLGUI") = "S"  and SESSION("RDCGUID") <> "S" THEN
          xscr = xscr & "FMtmp();"
 END IF
          xscr = xscr & "reinicia();"
          xscr = xscr & Chr(34)
    
          sHTML = OFV.MenuHeader("", xscr)

sHTML = sHTML & "<STYLE>"
sHTML = sHTML & " td.at{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:500;}"
if ofv.modtitcls <> "10" then
   sHTML = sHTML & " input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;" & XspEx("#dddddd",session("APLBGD"),0) & ";}"
else
   sHTML = sHTML & " input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;}"
end if
sHTML = sHTML & "</STYLE>"

sHTML = sHTML & "<CENTER>"

 if  SESSION("FULLGUI") = "S"  and SESSION("RDCGUID") <> "S" THEN

'''''''''''''''''''''


sHTML = sHTML & " <v:line from='0,0' to='1,0' id='LGM0' style='z-index:9;position:relative;top:60px;width:10px;height:10px;' > "
sHTML = sHTML & "   <v:fill on='True' color='#99ccdd' color2='#6699cc' method='sigma' type='gradient' angle='30' opacity='90%'  /> "
sHTML = sHTML & "   <v:stroke on='False' /> "
sHTML = sHTML & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
sHTML = sHTML & "   <v:path textpathok='True' /> "
sHTML = sHTML & "     <v:textpath on='True' string='DESCONECCION EN PROCESO ...' style=""font-size:'15pt';font-family:'verdana';font-weight:bold;""  /> "
sHTML = sHTML & "   <v:extrusion on='True' backdepth='10'  metal='True' color='#6699cc' rotationcenter='0,0,0'  rotationangle = '30,0' /> "
sHTML = sHTML & " </v:line>"


sHTML = sHTML & " <v:line from='0,0' to='1,0' id='LGM1' style='z-index:9;position:absolute;top:200px;left:550px;width:10px;height:10px;' > "
sHTML = sHTML & "   <v:fill on='True' color='#eeeeee' color2='#336699' method='sigma' type='gradient' angle='30' opacity='90%'  /> "
sHTML = sHTML & "   <v:stroke on='False' /> "
sHTML = sHTML & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
sHTML = sHTML & "   <v:path textpathok='True' /> "
sHTML = sHTML & "     <v:textpath on='True' string='&#xF080;' style=""font-size:'60pt';font-family:'webdings';font-weight:bold;""  /> "
sHTML = sHTML & "   <v:extrusion on='True' backdepth='7' color='#336699' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
sHTML = sHTML & " </v:line>"

sHTML = sHTML & " <v:line from='0,0' to='1,0' id='LGM2' style='z-index:3;position:absolute;top:195px;left:590px;width:10px;height:10px;' > "
sHTML = sHTML & "   <v:fill on='True' color='#eeeeee' color2='#993333' method='sigma' type='gradient' angle='30' opacity='90%'  /> "
sHTML = sHTML & "   <v:stroke on='False' /> "
sHTML = sHTML & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
sHTML = sHTML & "   <v:path textpathok='True' /> "
sHTML = sHTML & "     <v:textpath on='True' string='r' style=""font-size:'45pt';font-family:'webdings';font-weight:bold;""  /> "
sHTML = sHTML & "   <v:extrusion on='True' backdepth='7' color='#993333' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
sHTML = sHTML & " </v:line>"

sHTML = sHTML & " <v:line from='0,0' to='1,0' id='LGM3' style='z-index:3;position:absolute;top:260px;left:555px;width:10px;height:10px;' > "
sHTML = sHTML & "   <v:fill on='True' color='#FFFFFF' color2='#FFFFFF' method='sigma' type='gradient' angle='270' opacity='90%'  /> "
sHTML = sHTML & "   <v:stroke on='False' /> "
sHTML = sHTML & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
sHTML = sHTML & "   <v:path textpathok='True' /> "
sHTML = sHTML & "     <v:textpath on='True' string='FINAL DE SESION' style=""font-size:'10pt';font-family:'VERDANA';font-weight:bold;""  /> "
sHTML = sHTML & "   <v:extrusion on='True' backdepth='2'  metal='True' color='" & session("CMFPNTIT") & "' rotationcenter='0,0,0'  rotationangle = '0,0' /> "
sHTML = sHTML & " </v:line>"

sHTML = sHTML & " <v:line from='0,0' to='1,0' id='LGM4'  style='z-index:3;position:absolute;top:340px;left:500px;width:10px;height:10px;' > "
sHTML = sHTML & "   <v:fill on='True' color='#ffffff' color2='#f7aa22' method='sigma' type='gradient' angle='30' opacity='90%'  /> "
sHTML = sHTML & "   <v:stroke on='False' /> "
sHTML = sHTML & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
sHTML = sHTML & "   <v:path textpathok='True' /> "
sHTML = sHTML & "     <v:textpath on='True' string='6' style=""font-size:'50pt';font-family:'WINGdings';font-weight:bold;""  /> "
sHTML = sHTML & "   <v:extrusion on='True' backdepth='4' color='#eeeeee' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
sHTML = sHTML & " </v:line>"


 Whtml1 = ""


 Whtml1 = Whtml1 & "  <v:roundrect arcsize='55%' id='THM'  style='z-index:0;position:absolute;top:180;left:100;width:650px;height:220px;' >"
 Whtml1 = Whtml1 & "     <v:stroke on='True' color='#eeeeee' weight='2px' />"	
 Whtml1 = Whtml1 & "     <v:fill on='True' method='sigma' color2='"  & Session("CMFMENU") & "' color='"  & Session("CMFMENU") & "' type='gradient' angle='30' opacity='60%' focusposition=""95%, 15%""  />"
 Whtml1 = Whtml1 & "     <v:shadow on='True' offset='10px,10px' color='#eeeeee' weight='5px' />"
 Whtml1 = Whtml1 & "  </v:roundrect>"


 Whtml1 = Whtml1 & "     <v:oval id='O1' style='align:left;z-index:82;position:absolute;left:570px;top:350px;width:100px;height:20px;' >"
 Whtml1 = Whtml1 & "       <v:fill on='True' color='#f7aa22' color2='#cc0000' type='gradient' angle='70' opacity='80%' focus='10%' focusposition='45%, 50%' focussize='10%, 10%' />"
 Whtml1 = Whtml1 & "       <v:stroke on='false' />	"
 Whtml1 = Whtml1 & "         <v:path textpathok='True' />"
 Whtml1 = Whtml1 & "       <v:textpath on='True' string='E' style='font-size:20pt;font-family:tahoma;font-weight:bold;font-style:italic;'  />"
 Whtml1 = Whtml1 & "           <v:extrusion on='true' backdepth='0' metal='true'  rotationcenter='0,0,0'  rotationangle = '20,0'/>"
 Whtml1 = Whtml1 & "     </v:oval>"

 Whtml1 = Whtml1 & "     <v:oval id='O2' style='align:left;z-index:81;position:absolute;left:580px;top:350px;width:100px;height:20px;' >"
 Whtml1 = Whtml1 & "       <v:fill on='True' color='#aa0000' color2='#993333' type='gradient' angle='70' opacity='80%' focus='10%' focusposition='45%, 50%' focussize='10%, 10%' />"
 Whtml1 = Whtml1 & "       <v:stroke on='false' />	"
 Whtml1 = Whtml1 & "         <v:path textpathok='True' />"
 Whtml1 = Whtml1 & "       <v:textpath on='True' string='S' style='font-size:20pt;font-family:tahoma;font-weight:bold;font-style:italic;'  />"
 Whtml1 = Whtml1 & "           <v:extrusion on='true' backdepth='0' metal='true'  rotationcenter='0,0,0'  rotationangle = '20,0'/>"
 Whtml1 = Whtml1 & "     </v:oval>"

 Whtml1 = Whtml1 & "     <v:oval id='O3' style='align:left;z-index:80;position:absolute;left:590px;top:350px;width:100px;height:20px;' >"
 Whtml1 = Whtml1 & "       <v:fill on='True' color='#336699' color2='#6699cc' type='gradient' angle='70' opacity='80%' focus='10%' focusposition='45%, 50%' focussize='10%, 10%' />"
 Whtml1 = Whtml1 & "       <v:stroke on='false' />	"
 Whtml1 = Whtml1 & "         <v:path textpathok='True' />"
 Whtml1 = Whtml1 & "       <v:textpath on='True' string='S' style='font-size:20pt;font-family:tahoma;font-weight:bold;font-style:italic;'  />"
 Whtml1 = Whtml1 & "           <v:extrusion on='true' backdepth='0' metal='true'  rotationcenter='0,0,0'  rotationangle = '20,0'/>"
 Whtml1 = Whtml1 & "     </v:oval>"



sHTML = sHTML & Whtml1

''''''''''''''''''''''''''''''
else
''''''''''''''''''''''''''''''''''''''''''''

sHTML = sHTML & " <div id='LGM0' style=""z-index:9;position:relative;top:30px;background-color:#ffffff;"" > "
sHTML = sHTML & "<table style='" & XspSShw("#cccccc","125") & ";' ><td style='line-height:20pt;font-size:16pt;font-family:verdana;color:#6699CC;' ><b><i>&nbsp;DESCONECCION&nbsp;EN&nbsp;PROCESO&nbsp;...&nbsp;</i></b></td></table></div>"


sHTML = sHTML & " <p id='LGM1' style='z-index:9;position:absolute;top:200px;left:550px;width:10px;height:10px;" & XspDShw("2","2","#cccccc") & ";' > "
sHTML = sHTML & "<FONT  STYLE=""font-size:'50pt';font-family:'webdings';font-weight:bold;color:#336699;"" >&#0128;</FONT></p>"

sHTML = sHTML & " <p id='LGM2' style='z-index:3;position:absolute;top:190px;left:580px;width:10px;height:10px;" & XspDShw("2","2","#cccccc") & ";' > "
sHTML = sHTML & "<FONT  STYLE=""font-size:'50pt';font-family:'webdings';font-weight:bold;color:#AA0000;"" >r</FONT></p>"

sHTML = sHTML & " <p id='LGM3' style='z-index:3;position:absolute;top:260px;left:510px;width:10px;height:10px;line-height:20pt;" & XspDShw("2","1","#cccccc") & ";' > "
sHTML = sHTML & "<FONT  STYLE=""font-size:'10pt';font-family:'VERDANA';font-weight:bold;color:#FFFFFF;"" >FINAL&nbsp;DE&nbsp;SESION</FONT></p>"

 Whtml1 = ""



 Whtml1 = Whtml1 & "  <DIV id='THMX'  style='z-index:0;position:absolute;top:180;left:100;width:650px;height:220px;background-color:" & session("CMFPNTIT") & ";" & XspAlpha(40,100,1) & ";' >"
 Whtml1 = Whtml1 & "<P >&nbsp;</P>"	
 Whtml1 = Whtml1 & "  </DIV>"

 Whtml1 = Whtml1 & "     <div id='xx10' style='z-index:3;position:absolute;top:55;left:5;' >"
 Whtml1 = Whtml1 & "<FONT  STYLE=""font-size:'320pt';font-family:'wingdings';font-weight:100;color:#cccccc;"" >&#0162;</FONT></div>"

 Whtml1 = Whtml1 & "     <div id='xx14' style='z-index:3;position:absolute;top:-24;left:465;clip:rect(auto auto auto 180px);' >"
 Whtml1 = Whtml1 & "<FONT  STYLE=""font-size:'365pt';font-family:'verdana';font-weight:800;color:#FFFFFF;"" >D</FONT></div>"

 Whtml1 = Whtml1 & "     <div id='xx11' style='z-index:2;position:absolute;top:178;left:90;" & XspSShw("#dddddd","110") & ";' >"
 Whtml1 = Whtml1 & "<FONT  STYLE=""font-size:'158pt';font-family:'webdings';font-weight:100;color:#eeeeee;"" >&#0254;</FONT></div>"

sHTML = sHTML & " <div id='xx12' style=""z-index:4;position:absolute;top:250;left:70;"" > "
sHTML = sHTML & "<table cellspacing=0 cellpadding=0 style='" & XspSShw("#cccccc","125") & ";' >"
sHTML = sHTML & "<tr>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "</tr>"
sHTML = sHTML & "<tr>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "</tr>"
sHTML = sHTML & "<tr>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "</tr>"
sHTML = sHTML & "<tr>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#6699cc;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "<td style='line-height:13pt;font-size:18pt;font-family:wingdings;color:#993333;' >"
sHTML = sHTML & "n</td>"
sHTML = sHTML & "</tr>"
sHTML = sHTML & "</table></div>"


 Whtml1 = Whtml1 & "     <p id='OO1' style='align:left;z-index:82;position:absolute;left:590px;top:350px;line-height:24pt;" & XspDShw("2","1","#cccccc") & ";' >"
 Whtml1 = Whtml1 & "<FONT  STYLE=""font-size:22pt;font-family:tahoma;font-weight:bold;font-style:italic;color:#F7AA22;"" >E</FONT></p>"

 Whtml1 = Whtml1 & "     <p id='OO2' style='align:left;z-index:81;position:absolute;left:605px;top:350px;line-height:24pt;" & XspDShw("2","1","#cccccc") & ";' >"
 Whtml1 = Whtml1 & "<FONT  STYLE=""font-size:22pt;font-family:tahoma;font-weight:bold;font-style:italic;color:#AA0000;"" >S</FONT></p>"

 Whtml1 = Whtml1 & "     <p id='OO3' style='align:left;z-index:80;position:absolute;left:620px;top:350px;line-height:24pt;" & XspDShw("2","1","#cccccc") & ";' >"
 Whtml1 = Whtml1 & "<FONT  STYLE=""font-size:22pt;font-family:tahoma;font-weight:bold;font-style:italic;color:#336699;"" >S</FONT></p>"



sHTML = sHTML & Whtml1
'''''''''''''''''''''''''''''''''''''''''''
end if




sHTML = sHTML & "<script language=javascript >" & chr(13)

sHTML = sHTML & " var tt=0;          "  & chr(13)

sHTML = sHTML & " function FMtmp() { "  & chr(13)
 if  SESSION("FULLGUI") = "S"  and SESSION("RDCGUID") <> "S" THEN
sHTML = sHTML & " if (document.all.LGM4) {            "  & chr(13)
sHTML = sHTML & "  tt=tt + 10;       "  & chr(13)
sHTML = sHTML & " if (tt>=360) {tt=0;};"  & chr(13)
''sHTML = sHTML & " document.all.LGM4.extrusion.rotationangle = tt + ',0';"  & chr(13)
sHTML = sHTML & " document.all.LGM4.extrusion.rotationangle = '0,' + tt;"  & chr(13)
''sHTML = sHTML & " document.all.LGM4.extrusion.rotationangle = tt + ',' + tt;"  & chr(13)
sHTML = sHTML & "    }               "  & chr(13)    
sHTML = sHTML & " setTimeout(""FMtmp()"",100); "  & chr(13)
end if
sHTML = sHTML & " }                  "  & chr(13)      

sHTML = sHTML & " function reinicia() { "  & chr(13)
sHTML = sHTML & " document.parentWindow.top.navigate('../coress20/inicio.asp?opcion=IX'); "  & chr(13)
sHTML = sHTML & " } "  & chr(13)

sHTML = sHTML & "</script>"  & chr(13)

      End If

sHTML = sHTML & "</BODY></html>"
    
      LogFinal = sHTML
    

End Function




'''''''''''

%>