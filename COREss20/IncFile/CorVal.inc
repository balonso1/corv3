<%
Function Valida()


' devuelve:
' 0 - ok
' 1 - Usuario no valido
' 2 - Password no valida
' 3 - Usuario bloqueado por Intentos Fallidos
' 4 - Forzar a Cambio de password por Expiracion
' 5 - Forzar a Cambio de password por Password Forzada por Admininistrador
' 6 - Usuario bloqueado por Administrador
' 7 - Restriccion sabados y domingos
' 8 - Restriccion sabados, domingos y feriados
' 9 - fuera del rango horario
'10 -Trusted




    
      Valida = ""

     ' Session("Usuario") = ""
      Session("UsrOpcion") = ""
      Session("SkrDefT") = "0"
      Session("SkrDefE") = "0"
      Session("SkrDefI") = "0"
      Session("UsrPmin") = "0"
      Session("UsrPmax") = "0"
      Session("UsrHist") = "0"

      vu = OFV.ValidarUsuario()
      
      If vu = "OK" Then
          
       
'trusted
        
         sHTML = "0"
         
         


              Set cn = OFV.conectar(OFV.strconn0)


              StrSql = "Select * From skrdef"
              Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
              If Not rs.EOF Then
                 Session("UsrPmin") = CInt(rs("PassMin"))
                 Session("UsrPmax") = CInt(rs("PassMax"))
                 Session("UsrHist") = CInt(rs("PassHist"))
                 Session("SkrDefT") = CInt(rs("tipoesq"))
                If CInt(rs("expira")) = 0 Then
                    Session("SkrDefE") = 0
                 Else
                    Session("SkrDefE") = CInt(rs("diasexp"))
                 End If
                 Session("SkrDefI") = CInt(rs("intfall"))
              End If
              Call OFV.cerrarconsulta(rs)

              StrSql = "Select * From Usuarios Where id = '" & Session("Usuario") & "'"
           '   response.write strsql & " - " & Session("Usuario")
              Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
              

              sHTML = "0"
              If rs.EOF Then
                  sHTML = "1"
              Else
' verifica rangos horarios

                   rango = rs("rangos")
                   If Len(rango) = 0 Then rango = 1
                   rango = clng(rango)
                   If rango = 0 Then rango = 1
                                 
                   StrSql = "Select * From rangos Where codigo = " & rango
                   Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
                   If rs2.EOF Then
                      StrSql = "Select * From rangos Where codigo = 1"
                      Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
                   End If
                   If rs2.EOF Then sHTML = "999"
                   
                   If sHTML = "0" Then
                      Session("SDIAS") = rs2("restdias")
                      Session("SHORAS") = rs2("resthoras")
                      Session("SHD1") = rs2("hd1")
                      Session("SHH1") = rs2("hh1")
                      Session("SHD2") = rs2("hd2")
                      Session("SHH2") = rs2("hh2")
                      Session("SHD3") = rs2("hd3")
                      Session("SHH3") = rs2("hh3")
                      
                      swd = DatePart("w", Date)
' si SWD (dia de la semana) es 1 - domingo o 7 - sabado

                   rango = CLng(Date) - CLng(CDate("01-01-1900"))

' verifica si hoy fue informado como feriado

                   StrSql = "Select * From feriados Where fecha = " & rango
                   Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
                   If rs2.EOF Then
                      Session("SHOY") = "N"
                   Else
                      Session("SHOY") = "S"
                      Session("SHOYD") = rs2("observaciones")
                   End If
                   Call OFV.cerrarconsulta(rs2)
                   
' verifica restriccion dias

                    Select Case CInt(Session("SDIAS"))
                           Case 0
                                sHTML = "0"
                           
                           Case 1
                                If swd = 1 Or swd = 7 Then
                                   sHTML = "7"
                                End If
                           
                           Case 2
                                If swd = 1 Or swd = 7 Or Session("SHOY") = "S" Then
                                   sHTML = "8"
                                End If
                    
                    End Select
                    
' verifica restriccion horaria

                    If sHTML = "0" Then

                    shora = FormatDateTime(Now(), vbShortTime)
                    shora = CStr(Trim(shora))
                    
                    Select Case CInt(Session("SHORAS"))
                           Case 0
                                sHTML = "0"
                           
                           Case 1
                                sHTML = "0"
                                If Session("SHD1") <> "NO" Then
                                   
                                   If shora > Session("SHD1") And shora < Session("SHH1") Then
                                      sHTML = "0"
                                   Else
                                      sHTML = "9"
                                   End If
                                   
                                   If sHTML <> "0" Then
                                      If Session("SHD2") <> "NO" Then
                                         
                                         If shora > Session("SHD2") And shora < Session("SHH2") Then
                                            sHTML = "0"
                                         Else
                                            sHTML = "9"
                                         End If
                                         
                                         If sHTML <> "0" Then
                                            If Session("SHD3") <> "NO" Then
                                               If shora > Session("SHD3") And shora < Session("SHH3") Then
                                                  sHTML = "0"
                                               Else
                                                  sHTML = "9"
                                               End If
                                            End If
                                         End If
                                         
                                   
                                      End If
                                   End If
                                
                                End If
                           
                    
                    End Select
                   
                   
                   
                   End If
                   
                   End If
              
              
' fin verifica rangos horarios
                   If sHTML = "0" Then

                  If CInt(rs("defskr")) = 1 Then
                     Session("UsrPmin") = CInt(rs("PassMin"))
                     Session("UsrPmax") = CInt(rs("PassMax"))
                     Session("UsrHist") = CInt(rs("PassHist"))
                     Session("SkrDefT") = CInt(rs("tipoesq"))
                     If CInt(rs("expira")) = 0 Then
                        Session("SkrDefE") = 0
                     Else
                        Session("SkrDefE") = CInt(rs("diasexp"))
                     End If
                     Session("SkrDefI") = CInt(rs("intfall"))
                  End If
                  Session("UsrIntFail") = CInt(rs("LogFail"))
                  Session("UsrOpcion") = CInt(rs("opcion"))
                                   
               If Session("UsrOpcion") = 2 Then
                     Call OFV.cerrarconsulta(rs)
                     sHTML = "3"
               ElseIf Session("UsrOpcion") = 3 Then
                     Call OFV.cerrarconsulta(rs)
                     sHTML = "6"
               Else
        
                '  Set Encripter = Server.CreateObject("CorEncriptadorEx.SHA1")

                 ' sEncript = Encripter.Encrypt(xclavein)
                 ' If rs(2) = sEncript Then
                      Session("SegUsuario") = rs(1)
                      Session("SegNivel") = rs(4)
                      Session("SegGrp") = rs(3)

                      Session("SegPerfil") = Session("Usuario")

                      If Session("UsrOpcion") = 1 Then
                         Call OFV.cerrarconsulta(rs)
                         sHTML = "0"
                      Else
                          If Session("SkrDefT") = 1 And Session("SkrDefE") > 0 Then
                             If Not (CLng(rs("lchange")) + Session("SkrDefE")) > (CLng(Date) - CLng(CDate("01-01-1900"))) Then
                                Call OFV.cerrarconsulta(rs)
                                StrSql = "Update Usuarios set OPCION = 4 "
                                StrSql = StrSql & " Where id = '" & Session("Usuario") & "'"
                                Call OFV.crearconsultaEx(StrSql, cn, 1, "")
                                sHTML = "4"
                             Else
                                 Call OFV.cerrarconsulta(rs)
                             End If
                          Else
                             Call OFV.cerrarconsulta(rs)
                          End If
                      End If


                      

                  ' End If
                End If
                
                End If
                
              End If
            

       
      Else
          sHTML = vu
      End If

      Session("SkrOpc") = sHTML
      Valida = sHTML
    
      Exit Function


End Function
%>