VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_oOFV As Object
Private m_oServerVar As Object
Private m_oRequestVar As Object
Private m_oResponseVar As Object
Private m_oSessionVar As Object

Public Function UserV()

10      UserV = OFV.uv

End Function

Public Property Set OFV(ByVal oOFV As Object)
    
11      Set m_oOFV = oOFV

End Property

Public Property Get OFV() As Object
    
12      Set OFV = m_oOFV

End Property

Public Property Set Server(ByVal oServer As Object)
    
13      Set m_oServerVar = oServer

End Property

Public Property Get Server() As Object
    
14      Set Server = m_oServerVar

End Property

Public Property Set Request(ByVal oRequest As Object)
    
15      Set m_oRequestVar = oRequest

End Property

Public Property Get Request() As Object
    
16      Set Request = m_oRequestVar

End Property

Public Property Set Response(ByVal oResponse As Object)
    
17      Set m_oResponseVar = oResponse

End Property

Public Property Get Response() As Object
    
18      Set Response = m_oResponseVar

End Property

Public Property Set Session(ByVal oSession As Object)
    
19      Set m_oSessionVar = oSession

End Property

Public Property Get Session() As Object
    
20      Set Session = m_oSessionVar

End Property

Public Function Generar() As String

21      On Error GoTo GenerarErr

22      Dim sHTML$, sInput$, Celda$, Fila$, StrSql$, sOpcion$, sApli$, n%
23      Dim sHref$, sLink$, Formulario$, fondo$
        Dim xsel, cnm, atrok, rs1 As Object, rs2 As Object
24      Dim cn As Object, rs As Object, cx As Object
    
25      Generar = ""

26      Response.expires = 0

        
       
        fondo = " bgcolor=" & Session("CMFMENU")
        If Session("USRES") <> "B" Then fondo = fondo & " background='CorFondos/" & Session("CMIMENU") & "' "

27      sHTML = "<html>" & OFV.FormStylerCon()
28      sHTML = sHTML & "<body topmargin=1 leftmargin=1 rightmargin=1 bottommargin=1 bgcolor=" & Session("CMFMENU") & " style=" & Chr(34) & "border-style:outset;border-width:thin;" & Chr(34)
        If Session("USRES") <> "B" Then sHTML = sHTML & " background='CorFondos/" & Session("CMIMENU") & "' "
        sHTML = sHTML & " onload=" & Chr(34) & " document.initaplx.submit();" & Chr(34) & " >"

        sHTML = sHTML & OFV.GenTabla("", "aaa valign=middle ", "", "", "", "0", "", "-1", "0", "", "", "", "no")

        sHTML = sHTML & OFV.genrow("", "", "", "", "", "", "", "", "no")

        sHTML = sHTML & "<td  width=2% align=center valign=middle style='filter: Shadow(color=#999999,direction=112);' >"
        sHTML = sHTML & "<img src=" & Session("CliLogo") & " border=0  align=center></td>"
        sHTML = sHTML & "<td valign=middle  width=40%  >"
        sHTML = sHTML & "<font face=" & Session("CMFMEFTFC") & " size=1 color=" & Session("CMFMEFT") & " ><b>&nbsp;" & Session("Cliente") & "</b></font></td>"


29      sOpcion = Trim(Request.querystring("Opcion"))
        Celda = "N"
30      If sOpcion = "I" Then
31          sOpcion = ""

33          sHTML = sHTML & "<td  ><p align=center><font face=" & Session("CMFMEFTFC") & " size=1 color=" & Session("CMFMNFT") & " ><B>Bienvenido a COR ESolution Suite &reg - por favor ingrese su Usuario y Password</b></font></p></td>"
34          Session("CorApli") = ""
35          Session("CorPath") = ""
36      Else
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

          If sApli = "CORESS20" Then
              Session("CorPath") = ""
          Else
              StrSql = "Select DirApli,titapli From CorAplic where CorAplic.CodApli = '" & sApli & "'"
              Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
              If Not rs.EOF Then
                  Session("CorPath") = "../" & Trim(rs(0)) & "/"
                  Session("CorTitApli") = rs(1)

              Else
                  Session("CorPath") = ""
                  Session("CorTitApli") = " "

              End If
              Call OFV.cerrarconsulta(rs)
          End If

          sHTML = sHTML & OFV.gencelda("", "", "center", "1", "", "", "", "", "si")

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
              StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'A' "
              StrSql = StrSql & " And grupo = '" & Session("seggrp") & "' And "
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

              If atrok = "O" Then
              StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'S' "
              StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
              StrSql = StrSql & " And grupo = '" & Session("seggrp") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & sApli & "'"

              Set rs2 = OFV.crearconsultaEx(StrSql, cx, 1, "")
              If rs2.EOF Then
                 atrok = "O"
              ElseIf rs2(1) Then
                     atrok = "S"
                  Else
                     atrok = "N"
              End If
              Call OFV.cerrarconsulta(rs2)
              End If

              If atrok = "O" Or atrok = "S" Then
                            If Celda = "N" Then
                               Celda = ""
                               sInput = OFV.GenerarInput("boton2", "menu", "hidden", "", "10", "", "")
                               Formulario = OFV.GenForm("initaplx", "", "Submenus.asp?ID=" & rs(6) & "&Menu=" & Replace(UCase(rs(1)), " ", "_"), "Master", "post", "", sInput, "si")
                            End If
              
                          sInput = "<b>" & UCase(rs(1)) & "</B>"
                          sHref = "Submenus.asp?ID=" & rs(6) & "&Menu=" & Replace(UCase(rs(1)), " ", "_")
                          sLink = OFV.GenLink("", "ts", sHref, "Master", "", "", sInput, "si")
                          Celda = Celda & OFV.gencelda("", "ts", "center", "", "", "", "", sLink, "si")
                          Celda = Celda & OFV.gencelda("", "", "center", "1", "", "", "", "", "si")
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


      End If

         sHTML = sHTML & Celda
         sHTML = sHTML & "<td width=10% >&nbsp;</td>"
         sHTML = sHTML & "<td width=20% valign=middle align=center>"
         sHTML = sHTML & OFV.CorLogo
         sHTML = sHTML & "</td>"
         sHTML = sHTML & OFV.genFRow()
         sHTML = sHTML & OFV.GenFTabla()
         


107      sHTML = sHTML & Formulario & "</body></html>"

108      Generar = sHTML
    
109      Exit Function

110 GenerarErr:
111      Generar = Err.Description & " " & Err.Source & " " & CStr(Erl)

End Function


