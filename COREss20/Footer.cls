VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Footer"
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
    
22      Dim sHTML$, sInput$, Celda$, Fila$, StrSql$, sOpcion$, n%, XFondo$, Formulario$
        Dim ind, atrok, apllet$, Version$, AplMain$
23      Dim cn As Object, rs As Object, rs1 As Object, rs2 As Object
        Dim cx As Object, rx As Object

    
24      Generar = ""

25      Response.expires = 0

26      sOpcion = Request.querystring("Opcion")

        If Session("USRES") <> "B" Then XFondo = " background='CorFondos/" & Session("CMIMENU") & "' "

27      sHTML = "<html>"
            sHTML = sHTML & "<style>"

            sHTML = sHTML & " td.bt {color:" & Session("CMFMNFT") & ";font-family:" & Session("CMFMNFTFC") & ";font-size:8pt;}"
            sHTML = sHTML & " a.bt {font-family:" & Session("CMFMNFTFC") & ";font-size:8pt;color:" & Session("CMFMNFT") & ";}"
            sHTML = sHTML & "</style>"
            sHTML = sHTML & "<body bgcolor=" & Session("CMFMENU") & " " & XFondo & " style=" & Chr(34) & "border-style:outset;border-width:thin;" & Chr(34)



            Set cx = OFV.conectar(OFV.strconn1)
            StrSql = "select * from CorAplic where Codapli= '" & Session("TOPAPL") & "'"
            Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
            If Not rx.EOF Then
            
               AplMain = Replace(rx(2), "*", " ") & " " & rx(4)
               Version = rx(4)
               
            Else
               AplMain = "Cor E-Solution Suite 2.8"
               Version = "2.8"
            End If



28      sInput = "<img border=0 src=imagenes/corlogox1.jpg>"
        Celda = OFV.gencelda("", "", "", "", "", "", "", sInput, "si")
29      sInput = "<font color=" & Session("CMFMEFT") & " size=2 face=" & Session("CMFMEFTFC") & " ><b>" & Replace(AplMain, " ", "&nbsp;") & "</b></font>"
        Celda = Celda & OFV.gencelda("", "", "", "90%", "", "", "", sInput, "si")


32      If sOpcion = "I" Then
            sHTML = sHTML & " >"
33          Fila = OFV.genrow("", "", "", "", "", "", "", Celda, "si")
34          sHTML = sHTML & OFV.GenTabla("", "", "", "100%", "100%", "", "", "-1", "-2", "", "", Fila, "si")
35      Else
            XFondo = "N"
            sHTML = sHTML & " onload=" & Chr(34) & "document.initaplx.submit();" & Chr(34) & " >"

40          Dim aAplicacion()
            Dim aLinks()
            Dim aLinksImg()
            Dim aLinksLet()

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
                     
                     aAplicacion(ind) = rx(1)
                     aLinks(ind) = "<a class=bt title='" & rx(12)
                     aLinks(ind) = aLinks(ind) & "' href=" & rx(11)
                     If UCase(rx(11)) = UCase("Menus.asp") Then
                        aLinks(ind) = aLinks(ind) & "?Opcion=" & UCase(rx(1))
                     End If
                     aLinks(ind) = aLinks(ind) & " target=" & rx(10)
                     aLinks(ind) = aLinks(ind) & ">"
                     aLinksImg(ind) = "<img border=0 src=Imagenes/" & rx(6) & ">"
                     apllet = Replace(rx(2), "*", "")
                     aLinksLet(ind) = "<b>" & Replace(apllet, " ", "&nbsp;") & "</b>"
                     
'                     aLinks(ind) = aLinks(ind) & "</a>"
                     
                     ind = ind + 1
                     rx.movenext
                  Loop
                 
                  
               Else
                  
            ReDim Preserve aAplicacion(4)
41          aAplicacion(0) = "CORESS20"
42          aAplicacion(1) = "CORCHT20"
43          aAplicacion(2) = "COREP20"
44          aAplicacion(3) = "COREM20"
45          aAplicacion(4) = "COREC20"
        
46          ReDim Preserve aLinks(4)
47          aLinks(0) = "<a title='COR&nbsp;ESS&nbsp;2.0' href=Menus.asp?Opcion=CORESS20 target=Menus><img border=0 src=Imagenes/iCorESS.jpg></a>"
48          aLinks(1) = "<a title='COR&nbsp;Chat&nbsp;2.0' href=../CORChat20/chat.htm target=_blank><img border=0 src=Imagenes/iCorChat2.jpg></a>"
49          aLinks(2) = "<a title='COR&nbsp;EPlanning&nbsp;2.0' href=Menus.asp?Opcion=COREP20 target=Menus><img border=0 src=Imagenes/iCorplanning.jpg></a>"
50          aLinks(3) = "<a title='COR&nbsp;EManager&nbsp;2.0' href=Menus.asp?Opcion=COREM20 target=Menus><img border=0 src=Imagenes/iCormanager.jpg></a></td>"
51          aLinks(4) = "<a title='COR&nbsp;EComponent&nbsp;2.0' href=Menus.asp?Opcion=COREC20 target=Menus><img border=0 src=Imagenes/iCorComponent.jpg></a>"
        
            End If
            End If
            Call OFV.cerrarconsulta(rx)
            
            
        If Session("USRES") <> "B" Then
           sInput = "<a title='Ayuda' href=../COREss20/Ayuda/NormasHelp.asp target=info><img border=0 src=Imagenes/iCorSHelp.jpg></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a title='About' href=CorAbout.asp target=info><img border=0 src=Imagenes/iCorAbout.jpg></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a title='Usuario' href=InfoUser.asp target=info><img border=0 src=Imagenes/iCorSInfoUser.jpg></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
        Else
           sInput = "<a  class=bt title='Ayuda' href=../COREss20/Ayuda/NormasHelp.asp target=info><b>Ayuda</b></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a  class=bt title='About' href=CorAbout.asp target=info><b>Acerca&nbsp;de</b></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a  class=bt title='Usuario' href=InfoUser.asp target=info><b>Usuario</b></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
        
        End If
            
        
52          Set cn = OFV.conectar(OFV.strconn0)

53          For n = 0 To UBound(aAplicacion)

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
              StrSql = StrSql & " And grupo = '" & Session("seggrp") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & aAplicacion(n) & "'"
              Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
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
                            If XFondo = "N" Then
                               XFondo = ""
                               sInput = OFV.GenerarInput("boton2", "menu", "hidden", "", "10", "", "")
                               Formulario = OFV.GenForm("initaplx", "", "Menus.asp?Opcion=" & aAplicacion(n), "Menus", "post", "", sInput, "si")
                            End If
                            
                 If Session("USRES") <> "B" Then
                    sInput = aLinks(n) & aLinksImg(n) & "</a>"
                 Else
                    sInput = aLinks(n) & aLinksLet(n) & "</a>"
                 End If
              
                 Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
              End If

             Next

61          Fila = OFV.genrow("", "", "", "", "", "", "", Celda, "si")
           If Session("USRES") <> "B" Then
62             sHTML = sHTML & OFV.GenTabla("", "", "", "100%", "100%", "", "", "0", "-1", "", "", Fila, "si")
           Else
               sHTML = sHTML & OFV.GenTabla("", "", "", "100%", "100%", "", "", "1", "1", "", "", Fila, "si")
           End If
63      End If
64      sHTML = sHTML & Formulario & "</body></html>"

65      Generar = sHTML
    
66      Exit Function

67 GenerarErr:
68      Generar = Err.Description & " " & Err.Source & " " & CStr(Erl)

End Function
