VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Submenus"
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
    
22      Dim sHTML$, sInput$, Celda$, Fila$, StrSql$, sOpcion$, n%
23      Dim Menu$, Submenu$, sMenu$, sHref$, sLink$, sTarget$, sPath$
24      Dim Formulario$, spt%, sseg%, xesm, XFondo$, sAccion$
        Dim xsel, atrok
25      Dim cn As Object, rs As Object, cx As Object, rs1 As Object, rs2 As Object
    
26      Generar = ""

27      Response.expires = 0

        If Session("USRES") <> "B" Then XFondo = "background='CorFondos/" & Session("CMIMENU") & "' "

28      sHTML = OFV.Menuheader(Session("CMFMENU"), XFondo & " style='border-style:outset;border-width:1;' onload=document.nulo.submit();")


      Menu = Request.querystring("Menu")
      If Menu <> "I" Then
          Submenu = Request.querystring("ID")
          Menu = OFV.convertircar(Menu, "_", " ")
          sMenu = "Menu " & Menu
            Session("cormenu") = Menu
  
          sInput = "<font color=" & Session("CMFMEFT") & " face=" & Session("CMFMEFTFC") & " size=1><b>" & Menu & "</b></font>"
          Celda = OFV.gencelda("", "aaa  bgcolor=" & Session("CMFMENU") & " style='border-style:ridge;border-width:2;'", "center", "", "20", "", "", sInput, "si")
          Fila = OFV.genrow("", "", "", "", "", "", "", Celda, "si")

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

              If atrok = "O" Then
              StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'M' "
              StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
              StrSql = StrSql & " And grupo = '" & Session("seggrp") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & Session("CorApli") & "'"
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


                          sInput = Trim(rs(xsel))
                          sHref = Trim(rs(2))
                              sTarget = "info"
                              sPath = Session("CorPath")
                              If rs(3) <> "" Then sTarget = rs(3)
                              spt = InStr(rs(2), ".asp")
                              xesm = "N"
                              If spt > 1 Then
                                  If Left(rs(2), spt - 1) = "Submenus" Then
                                     sPath = ""
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
                              sLink = OFV.GenLink("", "ts", sPath & sHref, sTarget, "", "", UCase(rs(1)), "si")
                              Celda = OFV.gencelda("", "ts  ", "", "", "20", "", "", sLink, "si")
                              Fila = Fila & OFV.genrow("", "", "", "", "", "", "", Celda, "si")


                  End If

                          


              rs.movenext
          Loop
          Call OFV.cerrarconsulta(rs)
          Call OFV.cerrarconn(cn)
          Call OFV.cerrarconn(cx)



          sHTML = sHTML & OFV.GenTabla("", "", "", "99% ", "", "0", "", "0", "0", "#cccccc", "", Fila, "si")



      Else
          sHTML = sHTML & "<center><br><br>"

          Set cn = OFV.conectar(OFV.strconn1)
          StrSql = "Select * from coraplic"
          Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
          If Not rs.EOF Then
             If rs.fields.Count > 8 Then
                Call OFV.cerrarconsulta(rs)
                StrSql = "Select * from coraplic where codapli = '" & Session("CorApli") & "' "
                Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
                If Not rs.EOF Then
                   sHTML = sHTML & "<img src=Imagenes/" & rs(8) & ">"
                End If
          
             Else
          
              Select Case Session("CorApli")
                Case "CORESS20"
                  sHTML = sHTML & "<img src=Imagenes/COREssuite20.jpg>"
                Case "COREM20"
                  sHTML = sHTML & "<img src=Imagenes/norpcm2.jpg>"
                Case "COREP20"
                  sHTML = sHTML & "<img src=Imagenes/Planning.jpg>"
                Case "COREC20"
                  sHTML = sHTML & "<img src=Imagenes/Component.jpg>"
             End Select
              End If
         End If
             Call OFV.cerrarconsulta(rs)
             Call OFV.cerrarconn(cn)
         
          sHTML = sHTML & "</center>"
      End If



      sAccion = "inicio.asp?Opcion=I&ddat="


      Celda = OFV.gencelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Fila = OFV.genrow("", "", "", "", "20", "", "", Celda, "si")
      sInput = OFV.GenerarInput("boton", "Logoff", "submit", "", "20", "bt", "")
      Formulario = OFV.GenForm("logoff", "", sAccion, "", "post", "", sInput, "si")
      Celda = OFV.gencelda("", "rr", "center", "", "", "", "", Formulario, "si")
      Fila = Fila & OFV.genrow("", "", "", "", "10", "", "", Celda, "si")

      sInput = OFV.GenerarInput("boton", "", "hidden", "", "20", "it", "")
      Formulario = OFV.GenForm("nulo", "", "nulo.asp", "info", "post", "", sInput, "si")
      Celda = OFV.gencelda("", "rr", "center", "", "", "", "", Formulario, "si")
      Fila = Fila & OFV.genrow("", "", "", "", "10", "", "", Celda, "si")

      sHTML = sHTML & OFV.GenTabla("", "", "", "100%", "", "0", "", "0", "0", "", "", Fila, "si")

      shtml = shtml & "<SCRIPT language=javascript >"
      sHTML = sHTML & " var xaction = document.all.logoff.action; "
      sHTML = sHTML & " document.all.logoff.action = xaction + window.top.CorUsr + '___' "
      shtml = shtml & " + window.top.CorNetUsr  + '___' + window.top.CorDomain + '___' + window.top.CorLogid;  "
      shtml = shtml & "</SCRIPT>"

     shtml = shtml  & "</body></html>"

116      Generar = sHTML
    
117      Exit Function

118 GenerarErr:
119      Generar = Err.Description & " " & Err.Source & " " & CStr(Erl)

End Function
