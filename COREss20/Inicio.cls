VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Modal As String


Private m_oOFV As Object
Private m_oServerVar As Object
Private m_oRequestVar As Object
Private m_oResponseVar As Object
Private m_oSessionVar As Object


Private sHTML$, vu$, sInput$, Celda$, Fila$, Formulario$, xwsize%
Private ok$, uv$, StrSql$, sEncript$
Private rango As Long
Private swd As Integer
Private shora As String
Private cn As Object, rs As Object, rs2 As Object, Encripter As Object


    
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



Public Function getrdm() As Integer

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

Public Function getrdma(crdm) As String

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

Public Function getsid() As Integer

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

Function Getcrpmk(sopcion, sdato, sid) As String

Dim xdatoi, datoS, y, i, xy, mascx

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



Public Function InitSec() As String

21      On Error GoTo InitSecErr
    
22      InitSec = ""

23      vu = OFV.ValidarUsuario()
24      If vu = "OK" Then
25          sHTML = OFV.MenuHeader("#336699", "background=imagenes/bancor5.jpg onload='document.init.submit();document.initf.submit();document.initI.submit();' ") & Chr(13)

26          xwsize = 15
27          If OFV.explorador = "MSIE" Then xwsize = 12

28          sInput = "<p align=center><font color=#DCDCDC face=tahoma size=1><b>INICIO</b></font></p>"
29          Celda = OFV.gencelda("", "", "", "100%", "20", "#336699", "Imagenes/topban.jpg", sInput, "si")
30          Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

31          sInput = "Ingrese su Usuario<br> y Password:<br><br><br>"
32          Celda = OFV.gencelda("", "rr", "center", "80%", "20", "#336699", "", sInput, "si")
33          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

            Formulario = Fila




52          sInput = OFV.GenerarInput("boton", "menu", "hidden", "", xwsize, "", "")
53          Celda = OFV.gencelda("", "", "", "", "", "", "", sInput, "si")
54          Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
55          Formulario = Formulario & OFV.GenForm("init", "", "Menus.asp?Opcion=I", "Menus", "post", "", Fila, "si")

56          sInput = OFV.GenerarInput("boton", "menu", "hidden", "", xwsize, "", "")
57          Celda = OFV.gencelda("", "", "", "", "", "", "", sInput, "si")
58          Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
59          Formulario = Formulario & OFV.GenForm("initf", "", "Footer.asp?Opcion=I", "Footer", "post", "", Fila, "si")

            sInput = OFV.GenerarInput("boton", "menu", "hidden", "", xwsize, "", "")
            Celda = OFV.gencelda("", "", "", "", "", "", "", sInput, "si")
            Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
            Formulario = Formulario & OFV.GenForm("initI", "", "Inicio.asp?Opcion=IX", "info", "post", "", Fila, "si")

60          sHTML = sHTML & OFV.GenTabla("", "", "", "100% border", "", "0", "", "0", "0", "#cccccc", "", Formulario, "si")

61          sHTML = sHTML & "</BODY></html>"
62      Else
63          sHTML = vu
64      End If

65      InitSec = sHTML
    
66      Exit Function

67 InitSecErr:
68      InitSec = Err.Description & " " & Err.Source & " " & CStr(Erl)

End Function
Public Function InitSecT() As String

21      On Error GoTo InitSecErr
    
22      InitSecT = ""

23      vu = OFV.ValidarUsuario()
24      If vu = "OK" Then
25          sHTML = OFV.MenuHeader("#336699", "background=imagenes/bancor5.jpg onload='document.init.submit();document.initf.submit();document.initI.submit();' ") & Chr(13)

26          xwsize = 15
27          If OFV.explorador = "MSIE" Then xwsize = 12

28          sInput = "<p align=center><font color=#DCDCDC face=tahoma size=1><b>INICIO</b></font></p>"
29          Celda = OFV.gencelda("", "", "", "100%", "20", "#336699", "Imagenes/topban.jpg", sInput, "si")
30          Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

31          sInput = "Ingrese su Usuario<br> y Password:<br><br><br>"
32          Celda = OFV.gencelda("", "rr", "center", "80%", "20", "#336699", "", sInput, "si")
33          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

            Formulario = Fila




52          sInput = OFV.GenerarInput("boton", "menu", "hidden", "", xwsize, "", "")
53          Celda = OFV.gencelda("", "", "", "", "", "", "", sInput, "si")
54          Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
55          Formulario = Formulario & OFV.GenForm("init", "", "Menus.asp?Opcion=I", "Menus", "post", "", Fila, "si")

56          sInput = OFV.GenerarInput("boton", "menu", "hidden", "", xwsize, "", "")
57          Celda = OFV.gencelda("", "", "", "", "", "", "", sInput, "si")
58          Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
59          Formulario = Formulario & OFV.GenForm("initf", "", "Footer.asp?Opcion=I", "Footer", "post", "", Fila, "si")

            sInput = OFV.GenerarInput("boton", "menu", "hidden", "", xwsize, "", "")
            Celda = OFV.gencelda("", "", "", "", "", "", "", sInput, "si")
            Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
            Formulario = Formulario & OFV.GenForm("initI", "", "Inicio.asp?Opcion=IZ", "info", "post", "", Fila, "si")

60          sHTML = sHTML & OFV.GenTabla("", "", "", "100% border", "", "0", "", "0", "0", "#cccccc", "", Formulario, "si")

61          sHTML = sHTML & "</BODY></html>"
62      Else
63          sHTML = vu
64      End If

65      InitSecT = sHTML
    
66      Exit Function

67 InitSecErr:
68      InitSecT = Err.Description & " " & Err.Source & " " & CStr(Erl)

End Function


Public Function Inicio(sopcion)

      On Error GoTo InicioErr

      Dim xopcion$, serrordes$, i
      Dim UserNt$, AplMain$, DOMAINNT$
      Dim fondo As String
      xopcion = sopcion
      fondo = " TOPMARGIN=1 LEFTMARGIN=1 "
      vu = OFV.ValidarUsuario()
      If vu = "OK" Then

        

          sHTML = OFV.MenuHeader("#BBBBBB", fondo) & Chr(13)

'''          sHTML = OFV.MenuHeader("#eeeeee", " background='imagenes/corf5.jpg' ") & Chr(13)

sHTML = sHTML & "<script language=javascript >" & Chr(13)
sHTML = sHTML & "function cript()" & Chr(13)
sHTML = sHTML & "{ " & Chr(13)
sHTML = sHTML & "var sid = " & getsid() & "; " & Chr(13)
sHTML = sHTML & "var i = 0;" & Chr(13)
sHTML = sHTML & "var tab2 = new Array("
genMkt

For i = 0 To 255
    If i > 0 Then sHTML = sHTML & ","
    sHTML = sHTML & xtabla(i)
Next
sHTML = sHTML & ");" & Chr(13)
sHTML = sHTML & "var Y = 0;" & Chr(13)
sHTML = sHTML & "var i = 0;" & Chr(13)
sHTML = sHTML & "var dato = document.all.clave.value;" & Chr(13)
sHTML = sHTML & "var ldato = dato.length;" & Chr(13)
sHTML = sHTML & "var datos = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "var masc = 0;" & Chr(13)
sHTML = sHTML & "if (ldato < 31) " & Chr(13)
sHTML = sHTML & " {for (i=1;i<=ldato;i++)" & Chr(13)
sHTML = sHTML & "   {" & Chr(13)
sHTML = sHTML & "     Y = dato.charCodeAt(i-1);" & Chr(13)
sHTML = sHTML & "     masc = Math.round(i * 7) + Math.round(i * 31 / 17) + sid;" & Chr(13)
sHTML = sHTML & "     datos = datos + String(Math.round(tab2[Y]) - masc);" & Chr(13)
sHTML = sHTML & "   } " & Chr(13)
sHTML = sHTML & "  document.all.clave.value = datos;" & Chr(13)
sHTML = sHTML & "  document.all.Inicio.submit();" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & " else " & Chr(13)
sHTML = sHTML & " {" & Chr(13)
sHTML = sHTML & "  document.all.clave.value = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "  alert('La password no puede exceder de 30 caracteres');" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & "}" & Chr(13)


sHTML = sHTML & "function abrir()"
sHTML = sHTML & "{"
sHTML = sHTML & "document.all.u20.style.display='none';"

sHTML = sHTML & "document.all.u10.style.display='block';"
sHTML = sHTML & "document.all.u11.style.display='block';"
sHTML = sHTML & "document.all.u12.style.display='block';"
sHTML = sHTML & "document.all.u13.style.display='block';"
sHTML = sHTML & "document.all.u14.style.display='block';"
sHTML = sHTML & "}"

sHTML = sHTML & "function cerrar()"
sHTML = sHTML & "{"
sHTML = sHTML & "document.all.u20.style.display='block';"

sHTML = sHTML & "document.all.u10.style.display='none';"
sHTML = sHTML & "document.all.u11.style.display='none';"
sHTML = sHTML & "document.all.u12.style.display='none';"
sHTML = sHTML & "document.all.u13.style.display='none';"
sHTML = sHTML & "document.all.u14.style.display='none';"
sHTML = sHTML & "}"

sHTML = sHTML & "</script>" & Chr(13)

sHTML = sHTML & "<center>"
sHTML = sHTML & "<DIV  CLASS='aFilter' ID='oDiv' width=80% height=80% >"

          xwsize = 45
          If OFV.explorador = "MSIE" Then xwsize = 42

          sInput = "<BIG><b>LOG IN</b></BIG>"
          Celda = OFV.gencelda("", OFV.TitCls & " style='border-style:outset;'", "center", "100%", "20", "", "", sInput, "si")
          Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
          
          
          sInput = "<B>INGRESE SU USUARIO Y PASSWORD:</B>"
          Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", sInput, "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

1         UserNt = UCase(Request.ServerVariables("LOGON_USER"))
2         UserNt = Mid(UserNt, InStrRev(UserNt, "\") + 1, Len(UserNt))

           DOMAINNT = UCase(Request.ServerVariables("LOGON_USER"))
           DOMAINNT = Left(DOMAINNT, InStr(DOMAINNT, "\")) & "\"
           DOMAINNT = Replace(DOMAINNT, "\", "")
           
          If Modal = "MSNTT" Then
           
3             Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", "<B>USUARIO NT:</B>", "si")
4             Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("usuarNt readonly style='border-style:outset;border-width:2;BACKGROUND-COLOR:#F7EEDD;' ", UserNt, "agregar", "", "80", "at", "no")
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center ", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", "<B>DOMINIO NT:</B>", "si")
              Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("DOMAINNt readonly style='border-style:outset;border-width:2;BACKGROUND-COLOR:#F7EEDD;' ", DOMAINNT, "agregar", "", "80", "at", "no")
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center ", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
          
            Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);' onclick='abrir();' ", "center", "80%", "20", "", "", "<B><FONT face=wingdings size=2>1</font></B>", "si")
            Fila = Fila & OFV.GenRow("u20", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
            
            Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);' onclick='cerrar();' ", "center", "80%", "20", "", "", "<B><FONT face=wingdings size=2>0</font></B>", "si")
            Fila = Fila & OFV.GenRow("u10", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
          
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", "<B>USUARIO:</B>", "si")
              Fila = Fila & OFV.GenRow("u11", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("usuar", "", "agregar", "", "62", "at", "no")
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u12", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", "<B>PASSWORD:</B>", "si")
              Fila = Fila & OFV.GenRow("u13", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("clave", "", "password", "", "62", "at", "no")
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u14", "aaa style='display:none;cursor:hand;' ", "", "", "", "", "", Celda, "si")
          
          
          Else
          
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", "<B>USUARIO:</B>", "si")
              Fila = Fila & OFV.GenRow("u11", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("usuar", "", "agregar", "", "62", "at", "no")
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u12", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", "<B>PASSWORD:</B>", "si")
              Fila = Fila & OFV.GenRow("u13", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
    
              sInput = OFV.GenerarInput("clave", "", "password", "", "62", "at", "no")
              Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", sInput, "si")
              Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:hand;' ", "", "", "", "", "", Celda, "si")
          
          End If
          




          sInput = OFV.GenerarInput("boton", "INGRESAR", "button", "", xwsize, "it style='border-style:outset;border-width:3;' ", "cript();")
          Celda = OFV.gencelda("", "at style='color:#336699;filter:dropShadow(color=#bbbbbb,offX=3,offY=3);'", "center", "80%", "20", "", "", sInput, "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


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


              Celda = OFV.gencelda("", "ee", "center", "80%", "20", "#ffffff rowspan=4", "", serrordes, "si")
              Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
          End If

          Formulario = OFV.GenForm("Inicio", "", "Inicio.asp?Opcion=V", "", "post", "", Fila, "si")
          
          
  Set cn = OFV.conectar(OFV.strconn2)

  StrSql = "Select * From Clienteplan"
  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")

  Session("Cliente") = rs(1)
  Session("CliLogo") = "../COREss20/Empresa/" & rs(5)

  Call OFV.cerrarconsulta(rs)
  Call OFV.cerrarconn(cn)
  
  sHTML = sHTML & "<table  Valign=middle bgColor=#ffffff width=100% height=100% marginwidth=1 cellspacing=0 cellpadding=0 border=0 style='border-style:outset;border-width:2;' >"
  sHTML = sHTML & "<tr><td  width=100% height=40 align=center valign=top>"
  
  sHTML = sHTML & "<table  bgcolor=" & Session("CMFMENU") & " width=100%  marginwidth=1 cellspacing=0 cellpadding=2 border=0 style='border-style:outset;border-width:2;' >"
  sHTML = sHTML & "<tr><td  width=5% align=center style='filter: dropShadow(color=#999999,offX=2,offY=2);' >"
  sHTML = sHTML & "<img src=" & Session("CliLogo") & " border=0  align=center></td>"
  sHTML = sHTML & "<td align=left  class=at width=60% style='filter: dropShadow(color=#999999,offX=2,offY=2);' ><font color=" & Session("CMFMEFT") & " size=1 face=" & Session("CMFMEFTFC") & " ><b>" & Session("Cliente") & "</b></font></td>"
  sHTML = sHTML & "<td align=right  class=at width=40% ><font color=" & Session("CMFMEFT") & " size=1 face=" & Session("CMFMEFTFC") & " ><b>" & UCase(OFV.fechador()) & "</b></font></td></tr></table>"
          
  sHTML = sHTML & "</td></tr>"
  sHTML = sHTML & "<tr><td  width=100% align=center valign=middle style='filter:shadow(color=#aaaaaa,direction=135);' >"
  
  sHTML = sHTML & OFV.GenTabla("", "", "", "30% border", "", "0", "", "0", "0", "", "", Formulario, "si")
          
    Set cn = OFV.conectar(OFV.strconn1)
    StrSql = "select * from CorAplic where Codapli= '" & Session("TOPAPL") & "'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then
       AplMain = Replace(rs(2), "*", " ") & " " & rs(4)
    Else
       AplMain = "Cor E-Solution Suite 2.8"
    End If

    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

    sInput = "<img border=0 src=imagenes/corlogox1.jpg>"
    Celda = "<font color=" & Session("CMFMEFT") & " size=2 face=" & Session("CMFMEFTFC") & " ><b>&nbsp;" & Replace(AplMain, " ", "&nbsp;") & "</b></font>"

  sHTML = sHTML & "</td></tr>"
  sHTML = sHTML & "<tr><td  width=100% height=40 align=center valign=bottom>"
          
  sHTML = sHTML & "<table  bgcolor=" & Session("CMFMENU") & " width=100%  marginwidth=1 cellspacing=0 cellpadding=2 border=0 style='border-style:outset;border-width:2;' >"
  sHTML = sHTML & "<tr><td width=2%  >" & sInput & "</td>"
  sHTML = sHTML & "<td width=80%  >" & Celda & "</td>"
  sHTML = sHTML & "<td width=20%  valign=middle >"
  sHTML = sHTML & OFV.CorLogo
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr></table>"
          
  sHTML = sHTML & "</td></tr></table>"
  sHTML = sHTML & "</DIV> "
          sHTML = sHTML & "</BODY></html>"
          Inicio = ""
      Else
          sHTML = vu
      End If

      Inicio = sHTML
    
      Exit Function
InicioErr:
      Inicio = Err.Description & " " & Err.Source & " " & CStr(Erl)
End Function


Public Function InicioIz(sopcion)

      On Error GoTo InicioErr
      Dim sInput1$, sInput2$, sInput3$
      Dim xopcion$, serrordes$, i
      Dim UserNt$, DOMAINNT$
      xopcion = sopcion

      vu = OFV.ValidarUsuario()
      If vu = "OK" Then




          sHTML = OFV.MenuHeader("#eeeeee", " background='imagenes/corf5.jpg' ") & Chr(13)

sHTML = sHTML & "<script language=javascript >" & Chr(13)
sHTML = sHTML & "function cript()" & Chr(13)
sHTML = sHTML & "{ " & Chr(13)
sHTML = sHTML & "var sid = " & getsid() & "; " & Chr(13)
sHTML = sHTML & "var i = 0;" & Chr(13)
sHTML = sHTML & "var tab2 = new Array("
genMkt

For i = 0 To 255
    If i > 0 Then sHTML = sHTML & ","
    sHTML = sHTML & xtabla(i)
Next
sHTML = sHTML & ");" & Chr(13)
sHTML = sHTML & "var Y = 0;" & Chr(13)
sHTML = sHTML & "var i = 0;" & Chr(13)
sHTML = sHTML & "var dato = document.all.clave.value;" & Chr(13)
sHTML = sHTML & "var ldato = dato.length;" & Chr(13)
sHTML = sHTML & "var datos = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "var masc = 0;" & Chr(13)
sHTML = sHTML & "if (ldato < 31) " & Chr(13)
sHTML = sHTML & " {for (i=1;i<=ldato;i++)" & Chr(13)
sHTML = sHTML & "   {" & Chr(13)
sHTML = sHTML & "     Y = dato.charCodeAt(i-1);" & Chr(13)
sHTML = sHTML & "     masc = Math.round(i * 7) + Math.round(i * 31 / 17) + sid;" & Chr(13)
sHTML = sHTML & "     datos = datos + String(Math.round(tab2[Y]) - masc);" & Chr(13)
sHTML = sHTML & "   } " & Chr(13)
sHTML = sHTML & "  document.all.clave.value = datos;" & Chr(13)
sHTML = sHTML & "  document.all.Inicio.submit();" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & " else " & Chr(13)
sHTML = sHTML & " {" & Chr(13)
sHTML = sHTML & "  document.all.clave.value = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "  alert('La password no puede exceder de 30 caracteres');" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & "}" & Chr(13)
 sHTML = sHTML & "</script>" & Chr(13)


          xwsize = 45
          If OFV.explorador = "MSIE" Then xwsize = 42

          sInput = "<p align=center><font color=#ffffff face=tahoma size=1>"
          sInput = sInput & "<b>LOG IN</b></font></p>"
          Celda = OFV.gencelda("", "aaa style='border-style:outset;'", "", "100%", "20", "#6699cc", "", sInput, "si")
          Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
          
          
          sInput = "Ingrese su Usuario<br> y Password:<br><br><br>"
          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", sInput, "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

1          UserNt = UCase(Request.ServerVariables("LOGON_USER"))
2          UserNt = Mid(UserNt, InStrRev(UserNt, "\") + 1, Len(UserNt))

           DOMAINNT = UCase(Request.ServerVariables("LOGON_USER"))
           DOMAINNT = Left(DOMAINNT, InStr(DOMAINNT, "\")) & "\"
           DOMAINNT = Replace(DOMAINNT, "\", "")
           
3          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "Usuario NT:", "si")
4         Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
          sInput1 = OFV.GenerarInput("usuar ", "", "hidden", "", xwsize, "at", "no")
          sInput2 = OFV.GenerarInput("clave  ", "", "hidden", "", xwsize, "at", "no")
          sInput = OFV.GenerarInput("usuarNt readonly ", UserNt, "agregar", "", xwsize, "at", "no")
          sInput3 = OFV.GenerarInput("domainNt readonly ", DOMAINNT, "agregar", "", xwsize, "at", "no")
          Celda = OFV.gencelda("", "at style='color:#336699;'", "center ", "80%", "20", "", "", sInput & sInput3 & sInput1 & sInput2, "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
          


          sInput = OFV.GenerarInput("boton", "Ingreso", "button", "", xwsize, "it", "cript();")
          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", sInput, "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


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


              Celda = OFV.gencelda("", "ee", "center", "80%", "20", "#ffffff rowspan=4", "", serrordes, "si")
              Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
          End If

          Formulario = OFV.GenForm("Inicio", "", "Inicio.asp?Opcion=V", "", "post", "", Fila, "si")
          sHTML = sHTML & OFV.GenTabla("", "aaa Style='position:absolute;top:140;left:195;'", "", "55% border", "", "0", "", "0", "0", "#cccccc", "imagenes/headfon3.jpg", Formulario, "si")

          sHTML = sHTML & "</BODY></html>"
          InicioIz = ""
      Else
          sHTML = vu
      End If

      InicioIz = sHTML
    
      Exit Function
InicioErr:
      InicioIz = Err.Description & " " & Err.Source & " " & CStr(Erl)
End Function


Public Function Valida()

Dim mksid, xclavein
Dim UserNt$, DOMAINNT$, CKTUSR$, CKTCOD As Long
Dim UsrAcc As String
Dim UsrRes As String
Dim GrpAcc As String
Dim GrpRes As String

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
'11 - DOMINIO NO HABILITADO



      On Error GoTo ValidaErr
    
      Valida = ""

      Session("Usuario") = ""
      Session("UsrOpcion") = ""
      Session("SkrDefT") = "0"
      Session("SkrDefE") = "0"
      Session("SkrDefI") = "0"
      Session("UsrPmin") = "0"
      Session("UsrPmax") = "0"
      Session("UsrHist") = "0"

      vu = OFV.ValidarUsuario()
      If vu = "OK" Then
          
          Session("Usuario") = UCase(Request.Form("usuar"))
          Session("Clave") = Request.Form("clave")
          If Not IsNull(Session("Usuario")) Then
            Session("Usuario") = UCase(Session("Usuario"))
          End If
'trusted
        If Session("Usuario") = "" And Session("Clave") = "" Then
         sHTML = "0"
         Session("Usuario") = UCase(Request.Form("Usuarnt"))
         Session("UsuarNT") = UCase(Request.Form("Usuarnt"))
         Session("DOMINIO") = UCase(Request.Form("DOMAINnt"))
         Session("TRUSTEDUSR") = "S"

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

            ' CONTROL DEL DOMINIO

              CKTUSR = "N"

              StrSql = "Select * From DOMINIOS Where DESCRIPCION = '" & Session("DOMINIO") & "'"
              Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
              If rs.EOF Then
                 StrSql = "Select * From DOMINIOS Where ID = 1 "
                 Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
                 If Not rs.EOF Then

                    If rs(1) = "DEFAULT" Then
                       CKTUSR = "S"
                       CKTCOD = 1
                    End If
                 End If
              Else
                 CKTUSR = "S"
                 CKTCOD = rs(0)
              End If

              If CKTUSR = "S" Then
                 CKTUSR = "NU"

                 StrSql = "Select * From Usuarios Where netusuario = '" & Session("Usuario") & "' AND  DOMINIO = " & CKTCOD
                 Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
                 If Not rs.EOF Then
                    CKTUSR = "S"
                    Session("Usuario") = rs(0)
                 End If

              End If
              

              sHTML = "0"
             ''' If rs.EOF Then
              If CKTUSR <> "S" Then
                 If CKTUSR = "NU" Then sHTML = "1"
                 If CKTUSR = "N" Then sHTML = "11"
              Else
' verifica rangos horarios

                   UsrAcc = rs("acceso")
                   UsrRes = rs("resol")
                   GrpAcc = Session("CMACC")
                   GrpRes = Session("CMRES")
                   
                   
                   StrSql = "Select * From Grupos Where Id = '" & rs(3) & "' "
                   Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
                   If Not rs2.EOF Then
                      GrpAcc = rs2("acceso")
                      GrpRes = rs2("resol")
                        If GrpAcc = "D" Then
                           GrpAcc = Session("CMACC")
                        End If
                        If GrpRes = "D" Then
                           GrpRes = Session("CMRES")
                        End If

                   End If
                   
                   If UsrAcc = "D" Then
                      UsrAcc = GrpAcc
                   End If
                   If UsrRes = "D" Then
                      UsrRes = GrpRes
                   End If
                   
                   Session("USACC") = UsrAcc
                   Session("USRES") = UsrRes
                   
                   rango = rs("rangos")
                   If Len(rango) = 0 Then rango = 1
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


                    '  StrSql = "Update Usuarios set logfail = 0 "
                    '  StrSql = StrSql & " Where id = '" & Session("Usuario") & "'"
                     ' Call OFV.crearconsultaEx(StrSql, cn, 1, "")
                  'Else
                   '   Call OFV.cerrarconsulta(rs)
                  '    sHTML = "2"
                  '    Session("UsrIntFail") = Session("UsrIntFail") + 1

                   '   StrSql = "Update Usuarios set logfail = " & Session("UsrIntFail")
                    '  StrSql = StrSql & " Where id = '" & Session("Usuario") & "'"
                     ' Call OFV.crearconsultaEx(StrSql, cn, 1, "")

                      ' If (Session("UsrIntFail") > Session("SkrDefI")) And Session("SkrDefT") = 1 Then
                       '   sHTML = "3"
                        '  StrSql = "Update Usuarios set OPCION = 2 "
                         'StrSql = StrSql & " Where id = '" & Session("Usuario") & "'"
                          'Call OFV.crearconsultaEx(StrSql, cn, 1, "")
                     ' End If
                      

                  ' End If
                End If
                
                End If
                
              End If
              Call OFV.cerrarconsulta(rs)
              Call OFV.cerrarconn(cn)
        
         Else 'fin trusted
          Session("TRUSTEDUSR") = "N"
          If Session("Usuario") = "" Or Session("Clave") = "" Then
              sHTML = "1"
          Else
              mksid = getsid()
              genMkt
              genMkCh
              xclavein = Getcrpmk("D", Session("Clave"), mksid)

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
              Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
              

              sHTML = "0"
              If rs.EOF Then
                  sHTML = "1"
              Else

                   Session("UsuarNT") = rs("netusuario")
' verifica rangos horarios
                   UsrAcc = rs("acceso")
                   UsrRes = rs("resol")
                   GrpAcc = Session("CMACC")
                   GrpRes = Session("CMRES")
                   
                   
                   StrSql = "Select * From Grupos Where Id = '" & rs(3) & "' "
                   Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
                   If Not rs2.EOF Then
                      GrpAcc = rs2("acceso")
                      GrpRes = rs2("resol")
                        If GrpAcc = "D" Then
                           GrpAcc = Session("CMACC")
                        End If
                        If GrpRes = "D" Then
                           GrpRes = Session("CMRES")
                        End If

                   End If
                   
                   If UsrAcc = "D" Then
                      UsrAcc = GrpAcc
                   End If
                   If UsrRes = "D" Then
                      UsrRes = GrpRes
                   End If
                   
                   Session("USACC") = UsrAcc
                   Session("USRES") = UsrRes
                   
                   rango = rs("rangos")
                   If Len(rango) = 0 Then rango = 1
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
        
                  Set Encripter = Server.CreateObject("CorEncriptadorEx.SHA1")

                  sEncript = Encripter.Encrypt(xclavein)
                  If rs(2) = sEncript Then
                      Session("SegUsuario") = rs(1)
                      Session("SegNivel") = rs(4)
                      Session("SegGrp") = rs(3)

                      Session("SegPerfil") = Session("Usuario")

                      If Session("UsrOpcion") = 1 Then
                         Call OFV.cerrarconsulta(rs)
                         sHTML = "5"
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


                      StrSql = "Update Usuarios set logfail = 0 "
                      StrSql = StrSql & " Where id = '" & Session("Usuario") & "'"
                      Call OFV.crearconsultaEx(StrSql, cn, 1, "")
                  Else
                      Call OFV.cerrarconsulta(rs)
                      sHTML = "2"
                      Session("UsrIntFail") = Session("UsrIntFail") + 1

                      StrSql = "Update Usuarios set logfail = " & Session("UsrIntFail")
                      StrSql = StrSql & " Where id = '" & Session("Usuario") & "'"
                      Call OFV.crearconsultaEx(StrSql, cn, 1, "")

                       If (Session("UsrIntFail") > Session("SkrDefI")) And Session("SkrDefT") = 1 Then
                          sHTML = "3"
                          StrSql = "Update Usuarios set OPCION = 2 "
                         StrSql = StrSql & " Where id = '" & Session("Usuario") & "'"
                          Call OFV.crearconsultaEx(StrSql, cn, 1, "")
                      End If
                      

                   End If
                End If
                
                End If
                
              End If
              Call OFV.cerrarconsulta(rs)
              Call OFV.cerrarconn(cn)
          End If
          ' Session("Usuario") = "" and Session("Clave") = ""
       End If
      Else
          sHTML = vu
      End If

      Session("SkrOpc") = sHTML
      Valida = sHTML
    
      Exit Function
ValidaErr:
      Valida = Err.Description & " " & Err.Source & " " & CStr(Erl)


End Function

Public Function CambiaPass(sError)

      On Error GoTo CambiaPassErr
      Dim sopcion$, xerror$, xscr$, serrordes$, i

      sopcion = Session("SkrOpc")
      xerror = sError

      CambiaPass = ""

      ok = OFV.CheckUsuario()
      If ok <> "Y" Then
          uv = OFV.uv
          If uv = "NO" Then ok = OFV.ValidarUsuario
          sHTML = ok
      Else
          Response.expires = 0
          
          xscr = "    "
    
          sHTML = OFV.MenuHeader("", xscr)

sHTML = sHTML & "<script language=javascript >" & Chr(13)
sHTML = sHTML & "function cript()" & Chr(13)
sHTML = sHTML & "{ " & Chr(13)
sHTML = sHTML & "var sid = " & getsid() & "; " & Chr(13)
sHTML = sHTML & "var i = 0;" & Chr(13)
sHTML = sHTML & "var tab2 = new Array("
genMkt

For i = 0 To 255
    If i > 0 Then sHTML = sHTML & ","
    sHTML = sHTML & xtabla(i)
Next
sHTML = sHTML & ");" & Chr(13)
sHTML = sHTML & "var Y = 0;" & Chr(13)
sHTML = sHTML & "var i = 0;" & Chr(13)
sHTML = sHTML & "var dato1 = document.all.Npass.value;" & Chr(13)
sHTML = sHTML & "var dato2 = document.all.NpassC.value;" & Chr(13)
sHTML = sHTML & "var ldato1 = dato1.length;" & Chr(13)
sHTML = sHTML & "var ldato2 = dato2.length;" & Chr(13)
sHTML = sHTML & "var datos1 = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "var datos2 = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "var masc = 0;" & Chr(13)
sHTML = sHTML & "if (ldato1 < 31 && ldato2 < 31) " & Chr(13)
sHTML = sHTML & " {for (i=1;i<=ldato1;i++)" & Chr(13)
sHTML = sHTML & "   {" & Chr(13)
sHTML = sHTML & "     Y = dato1.charCodeAt(i-1);" & Chr(13)
sHTML = sHTML & "     masc = Math.round(i * 7) + Math.round(i * 31 / 17) + sid;" & Chr(13)
sHTML = sHTML & "     datos1 = datos1 + String(Math.round(tab2[Y]) - masc);" & Chr(13)
sHTML = sHTML & "   } " & Chr(13)
sHTML = sHTML & "  for (i=1;i<=ldato2;i++)" & Chr(13)
sHTML = sHTML & "   {" & Chr(13)
sHTML = sHTML & "     Y = dato2.charCodeAt(i-1);" & Chr(13)
sHTML = sHTML & "     masc = Math.round(i * 7) + Math.round(i * 31 / 17) + sid;" & Chr(13)
sHTML = sHTML & "     datos2 = datos2 + String(Math.round(tab2[Y]) - masc);" & Chr(13)
sHTML = sHTML & "   } " & Chr(13)
sHTML = sHTML & "  document.all.Npass.value = datos1;" & Chr(13)
sHTML = sHTML & "  document.all.NpassC.value = datos2;" & Chr(13)
sHTML = sHTML & "  document.all.Cambia.submit();" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & " else " & Chr(13)
sHTML = sHTML & " {" & Chr(13)
sHTML = sHTML & "  document.all.Npass.value = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "  document.all.NpassC.value = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "  alert('La password y la confirmacion no pueden exceder de 30 caracteres');" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & "}" & Chr(13)
 sHTML = sHTML & "</script>" & Chr(13)



             Select Case sopcion
                    Case "4"
                         serrordes = "Cambio de Password por expiracion"

                    Case "5"
                         serrordes = "Cambio de Password forzada por Administrador"

                    Case Else
                         serrordes = "Desea cambiar<br> su Password?<br><br><br>"
             End Select

        
          xwsize = 45
          If OFV.explorador = "MSIE" Then xwsize = 42

          Fila = OFV.GenForm("Cambia", "", "Inicio.asp?Opcion=CPOK", "", "post", "", "", "no")

          Celda = OFV.gencelda("", "tt style='border-style:outset;'", "", "100%", "20", "", "", "<b>PASSWORD</b>", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", serrordes, "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "Nueva Password:", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "<input class=at type=password size=" & xwsize & " name=Npass>", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "Confirme:", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "<input class=at type=password size=" & xwsize & " name=NpassC>", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "<input class=it type=button onclick='cript();' size=" & OFV.wsize & " name=boton value=OK>", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Fila = Fila & OFV.GenFForm()

          If sopcion = "0" Then
             Fila = Fila & OFV.GenForm("Inises", "", "nulo.asp", "", "post", "", "", "no")

             Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "<input class=it type=submit size=" & OFV.wsize & " name=botonC value=Cancelar>", "si")
             Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

             Fila = Fila & OFV.GenFForm()
          End If

          serrordes = ""
          If xerror <> "" Then
             Select Case xerror
                    Case "1"
                         serrordes = "SKR011 - Password invalida: Nueva password o password confirmacion "
                         serrordes = serrordes & "son nulas o no coinciden "
                    Case "2"
                         serrordes = "SKR012 - Password invalida: Nueva password no tiene un minimo de "
                         serrordes = serrordes & Session("UsrPmin")
                         serrordes = serrordes & " caracteres o excede de " & Session("UsrPmax")
                         serrordes = serrordes & " caracteres "
                    Case "3"
                         serrordes = "SKR013 - Password invalida: Nueva password tiene repeticion de "
                         serrordes = serrordes & " caracteres o numeros en forma consecutiva "
                    Case "4"
                         serrordes = "SKR014 - Password invalida: Nueva password tiene repeticion de "
                         serrordes = serrordes & " caracteres o numeros en la misma posicion "
                         serrordes = serrordes & " respecto de la password anterior "
                    Case "5"
                         serrordes = "SKR015 - Password invalida: Nueva password se compone de "
                         serrordes = serrordes & " una secuencia de numeros correlativos "
                    Case "6"
                         serrordes = "SKR016 - Password incorrecta: Nueva password se repite "
                         serrordes = serrordes & " dentro del historial de passwords para el "
                         serrordes = serrordes & " usuario "
                    Case Else
                         serrordes = "SKR900 - Password incorrecta: Error imprevisto en "
                         serrordes = serrordes & " cambio de passwords "

             End Select

              Celda = OFV.gencelda("", "ee", "center", "80%", "20", "#ffffff rowspan=4", "", serrordes, "si")
              Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
          End If
          If Session("USRES") = "B" Then
          sHTML = sHTML & OFV.GenTabla("panel", "aaa Style='position:absolute;top:140;left:195;'", "", "55% border", "", 0, 0, 0, 0, "#dddddd", "", Fila, "si") & "</BODY></html>"
          Else
          sHTML = sHTML & OFV.GenTabla("panel", "aaa Style='position:absolute;top:140;left:195;'", "", "55% border", "", 0, 0, 0, 0, "#dddddd", "", Fila, "si") & "</BODY></html>"
          End If
      End If
    
      CambiaPass = sHTML
    
      Exit Function
CambiaPassErr:
      CambiaPass = Err.Description & " " & Err.Source & " " & CStr(Erl)


End Function

Public Function CambiaPassOK()
      On Error GoTo CambiaPassOKErr

      Dim sopcion$

      sopcion = Session("SkrOpc")
      Dim Npass, NpassC, Plen, Alen, sEncript, xpass, xlcha, xHcha, xHcha1, xopcn, xind, xult
      Dim mksid, xclavein, xNpassin
      CambiaPassOK = ""

      ok = OFV.CheckUsuario()
      If ok <> "Y" Then
          uv = OFV.uv
          If uv = "NO" Then ok = OFV.ValidarUsuario
          sHTML = ok
      Else
          Response.expires = 0
          sHTML = "0"
          Npass = Request.Form(1)
          NpassC = Request.Form(2)

          sHTML = VerCambio("US", Session("Usuario"), Npass, NpassC)
      End If
    
      CambiaPassOK = sHTML
    
      Exit Function
CambiaPassOKErr:
     CambiaPassOK = Err.Description & " " & Err.Source & " " & CStr(Erl)



End Function

Public Function InicioSesion()

      On Error GoTo InicioSesionErr

     Dim xscr$

     InicioSesion = ""

      ok = OFV.CheckUsuario()
      If ok <> "Y" Then
          uv = OFV.uv
          If uv = "NO" Then ok = OFV.ValidarUsuario
          sHTML = ok
      Else
          Response.expires = 0
          xscr = "    onload=" & Chr(34)
          xscr = xscr & "document.Menuf.submit();" & Chr(34)
    
          sHTML = OFV.MenuHeader("", xscr)


        
          xwsize = 45
          If OFV.explorador = "MSIE" Then xwsize = 42

          Fila = ""

          Celda = OFV.gencelda("", "tt style='border-style:outset;'", "center", "100%", "20", "", "", "<b>INICIO DE SESION: " & Session.SESSIONID & "</b>", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "Usuario:", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", Session("usuario"), "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "Descripcion", "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

          Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", Session("segusuario"), "si")
          Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

'''             Fila = Fila & OFV.GenForm("cambia", "", "Inicio.asp?Opcion=CP", "", "post", "", "", "no")
'''             If Session("TINI") <> "Z" Then
'''              Celda = OFV.gencelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "<input class=it type=submit size=" & OFV.wsize & " name=botonC value='Cambiar Password?' >", "si")
'''              Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
'''             End If
'''             Fila = Fila & OFV.GenFForm()

           Fila = Fila & OFV.GenForm("Menuf", "", "Footer.asp?Opcion=A", "Footer", "post", "", "<input type=hidden size=" & OFV.wsize & " name=boton value=menu></input>", "si")


  Set cn = OFV.conectar(OFV.strconn2)

  StrSql = "Select * From Clienteplan"
  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")

  Session("Cliente") = rs(1)
  Session("CliLogo") = "../COREss20/Empresa/" & rs(5)

  Call OFV.cerrarconsulta(rs)
  Call OFV.cerrarconn(cn)

          If Session("USRES") = "B" Then
          sHTML = sHTML & OFV.GenTabla("panel", "aaa Style='position:absolute;top:140;left:195;'", "", "55% border", "", 0, 0, 0, 0, "#dddddd", "", Fila, "si") & "</BODY></html>"
          Else
          sHTML = sHTML & OFV.GenTabla("panel", "aaa Style='position:absolute;top:140;left:195;'", "", "55% border", "", 0, 0, 0, 0, "#dddddd", "", Fila, "si") & "</BODY></html>"
          End If
      End If
    
      InicioSesion = sHTML
    
      Exit Function
InicioSesionErr:
     InicioSesion = Err.Description & " " & Err.Source & " " & CStr(Erl)


End Function

Public Function VerCambio(stipo, susuario, Npass, NpassC) As String

      Dim sopcion$

      sopcion = Session("SkrOpc")
      Dim Plen, Alen, sEncript, xpass, xlcha, xHcha, xHcha1, xopcn, xind, xult
      Dim mksid, xclavein, xNpassin, xerr
      Dim SUsrPmin, SUsrPmax, SUsrHist, SSkrDefT

      VerCambio = ""
      sHTML = "0"

      xerr = "0"
' 1
          If Npass = "" Or NpassC = "" Then
              sHTML = "1"

          Else
' 2
              If Npass <> NpassC Then
                  sHTML = "1"
              Else
                  mksid = getsid()
                  genMkt
                  genMkCh
                   
                  If stipo = "US" Then
                     xclavein = Getcrpmk("D", Session("Clave"), mksid)
                     SUsrPmin = Session("UsrPmin")
                     SUsrPmax = Session("UsrPmax")
                     SUsrHist = Session("UsrHist")
                     SSkrDefT = Session("SkrDefT")
                  Else
                     Set cn = OFV.conectar(OFV.strconn0)
                     StrSql = "Select * From skrdef"
                     Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
                     If Not rs.EOF Then
                        SUsrPmin = CInt(rs("PassMin"))
                        SUsrPmax = CInt(rs("PassMax"))
                        SUsrHist = CInt(rs("PassHist"))
                        SSkrDefT = CInt(rs("tipoesq"))
                     End If
                     Call OFV.cerrarconsulta(rs)

                     StrSql = "Select * From Usuarios Where id = '" & susuario & "'"
                     Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
                     If rs.EOF Then
                        xerr = "1"
                     Else
                        If CInt(rs("defskr")) = 1 Then
                          SUsrPmin = CInt(rs("PassMin"))
                          SUsrPmax = CInt(rs("PassMax"))
                          SUsrHist = CInt(rs("PassHist"))
                          SSkrDefT = CInt(rs("tipoesq"))
                        End If
                     End If
                     Call OFV.cerrarconsulta(rs)
                     Call OFV.cerrarconn(cn)
                 End If
' 3
                 If xerr = "1" Then
                    sHTML = "1"
                 Else
                  xNpassin = Getcrpmk("D", Npass, mksid)
                  Plen = Len(xNpassin)
' 4
                  If (SUsrPmin > Plen) Or (SUsrPmax < Plen) Then
                      sHTML = "2"
                  Else
                       For xind = 1 To (Plen - 1)
                           If Mid(xNpassin, xind, 1) = Mid(xNpassin, xind + 1, 1) Then
                              sHTML = "3"
                              Exit For
                           End If
                       Next
' 5
                       If sHTML = "0" Then
                          If stipo = "US" Then
                             Alen = Len(xclavein)
                             If Alen > Plen Then Alen = Plen
                             For xind = 1 To Alen
                                 If Mid(xNpassin, xind, 1) = Mid(xclavein, xind, 1) Then
                                    sHTML = "4"
                                    Exit For
                                 End If
                             Next
                           End If
' 6
                          If sHTML = "0" Then
                             For xind = 1 To (Plen - 1)
                                 If IsNumeric(Mid(xNpassin, xind, 1)) Then
                                    If IsNumeric(Mid(xNpassin, xind + 1, 1)) Then
                                       If CInt(Mid(xNpassin, xind, 1)) = (CInt(Mid(xNpassin, xind + 1, 1)) + 1) Then
                                          sHTML = "5"
                                          Exit For
                                       ElseIf CInt(Mid(xNpassin, xind, 1)) = (CInt(Mid(xNpassin, xind + 1, 1)) - 1) Then
                                          sHTML = "5"
                                          Exit For
                                       End If
                                    End If
                                 End If
                             Next
' 7
                             If sHTML = "0" Then
                                Set Encripter = Server.CreateObject("CorEncriptadorEx.SHA1")
                                sEncript = Encripter.Encrypt(xNpassin)

                                Set cn = OFV.conectar(OFV.strconn0)
                                StrSql = "Select * From Usuarios Where id = '"
                                StrSql = StrSql & susuario & "'"
                                Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
                                If Not rs.EOF Then
                                   xpass = rs("Password")
                                   xlcha = rs("Lchange")
                                   xHcha = rs("Hchange")
                                   xopcn = rs("Opcion")
                                End If
                                Call OFV.cerrarconsulta(rs)
                                Call OFV.cerrarconn(cn)
                                If sEncript = xpass Then
                                   sHTML = "6"
                                Else
                                   xind = 0
                                   xult = 0
                                   Set cn = OFV.conectar(OFV.strconn0)
                                   StrSql = "select objid, password from histpwd "
                                   StrSql = StrSql & " Where usuario ='" & susuario & "' "
                                   StrSql = StrSql & " order by lchange desc, hchange desc "
                                   Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
                                   Do Until rs.EOF
                                      xind = xind + 1
                                      xult = rs(0)
                                      If rs(1) = sEncript Then
                                         sHTML = "6"
                                      End If
                                      rs.movenext
                                   Loop
                                   Call OFV.cerrarconsulta(rs)
                                   Call OFV.cerrarconn(cn)
                                End If
' 8
                                If sHTML = "0" Then
                                   Set cn = OFV.conectar(OFV.strconn0)
                                   If SSkrDefT = 1 Then
                                       StrSql = "insert into HistPwd "
                                       StrSql = StrSql & "(usuario,password,lchange,hchange,opcion) "
                                       StrSql = StrSql & "values ("
                                       StrSql = StrSql & "'" & susuario & "',"
                                       StrSql = StrSql & "'" & xpass & "',"
                                       StrSql = StrSql & "'" & xlcha & "',"
                                       StrSql = StrSql & "'" & xHcha & "',"
                                       StrSql = StrSql & "'" & xopcn & "')"
                                       Call OFV.crearconsultaEx(StrSql, cn, 1, "")
                                       xind = xind + 1
                                       If SUsrHist < xind Then
                                           StrSql = "select objid, password from histpwd "
                                           StrSql = StrSql & " Where usuario ='" & susuario & "' "
                                           StrSql = StrSql & " order by lchange , hchange  "
                                           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
                                           Do Until rs.EOF
    
                                              xult = rs(0)
                                              StrSql = "delete from HistPwd "
                                              StrSql = StrSql & " where objid = "
                                              StrSql = StrSql & xult
                                              Call OFV.crearconsultaEx(StrSql, cn, 1, "")
                                              xind = xind - 1
                                              If Not SUsrHist < xind Then Exit Do
                                           rs.movenext
                                           Loop
                                       End If


                                   End If
                                       xlcha = CLng(Date) - CLng(CDate("01-01-1900"))
                                       xHcha1 = Time()
                                       xHcha = ""
                                       For xind = 1 To Len(xHcha1)
                                           If IsNumeric(Mid(xHcha1, xind, 1)) Then
                                              xHcha = xHcha & Mid(xHcha1, xind, 1)
                                           End If
                                       Next
                                       xopcn = 1
                                       If stipo = "US" Then xopcn = 0
                                       
                                       StrSql = "Update Usuarios set password ='" & sEncript
                                       StrSql = StrSql & "', opcion = " & xopcn & ", lchange = '" & xlcha
                                       StrSql = StrSql & "', Hchange = '" & xHcha
                                       StrSql = StrSql & "' Where id ='" & susuario & "'"
                                       Call OFV.crearconsultaEx(StrSql, cn, 1, "")
                                       Call OFV.cerrarconn(cn)
                                       If stipo = "US" Then Session("Clave") = Npass
' 8
                                End If
' 7
                             End If
' 6
                         End If
' 5
                      End If
' 4
                  End If
' 3
              End If
' 2
          End If
' 1
        End If


    
      VerCambio = sHTML


End Function

Public Function Habilit() As String

  Dim sxopc1, h1, f1a, i

  If OFV.explorador = "MSIE" Then
    sHTML = OFV.MenuHeader("#cccccc", "")
  Else
    sHTML = OFV.MenuHeader("#ccbbaa", "")
  End If

sHTML = sHTML & "<script language=javascript >" & Chr(13)
sHTML = sHTML & "function cript()" & Chr(13)
sHTML = sHTML & "{ " & Chr(13)
sHTML = sHTML & "var sid = " & getsid() & "; " & Chr(13)
sHTML = sHTML & "var i = 0;" & Chr(13)
sHTML = sHTML & "var tab2 = new Array("
genMkt

For i = 0 To 255
    If i > 0 Then sHTML = sHTML & ","
    sHTML = sHTML & xtabla(i)
Next
sHTML = sHTML & ");" & Chr(13)
sHTML = sHTML & "var Y = 0;" & Chr(13)
sHTML = sHTML & "var i = 0;" & Chr(13)
sHTML = sHTML & "var dato1 = document.all.Npass.value;" & Chr(13)
sHTML = sHTML & "var dato2 = document.all.NpassC.value;" & Chr(13)
sHTML = sHTML & "var ldato1 = dato1.length;" & Chr(13)
sHTML = sHTML & "var ldato2 = dato2.length;" & Chr(13)
sHTML = sHTML & "var datos1 = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "var datos2 = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "var masc = 0;" & Chr(13)
sHTML = sHTML & "if (ldato1 < 31 && ldato2 < 31) " & Chr(13)
sHTML = sHTML & " {for (i=1;i<=ldato1;i++)" & Chr(13)
sHTML = sHTML & "   {" & Chr(13)
sHTML = sHTML & "     Y = dato1.charCodeAt(i-1);" & Chr(13)
sHTML = sHTML & "     masc = Math.round(i * 7) + Math.round(i * 31 / 17) + sid;" & Chr(13)
sHTML = sHTML & "     datos1 = datos1 + String(Math.round(tab2[Y]) - masc);" & Chr(13)
sHTML = sHTML & "   } " & Chr(13)
sHTML = sHTML & "  for (i=1;i<=ldato2;i++)" & Chr(13)
sHTML = sHTML & "   {" & Chr(13)
sHTML = sHTML & "     Y = dato2.charCodeAt(i-1);" & Chr(13)
sHTML = sHTML & "     masc = Math.round(i * 7) + Math.round(i * 31 / 17) + sid;" & Chr(13)
sHTML = sHTML & "     datos2 = datos2 + String(Math.round(tab2[Y]) - masc);" & Chr(13)
sHTML = sHTML & "   } " & Chr(13)
sHTML = sHTML & "  document.all.Npass.value = datos1;" & Chr(13)
sHTML = sHTML & "  document.all.NpassC.value = datos2;" & Chr(13)
sHTML = sHTML & "  document.all.Cambia.submit();" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & " else " & Chr(13)
sHTML = sHTML & " {" & Chr(13)
sHTML = sHTML & "  document.all.Npass.value = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "  document.all.NpassC.value = " & Chr(34) & Chr(34) & ";" & Chr(13)
sHTML = sHTML & "  alert('La password y la confirmacion no pueden exceder de 30 caracteres');" & Chr(13)
sHTML = sHTML & " }" & Chr(13)
sHTML = sHTML & "}" & Chr(13)
 sHTML = sHTML & "</script>" & Chr(13)



  If OFV.explorador = "MSIE" Then
    xwsize = 30
  Else
    xwsize = 15
  End If

  sxopc1 = "<option value=NO selected>[SELECCIONE USUARIO]</option>"

  Set cn = OFV.conectar(OFV.strconn0)

  StrSql = "Select * From Usuarios  order by descripcion"
  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")

  Do Until rs.EOF

     sxopc1 = sxopc1 & "<option value=" & rs(0) & ">" & rs(0) & " - " & rs(1) & "</option> "

  rs.movenext
  Loop

  Call OFV.cerrarconsulta(rs)
  Call OFV.cerrarconn(cn)

  sHTML = sHTML & "<BR><BR><CENTER>"
  sHTML = sHTML & OFV.GenTabla("panel", "", "", "60% border", "", 0, 0, 0, 0, "#cccccc", "", "", "no")

  h1 = OFV.gencelda("", "tt", "center", "100%", "20", "#336699", "", "Habilitacion de Usuario", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  h1 = OFV.gencelda("", "tt", "center", "80%", "20", "#336699", "", "Usuario", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  f1a = "HabilitPass.asp"
  sHTML = sHTML & OFV.GenForm("Cambia", "", f1a, "", "post", "", "", "no")

  h1 = OFV.gencelda("", "cc", "center", "", "", "", "", "<select class=at name=Nusr >" & sxopc1 & "</select>", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  h1 = OFV.gencelda("", "tt", "center", "80%", "20", "#336699", "", "Password:", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  h1 = OFV.gencelda("", "cc", "center", "80%", "20", "", "", "<input class=at type=password size=" & xwsize & " name=Npass>", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  h1 = OFV.gencelda("", "tt", "center", "80%", "20", "#336699", "", "Confirmacion:", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  h1 = OFV.gencelda("", "cc", "center", "80%", "20", "", "", "<input class=at type=password size=" & xwsize & " name=NpassC>", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  h1 = OFV.gencelda("", "tt", "center", "80%", "20", "#336699", "", "<input class=it type=submit onclick='cript();' size=" & OFV.wsize & " name=boton value=OK>", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  sHTML = sHTML & OFV.GenFForm()


  sHTML = sHTML & OFV.GenFTabla()

  sHTML = sHTML & "</CENTER>"
  sHTML = sHTML & "</BODY></html>"

  Habilit = sHTML

End Function

Public Function HabilitPass() As String

  Dim sxopc1, h1, f1a, i, xerr, Npass, NpassC, Nusr, Uerror, serrordes


  Nusr = Request.Form(1)
  Npass = Request.Form(2)
  NpassC = Request.Form(3)
  xerr = "0"

  xerr = VerCambio("UH", Nusr, Npass, NpassC)

  If xerr <> "0" Then
        Uerror = 1
  Else
        Uerror = 0
  End If

  If OFV.explorador = "MSIE" Then
    sHTML = OFV.MenuHeader("#cccccc", "")
  Else
    sHTML = OFV.MenuHeader("#ccbbaa", "")
  End If

  If OFV.explorador = "MSIE" Then
    xwsize = 10
  Else
    xwsize = 5
  End If

  Set cn = OFV.conectar(OFV.strconn0)
  StrSql = "Select * From Usuarios Where id = '" & Nusr & "'"
  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
  If Not rs.EOF Then sxopc1 = rs(0) & " - " & rs(1)
  Call OFV.cerrarconn(cn)

  sHTML = sHTML & "<BR><BR><CENTER>"
  sHTML = sHTML & OFV.GenTabla("panel", "", "", "40% border", "", 0, 0, 0, 0, "#cccccc", "", "", "no")

  h1 = OFV.gencelda("", "tt", "center", "100%", "20", "#336699", "", "Habilitacion de Usuario", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  h1 = OFV.gencelda("", "tt", "center", "80%", "20", "#336699", "", "Usuario", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")


  h1 = OFV.gencelda("", "cc", "center", "", "", "", "", sxopc1, "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  If Uerror = 1 Then

          serrordes = ""
          If xerr <> "" Then
             Select Case xerr
                    Case "1"
                         serrordes = "SKR021 - Usuario o Password invalidos: Nueva password o password confirmacion "
                         serrordes = serrordes & "son nulas o no coinciden, o el usuario no existe "
                    Case "2"
                         serrordes = "SKR012 - Password invalida: Nueva password no tiene un minimo de "
                         serrordes = serrordes & Session("UsrPmin")
                         serrordes = serrordes & " caracteres o excede de " & Session("UsrPmax")
                         serrordes = serrordes & " caracteres "
                    Case "3"
                         serrordes = "SKR013 - Password invalida: Nueva password tiene repeticion de "
                         serrordes = serrordes & " caracteres o numeros en forma consecutiva "
                    Case "4"
                         serrordes = "SKR014 - Password invalida: Nueva password tiene repeticion de "
                         serrordes = serrordes & " caracteres o numeros en la misma posicion "
                         serrordes = serrordes & " respecto de la password anterior "
                    Case "5"
                         serrordes = "SKR015 - Password invalida: Nueva password se compone de "
                         serrordes = serrordes & " una secuencia de numeros correlativos "
                    Case "6"
                         serrordes = "SKR016 - Password incorrecta: Nueva password se repite "
                         serrordes = serrordes & " dentro del historial de passwords para el "
                         serrordes = serrordes & " usuario "
                    Case Else
                         serrordes = "SKR900 - Password incorrecta: Error imprevisto en "
                         serrordes = serrordes & " cambio de passwords "

             End Select
          End If


  f1a = "Habilit.asp"
  sHTML = sHTML & OFV.GenForm("Cambia", "", f1a, "", "post", "", "", "no")

  h1 = OFV.gencelda("", "", "center", "80%", "20", "#993333", "", "<font face=tahoma size=2 color=#ffffff>" & serrordes & "</font>", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  h1 = OFV.gencelda("", "tt", "center", "80%", "20", "", "", "<input class=it type=submit size=" & OFV.wsize & " name=boton value=Volver>", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  sHTML = sHTML & OFV.GenFForm()

  Else

  h1 = OFV.gencelda("", "", "center", "100%", "20", "#800000", "", "<font face=tahoma size=2 color=#ffffff>Habilitacion OK !!</font>", "si")
  sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", h1, "si")

  End If

  sHTML = sHTML & OFV.GenFTabla()

  sHTML = sHTML & "</CENTER>"
  sHTML = sHTML & "</BODY></html>"

  HabilitPass = sHTML

End Function

