<!-- #INCLUDE FILE="Corstyles.inc" -->
<%

' ****************************************************************************************
' ** Cor E-Solution V3 - Esquema de Configuracion -----------------        Sept, 2006   **
' produccion
'  ****************************************************************************************



 set ofv = server.createobject("CorWebServerV34.FuncVarias")

Set ofv.Request = Request
Set ofv.Server = Server
Set ofv.Response = Response
Set ofv.Session = Session

Session("SMTPServer")="smtp.mail.blp.com.ar"
Session("SMTPPort")=25

session("MailComponent")="CDOSYS"

SESSION("TIPOFECHA")  =0
'Session("Sound") = "Imagenes/Logon.wav"


 Session("Modal") = "MSNTT"
 Session("TOPAPL") = "COREM20"
 Session("CMACC") = "L"
 Session("CMRES") = "A"
 Session("CONSET") = "0"
 Session("EMPMUN") = "N"
 Session("EXPLOG") = "S"


 Session("CORUNQ") = "S"
 Session("DELLOG")  = "N"
 session("USRPRV")  = "I"
 Session("FULLGUI") = "S"
 SESSION("RDCGUID") = "S"
 SESSION("FULLDOCN") = "1"
 SESSION("GLOSAUTOR")  = "N"

'------------------------------------------
SESSION("MAILOPC") = "TEXTO"
SESSION("URLTYP")  = "V"
SESSION("URLPORT")  = "Y"

SESSION("FILETIPO")  = "L"

SESSION("WKFETPU") = "D"
SESSION("WKFEUSR") = "COREMANAGER"
SESSION("WKFEMAI") = "@"
SESSION("TTCC")="BASIC"
'------------------------------------------

 if Session("SETLOG") = "S" then Session("EXPLOG") = "N"


if  Session("ESPLOG") = "S" then


     Session("EXPLOG") = "N"
     Session("DELLOG") = "N"
     Session("ESPLOG") = "S"

 end if


%>


<!-- #INCLUDE FILE="CorDbSources.inc" -->
<!-- #INCLUDE FILE="CorLOGF.inc" -->
<!-- #INCLUDE FILE="GetInfo.inc" -->

<%

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ** Cor E-Solution V3 - Esquema de Configuracion -----------------        Sept, 2006   **
'''  COR CONECT - Include
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ****************************************************************************************
' ** Cor E-Solution V3 - Esquema de Configuracion -----------------        Sept, 2006   **
' ****************************************************************************************

CALL GETFECH()

 session("XEXPIR") = XLHORDAT


%>

