<!-- #INCLUDE FILE="Corstyles.inc" -->


<%

' ****************************************************************************************
' ** Cor E-Solution V34 - Esquema de Configuracion -----------------       SEPT, 2008   **
' ****************************************************************************************



set ofv = server.createobject("CorWebServerV34.FuncVarias")

Set ofv.Request = Request
Set ofv.Server = Server
Set ofv.Response = Response
Set ofv.Session = Session

session("Mcorrecion")="6a"
session("MailComponent")="CDOSYS"

Session("SMTPServer")="smtp.iplannetworks.net"
Session("SMTPPort")=25
SESSION("MAILOPC") = "TEXTO"


SESSION("URLTYP")  = "V"
SESSION("URLPORT")  = "N"

SESSION("GLOSAUTOR")  = "S"

SESSION("TIPOFECHA")  =0  ' fechadorCor

SESSION("FULLDOCN") = "1"

Session("CORUNQ") = "S"

SESSION("WKFETPU") = "U"
SESSION("WKFEUSR") = "COREMANAGER"
SESSION("WKFEMAI") = "@"

Session("Sound") = "Imagenes/xLogon.wav"


 Session("Modal")   = "MSNTT"
 Session("TOPAPL")  = "COREM20"
 Session("CMACC")   = "L"
 Session("CMRES")   = "A"
 Session("CONSET")  = "0"
 Session("EMPMUN") = "N"

 Session("EXPLOG")  = "N"
 Session("DELLOG")  = "N"
 session("USRPRV")  = "I"
 Session("FULLGUI") = "S"
 SESSION("RDCGUID") = "S"
 SESSION("FILETIPO")  = "L"

 SESSION("SHWMOD")  = "S"

 if Session("SETLOG") = "S" then Session("EXPLOG") = "N"

%>
<!-- #INCLUDE FILE="CorDbSources.inc" -->
<!-- #INCLUDE FILE="CorLOGF.inc" -->
<!-- #INCLUDE FILE="GetInfo.inc" -->
<%

 CALL GETFECH()

 session("XEXPIR") = XLHORDAT



' ****************************************************************************************
' ** Cor E-Solution V34 - Esquema de Configuracion -----------------        Sept, 2008  **
' ****************************************************************************************

%>

