<%

' ****************************************************************************************
' ** Cor E-Solution v3 -    Esquema de Configuracion ------------------ marzo, 2007     **
' ****************************************************************************************


''XLCHARSET = "iso-8859-1"

XLCHARSET = "1252"


XLFECDAT   = 0
XLHORDAT   = 10
XLLOGID    = 0
XLLOGIDDET = 0
XLLOGIDEXT = 0

GLBSH      = 0
GLBSM      = 0
GLBSS      = 0

TAEMP      = 0 
IAEMP      = 0 
DIM ATREMP(50,5)

USREMP = 1

DIM ATXCOM(20,4)
GLBCOMBO    = 0
GLBCACTU    = 0


dim ATABGLOSLNKS(500,2)
TTLNKS = 0


 XGUTABLISKEY  = ""
 XGUTABLISKEY2 = ""
 XGUTABLISDAT  = ""

 XGUMASTER = "0"
 XGUFCOBJ  = "0"
 XGUFTOBJ  = "0"
 XGUFNIV1  = "0"
 XGUFNIV2  = "0"
 XGUFNIV3  = "0"
 XGUFNIV4  = "0"
 XGUFNIV5  = "0"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  Sets de configuraciones de Panel
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


XGBinput1   = ""
XGBinput2   = ""
XGBinput3   = ""
XGBinput4   = ""
XGBinput5   = ""
XGBinput6   = ""
XGBinput7   = ""

Styl    = ""
Sglos2  = ""
Sglos1  = ""
Sglos3  = ""

SfidkA  = ""
SfidkC  = "" 
SfidkB  = "" 

Sfidk0  = "" 
Sfidk1  = "" 
Sfidk2  = "" 
Sfidk3  = "" 
Sfidk4  = "" 

Sglcmp = ""
Sglcnt = ""
Sglprf = ""
Sgllnk = ""


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  Sets de configuraciones generales
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' Aplicacion


  if Session("CorApli") = "COREM20" then
     wkaplic = "EM"
  elseif Session("CorApli") = "COREP20" then
     wkaplic = "PL"
  elseif Session("CorApli") = "CLI916953113116518847748" then
     wkaplic = "GP"
  elseif Session("CorApli") = "CORESS20" then
     wkaplic = "CE"
  elseif Session("CorApli") = "CLI111557214260398409701" then
     wkaplic = "OR"
  else
     wkaplic = "RG"
  end if 



 'SmtpServer = "smtp.iplannetworks.net"
 'SmtpPort   = 25


if session("Usuario") = "" then
      session("LUID") = ofv.ADbUid
      session("LPWD") = ofv.ADbPwd
end if

ofv.ADbUidl = session("LUID")
ofv.ADbPwdl = session("LPWD")


  if request.querystring("MNN") <> "" then
     session("FormN") = replace(request.querystring("MNN"),"_"," ")
  end if

  if request.querystring("seg") <> "" then
     session("Form") = replace(request.querystring("seg"),"_"," ")
  end if

if not (Session("APLBGD") & "") <> "" then
   Conset(session("CONSET"))
   ofv.wstyles = CorStyler()
end if


ofv.CMBFiller   = "----- Sin Datos -----"
ofv.CMBFillOPC  = "1"



ofv.Browser()


bc = Request.servervariables("HTTP_USER_AGENT")

bcn = instr(bc,"; MSIE")
if bcn > 0 then
  xexplor = mid(bc,bcn + 7,3)
else
  xexplor = "4"
end if



if xexplor > "5.0" then
   if xexplor > "5.5" then
      Session("FULLGUI") = "S"
   ELSE
      Session("FULLGUI") = "M"
   END IF
else
   Session("FULLGUI") = "N"
end if

if  Session("USRES") = "B" then Session("FULLGUI") = "N"

 ' if session("SMPLMENU") = "S" then Session("FULLGUI") = "N"


  '''''''''''''''''''''''''''''' Session("FULLGUI") = "S"

if  Session("FULLGUI") = "S" then

    ofv.TitCls    = "ttr"
    ofv.ModTitCls = "1"
    ofv.ModBot    = "1"
    ofv.FBFcnt    = 0
    XFBClas       = "bt2"

else

    ofv.TitCls    = "ttr"
    ofv.ModTitCls = "2"
    ofv.ModBot    = "2"
    ofv.FBFcnt    = 0
    XFBClas       = "bt"

end if
 
   ofv.TitCls    = "ttx"
    ofv.ModTitCls = "2"


''''''' U - Botonera Arriba
''''''' D - Botonera Abajo

CorBarra          = "U"

ofv.wstyles       = CorStyler()
ofv.CorLogo       = CorLogo()
ofv.wbody         = cstr("bgcolor=" & Session("APLBGD") & " TOPMARGIN=1 LEFTMARGIN=5  ")

ofv.CorFormHeader = CorFormHeader2(session("FormN"),"S")

ofv.CorSetMHeader = CorSetMHeader()

ofv.CorSetFHeader = CorSetFHeader()

xxxx              = CorGenBoton(ofv.ModBot)
xxxx              = CorGenTit(ofv.ModTitCls)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Funciones de administracion de Windows
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function GTkey(XCodobJ)

GThms   = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses   = session.sessionid
GTsecu  = cstr(XCodobJ & GThms & GTses)

GTsecuf = ""

for GTi=1 to len(GTsecu)
    GTcs = mid(GTsecu,GTi,1)
    if  GTcs => "0" and GTcs <= "9" then
        GTsecuf = GTsecuf & GTcs
    end if
next

GTkey  = trim(GTsecuf)

end function

Function ConsCambios(sASP)

     Dim wxs
     wxs = wxs & "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " function AbrirCambios()" & Chr(13)
     wxs = wxs & " {" & Chr(13)
     wxs = wxs & " window.showModalDialog(" & Chr(34) & sASP & Chr(34) & ",window,"
     wxs = wxs &  Chr(34) & "dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"
     wxs = wxs & "dialogTop:250px;dialogLeft:300px;status:no" & Chr(34)
     wxs = wxs & " ); " & Chr(13)
     wxs = wxs & "}" & Chr(13)
     wxs = wxs & " </script>" & Chr(13)

 ConsCambios =      wxs

End Function


Function Ordenar(sASP)



     Dim wxs

     wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bor1", "ORdenar", "button", "", "10", "bt", "AbrirOrdenar()") & "<br><br>", "si")
     wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

     wxs = wxs & "<script language=JavaScript>" & Chr(13)
     wxs = wxs & "function AbrirOrdenar() { " & Chr(13)
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
     wxs = wxs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
     wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
     wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

    Ordenar =      wxs

End Function

Function Aprocesos(sASP)



     Dim wxs

     wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bor1", "Procesos", "button", "", "10", "bt", "AbrirProcesos()") & "<br><br>", "si")
     wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

     wxs = wxs & "<script language=JavaScript>" & Chr(13)
     wxs = wxs & "function AbrirProcesos() { " & Chr(13)
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
     wxs = wxs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
     wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
     wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

    Aprocesos =      wxs

End Function


Function Editar(sASP)



     Dim wxs

     wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bed1", "Editar", "button", "", "10", "bt", "AbrirVentana()") & "<br><br>", "si")
     wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

     wxs = wxs & "<script language=JavaScript>" & Chr(13)
     wxs = wxs & "function AbrirVentana() { " & Chr(13)
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
     wxs = wxs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
     wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
     wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

    Editar =      wxs

End Function

Function Eliminar(sASP)


    Dim      wxs

     wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bel1", "Eliminar", "button", "", "10", "bt", "AbrirEliminar()") & "<br><br>", "si")
     wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

     wxs = wxs & "<script language=JavaScript>" & Chr(13)
     wxs = wxs & "function AbrirEliminar() {" & Chr(13)
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
     wxs = wxs & "menubar=no,toolbar=no,height=250,width=550" & Chr(34)
     wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
     wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

     Eliminar =      wxs


End Function

Function Copiar(sASP)



     Dim      wxs

     wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bco1", "Copiar", "button", "", "10", "bt", "AbrirCopiar()") & "<br><br>", "si")
     wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

     wxs = wxs & "<script language=JavaScript>" & Chr(13)
     wxs = wxs & "function AbrirCopiar() {" & Chr(13)
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
     wxs = wxs & "menubar=no,toolbar=no,height=250,width=400" & Chr(34)
     wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
     wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

     Copiar =      wxs


End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Funciones de administracion de EMail 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function SendMail(ByVal strEmailBodyMessage, ByVal strRecipientEmailAddress , ByVal strFromEmailAddress, ByVal strSubject, strMailComponent, blnHTML,Byval Importance, byval notification)

	'Dimension variables
       Dim objCDOSYSMail		'Holds the CDOSYS mail object
       Dim objCDOMail			'Holds the CDONTS mail object
       Dim objJMail			'Holds the Jmail object
       Dim objAspEmail			'Holds the Persits AspEmail email object
       Dim objAspMail			'Holds the Server Objects AspMail email object
       Dim strEmailBodyAppendMessage	'Holds the appended email message
       Dim SmtpServer
       Dim SmtpPort

       SmtpServer = Session("SMTPServer")
       SmtpPort   = Session("SMTPPort")


       SHTMLMail = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">"
       SHTMLMail = SHTMLMail & "<html>"
       SHTMLMail = SHTMLMail & "<head>"
       SHTMLMail = SHTMLMail & "<title>" & subject & "</title>"
       SHTMLMail = SHTMLMail & "</head>"
       SHTMLMail = SHTMLMail & "<body bgcolor=""FFFFFF"">"
       SHTMLMail = SHTMLMail & strEmailBodyMessage

       SHTMLMail = SHTMLMail & "</body>"
       SHTMLMail = SHTMLMail & "</html>"


	'******************************************
	'***	        Mail components        ****
	'******************************************

	'Select which email component to use
	Select Case ucase(strMailComponent)


		'******************************************
		'***	  MS no envia     ****
		'******************************************

		'CDOSYS mail component
		Case "NOSEND"



		'******************************************
		'***	  MS CDOSYS mail component     ****
		'******************************************

		'CDOSYS mail component
		Case "CDOSYS"

			'Dimension variables
			Dim objCDOSYSCon

			'Create the e-mail server object
			    Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		    	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

		    	'Set and update fields properties

		        	'Out going SMTP server
		        	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SmtpServer
		        	'SMTP port
		        	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")  =SmtpPort
		        	'CDO Port
		        	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		        	
		        	'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 
				'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "soptec@corconsulting.com.ar"
				'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "bb4515"

		        	'Timeout
		        	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	        		objCDOSYSCon.Fields.Update


				'Update the CDOSYS Configuration
				Set objCDOSYSMail.Configuration = objCDOSYSCon


				'Who the e-mail is from
				'objCDOSYSMail.From = strFromEmailName & " <" & strFromEmailAddress & ">"
				objCDOSYSMail.From = strFromEmailAddress

				'Who the e-mail is sent to
				'objCDOSYSMail.To = strRecipientName & " <" & strRecipientEmailAddress & ">"
				 objCDOSYSMail.To =  strRecipientEmailAddress

				'The subject of the e-mail
				objCDOSYSMail.Subject = strSubject

				'Set the e-mail body format (HTMLBody=HTML TextBody=Plain)
				If UCASE(blnHTML)="SI" Then
				 	objCDOSYSMail.HTMLBody = SHTMLMail
				Else
					objCDOSYSMail.TextBody = SHTMLMail
				End If
                if UCASE(notification)="SI" then
                   objCDOSYSMail.fields("urn:schemas:mailheader:disposition-notification-to") =chr34 & strFromEmailAddress & chr(34) ' "<valid e-mail address>" 'ToDo: Type a valid e-mail address.
		   objCDOSYSMail.fields("urn:schemas:mailheader:return-receipt-to") = chr34 & strFromEmailAddress & chr(34) 			' "<valid e-mail address>"  'ToDo: Type a valid e-mail address.
                end if
                select case importance
                case 0 'BAJA
                 objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 5
                case 1 'NORMAL
                  objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 3
                case 2  'ALTA
                   objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 1

                end select

				objCDOSYSMail.Fields.Update

				'Send the e-mail
				objCDOSYSMail.Send


			'Close the server mail object
			Set objCDOSYSMail = Nothing




		'******************************************
		'***  	  MS CDONTS mail component     ****
		'******************************************

		'CDONTS mail component
		Case "CDONTS"






			'Create the e-mail server object
			Set objCDOMail = Server.CreateObject("CDONTS.NewMail")


				'Who the e-mail is from
				objCDOMail.From =  strFromEmailAddress

				'Who the e-mail is sent to
				objCDOMail.To = strRecipientEmailAddress

				'The subject of the e-mail
				objCDOMail.Subject = strSubject

				'The main body of the e-amil
				objCDOMail.Body = SHTMLMail

				'Set the e-mail body format (0=HTML 1=Text)
				If blnHTML ="SI" Then
					objCDOMail.BodyFormat = 0
				Else
					objCDOMail.BodyFormat = 1
				End If

				'Set the mail format (0=MIME 1=Text)
				objCDOMail.MailFormat = 0

				'Importance of the e-mail (0=Low, 1=Normal, 2=High)
				objCDOMail.Importance = importance

                               if UCASE(notification) = "SI" then
                                  strReply_To = "<" & From & ">"
                                  objCDOMail.Value("Disposition-Notification-To") = strReply_To
                               end if

				'Send the e-mail
				objCDOMail.Send


			'Close the server mail object
			  Set objCDOMail = Nothing




		'******************************************
		'***  	  w3 JMail mail component      ****
		'******************************************

		'JMail component
		Case "Jmail"

			'Create the e-mail server object
			Set objJMail = Server.CreateObject("JMail.SMTPMail")

			With objJMail
				'Out going SMTP mail server address
				.ServerAddress = strIncomingMailServer

				'Who the e-mail is from
				.Sender = strFromEmailAddress
				.SenderName = strFromEmailName

				'Who the e-mail is sent to
				.AddRecipient strRecipientEmailAddress

				'The subject of the e-mail
				.Subject = strSubject

				'Set the e-mail body format (BodyHTML=HTML Body=Text)
				If blnHTML = True Then
					.HTMLBody = strEmailBodyMessage & strEmailBodyAppendMessage
				Else
					.Body = strEmailBodyMessage & strEmailBodyAppendMessage
				End If

				'Importance of the e-mail
				.Priority = 3

				'Send the e-mail
				If NOT strIncomingMailServer = "" Then .Execute
			End With

			'Close the server mail object
			Set objJMail = Nothing




		'******************************************
		'*** Persits AspEmail mail component   ****
		'******************************************

		'AspEmail component
		Case ucase("AspEmail")
               response.write strMailComponent
			'Create the e-mail server object
			Set objAspEmail = Server.CreateObject("Persits.MailSender")


				'Out going SMTP mail server address
				objAspEmail.Host = SmtpServer

				'Who the e-mail is from
				objAspEmail.From = strFromEmailAddress
				'.FromName = strFromEmailName

				'Who the e-mail is sent to
				objAspEmail.AddAddress strRecipientEmailAddress
              '  response.write strRecipientEmailAddress
				'The subject of the e-mail
				objAspEmail.Subject = strSubject

				'Set the e-mail body format (BodyHTML=HTML Body=Text)
				If ucase(blnHTML) = "SI" Then
					objAspEmail.IsHTML = True
				End If

				'The main body of the e-mail
				objAspEmail.Body = SHTMLMail

				 select case importance
				   case 0 'BAJA
				    objAspEmail.Priority=5
				   case 1 'NORMAL
				    objAspEmail.Priority=3
				   case 2  'ALTA
				    objAspEmail.Priority=1
                 end select


                 if UCASE(notification)="SI" then
                  objAspEmail.AddCustomHeader "Return-Receipt-To: " & chr(34) & strFromEmailAddress & chr(34)
 				 end if


				'Send the e-mail
				objAspEmail.Send


			'Close the server mail object
			Set objAspEmail = Nothing




		'********************************************
		'*** ServerObjects AspMail mail component ***
		'********************************************

		'AspMail component
		Case "AspMail"

		   	'Create the e-mail server object
		   	Set objAspMail = Server.CreateObject("SMTPsvg.Mailer")

		   	With objAspMail
			   	'Out going SMTP mail server address
			   	.RemoteHost = strIncomingMailServer

			   	'Who the e-mail is from
			   	.FromAddress = strFromEmailAddress
			   	.FromName = strFromEmailName

			   	'Who the e-mail is sent to
			   	.AddRecipient " ", strRecipientEmailAddress

			   	'The subject of the e-mail
			   	.Subject = strSubject

			   	'Set the e-mail body format (BodyHTML=HTML Body=Text)
			   	If blnHTML = True Then
			    		.ContentType = "text/HTML"
			   	End If

			   	'The main body of the e-mail
			   	.BodyText = strEmailBodyMessage & strEmailBodyAppendMessage

			   	'Send the e-mail
			   	If NOT strIncomingMailServer = "" Then .SendMail
			   End With

		   	'Close the server mail object
		   	Set objAspMail = Nothing
	End Select

	'Set the returned value of the function to true
	SendMail = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' COR LOGF Include -  Sets de config. general y funciones - V 2.8.6   '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''
function CorGenTit(CModBot)

bc = Request.servervariables("HTTP_USER_AGENT")

bcn = instr(bc,"; MSIE")
if bcn > 0 then

  xexplor = mid(bc,bcn + 7,1)

else

  xexplor = "4"

end if

    if xexplor > "5" then
       xalign="center"
    else
       xalign="left"
    end if

 xfstilo = "  <v:roundrect arcsize='75%' id='AP'  style='z-index:4;position:RELATIVE;width:50px;height:20px;' >"

'''''''' 0. Tit.Columnas del Panel

        XinD   = 0

        Xinput = ""
        Xinput2 = "Titulo"
        Xinput3 = Session("CMFPNTITFTFC")
        Xinput4 = ofv.titcls
        Xinput5 = "Encabezados"
        Xinput6 = "no"
        Xinput7 = "Encabezados"

      '''  Xinput = Xinput & xfstilo


        Xinput = Xinput & " <v:line from='0,0' to='1,0' id='XXXIDXXX' style='z-index:9;align:left;position:relative;top:-1;width:10px;height:10px;' > "
        Xinput = Xinput & "   <v:fill on='True' color='" & session("CMFPNTIT") & "' color2='" & session("CMFPNTIT") & "' method='sigma' type='gradient' angle='180' opacity='95%'  /> "
        Xinput = Xinput & "   <v:stroke on='False' /> "
        Xinput = Xinput & "   <v:shadow on='True' color='#ffffff' weight='4px' /> "
        Xinput = Xinput & "   <v:path textpathok='True' /> "
        Xinput = Xinput & "     <v:textpath on='True' string='XXXDATOXXX' style=""font-size:'" & Session("CMFPNTITFTSZ") + 1 & "pt';font-family:'" & Session("CMFPNTITFTFC") & "';font-weight:bold;""  /> "
    '    Xinput = Xinput & "   <v:extrusion on='True' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
        Xinput = Xinput & " </v:line>"


        xxxx = OFV.TitBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))



'''''''' 1. Tit.Columnas del Panel

        XinD   = 1

        Xinput = ""
        Xinput2 = "Titulo"
        Xinput3 = Session("CMFPNTITFTFC")
        Xinput4 = ofv.titcls
        Xinput5 = "Encabezados"
        Xinput6 = "no"
        Xinput7 = "Encabezados"

      '''  Xinput = Xinput & xfstilo

        Xinput = Xinput & "  <v:roundrect arcsize='25%'  id='XXXIDXXX'  style='z-index:9;margin-top:0;position:relative;left:0px;top:0px;width:XXXWIDTHXXX;height:20px;'  >"
        Xinput = Xinput & "	<v:fill  type='gradientradial' method='sigma' color='" & session("CMFPNTIT") & "' color2='#eeeeee' colors='95% " & session("CMFPNTIT") & ", 0% #eeeeee' "
        Xinput = Xinput & "      opacity='100%' angle='45'  focusposition=""95%, 15%""  />"
        Xinput = Xinput & "       <v:stroke on='True' weight='2px' color='#eeeeee' />	"
        Xinput = Xinput & "       <v:shadow on='False' color='#dddddd' weight='3px' offset='3px, 2px' />"
        Xinput = Xinput & "        <v:textbox inset='10px,1px,0px,0px' style='font-size:" & Session("CMFPNTITFTSZ") & "pt;font-family:" & Session("CMFPNTITFTFC") & ";font-weight:900;color:" & Session("CMFPNTITFT") & ";text-transform:uppercase;' >XXXDATOXXX</v:textbox>"
      '  Xinput = Xinput & "        <v:extrusion on='True' metal='True'  type='perspective' backdepth='6'  rotationcenter='0,0,0'  rotationangle = '30, 0'/>"
        Xinput = Xinput & "    </v:roundrect>"




        xxxx = OFV.TitBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


'''''''' 5. Menus

        XinD   = 5

        Xinput = ""
        Xinput2 = "Menus"
        Xinput3 = Session("CMFPNTITFTFC")
        Xinput4 = ofv.titcls
        Xinput5 = "Encabezados"
        Xinput6 = "no"
        Xinput7 = "Encabezados"




        Xinput = Xinput & " <v:line from='0,0' to='1,0' id='XXXIDXXX' style='z-index:9;align:left;position:relative;top:-1;width:10px;height:10px;' > "
        Xinput = Xinput & "   <v:fill on='True' color='" & Session("CMFMNBOT") & "' color2='" & Session("CMFMNBOT") & "' method='sigma' type='gradient' angle='180' opacity='95%'  /> "
        Xinput = Xinput & "   <v:stroke on='False' /> "
        Xinput = Xinput & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
        Xinput = Xinput & "   <v:path textpathok='True' /> "
        Xinput = Xinput & "     <v:textpath on='True' string='XXXDATOXXX' style=""font-size:'" & Session("CMFMNFTSZ") + 1 & "pt';font-family:'" & Session("CMFMNFTFC") & "';font-weight:bold;""  /> "
        Xinput = Xinput & "   <v:extrusion on='True' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
        Xinput = Xinput & " </v:line>"


        xxxx = OFV.TitBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))





'''''''' 6. Menus

        XinD   = 6

        Xinput = ""
        Xinput2 = "Menus"
        Xinput3 = Session("CMFPNTITFTFC")
        Xinput4 = ofv.titcls
        Xinput5 = "Encabezados"
        Xinput6 = "no"
        Xinput7 = "Encabezados"



        Xinput = Xinput & "  <v:roundrect arcsize='75%'  id='XXXIDXXX'  style='z-index:9;margin-top:0;position:relative;left:0px;top:0px;width:XXXWIDTHXXX;height:20px;'  >"
        Xinput = Xinput & "	<v:fill  type='gradientradial' method='sigma' color='" & Session("CMFMNBOT") & "' color2='#eeeeee' colors='95% " & Session("CMFMNBOT") & ", 0% #eeeeee' "
        Xinput = Xinput & "      opacity='100%' angle='45'  focusposition=""95%, 15%""  />"
        Xinput = Xinput & "       <v:stroke on='True' weight='1px' color='#eeeeee' />	"
        Xinput = Xinput & "       <v:shadow on='True' color='#eeeeee' weight='2px'  />"
        Xinput = Xinput & "        <v:textbox inset='10px,1px,0px,0px' style='font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";font-weight:900;color:" & Session("CMFMNFT") & ";text-transform:uppercase;' >XXXDATOXXX</v:textbox>"
      '  Xinput = Xinput & "        <v:extrusion on='True' metal='True' type='perspective' backdepth='6'  rotationcenter='0,0,0'  rotationangle = '30, 0'/>"
        Xinput = Xinput & "    </v:roundrect>"




        xxxx = OFV.TitBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))




end function


'''''''''
function CorGenBoton(CModBot)

bc = Request.servervariables("HTTP_USER_AGENT")

bcn = instr(bc,"; MSIE")
if bcn > 0 then

  xexplor = mid(bc,bcn + 7,1)

else

  xexplor = "4"

end if

    if xexplor > "5" then
       xalign="center"
    else
       xalign="left"
    end if

''color='#aaCCEE' color2='#225588'

 xfstilo = "  <v:roundrect arcsize='75%' id='AP'  style='z-index:10;position:RELATIVE;width:45px;height:20px;' >"
'' xfsextr = "  <v:extrusion on='True' metal='True' type='perspective' backdepth='6'  rotationcenter='0,0,0'  rotationangle = '30, 0'/>"
'' xfsshdw = "  <v:shadow on='False' />"
 xfsshdw = "  <v:shadow on='True' color='#eeeeee' weight='1px' /><v:stroke on='True' color='#eeeeee' weight='1px' />"
 xfsfill = "  <v:fill on='true' color='XXXCOL1XXX' color2='XXXCOL2XXX' type='gradient' angle='30' opacity='100%'  />"
'' xfsfill = " <v:fill  type='gradientradial' method='sigma' color='XXXCOL2XXX' color2='XXXCOL1XXX' colors='95% XXXCOL2XXX, 0% XXXCOL1XXX' opacity='100%' angle='45'  focusposition=""95%, 15%""  />"

        XSCOLOR2 = Session("CMFPNBOT")
        XSCOLOR1 = Session("CMFMENU")

        xfsfillD = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
        xfsfillD = replace(xfsfillD,"XXXCOL2XXX",XSCOLOR2)

        Xtypm  = "D"
        Xweig  = "600" 


'''''''' 0. Bot.Help

        XinD   = 0

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "s"
        Xinput3 = "WebDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Ayuda On Line"
        Xinput6 = "_self"
        Xinput7 = "Ayuda"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#aaCCEE"
        XSCOLOR2 = "#225588"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

      '''  Xinput = Xinput & "     <v:fill on='true' color='#aaCCEE' color2='#225588' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"


        xxxx = OFV.AddBot(clng(XinD), "", cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


'''''''' 1. Bot.ANTERIOR

        XinD   = 1

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "3"
        Xinput3 = "WebDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Pagina Anterior"
        Xinput6 = "_self"
        Xinput7 = "Anterior"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#aaCCEE"
        XSCOLOR2 = "#225588"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

     ''   Xinput = Xinput & "     <v:fill on='true' color='#aaCCEE' color2='#225588' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"


        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


'''''''' 2. Bot.Siguiente

        XinD   = 2

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "4"
        Xinput3 = "WebDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Pagina Siguiente"
        Xinput6 = "_self"
        Xinput7 = "Siguiente"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#aaCCEE"
        XSCOLOR2 = "#225588"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

      ''  Xinput = Xinput & "     <v:fill on='true' color='#aaCCEE' color2='#225588' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"




        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))



'''''''' 3. Bot.Agregar

        XinD   = 3

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "5"
        Xinput3 = "WebDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Agregar Elemento"
        Xinput6 = "_self"
        Xinput7 = "Agregar"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#ff0000"
        XSCOLOR2 = "#770000"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

      ''  Xinput = Xinput & "     <v:fill on='true' color='#ff0000' color2='#770000' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


'''''''' 4. Bot.Buscar

        XinD   = 4

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "L"
        Xinput3 = "webdings"
        Xinput4 = XFBClas
        Xinput5 = "Buscar Elemento"
        Xinput6 = "_self"
        Xinput7 = "Buscar"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#0000ff"
        XSCOLOR2 = "#336699"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

   ''     Xinput = Xinput & "     <v:fill on='true' color='#0000ff' color2='#336699' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:12pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


'''''''' 5. Bot.Consultar

        XinD   = 5

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "1"
        Xinput3 = "WingDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Consultar - Lista"
        Xinput6 = "_blank"
        Xinput7 = "Consultar"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#ffbb22"
        XSCOLOR2 = "#bbaa22"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

       '' Xinput = Xinput & "     <v:fill on='true' color='#ffbb22' color2='#bbaa22' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


'''''''' 6. Bot.Test

        XinD   = 6

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "J"
        Xinput3 = "wingDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Test"
        Xinput6 = "_blank"
        Xinput7 = "Test"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#0000ff"
        XSCOLOR2 = "#336699"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

     ''   Xinput = Xinput & "     <v:fill on='true' color='#0000ff' color2='#336699' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-2px,0px,0px' style=""text-align:" & xalign & ";font-size:13pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))

'''''''' 7. Bot.Volver

        XinD   = 7

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "&#0231;"
        Xinput3 = "wingDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Volver"
        Xinput6 = "_self"
        Xinput7 = "Volver"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#005500"
        XSCOLOR2 = "#00cc00"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

     ''   Xinput = Xinput & "     <v:fill on='true' color='#005500' color2='#00cc00' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:12pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))

'''''''' 8. Bot.Eliminar

        XinD   = 8

        Xtyp   = "C"
      ''  Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "4"
        Xinput3 = "webDINGS"
        Xinput4 = "bte"
        Xinput5 = "Eliminar"
        Xinput6 = "_self"

        if ofv.modbot = "2" then
           Xinput7 = "4"
        else
           Xinput7 = "Eliminar"
        end if

       ''' Xinput = Xinput & xfstilo
        Xinput = Xinput & "  <v:roundrect arcsize='65%' id='AP'  style='z-index:4;position:RELATIVE;width:12px;height:12px;' >"
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#ff0000"
        XSCOLOR2 = "#770000"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

        Xinput = Xinput & "     <v:fill on='true' color='#ff0000' color2='#770000' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-4px,0px,0px' style=""text-align:" & xalign & ";font-size:7pt;font-family:'XXXFONTXXX';font-weight:400;color:#FFFFFF;"" >XXXDATOXXX"
        Xinput = Xinput & "     </v:textbox>"
        Xinput = Xinput & xfsshdw
       ' Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


'''''''' 11. Bot.ANTERIOR gris

        XinD   = 11

        Xtyp   = "C"
      ''  Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "3"
        Xinput3 = "WebDINGS"
        Xinput4 = XFBClas & "G"
        Xinput5 = "Pagina Anterior no Disponible"
        Xinput6 = "_self"
        Xinput7 = "Anterior"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#eeeeee"
        XSCOLOR2 = "#aaaaaa"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

      ''  Xinput = Xinput & "     <v:fill on='true' color='#eeeeee' color2='#aaaaaa' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))



'''''''' 12. Bot.Siguiente gris

        XinD   = 12

        Xtyp   = "C"
      ''  Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "4"
        Xinput3 = "WebDINGS"
        Xinput4 = XFBClas & "G"
        Xinput5 = "Pagina Siguiente no Disponible"
        Xinput6 = "_self"
        Xinput7 = "Siguiente"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#eeeeee"
        XSCOLOR2 = "#aaaaaa"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

   ''     Xinput = Xinput & "     <v:fill on='true' color='#eeeeee' color2='#aaaaaa' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))

'''''''' 13. Bot.Aceptar

        XinD   = 13

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "a"
        Xinput3 = "WebDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Aceptar Operacion"
        Xinput6 = "_self"
        Xinput7 = "Aceptar"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#0000ff"
        XSCOLOR2 = "#336699"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

     ''   Xinput = Xinput & "     <v:fill on='true' color='#0000ff' color2='#336699' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))

'''''''' 14. Bot.Cancelar

        XinD   = 14

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "r"
        Xinput3 = "WebDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Cancelar Operacion"
        Xinput6 = "_self"
        Xinput7 = "Cancelar"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#ff0000"
        XSCOLOR2 = "#770000"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

     ''   Xinput = Xinput & "     <v:fill on='true' color='#ff0000' color2='#770000' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))

'''''''' 15. Bot.Confirmar

        XinD   = 15

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "a"
        Xinput3 = "WebDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Confirmar"
        Xinput6 = "_self"
        Xinput7 = "Confirmar Operacion"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#0000ff"
        XSCOLOR2 = "#336699"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

     ''   Xinput = Xinput & "     <v:fill on='true' color='#0000ff' color2='#336699' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


'''''''' 16. Bot.Esp1

        XinD   = 16

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "2"
        Xinput3 = "webDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Consultar"
        Xinput6 = "_blank"
        Xinput7 = "Consultar"



        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#f7aa22"
        XSCOLOR2 = "#b79911"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

    ''    Xinput = Xinput & "     <v:fill on='true' color='#f7aa22' color2='#b79911' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-2px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))

'''''''' 17. Bot.Indice

        XinD   = 17

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "&#0105;"
        Xinput3 = "webDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Abrir/Cerrar indice"
        Xinput6 = "_self"
        Xinput7 = "Indice"



        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#f7aa22"
        XSCOLOR2 = "#b79911"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

    ''    Xinput = Xinput & "     <v:fill on='true' color='#f7aa22' color2='#b79911' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-2px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))

'''''''' 18. Bot.Glosario

        XinD   = 18

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "&#062;"
        Xinput3 = "webDINGS"
        Xinput4 = XFBClas
        Xinput5 = "Glosario"
        Xinput6 = "_self"
        Xinput7 = "Glosario"



        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#f7aa22"
        XSCOLOR2 = "#b79911"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF

    ''    Xinput = Xinput & "     <v:fill on='true' color='#f7aa22' color2='#b79911' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-2px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


'''''''' 20. Bot.INICIO

        XinD   = 20

      ''  Xtyp   = "C"
        Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "INGRESAR"
        Xinput3 = "verdana"
        Xinput4 = XFBClas
        Xinput5 = "Ingresar"
        Xinput6 = "_self"
        Xinput7 = "Ingresar"

        xfstilo = "  <v:roundrect arcsize='75%' id='AP'  style='z-index:10;position:RELATIVE;width:90px;height:25px;' >"
        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	

        XSCOLOR1 = "#aaCCEE"
        XSCOLOR2 = "#225588"

        IF  Xtyp   = "C" THEN
            xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
            xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
            Xinput = Xinput & xfsfillC
        ELSE
            Xinput = Xinput & xfsfillD
        END IF


        Xinput = Xinput & "	<v:textbox inset='0px,2px,0px,0px' style=""text-align:" & xalign & ";font-size:7pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;" & XspDShw("1","1",session("CMFPNTIT")) & ";"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"




        xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


end function


'''''''''
function CorGenBotonXp(XCPinD, XCPinput, XCPinput2, XCPinput3, XCPinput4, XCPinput5, XCPinput6, XCPinput7)


        xxxx = OFV.GetBot(clng(XCPinD))


    '    Xinput = ""
    '    Xinput2 = "2"               ' caracter
    '    Xinput3 = "webDINGS"        ' font
    '    Xinput4 = XFBClas
    '    Xinput5 = "Consultar"       ' hint
    '    Xinput6 = "_blank"          ' target
    '    Xinput7 = "Consultar"       ' nombre  


        Xinput  = OFV.FBotDat1
        Xinput2 = OFV.FBotDat2
        Xinput3 = OFV.FBotDat3
        Xinput4 = OFV.FBotDat4
        Xinput5 = OFV.FBotDat5
        Xinput6 = OFV.FBotDat6
        Xinput7 = OFV.FBotDat7


        if  XCPinput  <> "NO" then Xinput  = XCPinput
        if  XCPinput2 <> "NO" then Xinput2 = XCPinput2
        if  XCPinput3 <> "NO" then Xinput3 = XCPinput3
        if  XCPinput4 <> "NO" then Xinput4 = XCPinput4
        if  XCPinput5 <> "NO" then Xinput5 = XCPinput5
        if  XCPinput6 <> "NO" then Xinput6 = XCPinput6
        if  XCPinput7 <> "NO" then Xinput7 = XCPinput7

        xxxx = OFV.AddBot(clng(XCPinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))


end function


''''''''''''

function GenTitulo(xfDato,xfModo)

            

            xGenCont = xfDato

            If ofv.ModTitCls = "1" Or ofv.ModTitCls = "5" Or ofv.ModTitCls = "0" Then

               if xfModo = 1 then 

                  xfInd = clng(ofv.ModTitCls)
               else

                  xfInd = 0
               end if

               xGenConta = ofv.GetTit(xfInd + 5)
               xGenConta = ofv.FTitDat1

               xGenCont = xGenConta

           ''    xGenCont = FTitTab(CLng(ModTitCls), 0)

               If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then

                  xGenWidth = CStr(Len(xfDato) * 10) & "px"
                  xGenCont = Replace(xGenCont, "XXXDATOXXX", xfDato)
                  xGenCont = Replace(xGenCont, "XXXWIDTHXXX", xGenWidth)
                ''  xGenCont = Replace(xGenCont, "XXXIDXXX", "r" & xGenNom)

               End If

           End If

        GenTitulo = xGenCont

end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''' FUNCIONES DE MASTER-PANELES / PANELES DE MENUS / PANELES DE DOCUMENTOS Y COMPONENTES ''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Function CorPnlFnc(XCopc,XCdsc)

     dim Whtml
     dim Whtml1

     Whtml  = ""
     Whtml1 = ""


'''''''''''''''''''''''''

 Whtml = Whtml & "<style ID='oStyle' > span.gls {cursor:hand;background-color:#bbbbbb;color:white!important;"
 Whtml = Whtml & "border-style:outset;border-width:1;border-color:#888888;line-height: 25px;font-weight:bold;} </style>"



 Whtml = Whtml & "<public>" & chr(13)
 Whtml = Whtml & "<PROPERTY name='myWin' />"
 Whtml = Whtml & "<method name='creawin' />"
 Whtml = Whtml & " </public>" & chr(13)

 Whtml = Whtml & "<SCRIPT language='javascript' for='mt2' event='onmouseout' >" & chr(13) & chr(13)
 Whtml = Whtml & " if (document.selection.type=='Text') {;"
 Whtml = Whtml & "     myGlos = document.selection.createRange().text;"
 Whtml = Whtml & " };"
 Whtml = Whtml & "</SCRIPT>" & chr(13) & chr(13)

 Whtml = Whtml & "<SCRIPT language=javascript >             " & chr(13) & chr(13)
 Whtml = Whtml & "    var myWin;                            " & chr(13)
 Whtml = Whtml & "    var myGlos;                           " & chr(13)
 Whtml = Whtml & " oStyle.disabled = true;                  " & chr(13)
 Whtml = Whtml & " function creawin(spmk) {                 " & chr(13)
 Whtml = Whtml & " document.all[spmk].scrollIntoView(true); " & chr(13)
 Whtml = Whtml & " clIndice();                              " & chr(13)
 Whtml = Whtml & " }                                        " & chr(13)

 Whtml = Whtml & "function AbrirCons(Datos,nom) {" & Chr(13)
 Whtml = Whtml & " myWin= open(Datos"
 Whtml = Whtml & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," & Chr(34)
 Whtml = Whtml & "status=yes,menubar=no,toolbar=no,height=570,width=790" & Chr(34)
 Whtml = Whtml & " ); " & Chr(13) & " } " & Chr(13)

''''''
 Whtml = Whtml & "function setglos()  " & Chr(13)
 Whtml = Whtml & "{                   " & Chr(13)
 Whtml = Whtml & "   myGlos = document.selection.createRange().text;" & Chr(13)
 Whtml = Whtml & "}                   " & Chr(13)

 Whtml = Whtml & "function verglos()" & Chr(13)
 Whtml = Whtml & "{" & Chr(13)
 Whtml = Whtml & " var xtxt = myGlos;"
 Whtml = Whtml & " var sasp = 'GLOSCONS.ASP?DATOS=" & Chr(34) & "' + xtxt + '" & Chr(34) & "';"
' Whtml = Whtml & " alert(xtxt);"
 Whtml = Whtml & " window.showModalDialog(sasp,window,"
 Whtml = Whtml &  Chr(34) & "dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"
 Whtml = Whtml & "dialogTop:250px;dialogLeft:300px;status:no" & Chr(34)
 Whtml = Whtml & " ); " & Chr(13)
 Whtml = Whtml & " myGlos = '';"
 Whtml = Whtml & "}" & Chr(13)

 Whtml = Whtml & " function versup() { "
 Whtml = Whtml & " if (oStyle.disabled == false) "
 Whtml = Whtml & " {oStyle.disabled = true;} "
 Whtml = Whtml & " else "
 Whtml = Whtml & " {oStyle.disabled = false;} "
 Whtml = Whtml & " } "


 Whtml = Whtml & " function verglos2(xdata)" & Chr(13)
 Whtml = Whtml & " {" & Chr(13)
 Whtml = Whtml & " var xtxt = xdata;"
 Whtml = Whtml & " var sasp = 'GLOSCONS.ASP?DATOS=" & Chr(34) & "' + xtxt + '" & Chr(34) & "';"
 Whtml = Whtml & " window.showModalDialog(sasp,window,"
 Whtml = Whtml &  Chr(34) & "dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"
 Whtml = Whtml & "dialogTop:250px;dialogLeft:300px;status:no" & Chr(34)
 Whtml = Whtml & " ); " & Chr(13)
 Whtml = Whtml & "}" & Chr(13)


 Whtml = Whtml & "function clIndice() { " & chr(13)
 Whtml = Whtml & "  if  (document.all.DHEX1.style.display=='none')   " & chr(13)
 Whtml = Whtml & "      { document.all.DHEX1.style.display='block';} " & chr(13)
 Whtml = Whtml & "  else                                             " & chr(13)
 Whtml = Whtml & "      { document.all.DHEX1.style.display='none';}  " & chr(13)
 Whtml = Whtml & " } " & chr(13)

 Whtml = Whtml & "function DCEXRCLC() {"
 Whtml = Whtml & " var xrclc0 = document.body.clientWidth;"
 Whtml = Whtml & " var xrclc1 = document.body.clientHeight;"
 Whtml = Whtml & " var xrclc2 = document.all.TCEX1.clientHeight;"
 Whtml = Whtml & " var xrclc3 = document.all.TCEX1.offsetTop;"
 Whtml = Whtml & " document.all.mt2.style.height = xrclc1 - (xrclc3 + xrclc2) - 5;"
 Whtml = Whtml & " document.all.mt2.style.width  = xrclc0 - 10;"
 Whtml = Whtml & "}"

 Whtml = Whtml & "</SCRIPT>" & chr(13) & chr(13)


  styl    = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee;"
  styl    = styl & " SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa;"
  styl    = styl & " SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"

  sglos2  = "<b><font face=wingdings size=3 color=#ffffff>&</font></b>"
  sglos1  = "<b><font face=wingdings size=3 color=#ffffff>q</font></b>"
  sglos3  = "<b><font face=wingdings size=3 color=#ffffff>n</font></b>"

'  sfidkA  = "<font face=webdings size=1 color=#99ccee >&#060;</font>"

  sfidkA  = "<font face=webdings size=2 color=#99ccee >&#0157;</font>" 
  sfidkC  = "<font face=webdings size=2 color=#f7aa22 >4</font>" 
  sfidkB  = "<font face=wingdings size=2 color=#6699cc >&#0159;</font>"

  sfidk0  = "<font face=webdings size=3 color=" & session("CMFMNBOT") & " >&#0157;</font>" 
  sfidk1  = "<font face=webdings size=2 color=#f7aa22 >4</font>" 
  sfidk2  = "<font face=webdings size=2 color=#f7aa22 >4</font>" 
  sfidk3  = "<font face=webdings size=2 color=#00aa00 >4</font>" 
  sfidk4  = "<font face=webdings size=2 color=#00aa00 >&#060;</font>" 

  sglcmp = "<div class=btba style='width:1px;clear:right;' ><font class=btic >&#0157;</font></div>"
  sglcnt = "<div class=btbs style='width:1px;clear:right;' ><font class=btic >G</font></div>"
  sglprf = "<div class=btbn style='width:1px;clear:right;' ><font class=btic >&#0128;</font></div>"
  sgllnk = "<div class=btbn style='width:1px;clear:right;' ><font class=btic >&#0126;</font></div>"


''''''''''''''''''''''''
     CorPnlFnc = Whtml

 end function

 Function CorPnlIndx(XCopc,XCdsc)

     dim Whtml
     dim Whtml1

     Whtml  = ""
     Whtml1 = ""

''''''''''''''''''''''''

 Celdat  = "<td width=1% >&nbsp;</td>"
 Celdat  = Celdat & "<td width=99% >&nbsp;</td>"
 Filat   = filat & ofv.GenRow("","","","","5","","",Celdat,"si")

 Celdat  = "<td width=1% >&nbsp;</td>"
 Celdat  = Celdat & "<td width=99% >&nbsp;</td>"
 Filatt  = filatt & ofv.GenRow("","","","","5","","",Celdat,"si")



''''''''''''''''''''''
 Whtml = Whtml &  "<DIV class='DINDM' ID='DHEX1'  NAME='DHEX1'  style='display:block;position:absolute;top:168;left:180;HEIGHT=308px;WIDTH=558px;' >"

 Whtml = Whtml &  "<DIV class='DINDC' ID='DHEX1C' NAME='DHEX1C' style='position:relative;top:+2px;left:+2;HEIGHT=300px;WIDTH=550px;' >"
 Whtml = Whtml &  "<CENTER>"

 Whtml = Whtml &  "<table id=""MTINDX"" class='TWIND' width=100% cellspacing=0 cellpadding=2 >"
 Whtml = Whtml &  "<tr>"
 Whtml = Whtml &  "<td CLASS=mbt width=99%  >&nbsp;INDICE</td>"
 Whtml = Whtml &  "<td CLASS=mbt width=1% onclick=""clIndice();"" title='Ocultar indice.' >"
 Whtml = Whtml &  "<b><font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=webdings size=1   >r</font></b>" 
 Whtml = Whtml &  "</td>"
 Whtml = Whtml &  "</tr>"
 Whtml = Whtml &  "</table>"

 Whtml = Whtml &  "<DIV ID='DHEX2B' NAME='DHEX2B' style='display:none;position:relative;top:+1px;left:+2;HEIGHT=100px;WIDTH=544px;overflow:auto;border-bottom-style:inset;border-bottom-width:1;border-bottom-color:#dddddd;' >"
 Whtml = Whtml &  "<CENTER>"
 Whtml = Whtml &  ofv.GenTabla("TIND1","aaa style='display:block;' ","","95%","","0","0","2","0","","",Filatt,"si")
 Whtml = Whtml &  "</CENTER>"
 Whtml = Whtml &  "</div>"

 Whtml = Whtml &  "<DIV ID='DHEX2C' NAME='DHEX2C' style='position:relative;top:+1px;left:+2;HEIGHT=270px;WIDTH=544px;overflow:auto;' >"
 Whtml = Whtml &  "<CENTER>"
 Whtml = Whtml &  ofv.GenTabla("TIND2","aaa style='display:block;' ","","95%","","0","0","2","0","","",Filat,"si")
 Whtml = Whtml &  "</CENTER>"
 Whtml = Whtml &  "</div>"

 Whtml = Whtml &  "</CENTER>"
 Whtml = Whtml &  "</div>"
 Whtml = Whtml &  "</div>"

 Whtml1 =  ofv.GenCelda("","vt","center","","","","","<br><big><big>Cargando DOCUMENTO...</big><br><br>un momento por favor</big><br>","si")
 Whtml1 = ofv.GenRow("","","","","","","",Whtml1,"si")


 Whtml = Whtml & "<div id='mt1' style='display:block;' ><br><br><br><br><center>"
 Whtml = Whtml & ofv.GenTabla("","","","60% style='border-style:SOLID;border-width:2;border-color:#AAAAAA;' ","","0","","0","0","#ffffff","",Whtml1,"si")
 Whtml = Whtml & "</center></div>"


 Whtml = Whtml & "<script language=javascript >"
 Whtml = Whtml & "document.all.mt1.style.display='none';"
 Whtml = Whtml & "document.all.mt2.style.display='block';"
 Whtml = Whtml & " DCEXRCLC();"
 Whtml = Whtml & "</script>"



 if  XCopc = "S" then
     Whtml1 = ""
     Whtml1 = Whtml1 & "<script language=javascript >"
     if  SESSION("FULLGUI") = "S" THEN
         Whtml1 = Whtml1 & "   document.all.TBH1VML.textpath.string='" & XCdsc & "';" & chr(13)
     ELSE
         Whtml1 = Whtml1 & "   document.all.TBH1DV.innerHTML= '" & XCdsc & "';" & chr(13)
     END IF 
     Whtml1 = Whtml1 & "</script>"

     ofv.FBBarr = ofv.FBBarr &  Whtml1

 END IF 






''''''''''''''''''''''

     CorPnlIndx = Whtml

 end Function

'''' funciones del panel ejecutivo Cor On-Time ''''''''''''''''''''''''''''


 Function CorPnlCOT()

''''''''''''''''''''''
 Whtml = ""
 Whtml = Whtml &  "<DIV class='DINDM' ID='DHEJ1'  NAME='DHEX1'  style='display:NONE;position:absolute;top:160;left:180;HEIGHT=310px;WIDTH=558px;' >"

 Whtml = Whtml &  "<DIV class='DINDC' ID='DHEJ1C' NAME='DHEX1C' style='position:relative;top:+2px;left:+2;HEIGHT=300px;WIDTH=550px;' >"
 Whtml = Whtml &  "<CENTER>"

 Whtml = Whtml &  "<table id=""MTIJX"" class='TWIND' width=100% cellspacing=0 cellpadding=2 >"
 Whtml = Whtml &  "<tr>"
 Whtml = Whtml &  "<td CLASS=mbt width=99%  >&nbsp;Panel Ejecutivo</td>"
 Whtml = Whtml &  "<td CLASS=mbt width=1% onclick=""clpanel();"" title='Ocultar panel.' >"
 Whtml = Whtml &  "<b><font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=webdings size=1   >r</font></b>" 
 Whtml = Whtml &  "</td>"
 Whtml = Whtml &  "</tr>"
 Whtml = Whtml &  "</table>"

 Whtml = Whtml &  "<DIV ID='DHEXJB' NAME='DHEX2B' style='display:BLOCK;position:relative;top:+1px;left:+2;HEIGHT=270px;WIDTH=544px;overflow:auto;border-bottom-style:inset;border-bottom-width:1;border-bottom-color:#dddddd;' >"

 Whtml = Whtml &  "</div>"

 Whtml = Whtml &  "</CENTER>"
 Whtml = Whtml &  "</div>"
 Whtml = Whtml &  "</div>"

 Whtml = Whtml & "<SCRIPT language=javascript >" & chr(13) & chr(13)
 Whtml = Whtml & " var xspnl = 'NO'; "
 Whtml = Whtml & "function clpanel() { " & chr(13)
 Whtml = Whtml & "  if  (xspnl=='NO')   " & chr(13)
 Whtml = Whtml & "      { document.all.DHEXJB.innerHTML=window.top.verfrm4(); " & chr(13)
 Whtml = Whtml & "        xspnl='SI';} " & chr(13)

 Whtml = Whtml & "  if  (document.all.DHEJ1.style.display=='none')   " & chr(13)
 Whtml = Whtml & "      { document.all.DHEJ1.style.display='block';} " & chr(13)
 Whtml = Whtml & "  else                                             " & chr(13)
 Whtml = Whtml & "      { document.all.DHEJ1.style.display='none';}  " & chr(13)
 Whtml = Whtml & " } " & chr(13)

 Whtml = Whtml & "</SCRIPT>" & chr(13) & chr(13)

''''''''''''''''''''''

     CorPnlCOT = Whtml

 end Function

''''''''''''''' funciones del glosario ''''''''''''''''''''''''''''


function GenKeyLnks(gtlnks)

    TTLNKS = TTLNKS + 1

    sh = datepart("h",TIME())
    sm = datepart("n",TIME())
    ss = datepart("s",TIME())

    XCodobJ = sh & sm & ss

    GThms   = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
    GTses   = session.sessionid
    GTsecu  = cstr(XCodobJ & GThms & GTses)

    GTsecuf = "(XLNK:" & TTLNKS & ":"  


    for GTi=1 to len(GTsecu)
        GTcs = mid(GTsecu,GTi,1)
        if GTcs => "0" and GTcs <= "9" then
           GTsecuf = GTsecuf & GTcs
        end if
    next
 

    GTsecuf = GTsecuf & "XLNK)"

    ATABGLOSLNKS(TTLNKS,1) = trim(GTsecuf)
    ATABGLOSLNKS(TTLNKS,2) = gtlnks

    GenKeyLnks = trim(GTsecuf)

end function



function Capitalize (str)
  Capitalize=UCase(Left(str, 1)) & LCase(Mid(str, 2))
end function



function VerGlosario(XEhtml)

 dim ATABGLOS(200,5)
 ttglos = 0

 XEconvMl = "<span CLASS=gls onclick=" & chr(34) & "verglos2('XXARG')" & chr(34) & " >XXDATA</span>"

' glosario

  strsql = "Select KEYWORD,SHWWORD From GLOSARIO ORDER BY SHWWORD"
  set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsX1.eof then
     DO WHILE NOT RSX1.EOF

        TTGLOS = TTGLOS + 1
        ATABGLOS(TTGLOS,0) = RSX1(0) 

        ATABGLOS(TTGLOS,1) = RSX1(1) 

        if  UCASE(ATABGLOS(TTGLOS,1)) = ATABGLOS(TTGLOS,1) then
           ATABGLOS(TTGLOS,2) = "XXXXXX" 
        else
           ATABGLOS(TTGLOS,2) = UCASE(ATABGLOS(TTGLOS,1))
        end if

        if  LCASE(ATABGLOS(TTGLOS,1)) = ATABGLOS(TTGLOS,1) then
           ATABGLOS(TTGLOS,3) = "XXXXXX" 
        else
           ATABGLOS(TTGLOS,3) = LCASE(ATABGLOS(TTGLOS,1)) 
        end if

        if  Capitalize (ATABGLOS(TTGLOS,1)) = ATABGLOS(TTGLOS,1) then
           ATABGLOS(TTGLOS,4) = "XXXXXX" 
        else
           ATABGLOS(TTGLOS,4) = Capitalize (ATABGLOS(TTGLOS,1))
        end if



               

     RSX1.MOVENEXT
     LOOP 
       
  end if
  call ofv.cerrarconsulta(rsX1)

  IF TTGLOS > 0 THEN

     FOR IGLOS = 1 TO TTGLOS
         XEconvBs = replace(XEconvML,"XXARG","XXARG" & (999-IGLOS))
         FOR III = 1 TO 4
             if  ATABGLOS(IGLOS,III) <> "XXXXXX" then
                 XEconv = replace(XEconvBs,"XXDATA",ATABGLOS(IGLOS,III))
                 XEhtml = REPLACE(XEhtml,ATABGLOS(IGLOS,III),XEconv)
             end if
         NEXT
     NEXT

     FOR IGLOS = 1 TO TTGLOS
         XEhtml   = REPLACE(XEhtml,"XXARG" & (999-IGLOS),ATABGLOS(IGLOS,0))
     NEXT

  END IF

  for IGLOS = 1 TO TTLNKS
      XEhtml   = REPLACE(XEhtml,ATABGLOSLNKS(IGLOS,1),ATABGLOSLNKS(IGLOS,2))
  next

  xehtmln = instr(XEhtml,"<span CLASS=gls")

  xehtmlx = ""


  if xehtmln = 0 then 

     XEhtml = replace(XEhtml,"versup() { ","versup() { alert('No se encontraron entradas del GLOSARIO en la pagina actual'); " )

  end if


  VerGlosario = XEhtml

end function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



 Function CorLogo()

 dim Whtml

 Whtml = ""

 Whtml = Whtml & "<table border=0 cellspacing=0 cellpadding=0 bgcolor=#eeeeee width=100% height=100% style='border-style:ridge;border-width:3;' >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<td align=center style='border-style:ridge,border-width:2; " & XspSShw("#777777","135") & ";' >"
 Whtml = Whtml & "<IMG border=0 src='imagenes/corlogo.gif' style='filter: chroma(color=#ffffff);' >"
 Whtml = Whtml & "<FONT color=#ffffff face=verdana size=2>"
 Whtml = Whtml & "<B>&nbsp;Cor&nbsp;Consulting</B>"
 Whtml = Whtml & "</FONT>"
 Whtml = Whtml & "</td></tr></table>"

 CorLogo =  Whtml 

 end function

 Function CorPanelTop()

 dim Whtml
 dim WhtmlX

 WStyl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; "
 WStyl = WStyl & "SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; "
 WStyl = WStyl & "SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; "
 WStyl = WStyl & "SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"

 Whtml = ""
 WhtmlX = ""

 xsscpnm  = request.servervariables("script_name")
 xsscpnmi = instrrev(xsscpnm,"/")
 xsscpnmy = right(ucase(xsscpnm),len(xsscpnm)-xsscpnmi)

 if  xsscpnmy = "MENUS.ASP" or xsscpnmy = "SUBMENUS.ASP"  or xsscpnmy = "FOOTER.ASP" then
 
     WhtmlX = ""

 else 



'' if  xsscpnmy = "INICIO.ASP" or xsscpnmy = "COREXPRESS.ASP" then

     Whtml0i  = "<DIV id='XCPT000I'  ><DIV id='MDPTOP1' STYLE='z-index:0;display:block;' >"
     Whtml1i = ""

     Whtml1i = Whtml1i & " <DIV id='THM' STYLE='z-index:0;position:absolute;top:0;left:0;width:100%;height:40;' >"
     Whtml1i = Whtml1i & "<br style='line-height:45px;' ></DIV>"

'' else

     Whtml0o  = "<DIV id='XCPT000O'  ><DIV id='MDPTOP1' STYLE='z-index:0;display:block;' >"
     Whtml1o = ""
 
     IF  Session("FULLGUI") = "xS" THEN

         Whtml1o = Whtml1o & "  <v:roundrect arcsize='55%' id='THM'  style='z-index:0;position:absolute;top:0;left:-35;"
         Whtml1o = Whtml1o & "width:105%;height:40px;' >"
         Whtml1o = Whtml1o & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"	
         Whtml1o = Whtml1o & "     <v:fill on='True' method='sigma' color2='"  & Session("CMFMENU") & "' color='"
         Whtml1o = Whtml1o & Session("CMFMENU") & "' type='gradient' angle='30' opacity='60%' focusposition=""95%, 15%""  />"
         Whtml1o = Whtml1o & "     <v:shadow on='False' color='#dddddd' weight='10px' />"
         Whtml1o = Whtml1o & "        <v:extrusion on='True' metal='True'  type='perspective' backdepth='8' "
         Whtml1o = Whtml1o & " rotationcenter='0,0,0'  rotationangle = '30, 0'/>"
         Whtml1o = Whtml1o & "  </v:roundrect>"

     ELSE
         Whtml1o = Whtml1o & " <DIV id='THM' STYLE='z-index:0;position:absolute;top:0;left:0;width:100%;height:40;"
         Whtml1o = Whtml1o & "background-color:" & Session("CMFMENU") & ";"
         Whtml1o = Whtml1o & XspEx("#f0f0f0",Session("CMFMENU"),0) & ""
         Whtml1o = Whtml1o & "' >"
         Whtml1o = Whtml1o & "<br style='line-height:45px;' ></DIV>"
     END IF

'' end if
'''''''''
 Whtml = "" 

 Whtml = Whtml & "<TABLE id='MTH'  width=98% marginwidth=0  marginheight=0  "
 Whtml = Whtml & " cellspacing=2  cellpadding=2 "
 Whtml = Whtml & " style=" & Chr(34) & "z-index:90;position:relative;top:2;left:0;"

 Whtml = Whtml & Chr(34)
 Whtml = Whtml & " >"
 Whtml = Whtml & "<TR>"

 xscliente = "<font id='TXHIS' style='display:block;' face=" & Session("CMFMEFTFC") & " size=2 color=" & Session("CMFMEFT") & " ><b>&nbsp;" & Session("Cliente") & "</b></font>"

 Whtml = Whtml & "<TD  id='TXHI' CLASS=vt valign=middle width=4% align=center style='display:block;' >"

 Whtml3 =  Whtml 

'' if  xsscpnmy = "INICIO.ASP"  or xsscpnmy = "COREXPRESS.ASP" then
   ''''  Whtml  = Whtml & "&nbsp;"
     Whtml3i =  Whtml3 & "&nbsp;"

'' ELSE
  ''''   Whtml = Whtml & "<IMG id='TXHIM' border=0 src='" & Session("CliLogo") & "' style='display:block;' >"
     Whtml3o =  Whtml3 & "<IMG id='TXHIM' border=0 src='" & Session("CliLogo") & "' style='display:block;' >"

'' END IF 

 Whtml = ""

 Whtml = Whtml & "</TD>"

 Whtml = Whtml & "<TD  id='TXHI2' CLASS=vt valign=middle width=96% >"

 Whtml4 =  Whtml 

'' if  xsscpnmy = "INICIO.ASP" or xsscpnmy = "COREXPRESS.ASP" then
     '''''Whtml = Whtml & "&nbsp;"
     Whtml4i =  Whtml4 & "&nbsp;"
'' ELSE
     '''''Whtml = Whtml & xscliente
     Whtml4o =  Whtml4 & xscliente
'' END IF 


 Whtml = "" 

 Whtml = Whtml & "</TD>"


 xsmenumtr = xsmenumtr & "<TABLE  id=XMH  marginwidth=0  marginheight=0 "
 xsmenumtr = xsmenumtr & " cellspacing=0  cellpadding=0  "
 xsmenumtr = xsmenumtr & " style=" & Chr(34)
 xsmenumtr = xsmenumtr & "visibility:hidden;"
 xsmenumtr = xsmenumtr & Chr(34)
 xsmenumtr = xsmenumtr & " >"
 xsmenumtr = xsmenumtr & "<TR >"
 xsmenumtr = xsmenumtr & "<TD id=XMHZ  CLASS=ut style='font-size:6pt;' >&nbsp;</TD>"
 xsmenumtr = xsmenumtr & "</TR>"
 xsmenumtr = xsmenumtr & "</TABLE>"

'''''''' sublogo '''''''''''''''''''''

 if session("SMPLMENU") = "S"  then

    if session("SUBLOGO") <> "" then

       xsmenumtr = xsmenumtr & "<IMG id='TXHIMAD' border=0 src='../coress20/empresa/" & session("SUBLOGO") & "' style='display:none;' >"
    else
       xsmenumtr = xsmenumtr & "<FONT FACE='TAHOMA' id='TXHIMAD' style='display:none;' > </FONT>"

    end if

 end if 



'''''''''''''''''''''''''''''''''''''


 Whtml = Whtml & "<TD id='TXMH' CLASS=vt valign=middle style='visibility:visible;' width=3 >" & xsmenumtr & "</TD>"
 Whtml = Whtml & "</TR>"

 Whtml = Whtml & "<TR>"

 xsFuncion = ""

 Whtml = Whtml & "<TD id='TXHI3' colspan=3 CLASS=vt width=100% " & xsFuncion & " valign=middle   ><small>&nbsp;</small></TD>"
 Whtml = Whtml & "</TR>"

 Whtml = Whtml & "<TR id='TXHR4' style='display:none;' >"

 Whtml = Whtml & "<TD id='TXHI4' colspan=3 CLASS=TTT width=100% " & xsFuncion & " valign=top   >"
 Whtml = Whtml & "</TD>"


 Whtml = Whtml & "</TR>"

 Whtml = Whtml & "</TABLE>"

 Whtml5 =  Whtml 

'''''''''''''''''''''''''''''''''''''

 ''''''Whtml = Whtml &  Whtml1

 Whtml5i =  Whtml5 & Whtml1i
 Whtml5o =  Whtml5 & Whtml1o


'''''''''''''''''''''''''''''''''''''

 Whtml =  "" 

      xsAccion    = "../coress20/inicio.asp?Opcion=I&ddat="
      xsInput     = "<input type=text style='display:none;' value='ok' >"
      xFormulario = "<form action='" & xsaccion & "' method='post' id='logoff' name='logoff' style='display:none;'  >" & xsinput & "</form>"

 Whtml = Whtml & xFormulario

 Whtml = Whtml & "</DIV>"

 Whtml = Whtml & "</DIV>"

 Whtml6 =  Whtml 

''''''''''

'' if  xsscpnmy = "INICIO.ASP" or xsscpnmy = "COREXPRESS.ASP"  then

     Whtml =  ""

     Whtml = Whtml & "<script id='XCPT002I' language='javascript' for='TABLE' event='onBLUR' >" & chr(13)

     Whtml7 =  Whtml 

'' else

     Whtml =  ""
     Whtml = Whtml & "<DIV id='XCPT001O' >" 

''''''''''
      sImagenx   = "<B><FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >r</FONT></B>"
      sImagenx1  = "<B><FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >2</FONT></B>"
      sImagenx2a = "<B><FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >5</FONT></B>"
      sImagenx2b = "<B><FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >6</FONT></B>"
      sImagenx3  = "<B><FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >&#0204;</FONT></B>"
      sImagenx4  = "<B><FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >2</FONT></B>"
      xsimagena  = "<b><font FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >s</font></b>" 
      xsimagenb  = "<b><font FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >&#0128;</font></b>" 
      xsimagenc  = "<b><font FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >r</font></b>" 

'''''''
 Whtml = Whtml & "<DIV id='x1' STYLE='width:400px;display:none;z-index:99;margin-top:1;margin-left:1;position:absolute;left:5px;top:50px;border-style:outset;border-color:#ffffff;border-width:1;"

 IF  Session("FULLGUI") = "N" THEN
     Whtml = Whtml & "background-color:#aaaaaa;"
 else     
     Whtml = Whtml & XspEx("#aaaaaa","#dddddd",1) & ";"
 end if

 Whtml = Whtml & "'>"
 Whtml = Whtml & "<table id='HX1'   class='HX1CL' width=100% cellspacing=1 cellpadding=0 "
 Whtml = Whtml & " style='display:none;border-style:outset;border-width:1;border-color:#ffffff;"


 IF  Session("FULLGUI") = "N" THEN
     Whtml = Whtml & "background-color:#EEEEEE;"
 else     
     Whtml = Whtml & XspEx(Session("APLBGD"),"#AAAAAA",0) & ";"
 end if

 Whtml = Whtml & "' >"
 Whtml = Whtml & "<tr height=20 >"
 Whtml = Whtml & "<td  "
 Whtml = Whtml & "class=vT width=2 STYLE='" & XspSShw("#AAAAAA","125") & ";' >" & "<IMG valign=middle border=0 src='../coress20/imagenes/corlogo.gif' style='filter: chroma(color=#ffffff);' >"
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "<td id='HX1TP' onclick='submenu();' "
 Whtml = Whtml & "class=vT STYLE='" & XspDShw("1","1","#f9f9f9") & ";' ><font color=" & session("CMFMNBOT") & " >" & "&nbsp;" & Session("CorTitApli") & "/" & Session("cormenu") & "</font>"
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "<td id='SUB' CLASS=mbt onclick='submenu();' title='Menus' "
 Whtml = Whtml & " width=2 STYLE='display:block;' >" & sImagenx4
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "<td id='SOL' CLASS=mbt onclick='solmenu();' title='Solapas' "
 Whtml = Whtml & " width=2 STYLE='display:block;' >" & sImagenx3
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "<td id='MIN' CLASS=mbt onclick='minimenu();' title='Minimizar' "
 Whtml = Whtml & " width=2 STYLE='display:block;' >" & sImagenx2a
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "<td id='MAX' CLASS=mbt onclick='maximenu();' title='Maximizar' "
 Whtml = Whtml & " width=2 STYLE='display:none;' >" & sImagenx2b
 Whtml = Whtml & "</td>"
'''''''
 Whtml = Whtml & "<TD id=XMH1 CLASS=mbt width=2 onclick='mostrarHLP();' STYLE='display:block;;' >" & xsimagena & "</TD>"
 Whtml = Whtml & "<TD id=XMH2 CLASS=mbt width=2 onclick='mostrarUSR();' STYLE='display:block;' >" & xsimagenb & "</TD>"
 Whtml = Whtml & "<TD id=XMH3 CLASS=mbt width=2 onclick='document.logoff.submit();return false;' STYLE='display:block;' >" & xsimagenc & "</TD>"
''''''''
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"



 Whtml = Whtml & "<div id='MM0' style='display:block;overflow:auto;" & WStyl & "' >"
 Whtml = Whtml & "</div>"

 Whtml = Whtml & "<div id='MM1' style='display:block;" & WStyl
 Whtml = Whtml & "' >"
 Whtml = Whtml & "</div>"


 Whtml = Whtml & "<div id='MM2' style='display:block;"
 Whtml = Whtml & "' >"
 Whtml = Whtml & "</div>"
 Whtml = Whtml & "<div id='MM3' style='display:block;' >&nbsp;"
 Whtml = Whtml & "</div>"
 Whtml = Whtml & "</DIV>"
''''''''''




 Whtml = Whtml & "<DIV id='PDUSR' class='DINDM'  STYLE='display:none;z-index:99;margin-top:1;margin-left:1;position:absolute;left:60px;top:100px;"
 Whtml = Whtml & "'>"
 Whtml = Whtml & "<table id=""MTUSR1""  class='TWIND' width=400 cellspacing=0 cellpadding=2 >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<td class=mbt width=99%  >&nbsp;SESION&nbsp;&nbsp;" & Session.sessionid & "</td>"
 Whtml = Whtml & "<td CLASS=mbt width=1% onclick=""document.all.PDUSR.style.display='none';""  >"
 Whtml = Whtml & "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=webdings size=1  >r</font>" 
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"
 Whtml = Whtml & "<table id=""MTUSR1""  class='aaa' width=400 cellspacing=4 cellpadding=2 "
 Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<td class=tt width=25% >&nbsp;Usuario</td>"
 Whtml = Whtml & "<td class=vti width=75%  >"
 Whtml = Whtml & "<font color=#999999 >" & Session("SegUsuario") & "</font>"
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<td class=tt   >&nbsp;Empresa</td>"
 Whtml = Whtml & "<td class=vti    >"
 Whtml = Whtml & "<font color=#999999 >" & session("cliente") & "</font>"
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "</tr>"

 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<td class=tt  colspan=2 >&nbsp;Grupos</td>"
 Whtml = Whtml & "</tr>"

 wktgrupos = split(session("SegGrpdsc"),chr(9))

 for wkti=0 to ubound(wktgrupos)


     Whtml = Whtml & "<tr>"
     Whtml = Whtml & "<td class=vti  colspan=2  >"
     Whtml = Whtml & "<font color=#999999 >-&nbsp;" & wktgrupos(wkti) & "</font>"
     Whtml = Whtml & "</td>"
     Whtml = Whtml & "</tr>"

 next

 Whtml = Whtml & "</table>"
 Whtml = Whtml & "</DIV>"
''''''''''
 Whtml = Whtml & "<DIV id='PDHLP' class='DINDM'  STYLE='display:none;z-index:99;margin-top:1;margin-left:1;position:absolute;left:60px;top:100px;"
 Whtml = Whtml & "'>"
 Whtml = Whtml & "<table id=""MTHLP1""  class='TWIND' width=400 cellspacing=0 cellpadding=2 >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<td class=mbt width=99% >&nbsp;AYUDA&nbsp;ON-LINE</td>"
 Whtml = Whtml & "<td CLASS=mbt width=1% onclick=""document.all.PDHLP.style.display='none';""  >"
 Whtml = Whtml & "<font onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=webdings size=1  >r</font>" 
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"
 Whtml = Whtml & "<table id=""MTHLP2""  class=aaa width=400 cellspacing=4 cellpadding=2 "
 Whtml = Whtml & " style=""border-style:solid;border-width:3;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<td class=tt width=25% >&nbsp;</td>"
 Whtml = Whtml & "<td class=vti width=70%  >"
 Whtml = Whtml & "<font color=#999999 >en breve ......</font>"
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "<td class=btbs width=5% >"
 Whtml = Whtml & "<font class=btic >4</font>"
 Whtml = Whtml & "</td>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"
 Whtml = Whtml & "</DIV>"

 Whtml = Whtml & "</DIV>"

''''''''''

 Whtml = Whtml & "<script id='XCPT001' language=javascript>" & chr(13)

  if session("SMPLMENU") = "S" then
     Whtml = Whtml & " var smpmenu = 'S'; " & chr(13)
  ELSE
     Whtml = Whtml & " var smpmenu = 'N'; " & chr(13)
  end if


 Whtml = Whtml & " function verfrm() { "
 Whtml = Whtml & "    mostrarmenu(); " & chr(13)
 Whtml = Whtml & " } "

 Whtml = Whtml & "function mostrarmenu(){" & chr(13)

  if session("SMPLMENU") = "S" then

     Whtml = Whtml & " if (PCorType == 'D' || PCorType == 'K') {" & chr(13)
     Whtml = Whtml & "   document.getElementById('x1').style.display='block';" & chr(13)
 '    Whtml = Whtml & "   document.getElementById('TXHIMAD').style.display='none';" & chr(13)
     Whtml = Whtml & "                                     } else {          " & chr(13)
     Whtml = Whtml & "   document.getElementById('TXHIMAD').style.display='block';" & chr(13)
     Whtml = Whtml & "                                                              }" & chr(13)

  ELSE
     Whtml = Whtml & "   document.getElementById('x1').style.display='block';  " & chr(13)
  end if

 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function hdemenu(){" & chr(13)
 Whtml = Whtml & "   document.getElementById('x1').style.display='none';" & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function shwhdemenu(){" & chr(13)
 Whtml = Whtml & "   if (document.getElementById('MDPTOP1').style.display=='block') {" & chr(13)
 Whtml = Whtml & "       window.top.CorMMnu='H';                     " & chr(13)
 Whtml = Whtml & "       document.getElementById('MDPTOP1').style.display='none';  " & chr(13)

 if  SESSION("FULLGUI") = "S" THEN
     Whtml = Whtml & "       document.getElementById('tbhMAX').style.display='block';  " & chr(13)
 ELSE
     Whtml = Whtml & "       document.getElementById('tbhMAX').style.display='block';  " & chr(13)
 END IF 

 Whtml = Whtml & "        hdemenu();                                }" & chr(13)
 Whtml = Whtml & "   else {                                          " & chr(13)
 Whtml = Whtml & "       window.top.CorMMnu='S';                     " & chr(13)
 Whtml = Whtml & "       document.getElementById('MDPTOP1').style.display='block'; " & chr(13)

 if  SESSION("FULLGUI") = "S" THEN
     Whtml = Whtml & "       document.getElementById('tbhMAX').style.display='none';  " & chr(13)
 ELSE
     Whtml = Whtml & "       document.getElementById('tbhMAX').style.display='none';  " & chr(13)
 END IF 

 Whtml = Whtml & "       mostrarmenu();                             }" & chr(13)
 Whtml = Whtml & "   var Fbtnas = 0; "
 Whtml = Whtml & "   MoveBtnas(); " & chr(13)
 if  SESSION("FULLGUI") = "S" THEN
     Whtml = Whtml & " " & chr(13)
 ELSE
     Whtml = Whtml & " ubimenu(); " & chr(13)
 END IF 
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function mostrarUSR(){" & chr(13)
 Whtml = Whtml & "   document.getElementById('PDUSR').style.display='block';" & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function mostrarHLP(){" & chr(13)
 Whtml = Whtml & "   document.getElementById('PDHLP').style.display='block';" & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function solmenu(){" & chr(13)
 Whtml = Whtml & "   document.getElementById('MIN').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MAX').style.display='none';" & chr(13)
 Whtml = Whtml & "   document.getElementById('SOL').style.display='none';" & chr(13)
 Whtml = Whtml & "   document.getElementById('SUB').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM0').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM1').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM2').style.display='none';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM3').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('x1').style.borderStyle='outset';" & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function submenu(){" & chr(13)
 Whtml = Whtml & "   document.getElementById('MIN').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MAX').style.display='none';" & chr(13)
 Whtml = Whtml & "   document.getElementById('SOL').style.display='block';" & chr(13)
 Whtml = Whtml & "   if (PCorType == 'D' || PCorType == 'K' ) {       " & chr(13)
 Whtml = Whtml & "       document.getElementById('SOL').style.display='none';" & chr(13)
 Whtml = Whtml & "                                 }       " & chr(13)
 Whtml = Whtml & "   document.getElementById('SUB').style.display='none';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM0').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM2').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM3').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('x1').style.borderStyle='outset';" & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function despmenu(){" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM0').style.display='block';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM2').style.display='none';" & chr(13)
 Whtml = Whtml & "   document.getElementById('MM3').style.display='none';" & chr(13)
 Whtml = Whtml & "   if (PCorType == 'D' || PCorType == 'K') {       " & chr(13)
 Whtml = Whtml & "       document.getElementById('x1').style.width='400px';                            " & chr(13)
 Whtml = Whtml & "       document.getElementById('x1').style.left=document.getElementById('TXMH').offsetLeft-400;    " & chr(13)
 Whtml = Whtml & "                                 }                                     " & chr(13)
 Whtml = Whtml & "   else {                                                              " & chr(13)
 Whtml = Whtml & "       document.getElementById('x1').style.width='450px';                            " & chr(13)
 Whtml = Whtml & "       document.getElementById('x1').style.left=document.getElementById('TXMH').offsetLeft-430;    " & chr(13)
 Whtml = Whtml & "                                 }                                     " & chr(13)
 Whtml = Whtml & "   document.getElementById('x1').style.top=document.getElementById('TXMH').offsetTop +  document.getElementById('MTH').offsetTop - 4;" & chr(13)
 Whtml = Whtml & "   document.getElementById('HX1').cellSpacing=4;" & chr(13)
 Whtml = Whtml & "   document.getElementById('HX1').style.borderStyle='outset';" & chr(13)
 Whtml = Whtml & "   document.getElementById('x1').style.borderStyle='none';" & chr(13)
 Whtml = Whtml & "   if (document.getElementById('MDPTOP1').style.display=='none') {" & chr(13)
 Whtml = Whtml & "       hdemenu();                                  }" & chr(13)
 Whtml = Whtml & "                                                    " & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function minimenu(){" & chr(13)
 Whtml = Whtml & "   despmenu();" & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function maximenu(){" & chr(13)
 Whtml = Whtml & "  mostrarmenu();                         " & chr(13)
 Whtml = Whtml & "  document.getElementById('SOL').style.display='block';" & chr(13)
 Whtml = Whtml & "  document.getElementById('SUB').style.display='none';" & chr(13)
 Whtml = Whtml & "  if (PCorType == 'D' || PCorType == 'K') {       " & chr(13)
 Whtml = Whtml & "  document.getElementById('SOL').style.display='none';" & chr(13)
 Whtml = Whtml & "                                 }       " & chr(13)
 Whtml = Whtml & "  document.getElementById('MIN').style.display='block';" & chr(13)
 Whtml = Whtml & "  document.getElementById('MAX').style.display='none';" & chr(13)
 Whtml = Whtml & "  document.getElementById('MM0').style.display='block';" & chr(13)
 Whtml = Whtml & "  document.getElementById('MM1').style.display='none';" & chr(13)
 Whtml = Whtml & "  document.getElementById('MM2').style.display='block';" & chr(13)
 Whtml = Whtml & "  document.getElementById('MM3').style.display='block';" & chr(13)
 Whtml = Whtml & "   if (window.top.CorType == 'D' || window.top.CorType == 'K') {       " & chr(13)
 Whtml = Whtml & "       document.getElementById('x1').style.width='400px';                            " & chr(13)
 Whtml = Whtml & "                                 }                                     " & chr(13)
 Whtml = Whtml & "   else {                                                              " & chr(13)
 Whtml = Whtml & "       document.getElementById('x1').style.width='450px';                            " & chr(13)
 Whtml = Whtml & "                                 }                                     " & chr(13)
 Whtml = Whtml & "  document.getElementById('x1').style.left='5px';" & chr(13)
 Whtml = Whtml & "  document.getElementById('x1').style.top=document.getElementById('MTH').clientHeight;" & chr(13)
 Whtml = Whtml & "  document.getElementById('HX1').cellSpacing=4;" & chr(13)
 Whtml = Whtml & "  document.getElementById('HX1').style.borderStyle='outset';" & chr(13)
 Whtml = Whtml & "  document.getElementById('x1').style.borderStyle='outset';" & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function ubimenu(){" & chr(13)
 Whtml = Whtml & "   document.getElementById('TBHST').style.left=document.getElementById('TBH1').offsetLeft + 8;" & chr(13)
 Whtml = Whtml & "   document.getElementById('TBHST').style.top=document.getElementById('TBH').offsetTop + 5;" & chr(13)
 Whtml = Whtml & "}" & chr(13)


 Whtml = Whtml & "function locmenu(){" & chr(13)
 Whtml = Whtml & "   var frmxx = window.top.verInfrm2();"
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function locmenu2(){" & chr(13)
 Whtml = Whtml & "   alert(window.top.CorType);" & chr(13)
 Whtml = Whtml & "   alert('vis ' + document.all.XMH.style.visibility);" & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "function MoveBtnas(){" & chr(13)
 Whtml = Whtml & "  if (document.getElementById('btnas') ) { " & chr(13)
 if CorBarra = "U" then
    Whtml = Whtml & "     document.getElementById('btnas').style.position='absolute';" & chr(13)
    Whtml = Whtml & "     document.getElementById('btnas').style.top=document.getElementById('TBH').offsetTop - 2;" & chr(13)
    Whtml = Whtml & "     document.getElementById('btnas').style.left=document.getElementById('TBH2').offsetLeft + (document.getElementById('TBH2').clientWidth-document.getElementById('btnas').clientWidth);" & chr(13)
 end if
 Whtml = Whtml & "     document.getElementById('btnas').style.visibility='visible';"
 Whtml = Whtml & " } else { " & chr(13)
 Whtml = Whtml & "           if (Fbtnas > 10) {} else { Fbtnas++; " & chr(13)
 Whtml = Whtml & " setTimeout (""MoveBtnas()"", 350); } " & chr(13)
 Whtml = Whtml & " } " & chr(13)
 Whtml = Whtml & "}" & chr(13)

 Whtml = Whtml & "</script>" & chr(13)





 Whtml = Whtml & "<script id='XCPT007O' language='javascript' for='TABLE' event='onBLUR' >" & chr(13)





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Acciones realizadas al cargar el Panel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     Whtml = Whtml & " var PCorType;" & chr(13)
   
     Whtml = Whtml & " if (!window.top.CorType) {" & chr(13)
     Whtml = Whtml & "     PCorType = 'SA';" & chr(13)
     Whtml = Whtml & "                                                              }" & chr(13)
     Whtml = Whtml & " else" & chr(13)
     Whtml = Whtml & "     {" & chr(13)
     Whtml = Whtml & "     PCorType = window.top.CorType;" & chr(13)
     Whtml = Whtml & "                                                              }" & chr(13)

  if session("SMPLMENU") = "S" then

     Whtml = Whtml & " if (PCorType == 'D' || PCorType == 'K') {" & chr(13)
 '    Whtml = Whtml & "   document.getElementById('TXHIMAD').style.display='none';" & chr(13)
     Whtml = Whtml & "                                     } else {          " & chr(13)
     Whtml = Whtml & "   document.getElementById('TXHIMAD').style.display='block';" & chr(13)
     Whtml = Whtml & "                                                              }" & chr(13)
  end if

     Whtml = Whtml & " if (PCorType == 'SA') {" & chr(13)
     Whtml = Whtml & "                                                              }" & chr(13)
     Whtml = Whtml & " else" & chr(13)
     Whtml = Whtml & "     {" & chr(13)



     Whtml = Whtml & " var frmz = window.top.xclrstr();"

     Whtml = Whtml & " var frm0 = window.top.verfrm0();"
     Whtml = Whtml & " document.getElementById('MM0').innerHTML='<center>' + frm0 + '</center>'; " & chr(13)

     Whtml = Whtml & " var frm1 = window.top.verfrm1();"
     Whtml = Whtml & " document.getElementById('MM1').innerHTML='<center>' + frm1 + '</center>'; " & chr(13)

     Whtml = Whtml & " var frm2 = window.top.verfrm2();"
     Whtml = Whtml & " document.getElementById('MM2').innerHTML= '<center>' + frm2 + '</center>'; " & chr(13)

     Whtml = Whtml & " document.getElementById('x1').style.top=document.getElementById('MTH').clientHeight;" & chr(13)

 IF  Session("FULLGUI") = "S" THEN
     Whtml = Whtml & " document.getElementById('THM').style.height=document.getElementById('MTH').clientHeight - 4;" & chr(13)
 else
     Whtml = Whtml & " document.getElementById('THM').style.height=document.getElementById('MTH').clientHeight;" & chr(13)
 end if

     Whtml = Whtml & " if (PCorType == 'D' || PCorType == 'K') {" & chr(13)
     Whtml = Whtml & "     document.getElementById('XMH2').style.display='none';" & chr(13)
     Whtml = Whtml & "     document.getElementById('XMH3').style.display='none';" & chr(13)
     Whtml = Whtml & "     document.getElementById('SOL').style.display='none';" & chr(13)

     Whtml = Whtml & "     document.getElementById('TXHI4').innerHTML= '' + window.top.CorEsq + '';" & chr(13)

     Whtml = Whtml & "     document.getElementById('TXHR4').style.display='block';" & chr(13)
     Whtml = Whtml & "                                                              }" & chr(13)
     Whtml = Whtml & " else" & chr(13)
     Whtml = Whtml & "     {" & chr(13)
     Whtml = Whtml & "     document.getElementById('TXHI').style.display='block';" & chr(13)
     Whtml = Whtml & "                                                              }" & chr(13)


     Whtml = Whtml & " despmenu();    " & chr(13)

     Whtml = Whtml & " mostrarmenu(); " & chr(13)

     Whtml = Whtml & "                                                              }" & chr(13)


     Whtml = Whtml & " var Fbtnas = 0; "
     Whtml = Whtml & " MoveBtnas(); " & chr(13)

     Whtml8 =  Whtml


'' end if

 Whtml = ""

 Whtml = Whtml & " var xaction = document.logoff.action; "
 Whtml = Whtml & " document.logoff.action = xaction + window.top.CorUsr + '___' "
 Whtml = Whtml & " + window.top.CorNetUsr  + '___' + window.top.CorDomain + '___' + window.top.CorLogid;  " & chr(13)

 Whtml = Whtml & "</SCRIPT>" & chr(13) & chr(13)

 Whtml9 =  Whtml

''''''''''''' arma rutinas ''''''''''''''''

     Whtmlx = ""

     Whtmlx =  Whtmlx &  Whtml0i &  Whtml3i &  Whtml4i &  Whtml5i &  Whtml6 &  Whtml7 &  Whtml9 

     Whtmlx =  Whtmlx &  Whtml0o &  Whtml3o &  Whtml4o &  Whtml5o &  Whtml6 &  Whtml8 &  Whtml9 

'''''''''''''''''''''''''''''''''''''''''''

'' if  xsscpnmy  = "INICIO.ASP" or xsscpnmy = "COREXPRESS.ASP"  then
     Whtmlx = Whtmlx & ""
'' else
     Whtmlx = Whtmlx & "<SCRIPT id='XCPT003O' language='javascript' for='table' event='onresize' >" & chr(13) & chr(13)
     Whtmlx = Whtmlx & " if (PCorType == 'SA') {     " & chr(13)
     Whtmlx = Whtmlx & "                       }     " & chr(13)
     Whtmlx = Whtmlx & " else                        " & chr(13)
     Whtmlx = Whtmlx & "                       {     " & chr(13)
     Whtmlx = Whtmlx & "      minimenu();            " & chr(13)
     Whtmlx = Whtmlx & "                       }     " & chr(13)
     Whtmlx = Whtmlx & "  MoveBtnas();               " & chr(13)
     Whtmlx = Whtmlx & "</SCRIPT>" & chr(13) & chr(13)
'' end if


'''''''''
 end if


     CorPanelTop = Whtmlx

 end function

 Function CorPanelTop2()

     Whtmlx = ""

     xsscpnm  = request.servervariables("script_name")
     xsscpnmi = instrrev(xsscpnm,"/")
     xsscpnmy = right(ucase(xsscpnm),len(xsscpnm)-xsscpnmi)

     if  xsscpnmy = "MENUS.ASP" or xsscpnmy = "SUBMENUS.ASP"  or xsscpnmy = "FOOTER.ASP" then
 
         WhtmlX = ""

     elseif  xsscpnmy = "INICIO.ASP" or xsscpnmy = "COREXPRESS.ASP" or xsscpnmy = "BARRAENCIERRA2.ASP" then

         Whtmlx = Whtmlx & "<SCRIPT id='XCPT999I' language='javascript' >"
         Whtmlx = Whtmlx & " var xcptvar = window.top.verfuncion(); " 

         Whtmlx = Whtmlx & " document.body.innerHTML += xcptvar.getElementById('XCPT000I').innerHTML; "

         Whtmlx = Whtmlx & " 	var j2   = document.createElement('script'); "
         Whtmlx = Whtmlx & "	j2.type  = 'text/javascript'; "
         Whtmlx = Whtmlx & "	j2.text  = xcptvar.getElementById('XCPT002I').innerHTML; "
         Whtmlx = Whtmlx & "	document.body.appendChild(j2); "




         Whtmlx = Whtmlx & "</SCRIPT>" 

         WhtmlX = ""

     ELSE



         Whtmlx = Whtmlx & "<SCRIPT id='XCPT999O' language='javascript' >"
       '''  Whtmlx = Whtmlx & " var xcptvar = window.top.verfuncion(); " 
       '''  Whtmlx = Whtmlx & "	alert(window.top.CorMtopf.name + ' ' + 'corlogf'); "
         Whtmlx = Whtmlx & "  if (window.top.CorMtopm) { "
         Whtmlx = Whtmlx & "      var wcortopf = window.top.CorMtopf; } "
         Whtmlx = Whtmlx & "  else { "
         Whtmlx = Whtmlx & "      var wcortopf = window.opener.top.CorMtopf; } "


       '  Whtmlx = Whtmlx & "	alert(wcortopf.name + ' ' + 'corlogf'); "

         Whtmlx = Whtmlx & " var xcptvar = wcortopf.verfuncion(); " 
         Whtmlx = Whtmlx & " document.body.innerHTML += xcptvar.getElementById('XCPT000O').innerHTML; "
         Whtmlx = Whtmlx & " document.body.innerHTML += '<div>' + xcptvar.getElementById('XCPT001O').innerHTML + '</div>'; "


         Whtmlx = Whtmlx & " 	var j5   = document.createElement('script'); "
         Whtmlx = Whtmlx & "	j5.type  = 'text/javascript'; "
         Whtmlx = Whtmlx & "	j5.text  = xcptvar.getElementById('XCPT001').innerHTML; "
         Whtmlx = Whtmlx & "	document.body.appendChild(j5); "

         Whtmlx = Whtmlx & " 	var j6    = document.createElement('script'); "
         Whtmlx = Whtmlx & "	j6.type   = 'text/javascript'; "
         Whtmlx = Whtmlx & "	j6.text   = xcptvar.getElementById('XCPT007O').innerHTML; "
         Whtmlx = Whtmlx & "	document.body.appendChild(j6); "

         Whtmlx = Whtmlx & " 	var j3   = document.createElement('script'); "
         Whtmlx = Whtmlx & "	j3.type  = 'text/javascript'; "
         Whtmlx = Whtmlx & "	j3.htmlFor   = 'window'; "
         Whtmlx = Whtmlx & "	j3.event = 'onresize'; "
         Whtmlx = Whtmlx & "	j3.text  = xcptvar.getElementById('XCPT003O').innerHTML; "
         Whtmlx = Whtmlx & "	document.body.appendChild(j3); "

         Whtmlx = Whtmlx & "</SCRIPT>" 


     END IF


     CorPanelTop2 = Whtmlx

 end function

 Function CorSetMHeader()

     dim Whtml

     Whtml = ""

     Whtml = Whtml & "<html  xmlns:v='urn:schemas-microsoft-com:vml' >"
     Whtml = Whtml & "<HEAD><meta http-equiv='Content-Type' content='text/html; charset=" & XLCHARSET & "'></HEAD>"
     Whtml = Whtml & "<Title>Cor E-Solution Suite V3</title>"
     Whtml = Whtml & ofv.FormStyler()
     Whtml = Whtml & "<STYLE>v\:* {behavior:url(#default#VML);}</STYLE>"
     Whtml = Whtml & "<body " & ofv.wbody & " alink=#ffffff vlink=#ffffff link=#ffffff    XSFNK  >"

     Whtml = Whtml & CorPanelTop2

     CorSetMHeader =  Whtml 

 end function


 Function CorSetFHeader()

     dim Whtml

     Whtml = ""

     Whtml = Whtml & "<html xmlns:v='urn:schemas-microsoft-com:vml' >" 
     Whtml = Whtml & "<HEAD><META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " content='text/html; charset=" & XLCHARSET & "'></HEAD>"
     Whtml = Whtml & "<Title>Cor E-Solution Suite V3</title>"
     Whtml = Whtml & "<STYLE>v\:* {behavior:url(#default#VML);}</STYLE>" & ofv.FormStyler()
     Whtml = Whtml & "<body " & ofv.wbody & " alink=#ffffff vlink=#ffffff link=#ffffff    XSFNK  >"

     Whtml = Whtml & CorPanelTop2

     CorSetFHeader =  Whtml 

 end function



 Function CorFormHeader()

     dim Whtml

     Whtml = ""

     Whtml = Whtml & "<TABLE  name=headm id=headm width=100%  marginwidth=0  marginheight=0 "
     Whtml = Whtml & " cellspacing=0  cellpadding=0 style='BORDER-style:outset;border-width:2;' >"
     Whtml = Whtml & "<TR  align=top>"
     Whtml = Whtml & "<TD  CLASS=tt  >"

     Whtml = Whtml & "<table border=0 cellspacing=0 cellpadding=2 bgcolor=#eeeeee  style='border-style:ridge;border-width:3;' >"
     Whtml = Whtml & "<tr>"
     Whtml = Whtml & "<td valign=middle align=center style='border-style:ridge,border-width:2;" & XspSShw("#777777","135") & ";' >"
     Whtml = Whtml & "<IMG valign=middle border=0 src='imagenes/corlogo.gif' style='filter: chroma(color=#ffffff);' >"
     Whtml = Whtml & "<FONT color=#ffffff >"
     Whtml = Whtml & "<B>&nbsp;&nbsp;" & Session("FormN") & "&nbsp;&nbsp;</B>"
     Whtml = Whtml & "</FONT>"
     Whtml = Whtml & "</td></tr></table>"

     Whtml = Whtml & "</TD>"
     Whtml = Whtml & "<TD  CLASS=tt align=right>Usuario: " & Session("SegUsuario") & " </TD>"
     Whtml = Whtml & "</TR>"
     Whtml = Whtml & "</TABLE>"
     Whtml = Whtml & "<hr color=" & Session("CMFMENU") & " size=1>"


     CorFormHeader =  Whtml 

 end function

 Function CorFormHeader2(NForm,Nopt)

     dim Whtml

     sImagenx   = "<B><FONT FACE=WebDINGS COLOR=#005500 SIZE=3>1</FONT></B>"
     sImagenx1  = "<B><FONT FACE=WebDINGS COLOR=#993333 SIZE=3>&#0128;</FONT></B>"

     Whtml = ""
     Whtml1 = ""
     Whtml2 = ""
     Whtml3 = ""
     Whtml4 = ""

     whtcls = "TTT"
     whtclr = session("CMFMNBOT")

     Whtml = Whtml & "<TABLE class='tbhnn' id='TBH' width=100%  marginwidth=0  marginheight=0 "
     Whtml = Whtml & " cellspacing=0  cellpadding=0 "
     Whtml = Whtml & " style='background-color:#f0f0f0;border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#dddddd;"
     Whtml = Whtml & "border-right-style:solid;border-right-width:2;border-right-color:#dddddd;' "

 if  Session("FULLGUI") <> "S" then
     Whtml = Whtml & " style='position:relative;left:-4px;' "
end if

     Whtml = Whtml & "  >"
     Whtml = Whtml & "<TR height=30 >"



 if  Session("FULLGUI") = "xS" then

     Whtml1 = Whtml1 & " <v:line ID='TBH0VML' NAME='TBH0VML' from='0,0' to='1,0' title='Mostrar/esconder encabezado.' onclick='shwhdemenu();' style='cursor:hand;align:left;z-index:1;position:relative;top:-1;width:10px;height:10px;' > "
     Whtml1 = Whtml1 & "   <v:fill on='True' color='" & session("CMFMNBOT") & "' color2='" & session("CMFMNBOT") & "' method='sigma' type='gradient' angle='180' opacity='95%'  /> "
     Whtml1 = Whtml1 & "   <v:stroke on='False' /> "
     Whtml1 = Whtml1 & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
     Whtml1 = Whtml1 & "   <v:path textpathok='True' /> "
     Whtml1 = Whtml1 & "     <v:textpath on='True' string='¿' style=""font-size:'16pt';font-family:'webdings';font-weight:bold;""  /> "
     Whtml1 = Whtml1 & "   <v:extrusion on='True' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
     Whtml1 = Whtml1 & " </v:line>" & "&nbsp;"

     Whtml1 = Whtml1 &  Replace(GenTitulo(NForm,0), "XXXIDXXX", "TBH1VML")

   '  Whtml1 = ""

     Whtml4 = Whtml4 & " <v:line ID='TBHXVML' NAME='TBHXVML' from='0,0' to='1,0' title='Maximizar menu.' onclick='maximenu();' style='cursor:hand;align:left;z-index:1;position:relative;top:-1;width:10px;height:10px;' > "
     Whtml4 = Whtml4 & "   <v:fill on='True' color='" & session("CMFMNBOT") & "' color2='" & session("CMFMNBOT") & "' method='sigma' type='gradient' angle='180' opacity='95%'  /> "
     Whtml4 = Whtml4 & "   <v:stroke on='False' /> "
     Whtml4 = Whtml4 & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
     Whtml4 = Whtml4 & "   <v:path textpathok='True' /> "
     Whtml4 = Whtml4 & "     <v:textpath on='True' string='2' style=""font-size:'16pt';font-family:'webdings';font-weight:bold;""  /> "
     Whtml4 = Whtml4 & "   <v:extrusion on='True' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
     Whtml4 = Whtml4 & " </v:line>"

  ''   Whtml4 = ""

     Whtml3 = Whtml3 & " if (PCorType == 'D' || PCorType == 'K') {" & chr(13)
     Whtml3 = Whtml3 & "     document.all.TBH0VML.textpath.string='¥';" & chr(13)
     Whtml3 = Whtml3 & "     document.all.TBH1VML.textpath.string=window.top.CorSol;" & chr(13)
     Whtml3 = Whtml3 & " }" & chr(13)
     Whtml3 = Whtml3 & " else {" & chr(13)
     Whtml3 = Whtml3 & "     document.all.TBH0VML.textpath.string='¿';" & chr(13)
     Whtml3 = Whtml3 & " }" & chr(13)


 else

     if  Session("FULLGUI") = "M" then
         whtcls = "TTT"
         whtclr = session("CMFMNBOT")
     else
         whtcls = "TTB"
         whtclr = Session("CMFPNTITFT")
     end if 

         whtcls = "TTT"
         whtclr = session("CMFMNBOT")

     Whtml1 = Whtml1 & "&nbsp;"

     Whtml2 = Whtml2 & "<TABLE class='TBHST' id='TBHST' width=45%  marginwidth=0  marginheight=0"
     Whtml2 = Whtml2 & " cellspacing=0  cellpadding=0 style='z-index:2;position:absolute;top:2;left:2;' >"
     Whtml2 = Whtml2 & "<TD CLASS='" & whtcls & "' WIDTH=8%  >"
     Whtml2 = Whtml2 & "<FONT  ID='TBH0DV' NAME='TBH0DV' FACE='WEBDINGS' SIZE=5 COLOR='" & whtclr & "' >&#0191;</FONT>"
     Whtml2 = Whtml2 & "</TD>"
     Whtml2 = Whtml2 & "<TD CLASS='" & whtcls & "' WIDTH=90% >"
     Whtml2 = Whtml2 & "<BIG><DIV ID='TBH1DV' >" & NForm & "</DIV></BIG>"
     Whtml2 = Whtml2 & "</TD>"
     Whtml2 = Whtml2 & "</TR>"
     Whtml2 = Whtml2 & "</TABLE>"

     Whtml4 = Whtml4 & "<TABLE class='TBHST' id='TBHXDV' width=2%  marginwidth=0  marginheight=0"
     Whtml4 = Whtml4 & " cellspacing=0  cellpadding=0  >"
     Whtml4 = Whtml4 & "<TD CLASS='ut'  title='Maximizar menu.' onclick='maximenu();' style='cursor:hand;'  >"
     Whtml4 = Whtml4 & "<FONT  ID='TBHX0DV' NAME='TBHX0DV' FACE='WEBDINGS' SIZE=4 COLOR='" & whtclr & "' >2</FONT>"
     Whtml4 = Whtml4 & "</TD>"
     Whtml4 = Whtml4 & "</TR>"
     Whtml4 = Whtml4 & "</TABLE>"

     Whtml3 = Whtml3 & " if (PCorType == 'D' || PCorType == 'K') {" & chr(13)
     Whtml3 = Whtml3 & "     document.all.TBH0DV.innerHTML='&#0165;';" & chr(13)
     Whtml3 = Whtml3 & "     document.all.TBH1DV.innerHTML=window.top.CorSol;" & chr(13)
     Whtml3 = Whtml3 & " }" & chr(13)
     Whtml3 = Whtml3 & " else {" & chr(13)
     Whtml3 = Whtml3 & "     document.all.TBH0DV.innerHTML='&#0191;';" & chr(13)
     Whtml3 = Whtml3 & " }" & chr(13)
     Whtml3 = Whtml3 & " ubimenu(); " & chr(13)


 end if

     StyL2  = " STYLE='cursor:hand;display:none;"
  ''   StyL2  = StyL2 & "border-style:solid;border-width:1;border-color:#eeeeee;"
     StyL2  = StyL2 & "' "


     Whtml = Whtml & "<td id='tbhMAX'  title='Maximizar' class='" & whtcls & "' width=3% "
     Whtml = Whtml & StyL2 & " align=center >" & Whtml4
     Whtml = Whtml & "</td>"


     Whtml = Whtml & "<TD id='TBH1'  class='tbh1'  >" & Whtml1
     Whtml = Whtml & "</TD>"

     Whtml = Whtml & "<TD id='TBH2' width=50% class='tbh2' align=right  >"
     Whtml = Whtml & "&nbsp;"
     Whtml = Whtml & "</TD>"

     Whtml = Whtml & "</TR>"
     Whtml = Whtml & "</TABLE>"

     Whtml = Whtml & "<br>"
     Whtml = Whtml & Whtml2

     Whtml = Whtml & "<script language=javascript >" & chr(13)
     Whtml = Whtml & Whtml3
     Whtml = Whtml & " if (window.top.CorMMnu=='H')                                 {" & chr(13)
     Whtml = Whtml & "     shwhdemenu();" & chr(13)
     Whtml = Whtml & "                                                              }" & chr(13)
     Whtml = Whtml & "</script>" & chr(13)



     CorFormHeader2 =  Whtml 

 end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' FUNCION CONFECCION DEL EMAIL ( HTML O TEXTO)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 FUNCTION GenMDLMAIL (XGUPARA,XGUORIGEN,XGUMENSAGE)

''message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"


     XGUMENSAGEw = XGUMENSAGE & " "

     XGUMENSAGEw = replace(XGUMENSAGEw,"@@@NORMREQ@@@","<u>" & SESSION("NOMREQ") & "</u>")
     XGUMENSAGEw = replace(XGUMENSAGEw,"@@@FECREQ@@@",now())



     Whtml = ""

     SELECT CASE  SESSION("MAILOPC")

            CASE "HTML"

                  Whtml = Whtml & "<DIV STYLE='background-color:" & Session("CMFMENU") & ";height:40;width:600px;"
                  Whtml = Whtml & "BORDER-color:#E0E0E0;BORDER-STYLE:SOLID;BORDER-WIDTH:1;'  >"
                 ' Whtml = Whtml & "<br>"
                  Whtml = Whtml & "<P VALIGN=MIDDLE >&nbsp;&nbsp;<font size=4 face='" & Session("CMFMEFTFC") & "' color='" & Session("CMFMEFT") & "' ><B>"
                  Whtml = Whtml & Session("CorTitApli") & "</B></FONT></P>"
                  Whtml = Whtml & "</DIV>"
                  Whtml = Whtml & "<BR>"

                  Whtml = Whtml & "<DIV STYLE='BORDER-color:" & Session("CMFMENU") & ";BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
                  Whtml = Whtml & "height:50;width:600px;'  >"
                  Whtml = Whtml & "<CENTER>"

                  Whtml = Whtml & "<DIV STYLE='BORDER-color:#888888;BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
                  Whtml = Whtml & "background-color:#F0F0F0;height:25;width:100%;'  >"
                  Whtml = Whtml & "<P ALIGN=LEFT >"
                  Whtml = Whtml & "<font size=1 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B>"
                  Whtml = Whtml & "&nbsp;&nbsp;DESTINATARIO</B></FONT></P>"
                  Whtml = Whtml & "</DIV>"

                  Whtml = Whtml & "<br>"
                  Whtml = Whtml & "<DIV STYLE='width:80%;' >"
                  Whtml = Whtml & "<P WIDTH=450 ALIGN='JUSTIFY' >"
                  Whtml = Whtml & "<font size=2 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B><U>"
                  Whtml = Whtml & XGUPARA & "</U></B></FONT></P>"
                  Whtml = Whtml & "</DIV>"

                  Whtml = Whtml & "<br>"
                  Whtml = Whtml & "<DIV STYLE='BORDER-color:#888888;BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
                  Whtml = Whtml & "background-color:#F0F0F0;height:25;width:100%;'  >"
                  Whtml = Whtml & "<P ALIGN=LEFT >"
                  Whtml = Whtml & "<font size=1 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B>"
                  Whtml = Whtml & "&nbsp;&nbsp;DESCRIPCION</B></FONT></P>"
                  Whtml = Whtml & "</DIV>"

                  Whtml = Whtml & "<br>"
                  Whtml = Whtml & "<DIV STYLE='width:80%;' >"
                  Whtml = Whtml & "<P ALIGN='JUSTIFY' >"
                  Whtml = Whtml & "<font size=2 color=#888888 face='" & Session("CMFMEFTFC") & "' >"
                  Whtml = Whtml & XGUMENSAGEw
                  Whtml = Whtml & "</FONT></P>"
                  Whtml = Whtml & "</DIV>"
                  Whtml = Whtml & "<BR>"
                  Whtml = Whtml & "</DIV>"

                  Whtml = Whtml & "<br>"
                  Whtml = Whtml & "<DIV ALIGN='RIGHT' STYLE='width:200px;' >"
                  Whtml = Whtml & "<P  >"
                  Whtml = Whtml & "<font size=1 color=#6699cc face='" & Session("CMFMEFTFC") & "' ><B><i>"
                  Whtml = Whtml & "COR&nbsp;E-SOLUTION&nbsp;SUITE&nbsp;V3</i></B></FONT></P>"
                  Whtml = Whtml & "</DIV>"



            CASE ELSE


                  Whtml = Whtml & SESSION("NOMAPLI")  & "<BR>" & CHR(13) & CHR(13)

                  Whtml = Whtml & XGUPARA & ": <BR>" & CHR(13)

                  Whtml = Whtml & "------------------------------------------------------------------------"  & "<BR>" & CHR(13)

                  Whtml = Whtml & XGUMENSAGEw  & "<BR>" & CHR(13)
                  Whtml = Whtml & "------------------------------------------------------------------------"  & "<BR>" & CHR(13)




     END SELECT   


     GenMDLMAIL = Whtml



 END FUNCTION

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' FUNCIONES DE BUSQUEDAS MULTIPLES ( USUARIOS-GRUPOS-SECTORES-EMPRESAS)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


 function GenBuscaUsr(XGUM1,XGUM2,XGUM3,XGUM4,XGUM1F,XGUM2F,XGUM3F,XGUM4F)


 XGUM1FL = REPLACE(XGUM1F,CHR(9),"")
 XGUM2FL = REPLACE(XGUM2F,CHR(9),"")
 XGUM3FL = REPLACE(XGUM3F,CHR(9),"")
 XGUM4FL = REPLACE(XGUM4F,CHR(9),"")
 

 XGUFBURL = request.servervariables("SCRIPT_NAME")

 XGUFBURL = RIGHT(XGUFBURL,LEN(XGUFBURL) - (INSTRREV(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

 XGUMASTER = REQUEST.FORM("XGUMASTER")
 XGUFILTRO = REQUEST.FORM("XGUFILTRO")
 XGUITEMCD = REQUEST.FORM("XGUITEMCD")
 XGUCODIGO = REQUEST.FORM("XGUCODIGO")
 XGUCODDSC = REQUEST.FORM("XGUCODDSC")


 XGUMASTERW = XGUMASTER

 XGUMSBOT1  = "&nbsp;"
 XGUMSBOT2  = "&nbsp;"
 XGUMSBOT3  = "&nbsp;"
 XGUMSBOT4  = "&nbsp;"

 IF SESSION("EMPMUN") = "S" then

 IF XGUM1 = "S" THEN

    IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTERW = "E"
    END IF

    XGUMSBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='EMPRESAS'  CLASS=BT ONCLICK=""clactfrm('E','NO','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>" 
    XGUMSBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI1 & "</FORM>"

 END IF

 END IF

 IF XGUM2 = "S" THEN

    IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTERW = "S"
    END IF

    XGUMSBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTORES'  CLASS=BT ONCLICK=""clactfrm('S','NO','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>" 
    XGUMSBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI2 & "</FORM>"


 END IF

 IF XGUM3 = "S" THEN

    IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTERW = "G"
    END IF

    XGUMSBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPOS'  CLASS=BT ONCLICK=""clactfrm('G','NO','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>" 
    XGUMSBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI3 & "</FORM>"


 END IF

 IF XGUM4 = "S" THEN

    IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTERW = "U"
    END IF

    XGUMSBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='USUARIOS'  CLASS=BT ONCLICK=""clactfrm('U','NO','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>" 
    XGUMSBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI4 & "</FORM>"


 END IF


 IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTER = XGUMASTERW
 END IF


 XGUCONTEMS = ""

 XGUCONTEMS = XGUCONTEMS & "<table id=""MTBSUSR0B""  class='aaa'  cellspacing=1 cellPADDING=0 > "
 XGUCONTEMS = XGUCONTEMS & "<tr>"
 XGUCONTEMS = XGUCONTEMS & "<td class=ut WIDTH=60%  >&nbsp;</td>"

 IF XGUMSBOT1 <> "&nbsp;" THEN
    XGUCONTEMS = XGUCONTEMS & XGUMSBOT1
 END IF

 IF XGUMSBOT2 <> "&nbsp;" THEN
    XGUCONTEMS = XGUCONTEMS & XGUMSBOT2
 END IF

 IF XGUMSBOT3 <> "&nbsp;" THEN
    XGUCONTEMS = XGUCONTEMS & XGUMSBOT3
 END IF

 IF XGUMSBOT4 <> "&nbsp;" THEN
    XGUCONTEMS = XGUCONTEMS & XGUMSBOT4
 END IF

 XGUCONTEMS = XGUCONTEMS & "</tr>"
 XGUCONTEMS = XGUCONTEMS & "</table>"


 XGUFNC1 = ""
 XGUFNC2 = ""


 set xgucn = ofv.conectar(ofv.strconn2)

 XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#F7AA22;cursor:hand;' "
 XGUSTYLN = " STYLE='cursor:hand;' "

 XGUFBUSR = XGUCODIGO

 XGUFBTAB1  = ""
 XGUFBTAB2  = ""
 XGUFBTD1  = ""
 XGUFBTD2  = ""

 IF XGUMASTER = "" OR XGUMASTER = "NO" THEN XGUMASTER = "U"

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '' SELECCION DE BUSQUEDA MASTER - E:EMPRESAS / S:SECTORES / G:GRUPOS / U:USUARIOS              ''
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 SELECT CASE XGUMASTER

      CASE "E" 

 XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
 XGUFBBOT2  = ""
 XGUFBBOT3  = ""
 XGUFBBOT4  = ""

             XGUFILDSC = "&nbsp;EMPRESA"

             xTrColor = "#FFFFFF"

             IF XGUCODIGO <> "NO" THEN
                 XGUFNC1 = " document.all['***" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF 

             strsql = "Select * From clienteplan  order by cliente"
             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECIONAR EMPRESA.' ID='***" & xgurs(0) & "' "

                IF (XGUM1FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF 



                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM1FL & xgurs(0) & "','" & xgurs(1) & "');document.XGUFRM.submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)


 XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
 XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


 XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
 XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
 XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>" 

 XGUCONTENT = ""
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"


'''' Whtml = Whtml & " document.all[spmk].scrollIntoView(true); " & chr(13)




      CASE "S"

 XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
 XGUFBBOT2  = ""
 XGUFBBOT3  = ""
 XGUFBBOT4  = ""

             XGUFILDSC = "&nbsp;SECTOR"

             xTrColor = "#FFFFFF"

             IF XGUCODIGO <> "NO" THEN
                 XGUFNC1 = " document.all['***" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF 



             strsql    = "Select * From Eco  order by descri"
             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECIONAR SECTOR.'  ID='***" & xgurs(0) & "' "

                IF (XGUM2FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF 



                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM2FL & xgurs(0) & "','" & xgurs(1) & "');document.XGUFRM.submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)


 XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
 XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


 XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
 XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
 XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>" 

 XGUCONTENT = ""
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"


      CASE "G"

 XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
 XGUFBBOT2  = ""
 XGUFBBOT3  = ""
 XGUFBBOT4  = ""


             XGUFILDSC = "&nbsp;GRUPO"

             xTrColor = "#FFFFFF"

             IF XGUCODIGO <> "NO" THEN
                 XGUFNC1 = " document.all['***" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF 

             strsql = "Select * From GRUPOS  order by descriPCION "
             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECIONAR GRUPO.'  ID='***" & xgurs(0) & "' "

                IF (XGUM3FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF 



                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM3FL & xgurs(0) & "','" & xgurs(1) & "');document.XGUFRM.submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)


 XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
 XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


 XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
 XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
 XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>" 

 XGUCONTENT = ""
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"






      CASE "U"

 


 XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>" 
 XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>" 
 XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>" 
 XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>" 

 XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
 XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
 XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
 XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"



        XGUFILDSC = "&nbsp;USUARIO"

 SELECT CASE XGUFILTRO

        CASE "G"

             XGUFILDSC1 = ":&nbsp;GRUPO"

             xTrColor = "#FFFFFF"

             IF XGUITEMCD <> "NO" THEN
                 XGUFNC1 = " document.all['***" & XGUITEMCD & "'].scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF 


             IF XGUCODIGO <> "NO" THEN
                XGUFNC2 = " document.all['---" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
             ELSE
                XGUFNC2 = ""
             END IF 


             strsql = "Select * From GRUPOS  order by descriPCION "
             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECIONAR FILTRO.' ID='***" & xgurs(0) & "' "

                IF CSTR(xgurs(0)) = XGUITEMCD THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF 



                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.XGUFRM.submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)

             IF XGUITEMCD <> "NO" THEN

                CALL getUsrIngroup(XGUITEMCD)

                XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9)) 
                XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9)) 

                xTrColor = "#FFFFFF"



                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                 XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"

                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR USUARIO.'  ID='---" & XGUTABKEY(XGUINX) & "' "

                IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                ELSE
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                END IF 

                XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.XGUFRM.submit();return false;"" >"
                XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 




             END IF



        CASE "S"

             XGUFILDSC1 = ":&nbsp;SECTOR"

             xTrColor = "#FFFFFF"

             IF XGUITEMCD <> "NO" THEN
                 XGUFNC1 = " document.all['***" & XGUITEMCD & "'].scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF 


             IF XGUCODIGO <> "NO" THEN
                XGUFNC2 = " document.all['---" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
             ELSE
                XGUFNC2 = ""
             END IF 



             strsql = "Select * From Eco  order by descri"
             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

                IF CSTR(xgurs(0)) = XGUITEMCD THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF 

                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.XGUFRM.submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)


             IF XGUITEMCD <> "NO" THEN

                CALL getUsrInSECT(XGUITEMCD)

                XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9)) 
                XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9)) 

                xTrColor = "#FFFFFF"

                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                 XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"

                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                ELSE
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                END IF 

                XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.XGUFRM.submit();return false;"" >"
                XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 




             END IF




        CASE "L"

             XGUFILDSC1 = ":&nbsp;LETRA"

             xTrColor = "#F9F9F9"


             IF XGUCODIGO <> "NO" THEN
                XGUFNC2 = " document.all['---" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
             ELSE
                XGUFNC2 = ""
             END IF 


             dim ATABXGULTR(28)

             ATABXGULTR(1)  = "A" 
             ATABXGULTR(2)  = "B" 
             ATABXGULTR(3)  = "C" 
             ATABXGULTR(4)  = "D" 
             ATABXGULTR(5)  = "E" 
             ATABXGULTR(6)  = "F" 
             ATABXGULTR(7)  = "G" 
             ATABXGULTR(8)  = "H" 
             ATABXGULTR(9)  = "I" 
             ATABXGULTR(10) = "J" 
             ATABXGULTR(11) = "K" 
             ATABXGULTR(12) = "L" 
             ATABXGULTR(13) = "M" 
             ATABXGULTR(14) = "N" 
             ATABXGULTR(15) = "O" 
             ATABXGULTR(16) = "P" 
             ATABXGULTR(17) = "Q" 
             ATABXGULTR(18) = "R" 
             ATABXGULTR(19) = "S" 
             ATABXGULTR(20) = "T" 
             ATABXGULTR(21) = "U" 
             ATABXGULTR(22) = "V" 
             ATABXGULTR(23) = "W" 
             ATABXGULTR(24) = "X" 
             ATABXGULTR(25) = "Y" 
             ATABXGULTR(26) = "Z" 
             ATABXGULTR(27) = "NO"
             ATABXGULTR(28) = "NO"


             FOR XGUINX=1 TO UBOUND(ATABXGULTR) STEP 4

                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECIONAR FILTRO.' align=center "

                 IF ATABXGULTR(XGUINX) = XGUITEMCD THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF 

                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX) & "','NO','NO');document.XGUFRM.submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & "</TD>"


                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECIONAR FILTRO.' align=center "

                 IF ATABXGULTR(XGUINX+1) = XGUITEMCD THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF 

                 XGUFBTD1 = XGUFBTD1 & " >"
                 XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+1) & "','NO','NO');document.XGUFRM.submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & "</TD>"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECIONAR FILTRO.' align=center "

                 if  ATABXGULTR(XGUINX+2) = "NO" then

                    XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                 else
                    IF ATABXGULTR(XGUINX+2) = XGUITEMCD THEN
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                    ELSE
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                    END IF 
                    XGUFBTD1 = XGUFBTD1 & " >"
                    XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
                    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+2) & "','NO','NO');document.XGUFRM.submit();return false;"" >"

                 end if

                 XGUFBTD1 = XGUFBTD1 & "</TD>"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECIONAR FILTRO.' align=center "

                 if  ATABXGULTR(XGUINX+3) = "NO" then

                    XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                 else
                    IF ATABXGULTR(XGUINX+3) = XGUITEMCD THEN
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                    ELSE
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                    END IF 
                    XGUFBTD1 = XGUFBTD1 & " >"
                    XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
                    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+3) & "','NO','NO');document.XGUFRM.submit();return false;"" >"

                 end if

                 XGUFBTD1 = XGUFBTD1 & "</TD>"



                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                 if xTrColor = "#FFFFFF" then
                    xTrColor = "#F9F9F9" 
                 ELSE
                    xTrColor = "#FFFFFF" 
                 END IF

             NEXT 




             IF XGUITEMCD <> "NO" THEN

                CALL getUsrInTEXT(XGUITEMCD & "*")

                XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9)) 
                XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9)) 

                xTrColor = "#FFFFFF"

                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                 XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"

                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                ELSE
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                END IF 

                XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.XGUFRM.submit();return false;"" >"
                XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 




             END IF




        CASE "T"

             XGUFILDSC1 = ":&nbsp;TEXTO"



             xTrColor = "#FFFFFF"


             IF XGUCODIGO <> "NO" THEN
                XGUFNC2 = " document.all['---" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
             ELSE
                XGUFNC2 = ""
             END IF 



                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECIONAR FILTRO.'  "

                 IF XGUITEMCD <> "NO" THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF 

                 XGUFBTD1 = XGUFBTD1 & " >"
                 XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE='" & XGUITEMCD & "' CLASS=DT size=40 >&nbsp;&nbsp;"

                 XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.all.XGUBTEXTO.value,'NO','NO');document.XGUFRM.submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & "</TD>"

                 XGUFBTD1 = XGUFBTD1 & "</TR>"


             IF XGUITEMCD <> "NO" THEN

                CALL getUsrInTEXT("*" & XGUITEMCD & "*")

                XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9)) 
                XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9)) 

                xTrColor = "#FFFFFF"

                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                 XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"

                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                ELSE
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                END IF 

                XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.XGUFRM.submit();return false;"" >"
                XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 




             END IF


        CASE ELSE

             XGUFILDSC1 = "&nbsp;"

             XGUFILTRO = "NO"
             XGUITEMCD = "NO"
             XGUCODIGO = "NO"
             XGUCODDSC = "NO"


 END SELECT






 XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
 XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


 XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
 XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

 XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
 XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
 XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>" 

 XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
 XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
 XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

 XGUCONTENT = ""
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
 XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;USUARIOS</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
 XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"


 END SELECT


  call ofv.cerrarconn(xgucn)


 Whtml = ""

 Whtml = Whtml & "<center>"
 Whtml = Whtml & "<DIV id='BSUSR' class='DINDM'  STYLE='display:block;z-index:99;margin-top:1;margin-left:1;width:650;"
 Whtml = Whtml & "'>"
 Whtml = Whtml & "<CENTER>"

 Whtml = Whtml & "<table id=""MTBSUSR1""  class='TWIND' width=99% cellspacing=0 cellpadding=2 >"
 Whtml = Whtml & "<tr HEIGHT=25 >"
 Whtml = Whtml & "<td class=mbt width=99%  >&nbsp;&nbsp;BUSCAR" & XGUFILDSC & "&nbsp;</td>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"

 Whtml = Whtml & "<table id=""MTBSUSR0""  class='aaa' width=99%  cellspacing=0 cellpadding=0 "
 Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<tD ALIGN=RIGHT >"
 Whtml = Whtml & XGUCONTEMS
 Whtml = Whtml & "</tD>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"

 Whtml = Whtml & "<table id=""MTBSUSR2""  class='aaa' width=99%  cellspacing=1 cellpadding=1 "
 Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & XGUFBBOT1
 Whtml = Whtml & XGUFBBOT2
 Whtml = Whtml & XGUFBBOT3
 Whtml = Whtml & XGUFBBOT4

 if XGUCODDSC <> "NO" then
    Whtml = Whtml & "<td class=uti width=70% style='border-style:solid;border-width:1;border-color:#cccccc;' >" & XGUCODDSC & "</td>"
 else
    Whtml = Whtml & "<td class=uti width=70%  >&nbsp;</td>"
 end if 

 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"

 Whtml = Whtml & "<table id=""MTBSUSR3""  class='aaa' width=99% cellspacing=1 cellpadding=1 "
 Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"

 Whtml = Whtml & XGUCONTENT

 Whtml = Whtml & "</table>"
 Whtml = Whtml & "</CENTER>"
 Whtml = Whtml & "</DIV>"
 Whtml = Whtml & "</center>"


 Whtml = Whtml & "<FORM NAME='XGUFRM' ACTION='" & XGUFBURL & "' METHOD='POST' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUMASTER' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFILTRO' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUITEMCD' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUCODIGO' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUCODDSC' >"
 Whtml = Whtml & "</FORM>"

 Whtml = Whtml & "<script language=javascript >"
 Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4) { " & chr(13)
 Whtml = Whtml & "    document.all.XGUMASTER.value=XGU0; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFILTRO.value=XGU1; " & chr(13)
 Whtml = Whtml & "    document.all.XGUITEMCD.value=XGU2; " & chr(13)
 Whtml = Whtml & "    document.all.XGUCODIGO.value=XGU3; " & chr(13)
 Whtml = Whtml & "    document.all.XGUCODDSC.value=XGU4; " & chr(13)
 Whtml = Whtml & " } " & chr(13)
 Whtml = Whtml & XGUFNC1
 Whtml = Whtml & XGUFNC2
 Whtml = Whtml & "</script>"



 SESSION("FBUSR") = XGUCODIGO
 SESSION("FBSLC") = XGUMASTER



     GenBuscaUsr = Whtml


 end function 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' FUNCIONES DE BUSQUEDa de Documentos ( documento-solapa-1er nivel)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



 function GenBuscaObj(XGUCOBJI,XGUTOBJI,XGUN1I,XGUN2I,XGUN3I,XGUN4I,XGUN5I)


 XGUFBURL = request.servervariables("SCRIPT_NAME")

 XGUFBURL = RIGHT(XGUFBURL,LEN(XGUFBURL) - (INSTRREV(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

 XGUMASTER = REQUEST.FORM("XGUMASTER")

 IF  XGUMASTER = "" THEN

     XGUMASTER = "0"

     XGUFCOBJ = "0"
     XGUFTOBJ = "0"
     XGUFNIV1 = "0"
     XGUFNIV2 = "0"
     XGUFNIV3 = "0"
     XGUFNIV4 = "0"
     XGUFNIV5 = "0"

     if   XGUCOBJI <> "0" THEN 
          XGUFCOBJ = XGUCOBJI
          XGUFTOBJ = XGUTOBJI
     END IF

     if   XGUN1I <> "0" THEN 
          XGUFNIV1 = XGUN1I
          XGUMASTER = "1"
     END IF

     if   XGUN2I <> "0" THEN 
          XGUFNIV2 = XGUN2I
          XGUMASTER = "2"
     END IF

     if   XGUN3I <> "0" THEN 
          XGUFNIV3 = XGUN3I
          XGUMASTER = "3"
     END IF

     if   XGUN4I <> "0" THEN 
          XGUFNIV4 = XGUN4I
          XGUMASTER = "4"
     END IF

     if   XGUN5I <> "0" THEN 
          XGUFNIV5 = XGUN5I
          XGUMASTER = "5"
     END IF






 ELSE

     XGUFCOBJ = REQUEST.FORM("XGUFCOBJ")
     XGUFTOBJ = REQUEST.FORM("XGUFTOBJ")
     XGUFNIV1 = REQUEST.FORM("XGUFNIV1")
     XGUFNIV2 = REQUEST.FORM("XGUFNIV2")
     XGUFNIV3 = REQUEST.FORM("XGUFNIV3")
     XGUFNIV4 = REQUEST.FORM("XGUFNIV4")
     XGUFNIV5 = REQUEST.FORM("XGUFNIV5")

 END IF



 XGUCONTEMS = ""



 XGUFNC1 = ""
 XGUFNC2 = ""


 set xgucn = ofv.conectar(ofv.strconn2)

 XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#f7aa22;cursor:hand;background-color:#bbddfe;' "
 XGUSTYLN = " STYLE='cursor:hand;' "

 XGUFBUSR = XGUCODIGO

 XGUFBTAB1  = ""
 XGUFBTAB2  = ""
 XGUFBTD1   = ""
 XGUFBTD2   = ""

 XGUFBTDF0  = ""
 XGUFBTDF1  = ""
 XGUFBTDF2  = ""
 XGUFBTDF3  = ""
 XGUFBTDF4  = ""
 XGUFBTDF5  = ""


 XGUFILDSC = "&nbsp;"

 XGUImG    = "<font face='webdings' size=3 color='#f7aa22'>4</font>"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' ARMADO DE DIV SEGUN NIVEL DE PROFUNDIDAD ...
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 SELECT CASE XGUMASTER


           '  IF XGUCODIGO <> "NO" THEN
           '     XGUFNC2 = " document.all['---" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
           '  ELSE
           '     XGUFNC2 = ""
           '  END IF 


        CASE "0"

                CALL GETLISTOBJ("L","0","0")
                XGUTABKEY  = SPLIT(XGUTABLISKEY,CHR(9)) 
                XGUTABKEY2 = SPLIT(XGUTABLISKEY2,CHR(9)) 
                XGUTABDAT  = SPLIT(XGUTABLISDAT,CHR(9)) 

                xTrColor = "#FFFFFF"



                IF XGUFCOBJ <> "0" THEN
                   XGUFNC2 = " document.all['---" & XGUFCOBJ & "'].scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF 



                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR DOCUMENTO.'  ID='---" & XGUTABKEY(XGUINX) & "' "

                    IF CSTR(XGUTABKEY(XGUINX)) = cstr(XGUFCOBJ) THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF 



                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 




        CASE "1"


                CALL GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLP("L",XGUFCOBJ,XGUFTOBJ,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9)) 
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9)) 

                xTrColor = "#FFFFFF"


                IF XGUFNIV1 <> "0" THEN
                   XGUFNC2 = " document.all['---" & XGUFNIV1 & "'].scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF 



                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR SOLAPA.'  ID='---" & XGUTABKEY(XGUINX) & "' "


                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV1 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF 

                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUTABKEY(XGUINX) & "','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 



        CASE "2"

                CALL GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"





                CALL GETLISTNV2("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9)) 
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9)) 

                xTrColor = "#FFFFFF"


                IF XGUFNIV2 <> "0" THEN
                   XGUFNC2 = " document.all['---" & XGUFNIV2 & "'].scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF 


                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR 1ER. NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV2 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF 


                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUTABKEY(XGUINX) & "','0','0','0');document.XGUFRM.submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 





        CASE "3"

                CALL GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"



                CALL GETLISTNV2("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                XGUFBTDF2 = XGUFBTDF2 & " >"

                XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
                XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
                XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF2 = XGUFBTDF2 & "</TR>"





                CALL GETLISTNV3("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9)) 
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9)) 

                xTrColor = "#FFFFFF"

                IF XGUFNIV3 <> "0" THEN
                   XGUFNC2 = " document.all['---" & XGUFNIV3 & "'].scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF 



                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR 2DO.NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV3 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF 


                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('4','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUTABKEY(XGUINX) & "','0','0');document.XGUFRM.submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 


        CASE "4"

                CALL GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"



                CALL GETLISTNV2("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                XGUFBTDF2 = XGUFBTDF2 & " >"

                XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
                XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
                XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF2 = XGUFBTDF2 & "</TR>"



                CALL GETLISTNV3("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,XGUFNIV3)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF3 = XGUFBTDF3 & "<TR  "
                XGUFBTDF3 = XGUFBTDF3 & " >"

                XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT  "
                XGUFBTDF3 = XGUFBTDF3 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUFNIV3 & "','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF3 = XGUFBTDF3 & XGUImG & "&nbsp;"
                XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF3 = XGUFBTDF3 & "</TR>"





                XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"
                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI "
                XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                XGUFBTD2 = XGUFBTD2 & "  >"
                XGUFBTD2 = XGUFBTD2 & "----------------------------------------------------------------------</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"






 END SELECT



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'XGUMASTER

 IF XGUMASTER > "0" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR0""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF0 & "</TABLE>"


 END IF

 IF XGUMASTER > "1" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR1""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF1 & "</TABLE>"


 END IF

 IF XGUMASTER > "2" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR2""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF2 & "</TABLE>"


 END IF

 IF XGUMASTER > "3" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR3""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


 END IF

 IF XGUMASTER > "0" THEN

    XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;width:99%;background-color:#ffffff;"
    XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;'>"
    XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>" 

 END IF




 XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
 XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

 XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
 XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
 XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

 XGUCONTENT = ""
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  call ofv.cerrarconn(xgucn)


 Whtml = ""

 Whtml = Whtml & "<center>"
 Whtml = Whtml & "<DIV id='BSUSR' class='DINDM'  STYLE='display:block;z-index:99;margin-top:1;margin-left:1;width:650;"
 Whtml = Whtml & "'>"
 Whtml = Whtml & "<CENTER>"

 Whtml = Whtml & "<table id=""MTBSUSR1""  class='TWIND' width=99% cellspacing=0 cellpadding=2 >"
 Whtml = Whtml & "<tr HEIGHT=25 >"
 Whtml = Whtml & "<td class=mbt width=99%  >&nbsp;&nbsp;SELECCIONAR&nbsp;DOCUMENTO&nbsp;Y&nbsp;CONTENIDO&nbsp;</td>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"

 Whtml = Whtml & "<table id=""MTBSUSR9""  class='aaa' width=99%  cellspacing=0 cellpadding=0 "
 Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<tD ALIGN=RIGHT >"
 Whtml = Whtml & XGUCONTEMS
 Whtml = Whtml & "</tD>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"


 Whtml = Whtml & "<table id=""MTBSUSR3""  class='aaa' width=99% cellspacing=1 cellpadding=1 "
 Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"

 Whtml = Whtml & XGUCONTENT

 Whtml = Whtml & "</table>"
 Whtml = Whtml & "</CENTER>"
 Whtml = Whtml & "</DIV>"
 Whtml = Whtml & "</center>"



 Whtml = Whtml & "<FORM NAME='XGUFRM' ACTION='" & XGUFBURL & "' METHOD='POST' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUMASTER' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFCOBJ' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFTOBJ' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV1' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV2' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV3' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV4' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV5' >"
 Whtml = Whtml & "</FORM>"

 Whtml = Whtml & "<script language=javascript >"
 Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4,XGU5,XGU6,XGU7) { " & chr(13)
 Whtml = Whtml & "    document.all.XGUMASTER.value=XGU0; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFCOBJ.value=XGU1; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFTOBJ.value=XGU2; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV1.value=XGU3; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV2.value=XGU4; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV3.value=XGU5; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV4.value=XGU6; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV5.value=XGU7; " & chr(13)
 Whtml = Whtml & " } " & chr(13)
 Whtml = Whtml & XGUFNC1
 Whtml = Whtml & XGUFNC2
 Whtml = Whtml & "</script>"




     GenBuscaObj = Whtml


 end function 




 function GETLISTOBJ(XGUFOPC,XGUFCOBJF,XGUFTOBJF)

     XGUTABLISKEY  = "0"
     XGUTABLISKEY2 = "NO"
     XGUTABLISDAT  = "SIN DEFINIR"

     XGUTABSEP    = CHR(9)

     IF XGUFOPC = "E" THEN

        Strsql = "Select * From csObjetosRDO where CODOBJETO = 0" & XGUFCOBJF
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT Xrs2.eof THEN

           XGUTABLISKEY  = Xrs2(0)
           XGUTABLISKEY2 = Xrs2(8)
           XGUTABLISDAT  = Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"


        END IF
        call ofv.cerrarconsulta(Xrs2)


     ELSE

        Strsql = "Select * From csObjetosRDO where estado = 3  OR ESTADO = 1 order by descri, OBJDESCRI"
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

        do while not Xrs2.eof

           XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
           XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(8)
           XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"



        Xrs2.movenext
        loop
        call ofv.cerrarconsulta(Xrs2)

     END IF  

     GETLISTOBJ = "OK"


 end function 



 function GETLISTSLP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL DOCUMENTO"

     XGUTABSEP    = CHR(9)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' BUSCA LAS SOLAPAS DEL DOCUMENTO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  dim stabla(20)
  dim solord(20)
  dim Nombres(20)
  dim Solapas(20)

  solord(0) = 0
  solord(1) = 15
  solord(2) = 1
  solord(3) = 6
  solord(4) = 3
  solord(5) = 2
  solord(6) = 4
  solord(7) = 5
  solord(8) = 7
  solord(9) = 8
  solord(10) = 11
  solord(11) = 10
  solord(12) = 9
  solord(13) = 12
  solord(14) = 13
  solord(15) = 14  
  solord(16) = 16
  solord(17) = 17 
  solord(18) = 18 
  solord(19) = 19 
  solord(20) = 20 


    for xx = 1 to 20
        solord(xx) = 17
        Nombres(xx) = "N"
    next


  Strsql = "Select * From Objperfildoc Where codobjeto =" & XGUFCOBJF & " and tipobjeto ='" & XGUFTOBJF & "' " 
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Xrs.eof then 
     perfilsp = "N"
  end if

  if perfilsp = "N" then 
     call ofv.cerrarconsulta(Xrs)

        sPerfil = "OBJMAIN"

     Strsql = "Select PERFIL From TIPOBJETO Where tipoBJETO='" & XGUFTOBJF & "'"
     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not Xrs.eof then 
        sPerfil = XRS(0)
     end if

     Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  end if

  '' response.write strsql & "<br>"


    for xx = 2 to 16
        stabla (xx - 1) = Xrs(xx + 30)
    next


    yy = 0
    for xx = 1 to 15

        if cint(stabla (xx)) < 99 then
           yy = yy + 1
           solord (stabla (xx)) = xx
        end if
    next
    tyy = yy



  for z=0 to 20
      Solapas(z) = "N"
  next  


  X=0
  for z = 2 to 16

    if Xrs(z) Then 

      if (z <> 2 and z <> 4 and z <> 5 )  then
          x = x + 1    
          Solapas(z-1)  = XRS(z + 15)

          if  z = 6 then

             Solapas(16)  = "CIRCUITO COMPLETO"

             for xiord=0 to 19 

                 if solord(xiord) = 5 then

                    xmovey = 16

                    for xiord2=xiord + 1 to 19 
                        xmoven           = solord(xiord2) 
                        solord(xiord2)   = xmovey
                        xmovey           = xmoven
                    next



                    exit for
                 end if  
             next
          

          end if   

      end if
    end if

  next


  call ofv.cerrarconsulta(Xrs)

  xfinal = ubound(Solapas)






     IF XGUFOPC = "E" THEN

        Strsql = "Select * From csGenObjetos where CODOBJETO = 0" & XGUFCOBJF
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT Xrs2.eof THEN

           XGUTABLISKEY = Xrs2(0)
           XGUTABLISDAT = Xrs2(1)


        END IF
        call ofv.cerrarconsulta(Xrs2)

        for x = xstrsol to xfinal
            if solapas(solord(x)) <> "N" then
            '   response.write solord(x)
               if clng(XGUFNIV1F) = clng(solord(x)) then

                 XGUTABLISKEY = solord(x)
                 XGUTABLISDAT = solapas(solord(x))

                 EXIT FOR

               END IF 
             end if 
        next



     ELSE
        

        for x = xstrsol to xfinal
            if solapas(solord(x)) <> "N" then

                 XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & solord(x)
                 XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & solapas(solord(x))

            end if
        next




     END IF  

     GETLISTSLP = "OK"

 end function 



 function GETLISTNV2(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL CONTENIDO"

     XGUTABSEP    = CHR(9)

     XGUTABLAF    = "NO"


     SELECT CASE XGUFNIV1F

            CASE "2"   'CONSIDERACIONES 

                 XGUTABLAF    = "OBJNWEDT"   

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM OBJNWEDT WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then 
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM OBJNWEDT WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by consecuen "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(3)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF  







            CASE "5"   'PROCEDIMIENTOS 

                  XGUTABLAF    = "CSCEventoSectorPerfil"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM CSCEventoSectorPerfil WHERE CODPROCED = '" & XGUFNIV2F & "' "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then 
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(2) & "."  & Xrs(6) & ".&nbsp;" & Xrs(4) & "&nbsp;-&nbsp;"  & Xrs(3) & "&nbsp;-&nbsp;" & Xrs(5)


                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM CSCEventoSectorPerfil WHERE obcodobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by SECUEN, SECUENCIA "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(2) & "."  & Xrs(6) & ".&nbsp;" & Xrs(4) & "&nbsp;-&nbsp;"  & Xrs(3) & "&nbsp;-&nbsp;" & Xrs(5)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF  







            CASE "7"   'CONTROLES 

                  XGUTABLAF    = "cscontrolsel"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM cscontrolsel WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then 
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(1)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM cscontrolsel WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by DESCRI "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(1)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF  





            CASE "8"   'ESQUEMAS 

                  XGUTABLAF    = "csesquemasel"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM csesquemasel WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then 
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(1)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM csesquemasel WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by DESCRI "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(1)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF  






            CASE "9"   'TABLAS 

                  XGUTABLAF    = "ObjTabDef"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM ObjTabDef WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then 
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(4)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM ObjTabDef WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by tabsecuen "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(4)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF  





            CASE "11"   'CONTENIDO 

                  XGUTABLAF    = "OBJESTR"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM OBJESTR WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then 
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM OBJESTR WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by CONsecuen "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(3)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF  



            CASE ELSE

                  XGUTABLAF    = "NO"



     END SELECT




     GETLISTNV2 = "OK"

 end function 


 function GETLISTNV3(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F,XGUFNIV3F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL CONTENIDO"

     XGUTABSEP    = CHR(9)




     IF XGUFOPC = "E" THEN

        IF  XGUFNIV1F = "9" THEN


            Strsql = "SELECT * FROM ObjTabblk WHERE OBJID = 0" & XGUFNIV3F

            set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not Xrs.eof then 
               XGUTABLISKEY = Xrs(0)
               XGUTABLISDAT = Xrs(5)

            END IF
            call ofv.cerrarconsulta(Xrs)


        END IF

        IF  XGUFNIV1F = "11" THEN

            Strsql = "SELECT * FROM OBJESTDET WHERE OBJID = 0" & XGUFNIV3F
            set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not Xrs.eof then 
               XGUTABLISKEY = Xrs(0)
               XGUTABLISDAT = UCASE(Xrs(7))
            END IF
            call ofv.cerrarconsulta(Xrs)

        END IF

     ELSE

        IF  XGUFNIV1F = "9" THEN


                     Strsql = "SELECT * FROM ObjTabblk WHERE tabdef = 0" & XGUFNIV2F
                     Strsql = Strsql & " order by tabblksec "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(5)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)

        end if


        IF  XGUFNIV1F = "11" THEN

            strsql = " SELECT * FROM CSESTDET "   
            strsql = strsql & " where detcon = " & XGUFNIV2F & "  "
            strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
            strsql = strsql & " and not ( "
            strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) " 
            strsql = strsql & " ) and stmodi = '' AND DETTIPO < 'G' "
            strsql = strsql & " order by TSEC"

            set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            do while not Xrs.eof

               XGUINDN = "-&nbsp;"

               SELECT CASE XRS(5) 

                      CASE "B" 

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;"


                      CASE "C"

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"


                      CASE "D"

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"


                      CASE "E"

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"


                      CASE "F"

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"



               END SELECT

               XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
               XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & XGUINDN & UCASE(Xrs(7))

            Xrs.movenext
            loop
            call ofv.cerrarconsulta(Xrs)


        END IF


     END IF  

     GETLISTNV3 = "OK"

 end function 


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BUSCADOR DE NIVELES DE COMPONENTES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


 function GETLISTCMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF)

     XGUTABLISKEY  = "0"
     XGUTABLISKEY2 = "NO"
     XGUTABLISDAT  = "SIN DEFINIR"

     XGUTABSEP    = CHR(9)

     IF XGUFOPC = "E" THEN

        Strsql = "Select * From csacomple where CODCOMP = 0" & XGUFCOBJF
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT Xrs2.eof THEN

           XGUTABLISKEY  = Xrs2(0)
           XGUTABLISKEY2 = Xrs2(1)
           XGUTABLISDAT  = Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)


        END IF
        call ofv.cerrarconsulta(Xrs2)


     ELSE

        Strsql = "Select * From csacomple  order by descri, TIPODESCRI"
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

        do while not Xrs2.eof

           XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
           XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(1)
           XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)



        Xrs2.movenext
        loop
        call ofv.cerrarconsulta(Xrs2)

     END IF  

     GETLISTCMP = "OK"


 end function 


 function GETLISTSLPCMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL COMPONENTE"

     XGUTABSEP    = CHR(9)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' BUSCA LAS SOLAPAS DEL COMPONENTE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  dim stabla(20)
  dim solord(20)
  dim Nombres(20)


  solord(0) = 0
  solord(1) = 15
  solord(2) = 1
  solord(3) = 6
  solord(4) = 3
  solord(5) = 2
  solord(6) = 4
  solord(7) = 5
  solord(8) = 7
  solord(9) = 8
  solord(10) = 11
  solord(11) = 10
  solord(12) = 9
  solord(13) = 12
  solord(14) = 13
  solord(15) = 14  
  solord(16) = 16
  solord(17) = 17 
  solord(18) = 18 
  solord(19) = 19 
  solord(20) = 20 



'''''''''''''''''''''''

  dim Solapas(20)


  for z=0 to 20
      Solapas(z) = "N"
  next  


  Strsql = "Select * From tipocomp Where tipo = '" & XGUFTOBJF & "' "
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not Xrs.eof then 
     sPerfil  = Xrs(4)
  end if

  call ofv.cerrarconsulta(Xrs)



  Strsql = "Select * From comperfil Where tipoperfil='" & sPerfil & "'"
'' response.write strsql & "<br>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not Xrs.eof then 



    for xx = 1 to 20
        solord(xx) = 17
        Nombres(xx) = "N"
    next



    for xx = 2 to 16
        stabla (xx - 1) = Xrs(xx + 30)
    next



    yy = 0
    for xx = 1 to 15

        if cint(stabla (xx)) < 99 then
           yy = yy + 1
           solord (stabla (xx)) = xx
        end if
    next
    tyy = yy

    xstrsol     = 1



  X=0
  for z = 2 to 16

    if Xrs(z) Then 

      if  z <> 21 and z <> 4 and z <> 61 then
          x = x + 1    
          Solapas(z-1)  = REPLACE(XRS(z + 15)," ","*")

      end if
    end if

  next

  end if

  call ofv.cerrarconsulta(Xrs)


  xfinal = ubound(Solapas)










     IF XGUFOPC = "E" THEN

        Strsql = "Select * From CSACOMPLE where CODCOMP = 0" & XGUFCOBJF
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT Xrs2.eof THEN

           XGUTABLISKEY = Xrs2(0)
           XGUTABLISDAT = Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)


        END IF
        call ofv.cerrarconsulta(Xrs2)

        for x = xstrsol to xfinal
            if solapas(solord(x)) <> "N" then
            '   response.write solord(x)
               if clng(XGUFNIV1F) = clng(solord(x)) then

                 XGUTABLISKEY = solord(x)
                 XGUTABLISDAT = solapas(solord(x))

                 EXIT FOR

               END IF 
             end if 
        next



     ELSE
        

        for x = xstrsol to xfinal
            if solapas(solord(x)) <> "N" then

                 XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & solord(x)
                 XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP  & solord(x) & "&nbsp;" & solapas(solord(x))

            end if
        next




     END IF  

     GETLISTSLPCMP = "OK"

 end function 


 function GETLISTNV2CMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL CONTENIDO " & XGUFNIV1F

     XGUTABSEP    = CHR(9)

     XGUTABLAF    = "NO"


     SELECT CASE XGUFNIV1F

            CASE "2"   'estructura 

                 XGUTABLAF    = "COMPESTR"   

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then 
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE CODCOMP = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by OBJID "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(3)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF  







            CASE "4"   'imagenes 

                  XGUTABLAF    = "COMPIMAG"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then 
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(4)

                        IF XGUTABLISDAT = "" THEN XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE CODCOMP = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by SECUENCIA, OBJID "
                  ''   response.write strsql & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)

                        WKXGUTABLISDAT = Xrs(4)

                        IF WKXGUTABLISDAT = "" THEN WKXGUTABLISDAT = Xrs(3)

                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & WKXGUTABLISDAT



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF  







            CASE "6"   'archivos 

                  XGUTABLAF    = "COMPARC"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then 
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(4)

                        IF XGUTABLISDAT = "" THEN XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE CODCOMP = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by SECUENCIA, OBJID "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)

                        WKXGUTABLISDAT = Xrs(4)

                        IF WKXGUTABLISDAT = "" THEN WKXGUTABLISDAT = Xrs(3)

                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & WKXGUTABLISDAT



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF  








            CASE ELSE

                  XGUTABLAF    = "NO"



     END SELECT




     GETLISTNV2CMP = "OK"

 end function 


 function GETLISTNV3CMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F,XGUFNIV3F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL CONTENIDO"

     XGUTABSEP    = CHR(9)

     GETLISTNV3CMP = "OK"

 end function 



 function GenBuscaCMP(XGUCOBJI,XGUTOBJI,XGUN1I,XGUN2I,XGUN3I,XGUN4I,XGUN5I)


 XGUFBURL = request.servervariables("SCRIPT_NAME")

 XGUFBURL = RIGHT(XGUFBURL,LEN(XGUFBURL) - (INSTRREV(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

 XGUMASTER = REQUEST.FORM("XGUMASTER")

 IF  XGUMASTER = "" THEN

     XGUMASTER = "0"

     XGUFCOBJ = "0"
     XGUFTOBJ = "0"
     XGUFNIV1 = "0"
     XGUFNIV2 = "0"
     XGUFNIV3 = "0"
     XGUFNIV4 = "0"
     XGUFNIV5 = "0"

     if   XGUCOBJI <> "0" THEN 
          XGUFCOBJ = XGUCOBJI
          XGUFTOBJ = XGUTOBJI
     END IF

     if   XGUN1I <> "0" THEN 
          XGUFNIV1 = XGUN1I
          XGUMASTER = "1"
     END IF

     if   XGUN2I <> "0" THEN 
          XGUFNIV2 = XGUN2I
          XGUMASTER = "2"
     END IF

     if   XGUN3I <> "0" THEN 
          XGUFNIV3 = XGUN3I
          XGUMASTER = "3"
     END IF

     if   XGUN4I <> "0" THEN 
          XGUFNIV4 = XGUN4I
          XGUMASTER = "4"
     END IF

     if   XGUN5I <> "0" THEN 
          XGUFNIV5 = XGUN5I
          XGUMASTER = "5"
     END IF






 ELSE

     XGUFCOBJ = REQUEST.FORM("XGUFCOBJ")
     XGUFTOBJ = REQUEST.FORM("XGUFTOBJ")
     XGUFNIV1 = REQUEST.FORM("XGUFNIV1")
     XGUFNIV2 = REQUEST.FORM("XGUFNIV2")
     XGUFNIV3 = REQUEST.FORM("XGUFNIV3")
     XGUFNIV4 = REQUEST.FORM("XGUFNIV4")
     XGUFNIV5 = REQUEST.FORM("XGUFNIV5")

 END IF



 XGUCONTEMS = ""



 XGUFNC1 = ""
 XGUFNC2 = ""


 set xgucn = ofv.conectar(ofv.strconn2)

 XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#f7aa22;cursor:hand;background-color:#bbddfe;' "
 XGUSTYLN = " STYLE='cursor:hand;' "

 XGUFBUSR = XGUCODIGO

 XGUFBTAB1  = ""
 XGUFBTAB2  = ""
 XGUFBTD1   = ""
 XGUFBTD2   = ""

 XGUFBTDF0  = ""
 XGUFBTDF1  = ""
 XGUFBTDF2  = ""
 XGUFBTDF3  = ""
 XGUFBTDF4  = ""
 XGUFBTDF5  = ""


 XGUFILDSC = "&nbsp;"

 XGUImG    = "<font face='webdings' size=3 color='#f7aa22'>4</font>"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' ARMADO DE DIV SEGUN NIVEL DE PROFUNDIDAD ...
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 SELECT CASE XGUMASTER


           '  IF XGUCODIGO <> "NO" THEN
           '     XGUFNC2 = " document.all['---" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
           '  ELSE
           '     XGUFNC2 = ""
           '  END IF 


        CASE "0"

                CALL GETLISTCMP("L","0","0")
                XGUTABKEY  = SPLIT(XGUTABLISKEY,CHR(9)) 
                XGUTABKEY2 = SPLIT(XGUTABLISKEY2,CHR(9)) 
                XGUTABDAT  = SPLIT(XGUTABLISDAT,CHR(9)) 

                xTrColor = "#FFFFFF"



                IF XGUFCOBJ <> "0" THEN
                   XGUFNC2 = " document.all['---" & XGUFCOBJ & "'].scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF 



                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR COMPONENTE.'  ID='---" & XGUTABKEY(XGUINX) & "' "

                    IF CSTR(XGUTABKEY(XGUINX)) = cstr(XGUFCOBJ) THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF 



                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 




        CASE "1"


                CALL GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLPCMP("L",XGUFCOBJ,XGUFTOBJ,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9)) 
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9)) 

                xTrColor = "#FFFFFF"


                IF XGUFNIV1 <> "0" THEN
                   XGUFNC2 = " document.all['---" & XGUFNIV1 & "'].scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF 



                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR SOLAPA.'  ID='---" & XGUTABKEY(XGUINX) & "' "


                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV1 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF 

                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUTABKEY(XGUINX) & "','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 



        CASE "2"

                CALL GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLPCMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"





                CALL GETLISTNV2CMP("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9)) 
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9)) 

                xTrColor = "#FFFFFF"


                IF XGUFNIV2 <> "0" THEN
                   XGUFNC2 = " document.all['---" & XGUFNIV2 & "'].scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF 


                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR 1ER. NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV2 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF 


                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUTABKEY(XGUINX) & "','0','0','0');document.XGUFRM.submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 





        CASE "3"

                CALL GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLPCMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"



                CALL GETLISTNV2CMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                XGUFBTDF2 = XGUFBTDF2 & " >"

                XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
                XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
                XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF2 = XGUFBTDF2 & "</TR>"





                CALL GETLISTNV3CMP("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9)) 
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9)) 

                xTrColor = "#FFFFFF"

                IF XGUFNIV3 <> "0" THEN
                   XGUFNC2 = " document.all['---" & XGUFNIV3 & "'].scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF 



                FOR XGUINX=0 TO UBOUND(XGUTABKEY) 


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR 2DO.NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV3 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF 


                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('4','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUTABKEY(XGUINX) & "','0','0');document.XGUFRM.submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


                NEXT 


        CASE "4"

                CALL GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLPCMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"



                CALL GETLISTNV2CMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                XGUFBTDF2 = XGUFBTDF2 & " >"

                XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
                XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
                XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF2 = XGUFBTDF2 & "</TR>"



                CALL GETLISTNV3CMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,XGUFNIV3)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF3 = XGUFBTDF3 & "<TR  "
                XGUFBTDF3 = XGUFBTDF3 & " >"

                XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT  "
                XGUFBTDF3 = XGUFBTDF3 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUFNIV3 & "','0','0');document.XGUFRM.submit();return false;"" >"
                XGUFBTDF3 = XGUFBTDF3 & XGUImG & "&nbsp;"
                XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF3 = XGUFBTDF3 & "</TR>"





                XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"
                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI "
                XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                XGUFBTD2 = XGUFBTD2 & "  >"
                XGUFBTD2 = XGUFBTD2 & "----------------------------------------------------------------------</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"






 END SELECT



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'XGUMASTER

 IF XGUMASTER > "0" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR0""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF0 & "</TABLE>"


 END IF

 IF XGUMASTER > "1" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR1""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF1 & "</TABLE>"


 END IF

 IF XGUMASTER > "2" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR2""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF2 & "</TABLE>"


 END IF

 IF XGUMASTER > "3" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR3""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


 END IF

 IF XGUMASTER > "0" THEN

    XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;width:99%;background-color:#ffffff;"
    XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;'>"
    XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>" 

 END IF




 XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
 XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

 XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
 XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
 XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

 XGUCONTENT = ""
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"
 XGUCONTENT = XGUCONTENT & "<tr>"
 XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
 XGUCONTENT = XGUCONTENT & "</tr>"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  call ofv.cerrarconn(xgucn)


 Whtml = ""

 Whtml = Whtml & "<center>"
 Whtml = Whtml & "<DIV id='BSUSR' class='DINDM'  STYLE='display:block;z-index:99;margin-top:1;margin-left:1;width:650;"
 Whtml = Whtml & "'>"
 Whtml = Whtml & "<CENTER>"

 Whtml = Whtml & "<table id=""MTBSUSR1""  class='TWIND' width=99% cellspacing=0 cellpadding=2 >"
 Whtml = Whtml & "<tr HEIGHT=25 >"
 Whtml = Whtml & "<td class=mbt width=99%  >&nbsp;&nbsp;SELECCIONAR&nbsp;COMPONENTES&nbsp;Y&nbsp;CONTENIDO&nbsp;</td>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"

 Whtml = Whtml & "<table id=""MTBSUSR9""  class='aaa' width=99%  cellspacing=0 cellpadding=0 "
 Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<tD ALIGN=RIGHT >"
 Whtml = Whtml & XGUCONTEMS
 Whtml = Whtml & "</tD>"
 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"


 Whtml = Whtml & "<table id=""MTBSUSR3""  class='aaa' width=99% cellspacing=1 cellpadding=1 "
 Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"

 Whtml = Whtml & XGUCONTENT

 Whtml = Whtml & "</table>"
 Whtml = Whtml & "</CENTER>"
 Whtml = Whtml & "</DIV>"
 Whtml = Whtml & "</center>"



 Whtml = Whtml & "<FORM NAME='XGUFRM' ACTION='" & XGUFBURL & "' METHOD='POST' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUMASTER' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFCOBJ' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFTOBJ' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV1' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV2' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV3' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV4' >"
 Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV5' >"
 Whtml = Whtml & "</FORM>"

 Whtml = Whtml & "<script language=javascript >"
 Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4,XGU5,XGU6,XGU7) { " & chr(13)
 Whtml = Whtml & "    document.all.XGUMASTER.value=XGU0; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFCOBJ.value=XGU1; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFTOBJ.value=XGU2; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV1.value=XGU3; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV2.value=XGU4; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV3.value=XGU5; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV4.value=XGU6; " & chr(13)
 Whtml = Whtml & "    document.all.XGUFNIV5.value=XGU7; " & chr(13)
 Whtml = Whtml & " } " & chr(13)
 Whtml = Whtml & XGUFNC1
 Whtml = Whtml & XGUFNC2
 Whtml = Whtml & "</script>"




     GenBuscaCMP = Whtml


 end function 






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function SETEMPRE()

     xGRP = "S"
     xUSR = "S"

  set xcn = ofv.conectar(ofv.strconn2)

  Strsql = "select SITE from USUARIOS where ID = '" & SESSION("USUARIO") & "'"
  set Xrs = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if not Xrs.eof then
     USREMP = Xrs(0)
     USREMP = clng(USREMP)
  end if

  Strsql = "select * from empreatrm where ATRMTIPO = 'G' AND atrmusrgrp = '" & SESSION("SEGGRP") & "' ORDER BY EMPRE "
  set Xrs1 = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if Xrs1.eof then
     xGRP = "N"
  end if


  Strsql = "select * from empreatrm where ATRMTIPO = 'U' AND atrmusrgrp = '" & SESSION("USUARIO") & "' ORDER BY EMPRE "
  set Xrs2 = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if Xrs2.eof then
     xUSR = "N"
  end if

        ATREMP(0,0) = 0

  Strsql = "select * from CLIENTEPLAN ORDER BY ID"
  set Xrs = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if not Xrs.eof then
     DO WHILE NOT XRS.EOF
        IAEMP = IAEMP + 1
        TAEMP = IAEMP

        ATREMP(IAEMP,0) = XRS(0)
        ATREMP(IAEMP,1) = XRS(6)       
        ATREMP(IAEMP,4) = XRS(1)    
        ATREMP(IAEMP,5) = XRS(5)             
        ATREMP(IAEMP,2) = 1    
        ATREMP(IAEMP,3) = 1    

       FOR XX1=0 TO 50  
        IF XGRP <> "N" THEN
           IF  XRS1(0) = XRS(0) THEN
               ATREMP(IAEMP,2) = ABS(CINT(XRS1(5)))   
               ATREMP(IAEMP,3) = ABS(CINT(XRS1(6))) 
               XRS1.MOVENEXT
               IF XRS1.EOF THEN XGRP = "N" 
               EXIT FOR
           ELSEIF  XRS1(0) < XRS(0) THEN
               XRS1.MOVENEXT
               IF XRS1.EOF THEN XGRP = "N" 
           END IF
        ELSE
           EXIT FOR 
        END IF
       NEXT 

       FOR XX1=0 TO 50  
        IF XUSR <> "N" THEN
           IF  XRS2(0) = XRS(0) THEN
               ATREMP(IAEMP,2) = ABS(CINT(XRS2(5)))   
               ATREMP(IAEMP,3) = ABS(CINT(XRS2(6))) 
               XRS2.MOVENEXT
               IF XRS2.EOF THEN XUSR = "N"   
               EXIT FOR
           ELSEIF  XRS2(0) < XRS(0) THEN
               XRS2.MOVENEXT
               IF XRS2.EOF THEN XUSR = "N" 
           END IF
        ELSE
           EXIT FOR 
        END IF
       NEXT 


       IF XRS(0) = USREMP THEN
          IF  ATREMP(IAEMP,2) = 1 THEN

              ATREMP(0,0) = ATREMP(IAEMP,0)
              ATREMP(0,1) = ATREMP(IAEMP,1) 
              ATREMP(0,4) = ATREMP(IAEMP,4)   
              ATREMP(0,5) = ATREMP(IAEMP,5)           
              ATREMP(0,2) = ATREMP(IAEMP,2)   
              ATREMP(0,3) = ATREMP(IAEMP,3)   

          END IF 
       END IF


     XRS.MOVENEXT
     LOOP

     IF ATREMP(0,0) = 0 THEN

        FOR IAEMP = 1 TO TAEMP
          IF  ATREMP(IAEMP,2) = 1 THEN

              ATREMP(0,0) = ATREMP(IAEMP,0)
              ATREMP(0,1) = ATREMP(IAEMP,1) 
              ATREMP(0,4) = ATREMP(IAEMP,4)       
              ATREMP(0,5) = ATREMP(IAEMP,5)       
              ATREMP(0,2) = ATREMP(IAEMP,2)   
              ATREMP(0,3) = ATREMP(IAEMP,3)   
              EXIT FOR  
          END IF 


        NEXT
     END IF

    DIM WATREMPEL(5)
    DIM WATREMPTB(50)

         
        FOR IAEMP = 0 TO TAEMP
         FOR IAEMP2 = 0 TO 5
              WATREMPEL(IAEMP2) = cstr(ATREMP(IAEMP,IAEMP2) &"")
         NEXT

         WATREMPTB(IAEMP) = JOIN(WATREMPEL,CHR(09))

        NEXT   


    SESSION("ATREMP") = JOIN(WATREMPTB,CHR(10))
    SESSION("TAEMP")  = TAEMP
    SESSION("SETEMP") = "S"


  end if



end function

FUNCTION NEWEMPRE(XEMPRE)

    IF XEMPRE = "NO" THEN

       if SESSION("SETEMP") <> "S" THEN   xxxx = SETEMPRE()

       WATREMPTB = SPLIT(SESSION("ATREMP"),CHR(10))

       FOR IAEMP=0 TO UBOUND(WATREMPTB)

           WATREMPEL = SPLIT(WATREMPTB(IAEMP),CHR(09))

           FOR IAEMP2=0 TO UBOUND(WATREMPEL)

               ATREMP(IAEMP,IAEMP2) = WATREMPEL(IAEMP2)

           NEXT

       NEXT

       TAEMP  = SESSION("TAEMP")

       NEWEMPRE = "OK"

    ELSE

     NEWEMPRE = "N"

     IF ATREMP(0,0) <> XEMPRE THEN

        NEWEMPRE = "S"

        FOR IAEMP = 1 TO TAEMP
          IF  ATREMP(IAEMP,0) = XEMPRE THEN

              ATREMP(0,0) = ATREMP(IAEMP,0)
              ATREMP(0,1) = ATREMP(IAEMP,1) 
              ATREMP(0,4) = ATREMP(IAEMP,4)       
              ATREMP(0,5) = ATREMP(IAEMP,5)       
              ATREMP(0,2) = ATREMP(IAEMP,2)   
              ATREMP(0,3) = ATREMP(IAEMP,3)   
              EXIT FOR  
          END IF 


        NEXT

        DIM WATREMPEL(5)
        DIM WATREMPTB(50)

         
        FOR IAEMP = 0 TO TAEMP
         FOR IAEMP2 = 0 TO 5
              WATREMPEL(IAEMP2) = ATREMP(IAEMP,IAEMP2)
         NEXT

         WATREMPTB(IAEMP) = JOIN(WATREMPEL,CHR(09))

        NEXT   


        SESSION("ATREMP") = JOIN(WATREMPTB,CHR(10))
        SESSION("TAEMP")  = TAEMP

     END IF
    END IF

END FUNCTION


function GETGLBHMSW(XITIME)

 if not isnumeric(XITIME & "") then

 GETGLBHMSW = "000005"

 else 

XITIME=clng("0" & replace(cstr(XITIME & ""),"-",""))
GLBbb = XITIME MOD 3600
GLBbC = (GLBBB) MOD 60
GLBbD = (GLBBC) MOD 1

GLBSH = (XITIME - GLBBB) / 3600
GLBSM = (GLBBB - GLBBC) / 60
GLBSS = GLBBC / 1

GLBSH = cstr(GLBSH)
GLBSM = cstr(GLBSM)
GLBSS = cstr(GLBSS)

if len(GLBSH) = 1 then GLBSH = "0" & GLBSH
if len(GLBSM) = 1 then GLBSM = "0" & GLBSM
if len(GLBSS) = 1 then GLBSS = "0" & GLBSS

 XOSH = GLBSH & ":"
 XOSM = GLBSM & ":"
 XOSS = GLBSS & ""


 GETGLBHMSW = XOSH & XOSM & XOSS

 end if

END FUNCTION


function GETGLBHMS(XITIME)

 if not isnumeric(XITIME & "") then

 GETGLBHMS = "000005"

 else 
''response.write "-" & XITIME & "-" & "<br>"
XITIME=clng("0" & replace(cstr(XITIME & ""),"-",""))
GLBbb = XITIME MOD 3600
GLBbC = (GLBBB) MOD 60
GLBbD = (GLBBC) MOD 1

GLBSH = (XITIME - GLBBB) / 3600
GLBSM = (GLBBB - GLBBC) / 60
GLBSS = GLBBC / 1

IF GLBSH > 0 THEN XOSH = GLBSH & "h&nbsp;"
IF GLBSM > 0 THEN XOSM = GLBSM & "m&nbsp;"
IF GLBSS > 0 THEN XOSS = GLBSS & "s&nbsp;"


 GETGLBHMS = XOSH & XOSM & XOSS

 end if

END FUNCTION

function GETFECH()



    sh = datepart("h",TIME())
    sm = datepart("n",TIME())
    ss = datepart("s",TIME())

   XLHORDAT = (CINT(sh)*3600) + (CINT(sm)*60) + CINT(ss)
 ''   XLHORDAT = TIME()
  ''  RESPONSE.WRITE XLHORDAT & "<BR>"
    XLFECDAT = CLNG(DATE())



END FUNCTION

function GETNWEN(XNCODIGO,XNTIPO)


    sh = datepart("h",TIME())
    sm = datepart("n",TIME())
    ss = datepart("s",TIME())

XCodobJ = sh & sm & ss

GThms   = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses   = session.sessionid
GTsecu  = cstr(XCodobJ & GThms & GTses)

GTsecuf = XNCODIGO & XNTIPO


for GTi=1 to len(GTsecu)
    GTcs = mid(GTsecu,GTi,1)
    if GTcs => "0" and GTcs <= "9" then
       GTsecuf = GTsecuf & GTcs
    end if
next
 
GETNWEN = trim(GTsecuf)
 
end function

function GETLOGN(XLTYPE)

    IF XLTYPE = "" THEN XLTYPE = "L"

    sh = datepart("h",TIME())
    sm = datepart("n",TIME())
    ss = datepart("s",TIME())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

SELECT CASE XLTYPE
       CASE "L" 
             GTsecuf = "LOG"
       CASE "LE" 
             GTsecuf = "LXT"
       CASE "LD" 
             GTsecuf = "LDE"
END SELECT

for GTi=1 to len(GTsecu)
    GTcs = mid(GTsecu,GTi,1)
    if GTcs => "0" and GTcs <= "9" then
       GTsecuf = GTsecuf & GTcs
    end if
next
 
GETLOGN = trim(GTsecuf)

END FUNCTION


'''' LOG DE USUARIOS ''''''''''''''''''''''''''''''''''''''''''''''''''

FUNCTION GENUSRLOG()

   '      Session("Usuario") = UCase(Request.Form("Usuarnt"))
   '      Session("UsuarNT") = UCase(Request.Form("Usuarnt"))
   '      Session("DOMINIO") = UCase(Request.Form("DOMAINnt"))
   '      Session("TRUSTEDUSR") = "S"


  StrSql = "Select * From OBJALOG where CODOBJETO = 0 AND  TIPOBJETO = 'SKRLOG' "
  STRSQL = STRSQL & " AND USUARIO = '" & session("USUARIO") & "' AND  NETUSUARIO = '" & session("USUARNT") & "' "
  STRSQL = STRSQL & " AND DOMINIO = '" & session("DOMINIO") & "' "
 ''       response.write strsql & "<BR>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if Xrs.eof then

        STRSQL = "INSERT INTO OBJALOG (CODOBJETO,TIPOBJETO,FULTACC,AEDI,ACON,CACT,TTOTALACC,USUARIO,NETUSUARIO,DOMINIO,TRUSTED) VALUES ("  
        STRSQL = STRSQL & "'0','SKRLOG','NO','0','0','0','0','" & session("USUARIO") & "','" & session("USUARNT") & "','" & session("DOMINIO") & "','" & Session("TRUSTEDUSR") & "')"

        CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)
'        response.write strsql

  END IF



END FUNCTION

FUNCTION GETUSRLOGUPD(XLTIPO)

 '' RESPONSE.WRITE XLTIPO & "<BR>"
  XLOPC = "N"

  SELECT CASE XLTIPO

         CASE "LOGON" 
    ''  RESPONSE.WRITE XLTIPO & "0<BR>"
              XLOPC = "S"

              XLLOGIDDET = GETLOGN("L")
              SESSION("LOGIDUSRDET") = XLLOGIDDET
              XXXX = GETFECH()
              XXXX = GENUSRLOGDET(XLTIPO)
              XLUPD = " SET AEDI = AEDI + 1, FULTACC = '" & XLFECDAT & "' "



         CASE "LOGOFF" 
              XLOPC = "S"
              XXXX = GETFECH()
              XLHORAINI = GETUSRLOGDETTIME()
              XLUPD = " SET TTOTALACC = TTOTALACC + " & (cdbl("0" & XLHORDAT) - cdbl("0" & XLHORAINI))
              XXXX = GENUSRLOGDET(XLTIPO)


  END SELECT
     '' RESPONSE.WRITE SESSION("LOGIDUSRDET") & "1<BR>"
  IF XLOPC = "S" THEN
     '' RESPONSE.WRITE SESSION("LOGIDUSRDET") & "2<BR>"

     StrSql = "UPDATE OBJALOG " & XLUPD
     StrSql = StrSql & " where CODOBJETO = 0 AND  TIPOBJETO = 'SKRLOG' "
     STRSQL = STRSQL & " AND USUARIO = '" & session("USUARIO") & "' AND  NETUSUARIO = '" & session("USUARNT") & "' "
     STRSQL = STRSQL & " AND DOMINIO = '" & session("DOMINIO") & "' "
    ''    response.write strsql & "<BR>"
     CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)


   
  END IF



END FUNCTION

FUNCTION GENUSRLOGDET(XLTIPO)

  XLOPC = "N"

  XLDTACC = "NO"
  XLDDESD = Session("Form")
  XLDFINI = "NO"
  XLDHINI = "0"
  XLDFFIN = "NO"
  XLDHFIN = "0"
  XLDUSUA = SESSION("USUARIO")
  XLDNUSU = Session("UsuarNT")
  XLDGRUP = SESSION("SEGGRP")
  XLDSESS = SESSION.SESSIONID
  XLDDIRI = Ucase(Request.ServerVariables("REMOTE_ADDR"))
  XLDLOGI = SESSION("LOGIDUSRDET")
  XLDDOMI = Session("DOMINIO")
  XLDTRUS = Session("TRUSTEDUSR")

  SELECT CASE XLTIPO

         CASE "LOGON" 
              XLOPC = "S"
              XLDTACC = "LOGON"
              XLDFINI = XLFECDAT
              XLDHINI = XLHORDAT
              XLDFFIN = XLFECDAT
              XLDHFIN = XLHORDAT


         CASE "LOGOFF" 

             

           '   StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
            '  StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDUSRDET") & "' "
           ''  response.write strsql   & "<BR>"
           '   CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)


             


  END SELECT



  IF XLOPC = "S" THEN
     


        STRSQL = "INSERT INTO OBJALOGDET ("
        STRSQL = STRSQL & "CODOBJETO,TIPOBJETO,TACC,DESDEOBJ,FINICIO,HINICIO,FFINAL,"
        STRSQL = STRSQL & "HFINAL,USUARIO,NETUSUARIO,GRUPO,SESSIONID,DIRIP,LOGID,DOMINIO,TRUSTED,USRLOGID) VALUES ("  
        STRSQL = STRSQL & "'0','SKRLOG',"
        STRSQL = STRSQL & "'" & XLDTACC & "','" & XLDDESD & "',"
        STRSQL = STRSQL & "'" & XLDFINI & "','" & XLDHINI & "',"
        STRSQL = STRSQL & "'" & XLDFFIN & "','" & XLDHFIN & "',"
        STRSQL = STRSQL & "'" & XLDUSUA & "','" & XLDNUSU & "',"
        STRSQL = STRSQL & "'" & XLDGRUP & "','" & XLDSESS & "',"
        STRSQL = STRSQL & "'" & XLDDIRI & "','" & XLDLOGI & "','" & XLDDOMI & "','" & XLDTRUS & "','" & XLDLOGI & "')"
       ''      response.write strsql   & "<BR>"
        CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)



    
  END IF


              StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
              StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDUSRDET") & "' "
           ''  response.write strsql   & "<BR>"
              CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)



END FUNCTION

FUNCTION GETUSRLOGDETTIME()

  GETUSRLOGDETTIME = XLHORDAT

 


  StrSql = "Select HINICIO From OBJALOGDET where LOGID = '" & SESSION("LOGIDUSRDET") & "' "
          '   response.write strsql   & "<BR>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if NOT Xrs.eof then

     GETUSRLOGDETTIME = XRS(0)

  END IF

  

END FUNCTION

'''' LOG DE OBJETOS ''''''''''''''''''''''''''''''''''''''''''''''''''


FUNCTION GENOBJALOG()

  StrSql = "Select * From OBJALOG where CODOBJETO = " & session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "' "
 ''       response.write strsql & "<BR>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Xrs.eof then

        STRSQL = "INSERT INTO OBJALOG (CODOBJETO,TIPOBJETO,FULTACC,AEDI,ACON,CACT,TTOTALACC,USUARIO,NETUSUARIO,DOMINIO,TRUSTED) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & session("TCliente") & "','NO','0','0','0','0','NO','NO','NO','NO')"

        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

  END IF



END FUNCTION

FUNCTION GETOBJALOGUPD(XLTIPO)

 '' RESPONSE.WRITE XLTIPO & "<BR>"
  XLOPC = "N"

  SELECT CASE XLTIPO

         CASE "ED" 
    ''  RESPONSE.WRITE XLTIPO & "0<BR>"
              XLOPC = "S"

              XLLOGIDDET = GETLOGN("LD")
              SESSION("LOGIDDET") = XLLOGIDDET
              XXXX = GETFECH()
              XXXX = GENLOGDET(XLTIPO)
              XLUPD = " SET AEDI = AEDI + 1, FULTACC = '" & XLFECDAT & "' "

         CASE "CON" 
              XLOPC = "S"

              XLLOGIDDET = GETLOGN("LD")
              SESSION("LOGIDDET") = XLLOGIDDET
              XXXX = GETFECH()
              XXXX = GENLOGDET(XLTIPO)
              XLUPD = " SET ACON = ACON + 1, FULTACC = '" & XLFECDAT & "' "

         CASE "ACT" 
              XLOPC = "S"
              XLUPD = " SET CACT = CACT + 1 "

         CASE "FIN" 
              XLOPC = "S"
              XXXX = GETFECH()
              XLHORAINI = GETLOGDETTIME()
              XLUPD = " SET TTOTALACC = TTOTALACC + " & (clng(XLHORDAT) - clng(XLHORAINI))
              XXXX = GENLOGDET(XLTIPO)


  END SELECT
     '' RESPONSE.WRITE SESSION("LOGIDDET") & "1<BR>"
  IF XLOPC = "S" THEN
     '' RESPONSE.WRITE SESSION("LOGIDDET") & "2<BR>"

     StrSql = "UPDATE OBJALOG " & XLUPD
     StrSql = StrSql & " where CODOBJETO = " & session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "' "
    ''    response.write strsql & "<BR>"
     CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)


   
  END IF



END FUNCTION



FUNCTION GENLOGDET(XLTIPO)

  XLOPC = "N"

  XLDTACC = "NO"
  XLDDESD = Session("Form")
  XLDFINI = "NO"
  XLDHINI = "0"
  XLDFFIN = "NO"
  XLDHFIN = "0"
  XLDUSUA = SESSION("USUARIO")
  XLDNUSU = Session("UsuarNT")
  XLDGRUP = SESSION("SEGGRP")
  XLDSESS = SESSION.SESSIONID
  XLDDIRI = Ucase(Request.ServerVariables("REMOTE_ADDR"))
  XLDLOGI = SESSION("LOGIDDET")
  XLDDOMI = Session("DOMINIO")
  XLDTRUS = Session("TRUSTEDUSR")
  XLDULOG = SESSION("LOGIDUSRDET")

  SELECT CASE XLTIPO

         CASE "ED" 
              XLOPC = "S"
              XLDTACC = "EDEN"
              XLDFINI = XLFECDAT
              XLDHINI = XLHORDAT
              XLDFFIN = XLFECDAT
              XLDHFIN = XLHORDAT

         CASE "CON" 
              XLOPC = "S"
              XLDTACC = "CONSEN"
              XLDFINI = XLFECDAT
              XLDHINI = XLHORDAT



         CASE "FIN" 

             

     '         StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
     '         StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDDET") & "' "
           ''  response.write strsql   & "<BR>"
     '         CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)


             


  END SELECT



  IF XLOPC = "S" THEN
     


        STRSQL = "INSERT INTO OBJALOGDET ("
        STRSQL = STRSQL & "CODOBJETO,TIPOBJETO,TACC,DESDEOBJ,FINICIO,HINICIO,FFINAL,"
        STRSQL = STRSQL & "HFINAL,USUARIO,NETUSUARIO,GRUPO,SESSIONID,DIRIP,LOGID,DOMINIO,TRUSTED,USRLOGID) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & session("TCliente") & "',"
        STRSQL = STRSQL & "'" & XLDTACC & "','" & XLDDESD & "',"
        STRSQL = STRSQL & "'" & XLDFINI & "','" & XLDHINI & "',"
        STRSQL = STRSQL & "'" & XLDFFIN & "','" & XLDHFIN & "',"
        STRSQL = STRSQL & "'" & XLDUSUA & "','" & XLDNUSU & "',"
        STRSQL = STRSQL & "'" & XLDGRUP & "','" & XLDSESS & "',"
        STRSQL = STRSQL & "'" & XLDDIRI & "','" & XLDLOGI & "','" & XLDDOMI & "','" & XLDTRUS & "','" & XLDULOG & "')"
       ''      response.write strsql   & "<BR>"
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)



    
  END IF


              StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
              StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDDET") & "' "
           ''  response.write strsql   & "<BR>"
              CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)


END FUNCTION

FUNCTION GETLOGDETTIME()

  GETLOGDETTIME = XLHORDAT

 


  StrSql = "Select HINICIO From OBJALOGDET where LOGID = '" & SESSION("LOGIDDET") & "' "
          '   response.write strsql   & "<BR>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT Xrs.eof then

     GETLOGDETTIME = XRS(0)

  END IF

  

END FUNCTION

FUNCTION GENLOGEXT(XLTIPO)

  XLOPC = "N"

  XLDTACC = "NO"
  XXXX    = GETFECH()
  XLDFINI = XLFECDAT
  XLDHINI = XLHORDAT
  XLDLOGE = GETLOGN("LE")
  XLDLOGI = SESSION("LOGIDDET")
  XLDULOG = SESSION("LOGIDUSRDET")

  SELECT CASE XLTIPO

         CASE "EDIDET" 
              XLOPC   = "S"
              XLDTACC = XLTIPO
              XLDMOBJ = Session("Form2")
              XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
              XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

              IF (XLDCOBJ & "") = "" THEN XLOPC = "N"

             '' RESPONSE.WRITE REQUEST.QUERYSTRING & "<BR>"

         CASE "ACCDET" 
              XLOPC   = "S"
              XLDTACC = XLTIPO
              XLDMOBJ = Session("Form2")
              XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
              XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")


  END SELECT



  IF XLOPC = "S" THEN


        STRSQL = "INSERT INTO OBJALOGEXT ("
        STRSQL = STRSQL & "CODOBJETO, TIPOBJETO, "
        STRSQL = STRSQL & "LOGID, TOPER, "
        STRSQL = STRSQL & "TABLA, CLAVE, "
        STRSQL = STRSQL & "FOPER, HOPER, "
        STRSQL = STRSQL & "OBSERV, LOGIDEXT, "
        STRSQL = STRSQL & "CODOBJ, MASTEROBJ, "
        STRSQL = STRSQL & "MODOBJ, USRLOGID  "
        STRSQL = STRSQL & ") VALUES ("
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & session("TCliente") & "',"
        STRSQL = STRSQL & "'" & XLDLOGI & "','" & XLDTACC & "',"
        STRSQL = STRSQL & "'NO','NO',"
        STRSQL = STRSQL & "'" & XLDFINI & "','" & XLDHINI & "',"
        STRSQL = STRSQL & "'NO','" & XLDLOGE & "',"
        STRSQL = STRSQL & "'" & XLDCOBJ & "','" & XLDMOBJ & "',"
        STRSQL = STRSQL & "'" & XLDOOBJ & "','" & XLDULOG & "')"

  set cn = ofv.conectar(ofv.strconn2)
          ''  response.write strsql   & "<BR>"
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

  call ofv.cerrarconn(cn)

    
  END IF





END FUNCTION


Function FechadorCor(wkfecha,param1)

     
     'if wkfecha & "" <> "" and wkfecha <> "&nbsp;" and wkfecha <> "..." then
     
 SELECT CASE SESSION("TIPOFECHA")
 
 
  CASE 0
     if isdate(wkfecha) then 
    
       fec = cdate(wkfecha)
       ndia = day(fec)
       nMes = Month(fec)
       anio=year(fec)
      
       if ndia < 10 then ndia="0" & ndia
       if nMes < 10 then nMes="0" & nMes
      
       FechadorCor = ndia &"/" & nmes & "/" & anio
    else  
       FechadorCor=wkfecha 
    end if
    
   Case Else
     FechadorCor=wkfecha

  END SELECT   

End Function


function  getuserDb()

    Dim oContext
    Dim myUser 
    
   
     
     
    Dirc=  Request.ServerVariables("PATH_TRANSLATED") 
    Dirc=left(Dirc,instrRev(Dirc,"\",-1))
    
    Dirc=Dirc & "incfile\Application.xml" 
    
    'Dirc="C:\contextdelivery\Application.xml"
   
   ' response.write dirc
    Set oContext = Server.createobject("ContextDeliveryCOM.contextdelivery")
    set myUser	 =server.createobject("ContextDeliveryCOM.User")
    
    
   
    oContext.SetApplicationFile cstr(Dirc)
    
    Set myUser = OContext.GetUser("AliasAplicacion", "SQLALIAS", "DESA")
   ' response.write "Usuario=" & myUser.GetUserName() & "<br>"
   ' response.write "Password=" & myUser.GetPassword() & "<br>
  
    ofv.ADbUid=myUser.GetUserName()
    
    ofv.ADbPwd =ofv.getCript("E",myUser.GetPassword()&"")
   
    set oContext = nothing
    set myUser	 = nothing

end function



Function GTESTADODSC(GTESTADCOD)

      IF CSTR(GTESTADCOD & "") = "0" OR CSTR(GTESTADCOD & "") = "" THEN

         GTESTADODSC  = "-----"

      ELSE

         Set cnXX = OFV.conectar(OFV.strconn2)
    
           XXstrsql = "SELECT ESTADO FROM ESTADOS    "
           XXstrsql = XXstrsql & " WHERE OBJID = 0" & GTESTADCOD
       '''   response.write XXstrsql & "<br>"
           set XXRS = ofv.crearconsultaEx(XXStrSql,cnXX,1,parametros)
           if NOT XXRS.eof then
              XXWKESTADO = XXRS(0) 
            ELSE
              XXWKESTADO = GTESTADCOD 
           END IF

         GTESTADODSC  = XXWKESTADO

      END IF


End Function



Function GTFechador(GTWfecha)

      IF CSTR(GTWfecha & "") = "0" OR CSTR(GTWfecha & "") = "" THEN

         GTFechador  = "-----"

      ELSE


    
         GTWfec = cdate(GTWfecha + clng(cdate("01-01-1900")))
         GTWndia = ofv.cdia(Weekday(GTWfec))
         GTWnMes = ofv.cmes(Month(GTWfec))
    
         GTwFechas = "  " & GTWndia & ", " & Day(GTWfec) & " de " & GTWnMes & " de " & Year(GTWfec)

         GTFechador  = ucase(GTwFechas)

      END IF


End Function

Function GTFechadorsh(GTWfecha)

      IF CSTR(GTWfecha & "") = "0" OR CSTR(GTWfecha & "") = "" THEN

         GTFechadorSH  = "-----"

      ELSE

    
         GTwfec = cdate(GTWfecha + clng(cdate("01-01-1900")))

         GTWFDIAB = CSTR(Day(GTwfec))
         GTWFMESB = CSTR(Month(GTwfec))

         IF LEN(GTWFDIAB) = 1 THEN GTWFDIAB = "0" & GTWFDIAB
         IF LEN(GTWFMESB) = 1 THEN GTWFMESB = "0" & GTWFMESB

         GTwFechash = GTWFDIAB & "/" & GTWFMESB & "/" & Year(GTwfec)
         GTFechadorSH  = ucase(GTwFechash)

      END IF

End Function

function GTCDSPHCALR(GTCDATEW,GTCDATEWL,GTCDATEWS)

  GTWhtm = ""

  GTWhtm = GTWhtm & " onclick="" " & chr(13)
'  GTWhtm = GTWhtm & "   alert('" & GTCDATEW & "');" & chr(13)
  GTWhtm = GTWhtm & " DXCALDATE('" & GTCDATEW & "','" & GTCDATEWL & "','" & GTCDATEWS & "') ;" & chr(13)
  GTWhtm = GTWhtm & """ " & chr(13)

  GTCDSPHCALR = GTWhtm

end function

function GTCSETCALR(GTCFRM,GTCDAT,GTCSMT,GTCFCH,GTCFCHDAT,GTCIN,GTCYRAN,GTCYRDP,GTCMNAN,GTCMNDP,GTCTDIA)

  GTWhtm = ""

  GTWAccion = "ORCALENDAR.ASP?FCIN=" & GTCIN & "&FCYRAN=" & GTCYRAN & "&FCYRDP=" & GTCYRDP & "&FCMNAN=" & GTCMNAN & "&FCMNDP=" & GTCMNDP & "&FCTDIA=" & GTCTDIA & " "

  GTWhtm = GTWhtm & " onclick="" " & chr(13)
  GTWhtm = GTWhtm & " var dxformw1 = document.getElementById('FRMCALENDAR'); " & chr(13)
  GTWhtm = GTWhtm & " dxformw1.action = '" & GTWAccion & "'; " & chr(13)
  GTWhtm = GTWhtm & " DXCDATO  = '" & GTCDAT & "';" & chr(13)
  GTWhtm = GTWhtm & " DXCFORM  = '" & GTCFRM & "';" & chr(13)

  if GTCSMT = "N" then 
     GTWhtm = GTWhtm & " DXCSMT  = 'N';" & chr(13)
  end if

  if GTCFCH = "1"  or GTCFCH = "2" then 
     GTWhtm = GTWhtm & " DXCFCH     = '" & GTCFCH & "';" & chr(13)
     GTWhtm = GTWhtm & " DXCFCHDAT  = '" & GTCFCHDAT & "';" & chr(13)
  end if



  GTWhtm = GTWhtm & " dxformw1.submit(); " & chr(13)
  GTWhtm = GTWhtm & " document.getElementById('DXCALENDAR').style.display='block';" & chr(13)
  GTWhtm = GTWhtm & " return false;" & chr(13)
  GTWhtm = GTWhtm & """ " & chr(13)

  GTCSETCALR = GTWhtm

end function

function GTCCALENDAR(GTCTRG)

  if GTCTRG = "" then 
     GTCTRG = "fesquemas"
  elseif GTCTRG = "M" then 
     GTCTRG = "fesquemas"
  elseif GTCTRG = "D" then 
     GTCTRG = "WK4"
  end if

  GTWhtm = ""


  GTWhtm = GTWhtm & "<div id='DXCALENDAR' style='margin-top:1;margin-left:1;position:absolute;left:400px;top:150px;display:none;"
  GTWhtm = GTWhtm & "width:400;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
  GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
  GTWhtm = GTWhtm & "' >"
 '''''' GTWhtm = GTWhtm & TABLA
  GTWhtm = GTWhtm & "</div>"

  GTWAccion = "XXXX"
  GTWhtm = GTWhtm & ofv.genform("FRMCALENDAR","",GTWAccion,GTCTRG,"post","","","si")



  GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & chr(13)

  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & " var DXCDATO  = '';" & chr(13)
  GTWhtm = GTWhtm & " var DXCFORM  = '';" & chr(13)
  GTWhtm = GTWhtm & " var DXCSMT   = 'S';" & chr(13)
  GTWhtm = GTWhtm & " var DXCFCH   = 'N';" & chr(13)
  GTWhtm = GTWhtm & " var DXCFCHDAT  = '';" & chr(13)
  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & "function DXCALDATE(DXCDATE,DXCDATEL,DXCDATES) {" & chr(13)
  GTWhtm = GTWhtm & "    "
  GTWhtm = GTWhtm & "   var dxformw   = document.getElementById(DXCFORM);" & chr(13)
  GTWhtm = GTWhtm & "   var dxdatow   = dxformw.all[DXCDATO];" & chr(13)
''' ' GTWhtm = GTWhtm & "   var dxdatow   = dxformw.getElementById(DXCDATO);" & chr(13)
  GTWhtm = GTWhtm & "   dxdatow.value   = DXCDATE;" & chr(13)
 ' GTWhtm = GTWhtm & "   alert(dxformw.name);" & chr(13)
 ' GTWhtm = GTWhtm & "   alert(dxdatow.value); " & chr(13)

  GTWhtm = GTWhtm & "  if (DXCFCH=='N') {" & chr(13)
  GTWhtm = GTWhtm & "     } else {" & chr(13)
  GTWhtm = GTWhtm & "      var dxcellw   = document.getElementById(DXCFCHDAT);" & chr(13)
  GTWhtm = GTWhtm & "      if (DXCFCH=='1') { " & chr(13)
  GTWhtm = GTWhtm & "          dxcellw.innerHTML=DXCDATEL;" & chr(13)
  GTWhtm = GTWhtm & "    } else {" & chr(13)
  GTWhtm = GTWhtm & "          dxcellw.innerHTML=DXCDATES;" & chr(13)
  GTWhtm = GTWhtm & "               }" & chr(13)
  GTWhtm = GTWhtm & "                   }" & chr(13)

  GTWhtm = GTWhtm & "  if (DXCSMT=='S') {dxformw.submit();}" & chr(13)
  GTWhtm = GTWhtm & "  else {document.getElementById('DXCALENDAR').style.display='none';} " & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)

  GTWhtm = GTWhtm & "</SCRIPT>"

  GTCCALENDAR = GTWhtm


end function

function GTCOBS()

  GTWhtm = ""

  GTWhtm = GTWhtm & "<div id='DX4' style='position:absolute;left:50px;top:250px;display:none;"
  GTWhtm = GTWhtm & "width:500px;height:350px;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
  GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
  GTWhtm = GTWhtm & "' >"

  GTWhtm = GTWhtm & "&nbsp;"

  GTWhtm = GTWhtm & "</div>"

  GTWhtm = GTWhtm & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>"

  GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & chr(13)
  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSWM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='800px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSSS() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='500px';" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='350px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSHM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='600px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSCL() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.display='none';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "</SCRIPT>"

  GTCOBS = GTWhtm


end function

function GTCOBSRSK()

  GTWhtm = ""

  GTWhtm = GTWhtm & "<div id='DXA' style='position:absolute;left:50px;top:250px;display:none;"
  GTWhtm = GTWhtm & "width:500px;height:350px;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
  GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
  GTWhtm = GTWhtm & "' >"

  GTWhtm = GTWhtm & "&nbsp;"

  GTWhtm = GTWhtm & "</div>"

  GTWhtm = GTWhtm & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>"

  GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & chr(13)
  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSWM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.width='800px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSSS() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.width='500px';" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.height='350px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSHM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.height='600px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSCL() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.display='none';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "</SCRIPT>"

  GTCOBSRSK = GTWhtm


end function

function GTCPTOFUN()

  GTWhtm = ""

  GTWhtm = GTWhtm & "<div id='DX4' style='position:absolute;left:50px;top:250px;display:none;"
  GTWhtm = GTWhtm & "width:500px;height:400px;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
  GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
  GTWhtm = GTWhtm & "' >"

  GTWhtm = GTWhtm & "&nbsp;"

  GTWhtm = GTWhtm & "</div>"

  GTWhtm = GTWhtm & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>"

  GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & chr(13)
  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & "function GTCPTOFUNWM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='800px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCPTOFUNSS() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='500px';" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='400px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCPTOFUNHM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='600px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "</SCRIPT>"

  GTCPTOFUN = GTWhtm


end function


%>