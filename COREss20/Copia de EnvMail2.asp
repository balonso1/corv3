<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0
  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

    sHTML = sHTML & "<center>"



  celdas = ofv.GenCelda("","cc colspan=2","","","","","","Envio de Mail","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")


'' Grab Form Variables
from   = TRIM( Request( "from" ) )
impo   = cint( Request( "impo" ) )
aviso   = TRIM( Request( "aviso" ) )
toWho   = TRIM( Request( "toWho" ) )
toWho2   = TRIM( Request( "toWho2" ) )
subject = TRIM( Request( "subject" ) )
message = TRIM( Request( "message" ) )


    celdas = ofv.GenCelda("","tt","","","","","","Emisor:","si")
    celdas = celdas & ofv.GenCelda("","cc","","","","","",from,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","","","","","Destinatario Principal:","si")
    celdas = celdas & ofv.GenCelda("","cc","","","","","",toWho,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  if len(toWho2) > 0 then

    celdas = ofv.GenCelda("","tt","","","","","","Destinatario Adicional:","si")
    celdas = celdas & ofv.GenCelda("","cc","","","","","",toWho2,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  end if

    celdas = ofv.GenCelda("","tt","","","","","","Asunto:","si")
    celdas = celdas & ofv.GenCelda("","cc","","","","","",subject,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    select case impo
           case 0
                  impox = "BAJA"

           case 1
                  impox = "NORMAL"

           case 2
                  impox = "ALTA"

    end select 

    celdas = ofv.GenCelda("","tt","","","","","","Importancia:","si")
    celdas = celdas & ofv.GenCelda("","cc","","","","","",IMPOX,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","","","","","Aviso de Lectura:","si")
    celdas = celdas & ofv.GenCelda("","cc","","","","","",aviso,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt valign=top","","","","","","Mensaje:","si")
    celdas = celdas & ofv.GenCelda("","cc","","","","","",message,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","cc colspan=2","","","","","","&nbsp;","si")
    filas  = filas & ofv.GenRow("","","","","10","","",celdas,"si")


' Send the Mail
IF toWho <> "" THEN
  stowho =  toWho

  if len(toWho2) > 0 then  stowho = stowho & " , " & toWho2
  Set myMail = Server.CreateObject( "CDONTS.Newmail" )

'  myMail.From = "CorESolutionSuite@hotmail.com"

    HTML = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">"
    HTML = HTML & "<html>"
    HTML = HTML & "<head>" 
    HTML = HTML & "<title>" & subject & "</title>"
    HTML = HTML & "</head>"
    HTML = HTML & "<body bgcolor=""FFFFFF"">"
    HTML = HTML & message

    HTML = HTML & "</body>"
    HTML = HTML & "</html>"

  myMail.BodyFormat=0
  myMail.MailFormat=0
  myMail.From = from
  myMail.To = stowho
  myMail.Subject = subject
  myMail.Body = HTML

  myMail.importance = impo
  if aviso = "SI" then
'     strReply_To = "Disposition-Notification-To<" & from & ">"
     strReply_To = "<" & from & ">"
     myMail.Value("Disposition-Notification-To") = strReply_To
  end if
  myMail.Send
  Set myMail = Nothing



    celdas = ofv.GenCelda("","tt colspan=2 align=right","","","","","","El mensaje fue enviado","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


ELSE


    celdas = ofv.GenCelda("","tt colspan=2 align=right","","","","","","error en la direccion de mail","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
END IF

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</body></html>"
end if




response.write sHTML


%>
