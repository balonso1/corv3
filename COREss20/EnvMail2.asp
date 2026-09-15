<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0
  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

    sHTML = sHTML & "<center>"



  celdas = ofv.GenCelda("","ttt colspan=2","","","","","","Envio de Mail","si")
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
    celdas = celdas & ofv.GenCelda("","uti","","","","","",from,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","","","","","Destinatario Principal:","si")
    celdas = celdas & ofv.GenCelda("","uti","","","","","",toWho,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  if len(toWho2) > 0 then

    celdas = ofv.GenCelda("","tt","","","","","","Destinatario Adicional:","si")
    celdas = celdas & ofv.GenCelda("","uti","","","","","",toWho2,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  end if

    celdas = ofv.GenCelda("","tt","","","","","","Asunto:","si")
    celdas = celdas & ofv.GenCelda("","uti","","","","","",subject,"si")
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
    celdas = celdas & ofv.GenCelda("","uti","","","","","",IMPOX,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt","","","","","","Aviso de Lectura:","si")
    celdas = celdas & ofv.GenCelda("","uti","","","","","",aviso,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt valign=top","","","","","","Mensaje:","si")
    celdas = celdas & ofv.GenCelda("","uti","","","","","",message,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut colspan=2","","","","","","&nbsp;","si")
    filas  = filas & ofv.GenRow("","","","","10","","",celdas,"si")


' Send the Mail
IF toWho <> "" THEN
  stowho =  toWho

  if len(toWho2) > 0 then  stowho = stowho & " , " & toWho2

     Xmessage = GenMDLMAIL (stowho,FROM,message)


    CALL SendMail(Xmessage, stowho, FROM,subject, session("MailComponent"),"SI",impo, aviso)

    celdas = ofv.GenCelda("","uts colspan=2 align=right","","","","","","El mensaje fue enviado","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


ELSE


    celdas = ofv.GenCelda("","uts colspan=2 align=right","","","","","","error en la direccion de mail","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
END IF

    sHTML = sHTML & ofv.GenTabla("","aaa border=1 style='border-style:solid;border-width:1;border-color:#888888;' ","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</body></html>"
end if




response.write sHTML


%>
