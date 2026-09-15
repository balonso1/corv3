<html>
<head> 
<META http-equiv="refresh" content="10;URL=refrescar.asp">
</head> 

<body bgcolor="#ffffff"> 
    <%

function GETFECH()



    sh = datepart("h",TIME())
    sm = datepart("n",TIME())
    ss = datepart("s",TIME())

   GETFECH = (CINT(sh)*3600) + (CINT(sm)*60) + CINT(ss)

END FUNCTION

    XAHORA = GETFECH()
  
    XEXPIRA =  XAHORA - SESSION("XEXPIR")

    IF SESSION("USREXPIR") = "D" THEN
      ' RESPONSE.REDIRECT "NULO.HTM"
       
       xshtml = "<form name='ejec' method='post' action='corbase.htm' ><input type='submit' value='n'></form>"
       xshtml = xshtml & "<script language='javascript'>document.ejec.submit();</script>"
          
       Response.write xshtml      
       
    ELSEIF  SESSION("USREXPIR") > XEXPIRA THEN
            response.write SESSION("XEXPIR") & " - " & XEXPIRA & " / " & SESSION("USREXPIR") & "<BR>"
    ELSE

    SHTML = "<FORM name='LANZA'  TARGET='_top' ACTION='APLEXPIR.HTM' >"
    SHTML = SHTML & "<INPUT TYPE='SUBMIT' VALUE='INICIAR'>"
    SHTML = SHTML & "</FORM>"

    SHTML = SHTML & "<script language='javascript' > document.all.LANZA.submit();</script>"

       RESPONSE.write SHTML

    END IF
    %>
</body> 
</html> 
