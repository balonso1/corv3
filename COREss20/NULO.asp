<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires = 0


  Xbody = ofv.wbody & "  onresize='uu();'   "

    wmnnoved = "NOVED" & WKAPLIC 

 '''  session(wmnnoved) = "SI"



if (WKAPLIC = "EM" OR WKAPLIC = "OR" OR WKAPLIC = "PL") and session(wmnnoved) <> "SI" then

   session("FormN") = "NOVEDADES"
  ' response.redirect Session("CorPath") & "novedmenu.asp"
   
   xshtml = "<html><body ><form name='ejec' method='post' action='" & Session("CorPath") & "novedmenu.asp' style='visibility:hidden;'><input type='submit' value='n'></form>"
   xshtml = xshtml & "<script language='javascript'>document.ejec.submit();</script></body></html>"
             
   Response.write xshtml

elseif Session("CorApli") = "CLI916953113116518847748" and Session("novedgp") <> "SI" then

   session("FormN") = "NOVEDADES"
  ' response.redirect Session("CorPath") & "SEGNOVEDCIRCUITO.asp"
   
   xshtml = "<html><body ><form name='ejec' method='post' action='" & Session("CorPath") & "SEGNOVEDCIRCUITO.asp' style='visibility:hidden;'><input type='submit' value='n'></form>"
   xshtml = xshtml & "<script language='javascript'>document.ejec.submit();</script></body></html>"
             
   Response.write xshtml

else






if SESSION("FULLGUI") = "S"  then

  sHTML = replace(ofv.MenuHeader("",""),"<body ","<body " & Xbody & " ")  & chr(13)



sHTML = sHTML & "<script language=JavaScript>"
sHTML = sHTML & "function Init () "
sHTML = sHTML & "   {  }"
sHTML = sHTML & "function uu()"
sHTML = sHTML & " {  MT=1;z1=0;z2=0; "
sHTML = sHTML & " var xah2 = document.body.clientHeight  - 10; "
sHTML = sHTML & " var xaw2 = document.body.clientWidth   - 10; "


sHTML = sHTML & " document.all['L0'].style.left = 80; "
sHTML = sHTML & " document.all['L0'].style.top  = xah2  - 80; "

sHTML = sHTML & " document.all['L02'].style.left = 80 + 70 + 90; "
sHTML = sHTML & " document.all['L02'].style.top  = xah2  - 60; "

sHTML = sHTML & " document.all['L1'].style.left = xaw2  - document.all['L1'].clientWidth; "
sHTML = sHTML & " document.all['L1'].style.top  = xah2  - 60; "

sHTML = sHTML & " document.all['L2'].style.left = xaw2  - document.all['L1'].clientWidth; "
sHTML = sHTML & " document.all['L2'].style.top  = xah2  - 30; "


sHTML = sHTML & "  }"
sHTML = sHTML & "function ff() {"
sHTML = sHTML & "   } "
sHTML = sHTML & "</script>"

            Set cx = OFV.conectar(OFV.strconn1)
            StrSql = "select * from CorAplic where codapli = '" & Session("CorApli") & "' "
            Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
            if not rx.eof then
               saplic = rx(12)
            end if

            Set cx = OFV.conectar(OFV.strconn1)
            StrSql = "select * from CorAplic where codapli = 'CORESS20' "
            Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
            if not rx.eof then
               saplicM = rx(12)
            end if




%>



<%=shtml%>




<v:rect id='L0' style="position:absolute;top:400;left:200;width:60;height:40;" > 
  <v:fill on="True" type="frame" src="imagenes/logocorP3sm.jpg" opacity="95%"  /> 
  <v:extrusion on="True" backdepth="-8" color="#FFFFFF" color2="#99ccee" metal="true"   rotationcenter="0,0,0"  
   rotationangle = "0,0" opacity="10%" />
</v:rect> 


 <v:line id='L02' from='0,0' to='1,0' style='z-index:1;position:absolute;left:435px;top:485px;width:80px;height:20px;' > 
   <v:fill on='True' color='#990000' color2='#bb0000' method='sigma' type='gradient' angle='180' opacity='95%'  /> 
   <v:stroke on='False' /> 
   <v:shadow on='true' color='#eeeeee' weight='5px' /> 
   <v:path textpathok='True' /> 
     <v:textpath on='True' string='Cor Consulting' style="font-size:'20pt';font-family:'helvetica';font-weight:bold;"  /> 
   <v:extrusion on='True' backdepth='-8' color='#eeeeee' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> 
 </v:line>






 <v:line id='L1' from='0,0' to='1,0' style='z-index:1;position:absolute;left:135px;top:385px;width:80px;height:20px;' > 

   <v:fill on='True' color='#6699cc' color2='#99ccee' method='sigma' type='gradient' angle='180' opacity='95%'  /> 
   <v:stroke on='False' /> 
   <v:shadow on='False' color='#ffffff' weight='15px' /> 
   <v:path textpathok='True' /> 
     <v:textpath on='True' string='<%=saplicM%>' style="font-size:'18pt';font-family:'tahoma';font-style:italic;font-weight:bold;"  /> 
   <v:extrusion on='True' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> 
 </v:line>

 <a class=nt href='http://www.corconsulting.com.ar' target='_blank' >
 <v:line id='L2' from='0,0' to='1,0' style='z-index:1;position:absolute;left:135px;top:385px;width:80px;height:20px;' > 
   <v:fill on='True' color='#6699cc' color2='#6699cc' method='sigma' type='gradient' angle='180' opacity='95%'  /> 
   <v:stroke on='trueFalse' color='#ffffff' weight='5px'/> 
   <v:shadow on='true' color='#ffffff' weight='5px' /> 
   <v:path textpathok='True' /> 
     <v:textpath on='True' string='www.corconsultingasociados.com' style="font-size:'10pt';font-family:'garamont';font-weight:normal;"  /> 
   <v:extrusion on='true' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> 
 </v:line>
 </a>



<%

if ofv.explorador = "MSIE" then
  response.write "<script language='javascript' > uu();Init();submenu(); </script>"
else
  response.write "<script language='javascript' > uu();submenu(); </script>"
end if


else

  sHTML = ofv.MenuHeader("","") & chr(13)


'sHTML = "<html" & sHTML & " <BODY " & ofv.wbody & "  ><center>"

sImagen = "<img id=im1 name=im1 src=imagenes/COREssuite20.jpg border=0 >"

Celda = ofv.GenCelda("","","center","","","","",sImagen,"si")
Fila =  ofv.GenRow("","","","","","","",Celda,"si")

sHTML = sHTML & "<br><br><br><br><br><br><br>" & ofv.GenTabla("","","","","","","","-2","-1","","",Fila,"si")





sInput = ""
sInput = sInput & "<font face=verdana color=#000088 size=2>"
sInput = sInput & "<b>COR <font face=helvetica color=#f75500 size=4>"
sInput = sInput & "<i>e</i></font>Solution Suite </b></font>"
Celda = ofv.GenCelda("","","center","","","","",sInput,"si")
Fila =  ofv.GenRow("","","","","","","",Celda,"si")

            Set cx = OFV.conectar(OFV.strconn1)
            StrSql = "select * from CorAplic where codapli = '" & Session("CorApli") & "' "
            Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
            if not rx.eof then
               saplic = rx(12)
            end if

sInput = "<br><br>"
Celda = ofv.GenCelda("","","center","","","","",sInput,"si")
Fila =  fila & ofv.GenRow("","","","","","","",Celda,"si")

sInput = ""
sInput = sInput & "<font face=verdana color=#000088 size=3>"
sInput = sInput & "<b>" & saplic & "</b></font>"
Celda = ofv.GenCelda("","cc style='border-style:outset;border-width:2;' ","center","","","","",sInput,"si")
Fila =  fila & ofv.GenRow("","","","","","","",Celda,"si")


sHTML = sHTML & "<br><br>" & ofv.GenTabla("","","","50%","","4","","0","0","","",Fila,"si")
sHTML = sHTML & "</center>"

response.write shtml

if ofv.explorador = "MSIE" then
  response.write "<script language='javascript' > submenu(); </script>"
else
  response.write "<script language='javascript' > submenu(); </script>"
end if

end if





response.write "</body></html>"




end if

%>