<%

  VCOLOR   = request.querystring("VCOLOR")
  VVUELT   = request.querystring("VVUELT") & " "

  IF VVUELT <> "" THEN
     VVUELT = REPLACE(VVUELT,"_","&")
  ELSE
     VVUELT = "CHGCOLOR.ASP"
  END IF


  IF VCOLOR <> "" THEN
     XCOLOR = REPLACE(VCOLOR,"#","")
  ELSE
     XCOLOR = "7D7D7D"
  END IF

 ''     RESPONSE.WRITE XCOLOR & "<BR>"

  DIM ATABLA(3)
  DIM XRGB(3)



  PARTE = 1

  FOR IXX=1 TO 3

       ATABLA(IXX) = MID(XCOLOR,PARTE,2) 

      PARTE = PARTE + 2


      D1 =  LEFT(ATABLA(IXX),1) 
      D2 =  RIGHT(ATABLA(IXX),1) 

      FOR IXX2=1 TO 2
          IF IXX2=1 THEN
             DD = UCASE(D1)
          ELSE
             DD = UCASE(D2)
          END IF

          SELECT CASE DD
                 CASE "A"
                      DD = 10
                 CASE "B"
                      DD = 11
                 CASE "C"
                      DD = 12
                 CASE "D"
                      DD = 13
                 CASE "E"
                      DD = 14
                 CASE "F"
                      DD = 15 
                 CASE "0","1","2","3","4","5","6","7","8","9"
                      DD = CINT(DD)                   
                 CASE ELSE
                      DD = 0
          END SELECT  
        
          IF IXX2=1 THEN
             D1 = CINT(DD) * 16
          ELSE
             D2 = CINT(DD)
          END IF           

      NEXT

      XRGB(IXX) = D1 + D2


   ''   RESPONSE.WRITE ATABLA(IXX) & " " & D1 & " " & D2 & "<BR>"

  NEXT 




%>



<html >
<script language=javascript >

function inicia()
{

document.all.i1.value = document.all['xr1'].height - 15;
document.all.i2.value = document.all['xg1'].height - 15;
document.all.i3.value = document.all['xb1'].height - 15;

document.all['xs'].style.backgroundColor = rgbahexa(document.all.i1.value,document.all.i2.value,document.all.i3.value); 
var xsi = document.all['xs'].style.backgroundColor
document.all.xsi.value = xsi.toUpperCase();

}



var hD="0123456789ABCDEF";
function d2h(d) {
var h = hD.substr(d&15,1);
while(d>15) {d>>=4;h=hD.substr(d&15,1)+h;}

if (h.length==1) { h = '0' + h;}

return h;
}
function h2d(h) {return parseInt(h,16);} 

function rgbahexa(xro,xgr,xbl)
{
var dr = d2h(xro);
var dg = d2h(xgr);
var db = d2h(xbl);



//alert(dr.toString() + dg.toString() + db.toString());

 return  dr.toString() + dg.toString() + db.toString();


}

var xact = 0
var xheig = 0

function activa()
{
 if (xact==1) { 
     xact = 0;
     }
else {
     xact = 1;
     }
}

function desactiva()
{
 xact = 0;

}

var oldp = 0

function mueveb(barra,xdir)
{

               xact = 0;
               oldp = document.all[barra].height;

               if (xdir ==1) {
                  xheig = Number(oldp) + 1;
                  
                                               }
               else {
                  xheig = Number(oldp) - 1;
                                               }
        document.all.xs2.value = barra + ' ' + xheig;

               if (xheig > 270) {xheig = 270;} 
               if (xheig < 15)   {xheig = 15;} 

               document.all[barra].height = xheig;
               inicia();


}

function mueve()
{
if (xact==1)  {
  var xtag = window.event.srcElement.tagName;
  if (xtag == 'FONT') {

      if ((event.clientY > 17 ) && (event.clientY < 287) ) {

               var xobj = window.event.srcElement;
               var barra = 'x' + xobj.id;

               oldp = document.all[barra].height;

               if (oldp >  window.event.clientY) {
                  xheig = window.event.clientY - 30;
                                               }
               else {
                  xheig = window.event.clientY - 15;
                                               }

               document.all.xs2.value = event.clientY + ' ' + barra + ' ' + xheig + ' ' + oldp ;

               if (xheig > 270) {xheig = 270;} 
               if (xheig < 15)   {xheig = 15;} 

               document.all[barra].height = xheig;
               inicia();
            }
         }
      }
}


</script>



<body  onload='inicia();' onmousedown='activa();'  onmousemove='mueve();' >
<CENTER>
<table VALIGN=MIDDLE WIDTH=70%  CELLSPACING=2 CELLPADDING=10 >
<TD VALIGN=MIDDLE ALIGN=CENTER >
<table  RULES=ALL bgcolor=#D0D0D0 STYLE='border-style:OUTSET;border-width:2;border-color:#FFFFFF;' >
<tr height=280 >
<td  width=70% align=center ROWSPAN=2  >
    <table  id=xs  width=150 height=80 style='background-color:#eeeeee;border-style:outset;border-width:2;' ><tr><td>&nbsp;</td></tr></table>
    <br>      
    <FORM ID=FRMCOL ACTION='<% RESPONSE.WRITE VVUELT %>' TARGET=_self METHOD=POST >      
    <input id=xsi NAME=xsi type=text READONLY value=' ' size=12 style='height:25;font-family:verdana;font-size:9pt;font-weight:bold;color:#336699;border-style:ridge;border-width:2;border-color:#eeeeee;text-align:center;background-color:#ffffff;' >

    <br>            
    <input id=xs2 NAME=xs2 type=HIDDEN value=' ' size=40 style='font-family:tahoma;font-size:8pt;font-color:#336699;border-style:solid;border-width:1;border-color:#dddddd;text-align:center;background-color:#ffffff;' >

    <br> 

    <input type=SUBMIT value='ACEPTAR' size=40 style='font-family:VERDANA;font-size:8pt;color:#FFFFFF;text-align:center;background-color:#336699;' >

    </FORM>
</td>  

<td  width=10% align=center COLSPAN=2 >
              <table bgcolor=#EEEEEE id=xr  width=15 height=270 cellspacing=0 cellpadding=0 style='border-style:outset;border-width:1;' >
                 <tr id=xr1 height=<% RESPONSE.WRITE(XRGB(1) + 15) %> ><td bgcolor=#bb0000 valign=bottom >

<a ID=r1 href='javascript:void(0);'  title='modificar rojo' style='cursor:hand;background-color:#550000;' >
<font  ID=r1  face=webdings size=2 color=#ffffff>6</font>
</a>
                                 </td></tr>
                <tr ><td style='border-style:inset;border-width:1;' ></td></tr>
              </table>
</td>  

<td  width=10% align=center COLSPAN=2  >
              <table bgcolor=#EEEEEE id=xg  width=15 height=270 cellspacing=0 cellpadding=0 style='border-style:outset;border-width:1;' >
                 <tr  id=xg1 height=<% RESPONSE.WRITE(XRGB(2) + 15) %> ><td bgcolor=#00aa00 valign=bottom >

<a ID=g1 href='javascript:void(0);'  title='modificar verde' style='cursor:hand;background-color:#005500;'  >
<font ID=g1 face=webdings size=2 color=#ffffff>6</font>
</a>
                                 </td></tr>
                <tr ><td style='border-style:inset;border-width:1;' ></td></tr>
              </table>
</td>  

<td  width=10% align=center COLSPAN=2 >
              <table bgcolor=#EEEEEE id=xb  width=15 height=270 cellspacing=0 cellpadding=0 style='border-style:outset;border-width:1;' >
                 <tr  id=xb1 height=<% RESPONSE.WRITE(XRGB(3) + 15) %> ><td bgcolor=#0000aa valign=bottom >

<a ID=b1  href='javascript:void(0);'  title='modificar azul' style='cursor:hand;background-color:#000055;' >
<font  ID=b1 face=webdings size=2 color=#ffffff>6</font>
</a>
                                 </td></tr>
                <tr ><td style='border-style:inset;border-width:1;' ></td></tr>
              </table>
</td>  

</tr>

<tr>

<td  align=center  >
    <input id=Bi11 type=BUTTON  value='+' onclick="mueveb('xr1',1);" style='font-family:tahoma;font-size:7pt;color:#000000;background-color:#EEEEEE;' >
</td>
<td  align=center  >
    <input id=Bi12 type=BUTTON  value='-' onclick="mueveb('xr1',2);"  style='font-family:tahoma;font-size:7pt;color:#000000;background-color:#EEEEEE;' >
</td>

<td  align=center  >
    <input id=Bi21 type=BUTTON  value='+' onclick="mueveb('xg1',1);"  style='font-family:tahoma;font-size:7pt;color:#000000;background-color:#EEEEEE;' >
</td>
<td  align=center  >
    <input id=Bi22 type=BUTTON  value='-' onclick="mueveb('xg1',2);"  style='font-family:tahoma;font-size:7pt;color:#000000;background-color:#EEEEEE;' >
</td>

<td  align=center  >
    <input id=Bi31 type=BUTTON  value='+' onclick="mueveb('xb1',1);"  style='font-family:tahoma;font-size:7pt;color:#000000;background-color:#EEEEEE;' >
</td>
<td  align=center  >
    <input id=Bi32 type=BUTTON  value='-' onclick="mueveb('xb1',2);"  style='font-family:tahoma;font-size:7pt;color:#000000;background-color:#EEEEEE;' >
</td>

</tr>

<tr>
<td align=center  ><B><font size=2 face=verdana color=#336699>PALETA DE COLORES</FONT></B>
    
</td>
<td  align=center COLSPAN=2 >
    <input id=i1 type=text READONLY value=red size=2 style='font-family:tahoma;font-size:8pt;color:#336699;border-style:outset;border-width:1;border-color:#dddddd;text-align:center;background-color:#ffffff;' >
</td>
<td  align=center COLSPAN=2  >
    <input id=i2 type=text READONLY value=green size=2 style='font-family:tahoma;font-size:8pt;color:#336699;border-style:outset;border-width:1;border-color:#dddddd;text-align:center;background-color:#ffffff;' >
</td>
<td  align=center COLSPAN=2  >
    <input id=i3 type=text READONLY value=blue size=2 style='font-family:tahoma;font-size:8pt;color:#336699;border-style:outset;border-width:1;border-color:#dddddd;text-align:center;background-color:#ffffff;' >
</td>
</tr>
</table>

</td>
</table>

</body>
</html>