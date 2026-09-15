<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  shtml = ofv.MenuHeader("#eeeeee"," bacKground='imagenes/corf5.jpg' ")



  set cn = ofv.conectar(ofv.strconn1)

  Strsql = "Select * from CorConsulta"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then
     xLicencia = rs(2)
     xEmpresa = rs(3)
  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "<center>"

  sHTML = sHTML & "<table    width=650 marginwidth=4  cellspacing=2 cellpadding=2>"

  sHTML = sHTML & "<tr ><td align=center style='border-style:outset;border-width:1;' "
  sHTML = sHTML & " valign=middle >"
  sHTML = sHTML & "<img src=imagenes/COREssuite20.jpg border=0 "
  sHTML = sHTML & " >"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td   width=35%   valign=middle >"
  sHTML = sHTML & "<font face=verdana color=#000088 size=2 ><b>"
  sHTML = sHTML & "COR <font face=helvetica color=#f75500 size=4><i>e</i></font>Solution Suite 2.6</b></font>" 
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td   width=65%  valign=top >"
  sHTML = sHTML & "<font face=verdana color=#000088 size=1 ><b><br>Licencia:&nbsp;&nbsp;&nbsp;"
  sHTML = sHTML & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>" & xLicencia
  sHTML = sHTML & "</strong></b></font></td></tr>"

  sHTML = sHTML & "<tr><td align=center >&nbsp;</td><td align=center >&nbsp;</td>"
  sHTML = sHTML & "<td   width=65% ><font face=verdana color=#000088 size=1 ><b>"
  sHTML = sHTML & "Otorgada a:&nbsp; <strong>" & xEmpresa 
  sHTML = sHTML & "</strong><br></b></font></td>" 

  sHTML = sHTML & "</tr></table><br>"



  Set fso = CreateObject("Scripting.FileSystemObject")
  sPath = server.mappath("../coressdll")
'  response.write spath & "<br>"
  Set Folder = fso.getfolder(sPath)
  Set Files = Folder.getFiles(sPath,"s*.*")

  sHTML = sHTML & ofv.GenTabla("","aaa border","","650","","0","0","0","0","#f7eedd","","","no")
  sHTML = sHTML & "<font color=#DCDCDC face=tahoma size=1><b>"
 h1 = ofv.GenCelda("","tt","","","","","","Componente","si")
 h2 = ofv.GenCelda("","tt","","","","","","Modulo","si")
 h3 = ofv.GenCelda("","tt","","","","","","Fecha de Version","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1 & H2 & H3,"si")
  sHTML = sHTML & "</b></font>"

  For Each F in Files

     xVersion = ""

  '   if len(f.type) > 3 then
   '     if ucase(left(f.type,3)) = "APP" or ucase(left(f.type,3)) = "EXT" then  
           Set ver = nothing
           on error resume next 
       '    xnombre = left(f.name,instr(Ucase(f.name),".DLL") - 1)
        '   Set ver = CreateObject(xnombre & ".about")
       '    xVersion = "<img src='imagenes/corpnt2.jpg' border=0 align=middle > " & ver.version 
     '      if xVersion <> "" then
              h1 = ofv.GenCelda("","cc","","","","","",xVersion,"si")
              h2 = ofv.GenCelda("","cc","","","","","",f.name,"si")
              h3 = ofv.GenCelda("","cc","","","","","",f.datecreated,"si")
              sHTML = sHTML & ofv.GenRow("","","","","","","",h1 & H2 & H3,"si")
       '    end if
      '  end if
   '  end if
  Next

 xAviso = "Cor ESolution Suite&reg; es un producto registrado por Cor Consulting Asociados SRL."  
 xAviso = xAviso & "<br> Cor ESolution Suite&reg; y todos sus componentes son propiedad de Cor Consulting Asociados SRL."  
 xAviso = xAviso & "<br> &copy;1999 - 2001"

    sHTML = sHTML & ofv.GenFTabla()

    sHTML = sHTML & "<table width=650 marginwidth=2 cellspacing=2 cellpadding=2>"
    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td width=35% align=left rowspan=2><br><p align=center>"
    sHTML = sHTML & "<img src=imagenes/CorLogox1.jpg border=0 width=22 height=22>"
    sHTML = sHTML & "<br>"
    sHTML = sHTML & "<font face=helvetica,tahoma color=#000099 size=2 ><b>"
    sHTML = sHTML & "Cor Consulting<br>Asociados</b></font>" 
    sHTML = sHTML & "</p></td>"
    sHTML = sHTML & "<td  width=65% align=left colspan=2><font color=#336699>"
    sHTML = sHTML & xaviso & "</font></td>"
    sHTML = sHTML & "</tr>"
    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td  align=left >"
    sHTML = sHTML & "<img src=imagenes/reparg.jpg border=0  ></td>"
    sHTML = sHTML & "<td width=65% align=left >"
    sHTML = sHTML & "<font color=#336699>&nbsp;Producto desarrollado en Argentina</font>"
    sHTML = sHTML & "</td></tr>"
    sHTML = sHTML & "</table>"

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>