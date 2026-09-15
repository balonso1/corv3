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






  Set fso = CreateObject("Scripting.FileSystemObject")

  sPath = Request.ServerVariables("PATH_INFO")
  response.write spath & "<br>"
  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1) - 1)
 response.write Xruta & "<br>"
 sPath = server.mappath(sPath)
  response.write spath & "<br>"
    sPath = replace(sPath,"\cordscr.asp","") 
  Set Folder = fso.getfolder(sPath)
  set sfolder = Folder.SubFolders
  Set Files = Folder.Files

  sHTML = sHTML & ofv.GenTabla("","aaa border","","650","","0","0","0","0","#f7eedd","","","no")
  sHTML = sHTML & "<font color=#DCDCDC face=tahoma size=1><b>"
 h1 = ofv.GenCelda("","tt","","","","","","descripcion","si")
 h2 = ofv.GenCelda("","tt","","","","","","tipo","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1 & H2,"si")
  sHTML = sHTML & "</b></font>"

  For Each F in sfolder




              h1 = ofv.GenCelda("","cc","","","","","",f.name,"si")
              h2 = ofv.GenCelda("","cc","","","","","","dir","si")
              sHTML = sHTML & ofv.GenRow("","","","","","","",h1 & H2,"si")

  Next

  For Each F in Files

     xVersion = ""



              h1 = ofv.GenCelda("","cc","","","","","",f.name,"si")
              h2 = ofv.GenCelda("","cc","","","","","",f.datecreated,"si")
              sHTML = sHTML & ofv.GenRow("","","","","","","",h1 & H2,"si")

  Next



    sHTML = sHTML & ofv.GenFTabla()



  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>