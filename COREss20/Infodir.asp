<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

  response.expires=0

  Directorio = request.querystring("dir")

  shtml = "<html><Title>COR EManager</title>" & ofv.formstyler() & "<BODY " & ofv.wbody & ">"

  shtml = shtml & "<br><table width=100% cellspacing=0 cellpadding=0><tr><td>&nbsp;&nbsp;</td>"
  shtml = shtml & "</tr>"

  Dim fso,Folder,Files,F, s, n

  Set fso = CreateObject("Scripting.FileSystemObject")
  if Directorio = "" then
    sPath = server.mappath("1")
    sPath = left(sPath,len(sPath)-2) & Directorio
  else
    sPath = "C:\Inetpub\wwwroot\" & Directorio
  end if

  Set Folder = fso.getfolder(sPath)
  Set Files = Folder.Files
  
  redim s(0)
  n = 1
  For Each F in Files
    shtml = shtml & "<tr ><td>" & n & " - " & F.name & "</td></tr>"
    n = n + 1 
  Next
  shtml = shtml & "</table>"

  shtml = shtml & "</BODY></html>"

response.write shtml

%>