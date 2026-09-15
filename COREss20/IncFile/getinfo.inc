<%    

dim arrr (1000,1)



' esta url es donde se encuentra el fileserver con los asp para leer directorios ,archivos, y leer files
session("UrlGetdata")="http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & "/coresv3/FileServer/asp/Getdata.asp"
' esto es para el upload
session ("DirUpload")="http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & "/coresv3/Fileserver/"
' este el el sito web origen
session ("SiteCor")= "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT")



function getdata( XARCHIVO,op,folder,NEWFOLDER)

 Response.Charset="ISO-8859-1"

 ruta=server.mappath(request.servervariables("script_name"))
 
 pos=instrrev(ruta,"\")
 ruta=mid(ruta,1,pos)
 

' response.write ruta & folder & OP
 Set fso = SERVER.CreateObject("Scripting.FileSystemObject")
 Set xFolder = fso.getfolder(ruta  & folder)
 
 
  
 select case OP
 
 case "RF"
         'response.write ruta & folder & OP
        set sfolder = xFolder.SubFolders
  
          For Each F in sfolder
            getdata=getdata &  f.name & "|" '& "<BR>"
          next
          

case "AF"      
       
       set sfolder = xFolder.SubFolders
       sFolder.add(NEWFOLDER)
       
          
 case "RFILES"      
       
       Set Files = xFolder.Files
         
          For Each F in files
          '   response.write  f.name &"###" & f.DateCreated  & "|"
            getdata=getdata &  f.name &"###" & f.DateCreated  & "|"
          next
          
 CASE "F"         
 
       Xarchivo=RUTA & "CARPETASCMD\" & XARCHIVO
      ' RESPONSE.WRITE XARCHIVO
       if fso.FileExists(trim(Xarchivo)) then
              Set fff1 = fso.openTextFile(trim(Xarchivo), 1, -1)
              do while fff1.AtEndOfstream <> true
                 getdata=getdata & fff1.readline
              loop
              fff1.close
      end if 

 CASE "F1"         
 

       if fso.FileExists(trim(Xarchivo)) then
              Set fff1 = fso.openTextFile(trim(Xarchivo), 1, -1)
              do while fff1.AtEndOfstream <> true
                 getdata=getdata & fff1.readline
              loop
              fff1.close
      end if 
    
 End Select         
end function


function callUrl(ROOT,url)
  Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP")
   

   'url=session("UrlGetdata") &"?OP=RFILES&folder=" & root
   xmlHttp.Open "Get", url, false
   xmlHttp.Send()
   ss=xmlHttp.ResponseText
   callUrl = ss
  
  
end function


function getfolders(root,url)



select CASE SESSION("FILETIPO")

case "R"
  Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP")
   url=session("UrlGetdata") & "?OP=RF&folder=" & root
  ' response.write url
   xmlHttp.Open "Get", url, false
   xmlHttp.Send()
   ss=xmlHttp.ResponseText
   getfolders = ss

case "L"
'function getdata( XARCHIVO,op,folder,NEWFOLDER)
 getfolders=getdata("","RF",ROOT,"")
end select

  
end function

function getfilesinfolder(ROOT,url)

SELECT CASE SESSION("FILETIPO")

CASE  "R"

  Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP")
   

   url=session("UrlGetdata") &"?OP=RFILES&folder=" & root
   xmlHttp.Open "Get", url, false
   xmlHttp.Send()
   ss=xmlHttp.ResponseText
   getfilesinfolder = ss

CASE "L"
  getfilesinfolder = getdata("","RFILES",root,"")
END SELECT


end function

function getReadfile(ROOT,url)


SELECT CASE SESSION("FILETIPO")

CASE  "R"

  Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP")
   

   url=session("UrlGetdata") &"?OP=F&xarchivo=" & root 
   'response.write url
   xmlHttp.Open "Get", url, false
   xmlHttp.Send()
   ss=xmlHttp.ResponseText
   getReadfile = ss '& session("sitecor")

CASE "L"
   
   getReadfile=getdata(root,"F","","")
END SELECT
  
  
end function


function getReadfile2(ROOT,url)

SELECT CASE SESSION("FILETIPO")

CASE  "R"

  Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP")
   

   url2=session("UrlGetdata") &"?OP=F1&xarchivo=" & root & "&newfolder="& url
   'response.write url
   xmlHttp.Open "Get", url2, false
   xmlHttp.Send()
   ss=xmlHttp.ResponseText
   getReadfile2 = ss '& session("sitecor")

CASE "L"
   
   getReadfile2=getdata(root,"F1","","")
END SELECT
  
  
end function



FUNCTION ADDFOLDER(root,newfolder,url)

Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP")
   url=session("UrlGetdata") & "?OP=AF&folder=" & root & "&NEWFOLDER="&NEWFOLDER	
   xmlHttp.Open "Get", url, false
   xmlHttp.Send()
   'getfilesinfolder = xmlHttp.ResponseText
   
end function


function UserLogon()
   User=REQUEST.servervariables("LOGON_USER")
   User=mid(User,instr(1,User,"\")+1)
   UserLogon=User
   ' UserLogon="JGUZMAN"
    'UserLogon="DVARA"
end function


function UserAuth()
 User=REQUEST.servervariables("AUTH_USER")
 User=mid(User,instr(1,User,"\")+1)
 UserAuth=User
end function


Function getInfoECO(XGU_ECO)
data = ""  

GTUcase = session("ECOPRV")

IF GTUCASE = "" THEN GTUCASE = "I" 


  select case GTUcase

       case "I"

           IF ISNUMERIC(XGU_ECO) THEN
              WXGU_ECO = XGU_ECO
           ELSE
              WXGU_ECO = 1
           END IF 

           set xcn = ofv.conectar(ofv.strconn2)

           xStrSql = "Select ECO,DESCRI,CODINTER  From ECO WHERE ECO = 0" & WXGU_ECO
           Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
           IF NOT xRS.EOF THEN
     		session("SGU_SECO")      = xrs(0)
  		session("SGU_SDESCRI")   = xrs(1)
  		session("SGU_SCODINT")	 = xrs(2) & ""
           ELSE
     		session("SGU_SECO")      = XGU_ECO
  		session("SGU_SDESCRI")   = "NO"
  		session("SGU_SCODINT")	 = "NO"

           END IF

           call ofv.cerrarconsulta(xrs)
           call ofv.cerrarconn(xcn)


       case "E"

     		session("SGU_SECO")      = XGU_ECO
  		session("SGU_SDESCRI")   = "NO"
  		session("SGU_SCODINT")	 = "NO"



   end select

End Function   


'--------------------------------------------------------------------------------------------------



Function getInfoUser(Legajo)
data = ""  

GTUcase = session("USRPRV")

  select case GTUcase

       case "I"

           set xcn = ofv.conectar(ofv.strconn2)

           xStrSql = "Select ID,descripcion,email,sector, puesto  From USUARIOS WHERE ID = '" & Legajo & "' "
           Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
           IF NOT xRS.EOF THEN
     		session("SGU_userid")      = xrs(0)
  		session("SGU_nombre")      = xrs(1)
  		session("SGU_mail")	   = xrs(2) & ""
  		session("SGU_descrisector")= "MySect"
  		session("SGU_idcostcenter")= xrs(3)
  		session("SGU_descripuesto")= "MyPst"
  		session("SGU_codigopuesto")= xrs(4)
           ELSE
     		session("SGU_userid")      = Legajo
  		session("SGU_nombre")      = "NO"
  		session("SGU_mail")	   = "NO"
  		session("SGU_descrisector")= "NO"
  		session("SGU_idcostcenter")= "0"
  		session("SGU_descripuesto")= "NO"
  		session("SGU_codigopuesto")= "NO"
           END IF

           call ofv.cerrarconsulta(xrs)
           call ofv.cerrarconn(xcn)


       case "E"


  ' DATOS DE LOS WEBSERVICES
  SoapAction = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getUserWithoutManagers"
  Method = "getUserWithoutManagers" ' METODO DEL WEBSERVICE A UTILIZAR
  
  END_POINT_URL = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
  CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"

  
  'Legajo = request.querystring("Legajo")
  if Legajo<>"" then
  	Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
  	Connector.Property("EndPointURL") = END_POINT_URL
  	Connector.Connect
  	
  	Connector.Property("SoapAction") = SoapAction & Method
  	Connector.BeginMessage
  	
  	Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
  	Serializer.Init Connector.InputStream
  	
  	Serializer.StartEnvelope
  	Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
  	Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
  	Serializer.StartBody
  	Serializer.StartElement Method, CALC_NS
  	Serializer.StartElement "userId"
  	Serializer.WriteString  Legajo  ' LEGAJO DEL USUARIO A BUSCAR
  	Serializer.EndElement
  	Serializer.EndElement
  	Serializer.EndBody
  	Serializer.EndEnvelope
  	
  	Connector.EndMessage
  	    
  	Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
  	Reader.Load Connector.OutputStream
  	' OBTENCION DE LOS DATOS
  	
  	
  	Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
  	Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text 
  	Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text 
  	
  	set xml=Reader.Body
  	set raiz=xml.selectSingleNode("//getUserWithoutManagersResponse/getUserWithoutManagersReturn")
	   	   
	id=replace(raiz.getAttribute("href"),"#","")


	   ' tomo el idref de la persona
	    Set nList = xml.selectNodes("//multiRef[@id='"&ID &"']/person")

	    For Each n In nList
	     idpersona = replace(n.getAttribute("href"),"#","")
	    Next

	   ' tomo el idref del costcenter
	    Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/costCenter")

	    For Each n In nList
	     idcostCenter = replace(n.getAttribute("href"),"#","")
	    Next

  	
  	 ' tomo el descripcion del costcenter /sector
	      Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterDescription")
	          For Each n In nList
	            descrisector= n.text ' replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
	       Next
	       
	 ' tomo el id del costcenter
	   Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterCode")
	       For Each n In nList
	         idcostcenter= n.text 'replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
	       Next
       
        ' tomo el id del puesto
          Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/functionTitle")
              For Each n In nList
                idpuesto= replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
              Next
       
          'descripcion del puesto
           Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/titleDescription")
              For Each n In nList
                descripuesto=n.text   '  n.nodename &"=" & n.text & "<br>"
              Next
              
       
          'busco el codpuesto
           Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/comp_id")
              For Each n In nList
                idcomplemento= replace(n.getAttribute("href"),"#","")
              Next
         
         ' tomo el id del puesto
       
           Set nList = xml.selectNodes("//multiRef[@id='"&idcomplemento&"']/titleCode")
              For Each n In nList
                codigopuesto= n.text
              Next
  	
  	set Reader=nothing 
  	set Serializer=nothing 
  	set Connector=nothing 
  	'data= userid & "<br>"
  	'data=data &  nombre & "<br>"
  	'data=data &  mail & "<br>"
  	'data=data &  descrisector & "<br>"
  	'data=data &  idcostcenter & "<br>"
  	'data=data &  descripuesto & "<br>"
  	'data=data &  codigopuesto & "<br>"
  	
  	'getInfoUser=data
  	session("SGU_userid")       = userid
  	session("SGU_nombre")       = nombre
  	session("SGU_mail")	    = mail
  	session("SGU_descrisector") = descrisector
  	session("SGU_idcostcenter") = idcostcenter
  	session("SGU_descripuesto") = descripuesto
  	session("SGU_codigopuesto") = codigopuesto

   ELSE
     	session("SGU_userid")      = Legajo
  	session("SGU_nombre")      = "NO"
  	session("SGU_mail")	   = "NO"
  	session("SGU_descrisector")= "NO"
  	session("SGU_idcostcenter")= "0"
  	session("SGU_descripuesto")= "NO"
  	session("SGU_codigopuesto")= "NO"


   end if


   end select

End Function   

Function getInfoUserW(Legajo)
data = ""  

GTUcase = session("USRPRV")


  select case GTUcase

       case "I"

           set xcn = ofv.conectar(ofv.strconn2)

           xStrSql = "Select ID,descripcion,email,sector, puesto  From USUARIOS WHERE ID = '" & Legajo & "' "
           Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
           IF NOT xRS.EOF THEN
     		session("SGU_userid")      = xrs(0)
  		session("SGU_nombre")      = xrs(1)
  		session("SGU_mail")	   = xrs(2) & ""
  		session("SGU_descrisector")= "MySect"
  		session("SGU_idcostcenter")= xrs(3)
  		session("SGU_descripuesto")= "MyPst"
  		session("SGU_codigopuesto")= xrs(4)
           ELSE
     		session("SGU_userid")      = Legajo
  		session("SGU_nombre")      = "NO"
  		session("SGU_mail")	   = "NO"
  		session("SGU_descrisector")= "NO"
  		session("SGU_idcostcenter")= "0"
  		session("SGU_descripuesto")= "NO"
  		session("SGU_codigopuesto")= "NO"
           END IF

           call ofv.cerrarconsulta(xrs)
           call ofv.cerrarconn(xcn)


       case "E"


  ' DATOS DE LOS WEBSERVICES
  SoapAction = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getUserWithoutManagers"
  Method = "getUserWithoutManagers" ' METODO DEL WEBSERVICE A UTILIZAR
  
  END_POINT_URL = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
  CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"

  
  'Legajo = request.querystring("Legajo")
  if Legajo<>"" then
  	Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
  	Connector.Property("EndPointURL") = END_POINT_URL
  	Connector.Connect
  	
  	Connector.Property("SoapAction") = SoapAction & Method
  	Connector.BeginMessage
  	
  	Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
  	Serializer.Init Connector.InputStream
  	
  	Serializer.StartEnvelope
  	Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
  	Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
  	Serializer.StartBody
  	Serializer.StartElement Method, CALC_NS
  	Serializer.StartElement "userId"
  	Serializer.WriteString  Legajo  ' LEGAJO DEL USUARIO A BUSCAR
  	Serializer.EndElement
  	Serializer.EndElement
  	Serializer.EndBody
  	Serializer.EndEnvelope
  	
  	Connector.EndMessage
  	    
  	Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
  	Reader.Load Connector.OutputStream
  	' OBTENCION DE LOS DATOS
  	
  	
  	Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
  	Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text 
  	Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text 
  	
  	set xml=Reader.Body
  	set raiz=xml.selectSingleNode("//getUserWithoutManagersResponse/getUserWithoutManagersReturn")
	   	   
	id=replace(raiz.getAttribute("href"),"#","")


	   ' tomo el idref de la persona
	    Set nList = xml.selectNodes("//multiRef[@id='"&ID &"']/person")

	    For Each n In nList
	     idpersona = replace(n.getAttribute("href"),"#","")
	    Next

	   ' tomo el idref del costcenter
	    Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/costCenter")

	    For Each n In nList
	     idcostCenter = replace(n.getAttribute("href"),"#","")
	    Next

  	
  	 ' tomo el descripcion del costcenter /sector
	      Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterDescription")
	          For Each n In nList
	            descrisector= n.text ' replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
	       Next
	       
	 ' tomo el id del costcenter
	   Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterCode")
	       For Each n In nList
	         idcostcenter= n.text 'replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
	       Next
       
        ' tomo el id del puesto
          Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/functionTitle")
              For Each n In nList
                idpuesto= replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
              Next
       
          'descripcion del puesto
           Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/titleDescription")
              For Each n In nList
                descripuesto=n.text   '  n.nodename &"=" & n.text & "<br>"
              Next
              
       
          'busco el codpuesto
           Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/comp_id")
              For Each n In nList
                idcomplemento= replace(n.getAttribute("href"),"#","")
              Next
         
         ' tomo el id del puesto
       
           Set nList = xml.selectNodes("//multiRef[@id='"&idcomplemento&"']/titleCode")
              For Each n In nList
                codigopuesto= n.text
              Next
  	
  	set Reader=nothing 
  	set Serializer=nothing 
  	set Connector=nothing 
  	'data= userid & "<br>"
  	'data=data &  nombre & "<br>"
  	'data=data &  mail & "<br>"
  	'data=data &  descrisector & "<br>"
  	'data=data &  idcostcenter & "<br>"
  	'data=data &  descripuesto & "<br>"
  	'data=data &  codigopuesto & "<br>"
  	
  	'getInfoUser=data
  	session("SGU_userid")       = userid
  	session("SGU_nombre")       = nombre
  	session("SGU_mail")	    = mail
  	session("SGU_descrisector") = descrisector
  	session("SGU_idcostcenter") = idcostcenter
  	session("SGU_descripuesto") = descripuesto
  	session("SGU_codigopuesto") = codigopuesto

   ELSE
     	session("SGU_userid")      = Legajo
  	session("SGU_nombre")      = "NO"
  	session("SGU_mail")	   = "NO"
  	session("SGU_descrisector")= "NO"
  	session("SGU_idcostcenter")= "0"
  	session("SGU_descripuesto")= "NO"
  	session("SGU_codigopuesto")= "NO"


   end if


   end select

End Function   



'--------------------------------------------------------------------------------------------------

Function getGroupsUser(Legajo)
data = ""  


GTUcase = session("USRPRV")


  select case GTUcase

       case "I"

           set xcn = ofv.conectar(ofv.strconn2)

           xStrSql = "Select grupo  From USUARIOS WHERE ID = '" & Legajo & "' "
           Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
           IF NOT xRS.EOF THEN
              session("SGU_grupos")       = xrs(0)

           ELSE
              session("SGU_grupos")       = "NO"

           END IF

           call ofv.cerrarconsulta(xrs)
           call ofv.cerrarconn(xcn)


         session("SGU_grupos") 


       case "E"

  ' DATOS DE LOS WEBSERVICES
  SoapAction = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getGroups"
  Method = "getGroups" ' METODO DEL WEBSERVICE A UTILIZAR
  
  END_POINT_URL = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
  CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"

  
  'Legajo = request.querystring("Legajo")
  if Legajo<>"" then
  	Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
  	Connector.Property("EndPointURL") = END_POINT_URL
  	Connector.Connect
  	
  	Connector.Property("SoapAction") = SoapAction & Method
  	Connector.BeginMessage
  	
  	Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
  	Serializer.Init Connector.InputStream
  	
  	Serializer.StartEnvelope
  	Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
  	Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
  	Serializer.StartBody
  	Serializer.StartElement Method, CALC_NS
  	Serializer.StartElement "userId"
  	Serializer.WriteString  Legajo  ' LEGAJO DEL USUARIO A BUSCAR
  	Serializer.EndElement
  	Serializer.EndElement
  	Serializer.EndBody
  	Serializer.EndEnvelope
  	
  	Connector.EndMessage
  	    
  	Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
  	Reader.Load Connector.OutputStream

 ' prueba 	
  	'response.write Reader.Body.xml
  	Set nList = Reader.Body.selectNodes("//multiRef/name")
  	session("SGU_grupos")="DEF" & chr(9)
  	Dim iLoop
	   For Each n In nList
	       if n.text <>"" then
	       'response.write  iloop & " grupo=" & n.text & "<br>"
	        session("SGU_grupos")=session("SGU_grupos") & n.text & chr(9)
	       end if
	       iloop = iloop + 1
	Next
'---------------------------
  	'data =  Reader.Body.selectSingleNode("//multiRef/userId").Text & chr(13)
  	'data =  data & Reader.Body.selectSingleNode("//multiRef/displayName").Text & chr(13)
  	'data =  data & Reader.Body.selectSingleNode("//multiRef/workEMail").Text & chr(13)
  	'data =  data & Reader.Body.selectSingleNode("//multiRef/costCenterCode").Text & chr(13)
  	'data =  data & Reader.Body.selectSingleNode("//multiRef/costCenterDescription").Text
  	
  	
  	set Reader=nothing 
  	set Serializer=nothing 
  	set Connector=nothing 
  	
  	getGroupsUser=data

   end if


   end select


End Function   

Sub MuestraNodos ( Nodos )
 Dim oNodo 
  For Each oNodo In Nodos
   response.write oNodo.nodeName & "=" & oNodo.nodeValue & "<br>"
   
   If oNodo.hasChildNodes Then
       response.write chr(9)
       call   MuestraNodos (oNodo.childNodes)
   End If


  Next 

End Sub 



sub getUsrIngroup(grupo)
Dim xml,root,i

GTUcase = session("USRPRV")

  xgusep                   = ""
  session("SGU_UIGuserid") = ""
  session("SGU_UIGdescri") = ""

  select case GTUcase

       case "I"

           set xcn = ofv.conectar(ofv.strconn2)

           xStrSql = "Select ID,descripcion  From USUARIOS WHERE grupo = '" & grupo & "' order by descripcion "
           Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
           IF NOT xRS.EOF THEN

              DO WHILE NOT XRS.EOF

                session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
                session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & xrs(1)

                xgusep                   = chr(9)
             
              XRS.MOVENEXT
              LOOP

           ELSE

           END IF

           call ofv.cerrarconsulta(xrs)
           call ofv.cerrarconn(xcn)


       case "E"


   ' DATOS DE LOS WEBSERVICES
   SoapAction = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getMembersInGroup"
   Method = "getMembersInGroup" ' METODO DEL WEBSERVICE A UTILIZAR
   
   END_POINT_URL = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
   CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"
 
   
   	Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
   	Connector.Property("EndPointURL") = END_POINT_URL
   	Connector.Connect
   	
   	Connector.Property("SoapAction") = SoapAction & Method
   	Connector.BeginMessage
   	
   	Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
   	Serializer.Init Connector.InputStream
   	
   	Serializer.StartEnvelope
   	Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
   	Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
   	Serializer.StartBody
   	Serializer.StartElement Method, CALC_NS
   	Serializer.StartElement "cn"
   	Serializer.WriteString  grupo  ' LEGAJO DEL USUARIO A BUSCAR
   	Serializer.EndElement
   	Serializer.EndElement
   	Serializer.EndBody
   	Serializer.EndEnvelope
   	
   	Connector.EndMessage
   	    
   	Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
   	Reader.Load Connector.OutputStream
   	' OBTENCION DE LOS DATOS
   	
   	
   	
  	set xml=Reader.Body
   
   
    
   
   ' tomo el idref de la persona
    Set nList = raiz.selectNodes("item")
    iz        = 0
    xgutop    = 0
    for Each n In nList
       id            = replace(n.getAttribute("href") ,"#","")
       Set n2        = xml.selectSingleNode("//multiRef[@id='"&id&"']/name")
       idcomplemento = replace(n2.getAttribute("href"),"#","")

       Set n3        = xml.selectSingleNode("//multiRef[@id='"&idcomplemento&"']")
       nombre        = n3.selectSingleNode("employeeName").text
       apellido      = n3.selectSingleNode("employeeLastname").text
       legajo        = n3.selectSingleNode("userId").text

       xxx           = cstr(ucase(apellido & " "  & nombre ))

       arrr(iz,1)    = xxx
       arrr(iz,0)    = legajo
       iz            = iz + 1           
    Next

    xgutop    = iz

     if iz > 0 then
         call QSort(arrr,0,iz-1)

         for iz=0 to  xgutop

             session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & arrr(iz,0)
             session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & arrr(iz,1)

             xgusep                   = chr(9)

         next 


     end  if 



   end select

end sub

sub getUsrInSect(XGU_SECT)
Dim xml,root,i

GTUcase = session("USRPRV")

  xgusep                   = ""
  session("SGU_UIGuserid") = ""
  session("SGU_UIGdescri") = ""

  select case GTUcase

       case "I"

           IF ISNUMERIC(XGU_SECT) THEN
              WXGU_SECT = XGU_SECT
           ELSE
              WXGU_SECT = 1
           END IF 

           set xcn = ofv.conectar(ofv.strconn2)

           xStrSql = "Select ID,descripcion  From USUARIOS WHERE SECTOR = 0" & WXGU_SECT & " order by descripcion "
           Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
           IF NOT xRS.EOF THEN

              DO WHILE NOT XRS.EOF

                session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
                session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & xrs(1)

                xgusep                   = chr(9)
             
              XRS.MOVENEXT
              LOOP

           ELSE

           END IF

           call ofv.cerrarconsulta(xrs)
           call ofv.cerrarconn(xcn)


       case "E"




   end select

end sub

sub getUsrInTEXT(XGU_TEXT)
Dim xml,root,i

GTUcase = session("USRPRV")

  xgusep                   = ""
  session("SGU_UIGuserid") = ""
  session("SGU_UIGdescri") = ""

  select case GTUcase

       case "I"

           WXGU_TEXT = REPLACE(XGU_TEXT & "","*","%")

           set xcn = ofv.conectar(ofv.strconn2)

           xStrSql = "Select ID,descripcion  From USUARIOS WHERE DESCRIPCION LIKE '" & WXGU_TEXT & "' order by descripcion "
           Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
           IF NOT xRS.EOF THEN

              DO WHILE NOT XRS.EOF

                session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
                session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & xrs(1)

                xgusep                   = chr(9)
             
              XRS.MOVENEXT
              LOOP

           ELSE

           END IF

           call ofv.cerrarconsulta(xrs)
           call ofv.cerrarconn(xcn)


       case "E"




   end select

end sub


Sub QSort(aData, iLb, iUb)
	Dim lbStack(1000)
	Dim ubStack(1000)
	Dim sp              ' stack pointer
	Dim iLbx            ' current lower-bound
	Dim iUbx            ' current upper-bound
	Dim m
	Dim p               ' index to pivot
	Dim i
	Dim j
	Dim t               ' temp used for exchanges

	' Init stacks
	lbStack(0) = iLb
	ubStack(0) = iUb
	sp = 0

	Do While sp >= 0
		iLbx = lbStack(sp)
		iUbx = ubStack(sp)

		Do While (iLbx < iUbx)

			' select pivot and exchange with 1st element
			p = iLbx + (iUbx - iLbx) \ 2

			' exchange iLbx, p
			t = aData(iLbx,1)
			aData(iLbx,1) = aData(p,1)
			aData(p,1) = t
			
			
			t = aData(iLbx,0)
			aData(iLbx,0) = aData(p,0)
			aData(p,0) = t
			

			' partition into two segments
			i = iLbx + 1
			j = iUbx
			'RESPONSE.WRITE aData(iLbx,1) & "--" & aData(i,1) & "<BR>"
			Do
				Do While i < j
					If aData(iLbx,1) <= aData(i,1) Then Exit Do
					i = i + 1
				Loop

				Do While j >= i
					If aData(j,1) <= aData(iLbx,1) Then Exit Do
					j = j - 1
				Loop

				If i >= j Then Exit Do

				' exchange i, j
				t = aData(i,1)
				aData(i,1) = aData(j,1)
				aData(j,1) = t
				t = aData(i,0)
				aData(i,0) = aData(j,0)
				aData(j,0) = t

				j = j - 1
				i = i + 1
			Loop

			' pivot belongs in aData[j]
			' exchange iLbx, j
			t = aData(iLbx,1)
			aData(iLbx,1) = aData(j,1)
			aData(j,1) = t
			
			
			t = aData(iLbx,0)
			aData(iLbx,0) = aData(j,0)
			aData(j,0) = t
			

			m = j

			' keep processing smallest segment, and stack largest
			If m - iLbx <= iUbx - m Then
				If m + 1 < iUbx Then
					lbStack(sp) = m + 1
					ubStack(sp) = iUbx
					sp = sp + 1
				End If
				iUbx = m - 1
			Else
				If m - 1 > iLbx Then
					lbStack(sp) = iLbx
					ubStack(sp) = m - 1
					sp = sp + 1
				End If
				iLbx = m + 1
			End If
		Loop
		sp = sp - 1
	Loop
End Sub




%>    
