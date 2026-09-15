<%

 Function ConUsrset()


           set XCcn = ofv.conectar(ofv.strconn2)


           StrSql = "Select * From SYSDEFDSKTOPDET WHERE CODSET = 0" & Session("USRSET")
           Set XCrs = OFV.crearconsultaEx(StrSql, XCcn, 1, "")
           IF NOT XCRS.EOF THEN
              DO UNTIL XCRS.EOF

                 SESSION(XCRS(2)) = XCRS(3)

                 XCRS.MOVENEXT
                 LOOP 
           END IF

           call ofv.cerrarconn(XCcn)


 END FUNCTION


 Function Conset(Tipo)

  Select Case Tipo

    Case "0"

''' Definido por el usuario


 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#800000"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "fbeedd"


 Session("CMIMENU")        = "back10.gif"

 Session("CMFMENU")        = "#77aadd"
 Session("CMFMEFTFC")      = "verdana"
 Session("CMFMEFT")        = "#FFFFFF"
 Session("CMFMEFTSZ")      = "7"

 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"
 Session("CMFMNBOT")       = "#6699cc"


 Session("CMFPNBOT")       = "#6699cc"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#6699cc"
 Session("CMFPNTITFTFC")   = "Tahoma"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#fbeedd"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#888888"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#888888"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#888888"
 Session("CMFPNGRFTSZ")    = "7"

'' Session("CMFPNIP")        = "#f7eedd"
 Session("CMFPNIP")        = "fbeedd"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#888888"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#888888"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""
          
 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#888888"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""
          
 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#888888"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#f7eedd"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""
          
 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#888888"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""
          
 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#888888"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""
 
          
 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#888888"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""
          
 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#888888"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""
          
 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#888888"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""
          
 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#888888"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        =  XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 
 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#888888"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#888888"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#888888"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#888888"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#888888"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"
          

    Case "1"

''' Set 1 - rojo/rojo oscuro/verdana

 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#888888"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "fbeedd"

 Session("CMIMENU")        = "Xback10.gif"

 Session("CMFMENU")        = "#d68811"
 Session("CMFMEFTFC")      = "verdana"
 Session("CMFMEFT")        = "#FFFFFF"
 Session("CMFMEFTSZ")      = "7"

 Session("CMFMNBOT")       = "#880000"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"

 Session("CMFPNBOT")       = "#aa0000"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#aa0000"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#eeeeee"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#888888"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#888888"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"


 Session("CMFPNIP")        = "#DDDDDD"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#888888"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#888888"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""
          
 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#888888"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""
          
 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#888888"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#DDDDDD"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""
          
 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#888888"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""
          
 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#AA0000"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""
 
          
 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#888888"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""
          
 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#888888"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""
          
 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""
          
 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#888888"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 
 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#888888"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#888888"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#888888"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#888888"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#888888"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"
        
    Case "2"

 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#000080"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "fbeedd"

 Session("CMIMENU")        = "Xback10.gif"

 Session("CMFMENU")        = "#6699cc"
 Session("CMFMEFTFC")      = "verdana"
 Session("CMFMEFT")        = "#FFFFFF"
 Session("CMFMEFTSZ")      = "7"

 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"
 Session("CMFMNBOT")       = "#13B44D"

 Session("CMFPNBOT")       = "#13B44D"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#6699cc"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#F2F2F2"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#000080"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#000080"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"


 Session("CMFPNIP")        = "#F2F2F2"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#0000aa"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#000080"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""
          
 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#000080"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""
          
 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#000080"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#F2F2F2"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""
          
 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#000080"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""
          
 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#137ACF"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""
 
          
 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#000080"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""
          
 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#000080"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""
          
 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""
          
 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#000080"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 
 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#000080"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#000080"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#000080"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#000080"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#000080"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"
        
    Case "3"

''' Set 3 - Celeste/Azul/arial

 Session("APLBGD")         = "#ffffff"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#000080"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "f7eedd"


 Session("CMFMENU")        = "#993333"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"
 Session("CMIMENU")        = "xcorf5.jpg"
 Session("CMFMNBOT")       = "#f7aa22"

 Session("CMFPNBOT")       = Session("CMFMNBOT")
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = Session("CMFMNBOT")
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = Session("APLITBGD")
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = Session("CMFMENU")
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = Session("CMFMENU")
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = Session("CMFMENU")
 Session("CMFPNGRFTSZ")    = "7"

 Session("CMFPNIP")        = Session("APLITBGD")
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = Session("CMFMENU")
 Session("CMFPNIPFTSZ")    = "7"
 
 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#000080"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = ""
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""
          
 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#000080"
 Session("CMIETXFTSZ")     = "8"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""
          
 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#000080"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#ffffff"
 Session("CMISTTBL")       = ""
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""
          
 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#000080"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""
          
 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#000080"
 Session("CMIBTTFTSZ")     = "8"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = ""
 Session("CMIBTTIT")       = ""
 Session("CMIBTTUL")       = ""
 
          
 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#000080"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""
          
 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#000080"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""
          
 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#000080"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = ""
 Session("CMICTRUL")       = ""
          
 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#000080"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = ""
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspDShw("2","2",Session("APLSHW")) & ";"
''' Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 
 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = Session("CMFMNBOT")
 Session("CORGGFMFTSZ")     = "7"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = Session("CMFMNBOT")
 Session("CORGGFSFTSZ")     = "7"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = Session("CMFMNBOT")
 Session("CORGGFNFTSZ")     = "7"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = Session("CMFMNBOT")
 Session("CORGRFTFTSZ")     = "7"
 Session("CORGRFT")         = ""

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = Session("CMFMNBOT")
 Session("CORGRFCFTSZ")     = "7"
 Session("CORGRFC")         = ""

    Case "4"

''' Set 4 - Celeste/Azul/tahoma

 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#000080"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#EEEEEE"
 Session("APLITBGD")       = "#99CCEE"

 Session("CMFMENU")        = "#6495ED"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"
 Session("CMIMENU")        = "Xback10.gif"
 Session("CMFMNBOT")       = "#DAA520"

 Session("CMFPNBOT")       = "#DAA520"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#DAA520"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#F7EEDD"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#000080"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#000080"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"


 Session("CMFPNIP")        = "#F7EEDD"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#0000aa"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#000080"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""
          
 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#000080"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""
          
 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#000080"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#DDDDDD"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""
          
 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#000080"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""
          
 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#DAA520"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""
 
          
 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#000080"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""
          
 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#000080"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""
          
 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""
          
 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#000080"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 
 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#000080"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#000080"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#000080"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#000080"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#000080"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"

    Case "5"

''' Definido por el usuario


 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#000080"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "eeeeee"

 Session("CMIMENU")        = "xback10.gif"

 Session("CMFMENU")        = "#e0dbd5"
 Session("CMFMEFTFC")      = "verdana"
 Session("CMFMEFT")        = "#1e93a7"
 Session("CMFMEFTSZ")      = "7"

 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#ffffff"
 Session("CMFMNFTSZ")      = "7"
 Session("CMFMNBOT")       = "#1e93a7"

 Session("CMFPNBOT")       = "#1e93a7"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#1e93a7"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#eeeeee"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#000080"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#000080"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"

 Session("CMFPNIP")        = "eeeeee"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#0000aa"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#000080"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""
          
 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#000080"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""
          
 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#000080"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#f7eedd"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""
          
 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#000080"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""
          
 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#993333"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""
 
          
 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#000080"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""
          
 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#000080"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""
          
 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""
          
 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#000080"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 
 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#000080"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#000080"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#000080"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#000080"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#000080"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"
          

    Case "6"

''' Definido por el usuario


 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#4D4B3D"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "fbeedd"

 Session("CMIMENU")        = ""

 Session("CMFMENU")        = "#B3B2A0"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"

 Session("CMFMNBOT")       = "#828278"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"

 Session("CMFPNBOT")       = "#828278"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#828278"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#E7E7E0"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#4D4B3D"
 Session("CMFPNITFTSZ")    = "7"



 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#4D4B3D"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"


 Session("CMFPNIP")        = "#E7E7E0"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#4D4B3D"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#4D4B3D"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""
          
 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#4D4B3D"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""
          
 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#4D4B3D"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#f7eedd"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""
          
 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#4D4B3D"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""
          
 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#993333"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""
 
          
 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#4D4B3D"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""
          
 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#4D4B3D"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""
          
 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""
          
 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#4D4B3D"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 
 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#4D4B3D"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#4D4B3D"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#4D4B3D"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#4D4B3D"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#4D4B3D"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"
          

        
    Case Else 

        Conset("1")


  End Select




 End Function

 Function XspEx(XspCl1,XspCl2,XspTyp)

     if   Session("FULLGUI") = "S" then 

          XspEx = "filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & XspCl1 & ", endColorStr=" & XspCl2 & ", gradientType=" & XspTyp & ")"

     elseif  Session("FULLGUI") = "M" then 

          XspEx = "filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & XspCl1 & ", endColorStr=" & XspCl2 & ", gradientType=" & XspTyp & ")"

     else

          XspEx = "background-color:" & XspCl2 & ";" & XspAlpha(0,100,1) & ""

     end if

 End Function

 Function XspAlpha(XspCl1,XspCl2,XspTyp)

     if   Session("FULLGUI") = "S" then 

          XspAlpha = "filter: alpha(opacity=" & XspCl1 & ",finishOpacity=" & XspCl2 & ",STYLE=" & XspTyp & ")"

     elseif  Session("FULLGUI") = "M" then 

          XspAlpha = "filter: alpha(opacity=" & XspCl1 & ",finishOpacity=" & XspCl2 & ",STYLE=" & XspTyp & ")"

     else

        '''  XspAlpha = "filter: alpha(opacity=100,finishOpacity=90,STYLE=" & XspTyp & ")"

          XspAlpha = ""

     end if

 End Function

 Function XspSShw(XspCl1,XspAng)

     if   Session("FULLGUI") = "S" then 

          XspSShw = "filter:shadow(color=" & XspCl1 & ",direction=" & XspAng & ")"

     elseif  Session("FULLGUI") = "M" then 

          XspSShw = "filter:shadow(color=" & XspCl1 & ",direction=" & XspAng & ")"

     else

        ''  XspSShw = "border-right-style:solid;border-right-width:1;border-right-color:" & XspCl1 & ";border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & XspCl1 & ""
          XspSShw = ""

     end if

 End Function

 Function XspDShw(XspOFX,XspOFY,XspCl1)

     if   Session("FULLGUI") = "S" then 

          XspDShw = "filter: dropShadow(Color=" & XspCl1 & ",OFFX=" & XspOFX & ",OFFY=" & XspOFY & ")"

     elseif  Session("FULLGUI") = "M" then 

          XspDShw = "filter: dropShadow(Color=" & XspCl1 & ",OFFX=" & XspOFX & ",OFFY=" & XspOFY & ")"

     else

        ''  XspDShw = "border-right-style:solid;border-right-width:1;border-right-color:" & XspCl1 & ";border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & XspCl1 & ""
          XspDShw = ""
     end if

 End Function

 Function CorStyler()

   txsscpnm  = request.servervariables("script_name")
   txsscpnmi = instrrev(txsscpnm,"/")
   txsscpnmy = right(ucase(txsscpnm),len(txsscpnm)-txsscpnmi)

   if  txsscpnmy = "INICIO.ASP" then
 
       CorStyler = " "

   else 

      wcorstyl = "<script language='javascript' >"
      wcorstyl = wcorstyl & "  if (window.top.CorMtopm) { "
      wcorstyl = wcorstyl & "      var wcortopf = window.top.CorMtopf; } "
      wcorstyl = wcorstyl & "  else { "
      wcorstyl = wcorstyl & "         if (window.opener.top.CorMtopf) { "
      wcorstyl = wcorstyl & "             var wcortopf = window.opener.top.CorMtopf; } "
      wcorstyl = wcorstyl & "         else { "
      wcorstyl = wcorstyl & "             var wcortopf = window.opener; "
'''sHTML = sHTML & " alert(CorMtopm.name);" & chr(13)
      wcorstyl = wcorstyl & "             for (ii=0;ii<10;ii++) {" & chr(13)
'''''sHTML = sHTML & " alert(CorMtopm.top.name + ' ' + ii);" & chr(13)
      wcorstyl = wcorstyl & "               if (wcortopf.top.name=='superm') {" & chr(13)
      wcorstyl = wcorstyl & "                   wcortopf = wcortopf.top;" & chr(13)
''sHTML = sHTML & " alert(CorMtopf.name);" & chr(13)
''sHTML = sHTML & " alert(verfuncion());" & chr(13)
      wcorstyl = wcorstyl & "               break; " & chr(13)
      wcorstyl = wcorstyl & "               } else {" & chr(13)
      wcorstyl = wcorstyl & "                 wcortopf = wcortopf.top.opener;" & chr(13)
      wcorstyl = wcorstyl & "               }" & chr(13)
      wcorstyl = wcorstyl & "             }" & chr(13)
      wcorstyl = wcorstyl & "          } "
      wcorstyl = wcorstyl & "   } "
      wcorstyl = wcorstyl & "document.write ('<style>'); "
      wcorstyl = wcorstyl & "document.write (wcortopf.verestilo()); "
      wcorstyl = wcorstyl & "document.write ('</style>');"
      wcorstyl = wcorstyl & "</script>"

       CorStyler = wcorstyl

   end if

 End Function

 Function CorStyler2()

         if Session("CMFMEFTFC") = "" then
            Session("CMFMEFTFC")      = Session("CMFMNFTFC")
            Session("CMFMEFT")        = Session("CMFMNFT")
            Session("CMFMEFTSZ")      = Session("CMFMNFTSZ")
         end if

'''filter: gray;
'''filter: xray;
      
         s = "<META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " CONTENT=" & Chr(34) & "text/html; charset=windows-1252" & Chr(34) & ">"
         s = "<style id='estil' >"

         s = s & "BODY {SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;}"

    '''     s = s & "BODY {SCROLLBAR-FACE-COLOR:" & session("CMFMNBOT") & "; SCROLLBAR-HIGHLIGHT-COLOR: #eeeeee; SCROLLBAR-SHADOW-COLOR: #aaaaaa; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:" & session("CMFMNFT") & "; SCROLLBAR-TRACK-COLOR: #eeeeee; SCROLLBAR-DARKSHADOW-COLOR: #888888}"

         s = s & "p{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
         s = s & "a{text-decoration:none;}"
         s = s & "a.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;}"
         s = s & "a.link{color:#ffffff;text-decoration:none;}"
         s = s & "a.nn{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:#ffffff;text-transform:capitalize;font-weight:500;text-decoration:none;}"
         s = s & "a.nt{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-weight:500;text-decoration:none;}"
         s = s & "a.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;text-decoration:none;}"
         s = s & "a.utg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;text-decoration:none;}"
         s = s & "a.cc{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";text-transform:uppercase;font-weight:800;text-decoration:none;}"
         s = s & "a.tt{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";text-transform:uppercase;font-weight:800;text-decoration:none;}"
         s = s & "a.ts{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";text-transform:uppercase;font-weight:800;text-decoration:none;}"
         s = s & "a.cmb{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-weight:normal;text-decoration:none;}"

         s = s & "a.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;text-decoration:none;}"

         s = s & "font.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1;border-color:#f9f9f9;}"

         s = s & "font.bt2{font-family:wingdings;color:#ffffff;font-size:11pt;font-weight:bold;background-color:#dd0000;border-style:inset;border-width:1;border-color:#ffffff;}"
         s = s & "font.bt1{font-family:wingdings;color:#ffffff;font-size:9pt;font-weight:bold;background-color:#00bb00;border-style:outset;border-width:1;border-color:#ffffff;}"
         s = s & "font.btc1{font-family:webdings;color:#ffffff;font-size:7pt;font-weight:bold;}"


         s = s & "input{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("aplbgd") & ";}"
         s = s & "input.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:#ffffff;}"
         s = s & "input.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#dddddd",session("CMFPNBOT"),0) & ";border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;}"
         s = s & "input.bte{font-family:'webdings';color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#990000","#dddddd",0) & ";border-style:solid;border-width:1;border-color:#990000;height:20px;cursor:hand;}"
         s = s & "input.btG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#bbbbbb;font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#eeeeee","#dddddd",0) & ";border-style:solid;border-width:1;border-color:#eeeeee;height:20px;cursor:hand;}"
         s = s & "input.btR{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#eeeeee","#550000",0) & ";border-style:solid;border-width:1;border-color:#550000;height:20px;cursor:hand;}"
'''background-color:" & Session("CMFPNIP") & ";

    '  if ofv.modtitcls <> "10" then
    '     s = s & "input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;" & XspEx("#dddddd",session("APLBGD"),0) & ";}"
    '     s = s & "input.dtb{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;font-weight:bold;border-style:solid;border-width:1pt;border-color:#dddddd;" & XspEx("#dddddd",session("APLBGD"),0) & ";}"
    '     s = s & "input.dtn{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;text-align=right;border-style:solid;border-width:1pt;border-color:#dddddd;" & XspEx("#dddddd",session("APLBGD"),0) & ";}"
    '     s = s & "input.dts{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;background-color:transparent;}"
    '     s = s & "input.dtbs{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;font-weight:bold;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;background-color:transparent;}"
    '     s = s & "input.dtns{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;text-align=right;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;background-color:transparent;}"
    '  else
         s = s & "input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#f0f0f0;}"
         s = s & "input.dtb{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;font-weight:bold;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#f0f0f0;}"
         s = s & "input.dtn{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;text-align=right;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#f0f0f0;}"
         s = s & "input.dts{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFPNTIT") & ";background-color:#f0f0f0;}"
         s = s & "input.dtbs{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;font-weight:bold;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFPNTIT") & ";background-color:#f0f0f0;}"
         s = s & "input.dtns{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;text-align=right;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFPNTIT") & ";background-color:#f0f0f0;}"
     ' end if


         s = s & "input.it{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("CMFPNIP") & ";}"
         s = s & "input.rd{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("CMFPNIP") & ";}"
         s = s & "input.tt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:600;background-color:" & session("CMFPNBOT") & ";}"
         s = s & "input.tt3{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:600;background-color:#993333;}"

         s = s & "option{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;}"
         s = s & "option.OCM{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;}"

         s = s & "select{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("APLBGD") & ";border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";}"
         s = s & "select.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";}"

         s = s & "span.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("CMFPNIP") & ";}"
         s = s & "span.mt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNIP") & ";width:80%;}"

         s = s & "th{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"

         s = s & "tr.st{background-color:" & Session("CMFPNTIT") & ";}"
         s = s & "tr.cc{background-color:" & Session("CMFPNIT") & ";border-style:ridge;border-width:1pt;}"
         s = s & "tr.XFK{height:20px;}"

         s = s & "td.aa{text-transform: capitalize;}"
         s = s & "td{font-family:" & Session("CMFPNITFTFC") & ";font-size:" & Session("CMFPNITFTSZ") & "pt;}"
         s = s & "td.at{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:500;}"
         s = s & "td.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:" & session("CMFPNBOT") & ";text-decoration:underline;border-style:outset;border-width:1pt;}"
         s = s & "td.btG{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#dddddd;border-style:inset;border-width:1pt;}"
         s = s & "td.cc{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";text-indent:4px;font-size:" & Session("CMFPNITFTSZ") & "pt;background-color:" & Session("CMFPNIT") & ";border-style:solid;border-width:1;border-color:#eeeeee;}"
         s = s & "td.ee{font-size:" & Session("CMFPNITFTSZ") & "pt;font-family:" & Session("CMFPNITFTFC") & ";border-width:1pt;border-color:" & Session("CMFPNTIT") & ";color:" & Session("CMFPNITFT") & ";font-weight:500;}"
         s = s & "td.lt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;text-indent:10;}"
         s = s & "td.dt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;}"
         s = s & "td.dt2{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;}"
         s = s & "td.et{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;}"
         s = s & "td.mt{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";font-size:" & Session("CMFPNITFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNIT") & ";}"
         s = s & "td.qt{font-family:" & Session("CMFPNGRFTFC") & ";color:" & Session("CMFPNGRFT") & ";font-size:" & Session("CMFPNGRFTSZ") & "pt;font-weight:600;}"
         s = s & "td.rr{font-size:" & Session("CMFPNTITFTSZ") & "pt;font-family:" & Session("CMFPNTITFTFC") & ";border-width:1pt;border-color:" & Session("CMFMNBOT") & ";color:" & Session("CMFPNTITFT") & ";font-weight:500;}"
         s = s & "td.st{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"

''         s = s & "td.tdcm1{cursor:hand;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;text-indent:5;" & XspEx("#dddddd",session("APLBGD"),0) & ";}"
''         s = s & "td.tdcm2{cursor:hand;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;" & XspEx("#EEEEEE",session("CMFPNTIT"),0) & ";}"

         s = s & "td.tdcm1{cursor:hand;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;text-indent:5;background-color:#f0f0f0;}"
         s = s & "td.tdcm2{cursor:hand;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#0000dd;background-color:#0000dd;}"

         s = s & "td.ts{font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";font-size:" & Session("CMFMNFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1;background-color:" & session("CMFMNBOT") & ";" & XspAlpha(100,60,1) & ";}"



         s = s & "td.btbs{color:#6699cc;cursor:hand;}"
         s = s & "td.btbn{color:#f7aa22;cursor:hand;}"
         s = s & "td.btba{color:#005500;cursor:hand;}"

         s = s & "td.btbsr{background-color:#6699cc;color:#ffffff;border-style:solid;border-width:1;}"
         s = s & "td.btbnr{background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;}"
         s = s & "td.btbar{background-color:#005500;color:#ffffff;border-style:solid;border-width:1;}"

         s = s & "div.ut  {border-style:none;border-width:1;border-color:#ffffff;}"
         s = s & "div.btbs{border-style:solid;border-width:1;border-color:#6699cc;}"
         s = s & "div.btbn{border-style:solid;border-width:1;border-color:#f7aa22;}"
         s = s & "div.btba{border-style:solid;border-width:1;border-color:#005500;}"

         s = s & "a.btbs{color:#6699cc;text-decoration:none;}"
         s = s & "a.btbn{color:#f7aa22;text-decoration:none;}"
         s = s & "a.btba{color:#005500;text-decoration:none;}"

         s = s & "font.btic{font-family:webdings;font-size:10pt;line-height:12pt;font-weight:normal;}"
         s = s & "font.btiw{font-family:wingdings;font-size:10pt;line-height:12pt;font-weight:normal;}"
         s = s & "font.btic2{font-family:" & Session("CMFPNIPFTFC") & ";font-size:9pt;line-height:12pt;font-weight:normal;}"


         s = s & "a.ft1n{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#f7aa22;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#f7aa22;}"
         s = s & "a.ft1o{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;text-indent:7pt;letter-spacing:3pt;color:#f7aa22;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#f7aa22;background-color:#ffffff;}"

         s = s & "a.ft2n{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#aa0000;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#aa0000;}"
         s = s & "a.ft2o{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;text-indent:7pt;letter-spacing:3pt;color:#aa0000;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#aa0000;background-color:#ffffff;}"

         s = s & "a.ft3n{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#00aa00;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#00aa00;}"
         s = s & "a.ft3o{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;text-indent:5pt;letter-spacing:3pt;color:#00aa00;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#00aa00;background-color:#ffffff;}"

 

   ''''      s = s & "td.tt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-style:solid;border-width:1;border-color:" & Session("CMFPNBOT") & ";" & XspEx(Session("CMFPNTIT"),"#EEEEEE",1) & ";}"

         s = s & "td.ttb{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-indent:4px;text-transform:uppercase;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"
   ''''      s = s & "td.tt {font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-indent:4px;text-transform:uppercase;font-weight:600;border-style:solid;border-width:1;border-color:#eeeeee;" & XspAlpha(40,100,3) & ";background-color:" & Session("CMFPNTIT") & ";}"
         s = s & "td.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-bottom-style:solid;border-bottom-width:1pt;border-bottom-color:#eeeeee;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:10pt;text-indent:4pt;padding-right:4pt;border-style:none;border-width:1;border-color:#eeeeee;" & XspEx("#EEEEEE","#888888",0) & ";}"
         s = s & "td.ttx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
           s = s & "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-right-style:solid;border-right-width:1;border-right-color:#aa0000;background-color:#BCBCBC;}"

         s = s & "td.tt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
           s = s & "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-right-style:solid;border-right-width:1;border-right-color:#aa0000;background-color:#BCBCBC;}"



         s = s & "td.mbt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & session("CMFMNBOT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:600;" & XspDShw("1","1","#f9f9f9") & ";border-style:solid;border-width:1;border-color:#f9f9f9;cursor:hand;}"

     '''    s = s & "td.tt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;background-color:" & session("CMFPNTIT") & ";border-style:outset;border-width:1pt;}"
         s = s & "td.tu{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") + 2 & "px;font-weight:bold;}"
         s = s & "td.tx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNTITFT") & ";}"
         s = s & "td.ft{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;" & XspDShw("2","3","#cccccc") & ";}"
         s = s & "td.fti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;" & XspDShw("2","3","#cccccc") & ";padding-left:20px;}"
         s = s & "td.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
         s = s & "td.uts{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.uti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;}"
         s = s & "td.utis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.utg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;}"
         s = s & "td.utgi{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;}"
         s = s & "td.utgs{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.utgis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.ut2{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;font-weight:none;}"
         s = s & "td.vt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;}"
         s = s & "td.vtg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:600;}"
         s = s & "td.vti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;padding-left:20px;}"
         s = s & "td.vtig{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:600;padding-left:20px;}"
         s = s & "td.wt{font-family:" & Session("CMFPNGNFTFC") & ";color:#550000;font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;}"
         s = s & "td.xt{font-family:" & Session("CMFPNGNFTFC") & ";color:#dddddd;font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"

         s = s & "td.za{font-family:" & Session("CMFPNTITFTFC") & ";color:#ff0000;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;text-transform:uppercase;}"
         s = s & "td.zb{font-family:" & Session("CMFPNTITFTFC") & ";color:#0000ff;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;text-transform:uppercase;}"
         s = s & "td.zc{font-family:" & Session("CMFPNTITFTFC") & ";color:#000000;font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;}"
         s = s & "td.zd{font-family:" & Session("CMFPNTITFTFC") & ";color:#000099;font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;}"
         s = s & "td.ze{font-family:" & Session("CMFPNTITFTFC") & ";color:#000033;font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;}"
         s = s & "td.tbh1{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;}"
      '   s = s & "td.tbh2{" & XspEx(Session("APLBGD"),Session("CMFMENU"),1) & ";}"

''         s = s & "textarea{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") + 1 & "pt;background-color:" & Session("CMFPNIP") & ";border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";}
'''filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & session("APLBGD") & ", endColorStr=#dddddd, gradientType=0);
         s = s & "textarea{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") + 1 & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#F9F9F9;}"
         s = s & "textarea.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") + 1 & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#F9F9F9;}"

''' Estilos de Impresion......
         s = s & "div{page-break-after:always;}"
         s = s & "div.DINDM{background-color:" & Session("APLBGD") & ";border-style:SOLID;border-width:2;border-color:" & session("CMFMNBOT") & ";}"
         s = s & "div.DINDC{border-style:solid;border-width:1;border-color:#AAAAAA;background-color:#f9f9f9;}"

         s = s & "td.iett{font-family:" & Session("CMIETTFTFC") & ";color:" & Session("CMIETTFT") & ";font-size:" & Session("CMIETTFTSZ") & "pt;background-color:" & Session("CMIETT") & ";" & replace(Session("CMIETTBL"),"1","font-weight:bold;") & replace(Session("CMIETTIT"),"1","font-style:italic;") & replace(Session("CMIETTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ietx{font-family:" & Session("CMIETXFTFC") & ";color:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;background-color:" & Session("CMIETX") & ";" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.istt{font-family:" & Session("CMISTTFTFC") & ";color:" & Session("CMISTTFT") & ";font-size:" & Session("CMISTTFTSZ") & "pt;background-color:" & Session("CMISTT") & ";" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.istx{font-family:" & Session("CMISTXFTFC") & ";color:" & Session("CMISTXFT") & ";font-size:" & Session("CMISTXFTSZ") & "pt;background-color:" & Session("CMISTX") & ";" & replace(Session("CMISTXBL"),"1","font-weight:bold;") & replace(Session("CMISTXIT"),"1","font-style:italic;") & replace(Session("CMISTXUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ibtt{font-family:" & Session("CMIBTTFTFC") & ";color:" & Session("CMIBTTFT") & ";font-size:" & Session("CMIBTTFTSZ") & "pt;background-color:" & Session("CMIBTT") & ";" & replace(Session("CMIBTTBL"),"1","font-weight:bold;") & replace(Session("CMIBTTIT"),"1","font-style:italic;") & replace(Session("CMIBTTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ictt{font-family:" & Session("CMICTTFTFC") & ";color:" & Session("CMICTTFT") & ";font-size:" & Session("CMICTTFTSZ") & "pt;background-color:" & Session("CMICTT") & ";" & replace(Session("CMICTTBL"),"1","font-weight:bold;") & replace(Session("CMICTTIT"),"1","font-style:italic;") & replace(Session("CMICTTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ictx{font-family:" & Session("CMICTXFTFC") & ";color:" & Session("CMICTXFT") & ";font-size:" & Session("CMICTXFTSZ") & "pt;background-color:" & Session("CMICTX") & ";" & replace(Session("CMICTXBL"),"1","font-weight:bold;") & replace(Session("CMICTXIT"),"1","font-style:italic;") & replace(Session("CMICTXUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ictr{font-family:" & Session("CMICTRFTFC") & ";color:" & Session("CMICTRFT") & ";font-size:" & Session("CMICTRFTSZ") & "pt;background-color:" & Session("CMICTR") & ";" & replace(Session("CMICTRBL"),"1","font-weight:bold;") & replace(Session("CMICTRIT"),"1","font-style:italic;") & replace(Session("CMICTRUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.icxr{font-family:" & Session("CMICXRFTFC") & ";color:" & Session("CMICXRFT") & ";font-size:" & Session("CMICXRFTSZ") & "pt;background-color:" & Session("CMICXR") & ";" & replace(Session("CMICXRBL"),"1","font-weight:bold;") & replace(Session("CMICXRIT"),"1","font-style:italic;") & replace(Session("CMICXRUL"),"1","text-decoration:underline;") & "}"

      '''   s = s & "tr.XFK{background-color:#eeeeee;" & XspAlpha(50,100,1) & ";}"

         s = s & "table.tbb{border-style:none;visibility:hidden;}"
       '  s = s & "table.TBCM{WIDTH:'70%';}"

         s = s & "table.tbh2{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";}"
         s = s & "table.tbh{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";background-color:" & Session("CMFMNBOT") & ";" & XspAlpha(0,100,1) & ";}"
         s = s & "table.tbp{border-style:SOLID;border-width:1;border-color:#EEEEEE;" & XspEx(Session("APLBGD"),"#CCCCCC",0) & ";}"       
         s = s & "table.TDCM{border-style:outset;border-width:1;border-color:#ffffff;" & XspEx(Session("APLBGD"),"#CCCCCC",0) & ";}"     
         s = s & "table.TWIND{border-style:solid;border-width:1;border-color:#eeeeee;" & XspEx(Session("APLBGD"),"#AAAAAA",0) & ";}"    
         s = s & "table.TWINE{border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f9f9f9;}"        
         s = s & "table.TWINE2{border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f9f9f9;}"                                            
''' Fin Estilos ................

         s = s & "</style>"
   
      CorStyler2 = s
    

End Function
%>