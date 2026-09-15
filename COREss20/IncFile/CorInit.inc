<%
' ****************************************************************************************
' ** Cor E-Solution 2.2 - Inicializacion acceso a dbases ----------------- Julio, 2001  **
' ****************************************************************************************
 
Function CorInit()

'      session("CMACC") = "L"
'      session("CMRES") = "A"
'      session("USACC") = "L"
'      session("USRES") = "A"

  ' Conset(session("CONSET"))
  ' ofv.wstyles = CorStyler()

   if ofv.ATrusted <> "S" then

      session("LUID") = ofv.ADbUid
      session("LPWD") = ofv.ADbPwd
      ofv.ADbUidl = ofv.ADbUid
      ofv.ADbPwdl = ofv.ADbPwd



   end if 

end function

Function UsrInit()
   if ofv.ATrusted <> "S" then
      set cn = ofv.conectar(ofv.strconn0)

      StrSql = "Select * From Usuarios where ID ='" & session("usuario") & "' "
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rs.eof then
         if cint(rs("dbdef")) = 1 then
            session("LUID") = cstr(rs("dbuid"))
            session("LPWD") = cstr(rs("dbpwd"))
            ofv.ADbUidl = cstr(rs("dbuid"))
            ofv.ADbPwdl = cstr(rs("dbpwd"))
         end if
      end if
      call ofv.cerrarconsulta(rs)
      call ofv.cerrarconn(cn)

   end if 

end function


%>