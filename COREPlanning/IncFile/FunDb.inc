<%

Public Conn
public  strconn
strconn1="dsn=DSNCOR01CEP"
strconn2="dsn=DSNCOR02CEP"

function conectar (strconn)

  set conn = server.createobject("adodb.connection")
  conn.open strconn
  set conectar=conn
end function


function CrearConsultaCn(strsql)
dim rs
  set rs=conn.execute(strsql)
  set crearconsultaCn=rs
  
end function

function CrearConsulta(strsql,strconn)

dim rs
 set rs=server.createobject("ador.recordset")

 rs.open strsql,strconn
 set crearconsulta=rs

end function

function CrearConsultaEx(byref strsql,byref conn,TipoConsulta, byref a) 
dim t
dim rs
dim errloop
dim errores
 set rs=server.createobject("adodb.command")



  rs.activeconnection = conn
  rs.commandtext = strsql
  rs.commandtype = TipoConsulta
 

select case TipoConsulta
 Case 1
   on error resume next
   set crearconsultaEx= rs.execute() 
   set errores=conn.errors
   if errores.count > 0 then
'     response.write "<script> window.open('../asp/errordb.asp?volver=../PerfilesdeEsquemaNormativo.asp', 'displayWindow','width=255,height=100,status=no,toolbar=no,menubar=no');</script>"
   '   response.redirect "../asp/errordb.asp?volver="& a
     '  response.write errores(1).description
   end if 
 Case 4
   convertirArregloParam parametros,carblanco," "
   set crearconsultaEx= rs.execute( ,a) 
End Select
end function

sub cerrarconsulta(byref consulta)
   consulta.close
   set consulta = nothing
end sub

function ConsultaVacia(Byref RsRecordset)
  ConsultaVacia =(RsRecordset.Eof and RsRecordset.bof)
end function

sub cerrarconn(byref conn)
 conn.close
 set conn = nothing
end sub
%>