<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

'---------------------数据库类型及路径定义---------------------
dim conn,connstr
connstr="DBQ="+server.mappath("data/ip_address.mdb")+";DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr

sub close_conn()
  conn.close
  set conn=nothing
end sub
%>