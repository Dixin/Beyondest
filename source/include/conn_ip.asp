<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
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