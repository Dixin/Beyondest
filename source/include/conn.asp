<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

if web_login<>1 and web_login<>2 then
  call cookies_type("loading")
  response.end
end if

'---------------------数据库类型及路径定义---------------------
dim conn,connstr
'connstr="DBQ="&server.mappath(web_var(web_config,6))&";DRIVER={Microsoft Access Driver (*.mdb)};"
connstr="DSN=Beyondest"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr

sub close_conn()
  conn.close
  set conn=nothing
end sub
%>