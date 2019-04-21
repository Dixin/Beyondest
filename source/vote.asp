<!--#include file="include/config.asp"-->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim id,types,mcolor,bgcolor,counter,ttt,j,c,w,h:j=0:c=0
w=350:h=220
id=trim(request.querystring("id"))
if not(isnumeric(id)) then id=0
types=trim(request.querystring("types"))
if not(isnumeric(types)) then types=1
if types=1 then
  ttt="radio"
else
  ttt="checkbox"
end if
mcolor=code_form(request.querystring("mcolor"))
if len(mcolor)<>6 then mcolor="CC3300"
bgcolor="#"&code_form(request.querystring("bgcolor"))
if len(bgcolor)<>7 then bgcolor=web_var(web_color,6)

sql="select id,vname,counter from vote where vid="&id&" order by id"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  rs.close:set rs=nothing
  call close_conn()
  response.write "document.write(""没有此调查列表！"");"
  response.end
end if

response.write vbcrlf&"document.write(""<table border=0 cellspacing=0 cellpadding=2>"");"
response.write vbcrlf&"document.write(""<form action='votetype.asp?action=save&vid="&id&"' method=POST target='vote_view'>"");"
do while not rs.eof
  if j=0 then
    response.write vbcrlf&"document.write(""<tr><td align=center height=25><font color=#"&mcolor&"><b>"&code_html(rs("vname"),1,0)&"</b></font></td></tr>"");"
  else
    response.write vbcrlf&"document.write(""<tr><td><input type="&ttt&" name=vote_id value='"&rs("id")&"' style='background-color:"&web_var(web_color,6)&"'>"&code_html(rs("vname"),1,0)&"</td></tr>"");"
  end if
  j=j+1:c=c+rs("counter")
  rs.movenext
loop
rs.close:set rs=nothing
response.write vbcrlf&"document.write(""<tr><td align=center height=25><input onclick=\""javascript:open_win('','vote_view',"&w&","&h&",'no');\"" type=submit value='投票'>&nbsp;&nbsp;<a href='javascript:;' onclick=\""javascript:open_win('votetype.asp?action=view&vid="&id&"','vote_view',"&w&","&h&",'no');\"">查看结果</a><font class=gray>(共<font class=blue>"&c&"</font>票)</font></td></tr>"");"
response.write vbcrlf&"document.write(""</form></table>"");"

call close_conn()
%>