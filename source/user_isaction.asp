<!-- #include file="include/config.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim cancel,old_url,username
cancel=trim(request.querystring("cancel"))
old_url=request.servervariables("http_referer")
if len(old_url)<3 then old_url="user_main.asp"
username=trim(request.querystring("username"))
if symbol_name(username)<>"yes" or (action<>"locked" and action<>"shield") then
  response.redirect old_url
  response.end
end if
%>
<!-- #include file="include/skin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
call web_head(2,2,0,0,0)
if format_user_power(login_username,login_mode,"")<>"yes" then call close_conn():call cookies_type("power")

sql="select power,popedom from user_data where username='"&username&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
response.write username
response.end
  rs.close:set rs=nothing
  call close_conn()
  response.redirect old_url
  response.end
end if
dim user_popedom,u_power,aname,fname,popedom_true
u_power=rs("power")
user_popedom=rs("popedom")
rs.close:set rs=nothing

if int(format_power(u_power,2))=1 then
  call close_conn()
  call cookies_type("power")
  response.end
end if

popedom_true="yes"
if cancel="yes" then fname="���"

select case action
case "shield"
  aname="����"
  call useres_popedom(42)
case "locked"
  aname="����"
  call useres_popedom(41)
end select

call useres_msg()

call close_conn()
'response.redirect old_url
'response.end

sub useres_popedom(pn)
  dim temp1,temp2,temp3
  if len(user_popedom)<>50 or pn>len(user_popedom) then popedom_true="no":exit sub
  temp1=left(user_popedom,pn-1)
  temp2=popedom_format(user_popedom,pn)
  temp3=right(user_popedom,len(user_popedom)-pn)
  if cancel="yes" then
    temp2="0"
  else
    temp2="1"
  end if
  sql="update user_data set popedom='"&temp1&temp2&temp3&"' where username='"&username&"'"
  conn.execute(sql)
end sub

sub useres_msg()
  if popedom_true="yes" then
    response.write "<script language=javascript>alert(""�ѳɶ��û���"&username&"�����������²�����\n\n"&fname&" "&aname&"\n\n������أ�"");location.href='"&old_url&"';</script>"
  else
    response.write "<script language=javascript>alert(""�ڶ��û���"&username&"�����в���ʱ���������ش���\n\n����վ����ϵ��\n\n������أ�"");location.href='"&old_url&"';</script>"
  end if
end sub
%>