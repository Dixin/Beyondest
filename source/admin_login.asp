<!-- #include file="include/config.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

if action="logout" then
  session.abandon
  format_redirect("main.asp")
  response.end
end if

if web_login=0 then web_login=2
if session("beyondest_online_admin")="beyondest_admin" and session("beyondest_online_admines")=login_username then
  format_redirect("admin.asp")
  response.end
end if
%>
<!-- #include file="include/jk_md5.asp" -->
<!-- #include file="include/conn.asp" -->
<html>
<head>
<title><%response.write web_var(web_config,1)%> - 管理后台登陆</title>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel=stylesheet href='include/beyondest.css' type=text/css>
</head>
<body leftmargin=0 topmargin=0 bgcolor=#ededed>
<table border=0 width=100% height=100%>
<tr>
<td width=100% align=center height=100%>
<%
dim achk
if trim(request.form("admin_log"))="ok" then
  achk=admin_chk()
  if achk="yes" then
    close_conn
    response.redirect "admin.asp"
    response.end
  else
    response.write admin_login()
  end if
else
  response.write admin_login()
end if

close_conn

function admin_login()
%><table border=0 width=350><tr>
<td align=left height=50 valign=top>
<font class=red><b>管 理 员 登 陆</b></font>
</td>
</tr>
<tr height=25 align=right>
<form action='admin_login.asp' method=post>
<input type=hidden name=admin_log value='ok'>
<td width="30%">
用户名&nbsp;&nbsp;&nbsp;&nbsp;<input type=text name=username value='<%response.write login_username%>' size=20>
</td></tr>
<tr height=25 align=right>
<td width="30%">
密&nbsp;&nbsp;&nbsp;码&nbsp;&nbsp;&nbsp;&nbsp;<input type=password name=password size=20 maxlength=20>
</td></tr>
<tr><td align=right height=30>
<input type=submit value="确 定">&nbsp;&nbsp;&nbsp;
<input type=button value="取 消">
</td></form></tr>
<tr><td align=center height=60 align=bottom>
<font class=red_4>本次登陆在无活动状态20分钟后将自动注销</font>
</td></tr>
<tr><td align=center><%response.write web_var(web_error,4)%></td></tr>
</table><%
end function

function admin_chk()
  dim username,password,founderr,rs,sql,id,power,hidden,nname,face
  username=trim(request.form("username"))
  password=trim(request.form("password"))
  founderr="no"
  if symbol_name(username)="no" then
    admin_chk=founderr:exit function
  end if
  if symbol_ok(password)="no" then
    admin_chk=founderr:exit function
  end if
  
  if founderr="no" then
    password=jk_md5(password,"short")
    sql="select popedom from user_data where username='"&username&"' and password='"&password&"' and power='"&format_power2(1,1)&"' and hidden=1"
    set rs=conn.execute(sql)
    if rs.eof and rs.bof then
      rs.close:set rs=nothing
      admin_chk=founderr:exit function
    else
      if login_username="" then
        response.cookies("beyondest_online")("login_username")=username
        response.cookies("beyondest_online")("login_password")=password
        call cookies_yes()
      end if
      session("beyondest_online_admin")="beyondest_admin"
      session("beyondest_online_admines")=username
      session("beyondest_online_popedom")=rs("popedom")
      
      rs.close:set rs=nothing
      admin_chk="yes":exit function
    end if
    rs.close:set rs=nothing
  end if
end function
%></td></tr></table>
</body>
</HTML>