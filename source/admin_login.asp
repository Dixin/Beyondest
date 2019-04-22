<!-- #include file="include/config.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

If action = "logout" Then
    Session.abandon
    format_redirect("main.asp")
    Response.End
End If

If web_login = 0 Then web_login = 2

If Session("beyondest_online_admin") = "beyondest_admin" And Session("beyondest_online_admines") = login_username Then
    format_redirect("admin.asp")
    Response.End
End If %>
<!-- #include file="include/jk_md5.asp" -->
<!-- #include file="include/conn.asp" -->
<html>
<head>
<title><% Response.Write web_var(web_config,1) %> - 管理后台登陆</title>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel=stylesheet href='include/beyondest.css' type=text/css>
</head>
<body leftmargin=0 topmargin=0 bgcolor=#ededed>
<table border=0 width=100% height=100 %>
<tr>
<td width=100% align=center height=100 %>
<%
Dim achk

If Trim(Request.form("admin_log")) = "ok" Then
    achk = admin_chk()

    If achk = "yes" Then
        close_conn
        Response.redirect "admin.asp"
        Response.End
    Else
        Response.Write admin_login()
    End If

Else
    Response.Write admin_login()
End If

close_conn

Function admin_login() %><table border=0 width=350><tr>
<td align=left height=50 valign=top>
<font class=red><b>管 理 员 登 陆</b></font>
</td>
</tr>
<tr height=25 align=right>
<form action='admin_login.asp' method=post>
<input type=hidden name=admin_log value='ok'>
<td width="30%">
用户名&nbsp;&nbsp;&nbsp;&nbsp;<input type=text name=username value='<% Response.Write login_username %>' size=20>
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
<tr><td align=center><% Response.Write web_var(web_error,4) %></td></tr>
</table><%
End Function

Function admin_chk()
    Dim username
    Dim password
    Dim founderr
    Dim rs
    Dim sql
    Dim id
    Dim power
    Dim hidden
    Dim nname
    Dim face
    username      = Trim(Request.form("username"))
    password      = Trim(Request.form("password"))
    founderr      = "no"

    If symbol_name(username) = "no" Then
        admin_chk = founderr:Exit Function
    End If

    If symbol_ok(password) = "no" Then
        admin_chk = founderr:Exit Function
    End If

    If founderr = "no" Then
        password = jk_md5(password,"short")
        sql      = "select popedom from user_data where username='" & username & "' and password='" & password & "' and power='" & format_power2(1,1) & "' and hidden=1"
        Set rs   = conn.execute(sql)

        If rs.eof And rs.bof Then
            rs.Close:Set rs = Nothing
            admin_chk = founderr:Exit Function
        Else

            If login_username = "" Then
                Response.cookies("beyondest_online")("login_username") = username
                Response.cookies("beyondest_online")("login_password") = password
                Call cookies_yes()
            End If

            Session("beyondest_online_admin") = "beyondest_admin"
            Session("beyondest_online_admines") = username
            Session("beyondest_online_popedom") = rs("popedom")

            rs.Close:Set rs = Nothing
            admin_chk = "yes":Exit Function
        End If

        rs.Close:Set rs = Nothing
    End If

End Function %></td></tr></table>
</body>
</HTML>