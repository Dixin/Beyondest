<!-- #include file="include/onlogin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim id:id = Trim(Request.querystring("id"))

If Not(IsNumeric(id)) Then
    Response.redirect "admin_user_list.asp"
    Response.End
End If %>
<!-- #include file="include/conn.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/jk_md5.asp" -->
<%
Dim admin_menu
Dim udim
Dim unum
admin_menu     = "<a href='admin_user_list.asp'>用户管理</a>　　┋"
udim           = Split(user_power,"|"):unum = UBound(udim) + 1

For i = 0 To unum - 1
    admin_menu = admin_menu & "<a href='admin_user_list.asp?power=" & Left(udim(i),InStr(udim(i),":") - 1) & "'>" & Right(udim(i),Len(udim(i)) - InStr(udim(i),":")) & "</a>┋"
Next

admin_menu     = admin_menu & "　　<a href='admin_user_list.asp?hidden=true'>正常用户</a>┋" & _
"<a href='admin_user_list.asp?hidden=false'>锁定用户</a>"

Response.Write header(1,admin_menu) %>
<table border=0 width='98%' cellspacing=0 cellpadding=2 align=center>
<tr><td align=center height=350>
<%
Set rs = Server.CreateObject("adodb.recordset")
sql    = "select * from user_data where id=" & id
rs.open sql,conn,1,3

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing
    Call close_conn()
    Response.redirect "admin_user_list.asp"
    Response.End
End If

If rs("username") = "笼民" Then
    rs.Close:Set rs = Nothing
    Call close_conn()
    Response.redirect "admin_user_list.asp"
    Response.End
End If

If Trim(Request("edit")) = "ok" Then
    Response.Write user_chk()
Else
    Response.Write user_type()
End If

rs.Close:Set rs = Nothing
Call close_conn() %>
</td></tr></table>
<%
Response.Write ender()

Function user_chk()
    Dim password
    Dim password2
    Dim passwd
    Dim passwd2
    Dim bbs_counter
    Dim counter
    Dim integral
    Dim emoney
    Dim power
    Dim hidden
    password           = Trim(Request.form("password"))
    password2          = Trim(Request.form("password2"))
    passwd             = Trim(Request.form("passwd"))
    passwd2            = Trim(Request.form("passwd2"))
    power              = Trim(Request.form("power"))
    hidden             = Trim(Request.form("hidden"))

    If password <> password2 Then
        rs("password") = jk_md5(password,"short")
    End If

    If passwd <> passwd2 Then
        rs("passwd") = jk_md5(passwd,"short")
    End If

    bbs_counter        = Trim(Request.form("bbs_counter"))
    counter            = Trim(Request.form("counter"))
    integral           = Trim(Request.form("integral"))
    emoney             = Trim(Request.form("emoney"))
    '-2147483648 +2147483647

    If IsNumeric(bbs_counter) Then
        bbs_counter        = Int(bbs_counter)

        If bbs_counter <> Int(Request.form("bbs_counter2")) And bbs_counter > 0 And bbs_counter <= 2147483647 Then
            rs("bbs_counter") = bbs_counter
        End If

    End If

    If IsNumeric(counter) Then
        counter            = Int(counter)

        If counter <> Int(Request.form("counter2")) And counter > 0 And counter <= 2147483647 Then
            rs("counter") = counter
        End If

    End If

    If IsNumeric(integral) Then
        integral           = Int(integral)

        If integral <> Int(Request.form("integral2")) And integral > 0 And integral <= 2147483647 Then
            rs("integral") = integral
        End If

    End If

    If IsNumeric(emoney) Then
        emoney             = Int(emoney)

        If emoney <> Int(Request.form("emoney2")) And emoney > 0 And emoney <= 2147483647 Then
            rs("emoney") = emoney
        End If

    End If

    rs("power") = power
    rs("hidden") = hidden
    rs.update
    Response.Write "<font class=red>用户信息修改成功！</font><br><br><a href='admin_user_list.asp'>点击返回</a>"
End Function

Function user_type() %>
<table border=0 width=300>
<form action='admin_user_edit.asp?edit=ok&id=<% Response.Write id %>' method=post>
  <tr>
    <td colspan=2 align=center height=50><font class=red>用户管理修改</font></td>
  </tr>
  <tr>
    <td width='30%'>用户名称：</td>
    <td width='70%'><input type=text value='<% Response.Write rs("username") %>' readonly size=25></td>
  </tr>
  <tr>
    <td>用户密码：</td>
    <td><input type=text name=password value='<% Response.Write rs("password") %>' size=25 maxlength=20><input type=hidden name=password2 value='<% Response.Write rs("password") %>'></td>
  </tr>
  <tr>
    <td>密码钥匙：</td>
    <td><input type=text name=passwd value='<% Response.Write rs("passwd") %>' size=25 maxlength=20><input type=hidden name=passwd2 value='<% Response.Write rs("passwd") %>'></td>
  </tr>
  <tr>
    <td>论坛发贴：</td>
    <td><input type=text name=bbs_counter value='<% Response.Write rs("bbs_counter") %>' size=15 maxlength=10></td>
  </tr><input type=hidden name=bbs_counter2 value='<% Response.Write rs("bbs_counter") %>'>
  <tr>
    <td>文栏发贴：</td>
    <td><input type=text name=counter value='<% Response.Write rs("counter") %>' size=15 maxlength=10></td>
  </tr><input type=hidden name=counter2 value='<% Response.Write rs("counter") %>'>
  <tr>
    <td>用户积分：</td>
    <td><input type=text name=integral value='<% Response.Write rs("integral") %>' size=15 maxlength=10></td>
  </tr><input type=hidden name=integral2 value='<% Response.Write rs("integral") %>'>
  <tr>
    <td>用户金钱：</td>
    <td><input type=text name=emoney value='<% Response.Write rs("emoney") %>' size=15 maxlength=10></td>
  </tr><input type=hidden name=emoney2 value='<% Response.Write rs("emoney") %>'>
  <tr>
    <td>用户类型：</td>
    <td><select name=power size=1><%
    Dim power
    Dim pi
    Dim hidden
    Dim h1
    Dim h2
    power = rs("power")

    For pi = 1 To unum
        Response.Write vbcrlf & "<option value='" & format_power2(pi,1) & "'"
        If power = format_power2(pi,1) Then Response.Write " selected"
        Response.Write ">" & format_power2(pi,2) & "</option>"
    Next %></select>（<% Response.Write power %>）</td>
  </tr>
  <tr>
    <td>类型状态：</td>
    <td><%
    hidden = rs("hidden")

    If hidden = True Then
        h1 = " checked"
        h2 = ""
    Else
        h1 = ""
        h2 = " checked"
    End If %><input type=radio name=hidden value=true<% Response.Write h1 %>>正常<input type=radio name=hidden value=false<% Response.Write h2 %>>锁定</td>
  </tr>
  <tr>
    <td colspan=2 align=center height=30><input type=submit value=' 提 交 修 改 '></td>
  </tr>
</form>
</table>
<%
End Function %>