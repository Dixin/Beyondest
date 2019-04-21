<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim tit_menu
Dim popedoms
Dim usernames
Dim username
Dim frm_view
Dim popedom
tit_menu = "<a href='?'>权限管理</a>"
Response.Write header (6,tit_menu)

frm_view      = "no"
username      = Trim(Request.querystring("username")) %>
<table border=1 width=400 cellspacing=0 cellpadding=2<% Response.Write table1 %>>
<tr align=center bgcolor=#ffffff>
<td width='30%' class=red_3>现任管理员</td>
<td width='35%' class=red_3>系统管理权限</td>
<td width='35%' class=red_3>版面管理权限</td>
</tr>
<tr align=center valign=top><td>
<table border=0>
<%
sql           = "select username,popedom from user_data where power='" & format_power2(1,1) & "' order by id"
Set rs        = conn.execute(sql)

Do While Not rs.eof
    usernames = rs("username")
    Response.Write vbcrlf & "<tr><td align=center><a href='?username=" & Server.urlencode(usernames) & "'"

    If username = usernames Then
        popedoms = rs("popedom")
        popedom  = Right(popedoms,30)
        Response.Write " class=red"
        frm_view = "yes"
    End If

    Response.Write ">" & usernames & "</a></td></tr>"
    rs.movenext
Loop

rs.Close:Set rs = Nothing

If frm_view = "yes" And Trim(Request.querystring("chk")) = "yes" Then
    popedom  = popedom_frm(1) & popedom_frm(2) & popedom_frm(3) & popedom_frm(4) & popedom_frm(5) & popedom_frm(6) & popedom_frm(7) & popedom_frm(8) & popedom_frm(9) & popedom_frm(10) & _
    popedom_frm(11) & popedom_frm(12) & popedom_frm(13) & popedom_frm(14) & popedom_frm(15) & popedom_frm(16) & popedom_frm(17) & popedom_frm(18) & popedom_frm(19) & popedom_frm(20) & popedom
    popedoms = popedom
    sql      = "update user_data set popedom='" & popedom & "' where username='" & username & "' and power='" & format_power2(1,1) & "'"
    conn.execute(sql)
    Response.Write "<script language=javascript>alert(""" & username & " 的权限修改成功！"");</script>"
End If %>
</table>
</td>
<%

If frm_view = "yes" Then
    Response.Write "<form action='?username=" & Server.urlencode(username) & "&chk=yes' method=post>"
End If %>
<td>
<table border=0>
<tr><td><input type=checkbox name=popedom_cb1 value='1'<% If popedom_formated(popedoms,1,0) = 1 Then Response.Write " checked" %>></td><td>用户管理</td></tr>
<tr><td><input type=checkbox name=popedom_cb2 value='1'<% If popedom_formated(popedoms,2,0) = 1 Then Response.Write " checked" %>></td><td>执行SQL</td></tr>
<tr><td><input type=checkbox name=popedom_cb3 value='1'<% If popedom_formated(popedoms,3,0) = 1 Then Response.Write " checked" %>></td><td>配置修改</td></tr>
<input type=hidden name=popedom_cb4 value='0'>
<tr><td><input type=checkbox name=popedom_cb5 value='1'<% If popedom_formated(popedoms,5,0) = 1 Then Response.Write " checked" %>></td><td>分类管理</td></tr>
<tr><td><input type=checkbox name=popedom_cb6 value='1'<% If popedom_formated(popedoms,6,0) = 1 Then Response.Write " checked" %>></td><td>权限管理</td></tr>
<tr><td><input type=checkbox name=popedom_cb7 value='1'<% If popedom_formated(popedoms,7,0) = 1 Then Response.Write " checked" %>></td><td>更新公告</td></tr>
<tr><td><input type=checkbox name=popedom_cb8 value='1'<% If popedom_formated(popedoms,8,0) = 1 Then Response.Write " checked" %>></td><td>调查管理</td></tr>
<tr><td><input type=checkbox name=popedom_cb9 value='1'<% If popedom_formated(popedoms,9,0) = 1 Then Response.Write " checked" %>></td><td>上传管理</td></tr>
<input type=hidden name=popedom_cb10 value='0'>
</table>
</td><td>
<table border=0>
<tr><td><input type=checkbox name=popedom_cb11 value='1'<% If popedom_formated(popedoms,11,0) = 1 Then Response.Write " checked" %>></td><td>论坛管理</td></tr>
<tr><td><input type=checkbox name=popedom_cb12 value='1'<% If popedom_formated(popedoms,12,0) = 1 Then Response.Write " checked" %>></td><td>行业动态</td></tr>
<tr><td><input type=checkbox name=popedom_cb13 value='1'<% If popedom_formated(popedoms,13,0) = 1 Then Response.Write " checked" %>></td><td>文栏管理</td></tr>
<tr><td><input type=checkbox name=popedom_cb14 value='1'<% If popedom_formated(popedoms,14,0) = 1 Then Response.Write " checked" %>></td><td>下载栏目</td></tr>
<tr><td><input type=checkbox name=popedom_cb15 value='1'<% If popedom_formated(popedoms,15,0) = 1 Then Response.Write " checked" %>></td><td>网站推荐</td></tr>
<tr><td><input type=checkbox name=popedom_cb16 value='1'<% If popedom_formated(popedoms,16,0) = 1 Then Response.Write " checked" %>></td><td>图库管理</td></tr>
<tr><td><input type=checkbox name=popedom_cb17 value='1'<% If popedom_formated(popedoms,17,0) = 1 Then Response.Write " checked" %>></td><td>友情链接</td></tr>
<input type=hidden name=popedom_cb18 value='0'>
<input type=hidden name=popedom_cb19 value='0'>
<input type=hidden name=popedom_cb20 value='0'>
</table>
</td>
</tr>
<tr><td colspan=3 align=center height=30>
<% If frm_view = "yes" Then %>
<input type=submit value='提 交 修 改'>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=reset value=' 重 置 '>
</td></form>
<% Else %>
<font class=red>请点选左边的管理员以进行下一步操作</font></td>
<% End If %>
</tr>
</table>
<%

close_conn
Response.Write ender()

Function popedom_frm(pnums)
Dim pnum:pnum = Trim(Request.form("popedom_cb" & pnums))
If Not(IsNumeric(pnum)) Then pnum = 0
If Int(pnum) <> 0 Then pnum = 1
popedom_frm = pnum
End Function %>