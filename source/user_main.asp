<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit     = tit_fir & "（" & login_username & "）"
tit_fir = ""

Call web_head(0,0,0,0,0)

If login_mode = "" Then
    Set rs = Nothing
    Call close_conn()
    Response.redirect "login.asp"
    Response.End
End If

'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------

Call user_my()

'---------------------------------center end-------------------------------
Call web_end(0)

Sub user_my()
    Dim tim_login
    sql           = "select l_tim_login from user_login where l_username='" & login_username & "'"
    Set rs        = conn.execute(sql)

    If rs.eof And rs.bof Then
        tim_login = 0
    Else
        tim_login = rs(0)
    End If

    rs.Close

    sql    = "select * from user_data where username='" & login_username & "'"
    Set rs = conn.execute(sql)
    Response.Write ukong & vbcrlf & table1 %>
<tr<% Response.Write table2 %>>
<td colspan=3 background='images/<% = web_var(web_config,5) %>/bar_3_bg.gif'>&nbsp;<% Response.Write img_small(us) %>&nbsp;&nbsp;<font class=end><b>用户个人信息</b></font></td>
</tr>
<tr<% Response.Write table3 %> height=30>
<td width='20%' bgcolor=<% = web_var(web_color,6) %> align=center>用户名称：</td>
<td width='40%'>&nbsp;<font class=blue><b><% Response.Write login_username %></b></font></td>
<td bgcolor=<% = web_var(web_color,6) %>  width='40%' align=center>您已登陆 <font class=red><% Response.Write DateDiff("n",tim_login,Now()) %></font> 分钟</td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %>  align=center>用户类型：</td>
<td>&nbsp;<font class=red_3><% Response.Write format_power(rs("power"),1) %></font></td>
<td rowspan=8 align=center><img src='images/face/<% Response.Write rs("face") %>.gif' border=0></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %>  align=center>用户头衔：</td>
<td>&nbsp;<%
    tit = rs("nname")

    If var_null(tit) = "" Then
        Response.Write "<font color=#c0c0c0>没有</font>"
    Else
        Response.Write "" & code_html(tit,1,0)
    End If %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %>  align=center>来自：</td>
<td>&nbsp;<% Response.Write code_html(rs("whe"),1,0) %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %>  align=center>论坛发贴：</td>
<td>&nbsp;<font class=red><% Response.Write rs("bbs_counter") %></font></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %>  align=center>社区积分：</td>
<td>&nbsp;<font class=red_3><% Response.Write rs("integral") %></font></td>
</tr>

<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %> align=center>用户性别：</td>
<td>&nbsp;<%
    tit = rs("sex")

    If tit = False Then
        Response.Write "<img src='images/small/forum_girl.gif' align=absMiddle border=0>&nbsp;&nbsp;女孩"
    Else
        Response.Write "<img src='images/small/forum_boy.gif' align=absMiddle border=0>&nbsp;&nbsp;男孩"
    End If %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %> align=center>出生年月：</td>
<td>&nbsp;<% Response.Write rs("birthday") %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>用户QQ：</td>
<td>&nbsp;<%
    tit = rs("qq")

    If Not(IsNumeric(tit)) Or Len(tit) < 2 Then
        Response.Write "<font class=gray>没有</font>"
    Else
        Response.Write "<img src='images/small/qq.gif' align=absMiddle border=0>&nbsp;<a href='http://search.tencent.com/cgi-bin/friend/user_show_info?ln=" & tit & "' target=_blank>" & tit & "</a>"
    End If %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %> align=center>E - mail：</td>
<td>&nbsp;<%
    tit      = code_html(rs("email"),1,0)
    Response.Write "<img src='images/small/email.gif' align=absMiddle border=0>&nbsp;<a href='mailto:" & tit & "' title=''>" & tit & "</a>" %></td>
<td bgcolor=<% = web_var(web_color,6) %> align=center><a href='forum_action.asp?action=my'>查看我所参与过的主题</a></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %> align=center>个人主页：</td>
<td colspan=2>&nbsp;<%
    tit      = code_html(rs("url"),1,0)

    If var_null(tit) = "" Then
        Response.Write "<font color=#c0c0c0>没有</font>"
    Else
        Response.Write "<img src='images/small/url.gif' align=absMiddle border=0>&nbsp;<a href='" & tit & "' target=_blank>" & tit & "</a>"
    End If %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td bgcolor=<% = web_var(web_color,6) %> align=center>个人介绍：</td>
<td colspan=2 align=center><table border=0 width='100%' class=tf><tr><td class=bw><% Response.Write code_jk2(rs("remark")) %></td></tr></table></td>
</tr>
<% rs.Close:Set rs = Nothing %>
<tr<% Response.Write table2 %>>
<td colspan=3 background='images/<% = web_var(web_config,5) %>/bar_3_bg.gif'>&nbsp;<% Response.Write img_small(us) %>&nbsp;&nbsp;<font class=end><b>论坛最新发贴</b></font></td>
</tr>
<tr<% Response.Write table3 %>>
<td colspan=3 align=center>
  <table border=0 width='96%'>
<%
    sql = "select top 10 id,forum_id,topic,tim from bbs_topic where username='" & login_username & "' order by id desc"
    Set rs = conn.execute(sql)

    Do While Not rs.eof
        Response.Write "  <tr><td>" & img_small("jt0") & "<a href='forum_view.asp?forum_id=" & rs("forum_id") & "&view_id=" & rs("id") & "' target=_blank>" & code_html(rs("topic"),1,30) & "</a>" & format_end(1,time_type(rs("tim"),8)) & "</td></tr>"
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing %>
  </table>
</td>
</tr>
</table>
<br>
<%
End Sub %>