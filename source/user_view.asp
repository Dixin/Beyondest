<!-- #include file="include/config_user.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim username
Dim view_username
Dim userp
Dim login1
Dim login2
username = code_form(Trim(Request.querystring("username")))
tit      = "�鿴�û���Ϣ��" & username & "��"

Call web_head(2,0,0,0,0)
userp = Int(format_power(login_mode,2))
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Call user_view()
'---------------------------------center end-------------------------------
Call web_end(0)

Sub user_view()
    Dim tim_login
    Dim user_popedom
    Dim user_p
    sql        = "select l_where,l_tim_login from user_login where l_username='" & username & "'"
    Set rs     = conn.execute(sql)

    If rs.eof And rs.bof Then
        login1 = "<font class=gray>���û�����û�е�½����������״̬</font>"
        login2 = login1
    Else
        login1 = "����ʱ�� <font class=red>" & DateDiff("n",rs(1),Now()) & "</font> ����"
        login2 = "��ǰλ�ã�<font class=blue>" & rs(0) & "</font>"
    End If

    rs.Close

    sql    = "select * from user_data where username='" & username & "'"
    Set rs = conn.execute(sql)

    If rs.eof And rs.bof Then
        rs.Close:Set rs = Nothing
        Call close_conn()
        format_redirect("user_main.asp")
        Response.End
    End If

    user_popedom = rs("popedom")
    user_p       = Int(format_power(rs("power"),2))

    If user_p = 3 Then

        If Int(userp) > Int(user_p) Then
            rs.Close:Set rs = Nothing
            Call close_conn()
            Call cookies_type("power")
            Response.End
        End If

    End If

    Response.Write ukong & vbcrlf & table1 %>
<tr<% Response.Write table2 %> height=25>
<td colspan=3 background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small(us) %>&nbsp;&nbsp;<font class=end><b>�鿴�û���Ϣ��<% Response.Write username %>��</b></font></td>
</tr>
<tr<% Response.Write table3 %> height=30>
<td width='20%' align=center bgcolor=<% = web_var(web_color,6) %>>�û����ƣ�</td>
<td width='40%'>&nbsp;<font class=blue><b><% Response.Write username %></b></font>&nbsp;&nbsp;<a href='user_message.asp?action=write&accept_uaername=<% Response.Write Server.urlencode(username) %>'><img src='IMAGES/MAIL/MSG.GIF' border=0 align=absmiddle title='�� <% Response.Write username %> ����վ�ڶ���'></a></td>
<td width='40%' align=center bgcolor=<% = web_var(web_color,6) %>><% Response.Write login1 %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>�û����ͣ�</td>
<td>&nbsp;<font class=red_3><% Response.Write format_power(rs("power"),1) %></font></td>
<td rowspan=8 align=center><img src='images/face/<% Response.Write rs("face") %>.gif' border=0></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>�û�ͷ�Σ�</td>
<td>&nbsp;<%
    tit = rs("nname")

    If var_null(tit) = "" Then
        Response.Write "<font class=gray>û��</font>"
    Else
        Response.Write "" & code_html(tit,1,0)
    End If %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>�������</td>
<td>&nbsp;<% Response.Write code_html(rs("whe"),1,0) %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>��̳������</td>
<td>&nbsp;<font class=red><% Response.Write rs("bbs_counter") %></font></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>�������֣�</td>
<td>&nbsp;<font class=red_4><% Response.Write rs("integral") %></font></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>�û��Ա�</td>
<td>&nbsp;<%
    tit = rs("sex")

    If tit = False Then
        Response.Write "<img src='images/small/forum_girl.gif' align=absmiddle border=0>&nbsp;&nbsp;�ഺŮ��"
    Else
        Response.Write "<img src='images/small/forum_boy.gif' align=absmiddle border=0>&nbsp;&nbsp;�����к�"
    End If %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>�������£�</td>
<td>&nbsp;<% Response.Write rs("birthday") %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>�û��ѣѣ�</td>
<td>&nbsp;<%
    tit = rs("qq")

    If Not(IsNumeric(tit)) Or Len(tit) < 2 Then
        Response.Write "<font class=gray>û��</font>"
    Else
        Response.Write "<img src='images/small/qq.gif' align=absmiddle border=0>&nbsp;<a href='http://search.tencent.com/cgi-bin/friend/user_show_info?ln=" & tit & "' target=_blank>" & tit & "</a>"
    End If %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>����½��</td>
<td>&nbsp;<% Response.Write time_type(rs("last_tim"),88) %></td>
<td align=center bgcolor=<% = web_var(web_color,6) %>><% Response.Write login2 %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>E - mail��</td>
<td colspan=2>&nbsp;<%
    tit = code_html(rs("email"),1,0)
    Response.Write "<img src='images/small/email.gif' align=absmiddle border=0>&nbsp;<a href='mailto:" & tit & "' title=''>" & tit & "</a>" %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>������ҳ��</td>
<td colspan=2>&nbsp;<%
    tit = code_html(rs("url"),1,0)

    If var_null(tit) = "" Then
        Response.Write "<font class=gray>û��</font>"
    Else
        Response.Write "<img src='images/small/url.gif' align=absmiddle border=0>&nbsp;<a href='" & tit & "' target=_blank>" & tit & "</a>"
    End If %></td>
</tr>
<tr<% Response.Write table3 %> height=25>
<td align=center bgcolor=<% = web_var(web_color,6) %>>���˽��ܣ�</td>
<td colspan=2 align=center><table border=0 width='100%' class=tf><tr><td class=bw><% Response.Write code_jk2(rs("remark")) %></td></tr></table></td>
</tr>
<% rs.Close:Set rs = Nothing %>
<tr<% Response.Write table2 %> height=25>
<td colspan=3  background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small(us) %>&nbsp;&nbsp;<font class=end><b>��̳���·���</b>&nbsp;&nbsp;-</font>&nbsp;&nbsp;<a href='forum_action.asp?action=user&username=<% Response.Write Server.urlencode(username) %>' class=menu>�鿴 <% Response.Write username %> �����������</a></td>
</tr>
<tr<% Response.Write table3 %>>
<td colspan=3 align=center>
  <table border=0 width='96%'>
<%
    sql = "select top 10 id,forum_id,topic,tim from bbs_topic where username='" & username & "' order by id desc"
    Set rs = conn.execute(sql)

    Do While Not rs.eof
        Response.Write "  <tr><td>" & img_small("jt0") & "<a href='forum_view.asp?forum_id=" & rs("forum_id") & "&view_id=" & rs("id") & "' target=_blank>" & code_html(rs("topic"),1,30) & "</a>" & format_end(1,time_type(rs("tim"),8)) & "</td></tr>"
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing %>
  </table>
</td>
</tr>
<tr<% Response.Write table3 %>><td colspan=3 height=30 bgcolor=<% = web_var(web_color,6) %>>
&nbsp;&nbsp;�û����������&nbsp;&nbsp;<font class=gray>[<a href='user_isaction.asp?username=<% Response.Write Server.urlencode(username) %>&action=locked<%

    If Int(popedom_format(user_popedom,41)) = 0 Then
        Response.Write "'>����"
    Else
        Response.Write "&cancel=yes' class=red_3>�������"
    End If %></a>]&nbsp;&nbsp;[<a href='user_isaction.asp?username=<% Response.Write Server.urlencode(username) %>&action=shield<%

    If Int(popedom_format(user_popedom,42)) = 0 Then
        Response.Write "'>����"
    Else
        Response.Write "&cancel=yes' class=red_3>�������"
    End If %></a>]</font>
</td></tr>
</table>
<br>
<%
End Sub %>