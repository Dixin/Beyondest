<!-- #include file="INCLUDE/config_forum.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim sqladd,nummer,user_temp,rssum,viewpage,thepages,page,pageurl
rssum  = 0:thepages = 0:viewpage = 1:nummer = web_var(web_num,1)
sqladd = "":user_temp = ""

Select Case action
    Case "top"
        tit    = "发贴排行"
        sqladd = "bbs_counter desc,id desc"
    Case "emoney"
        tit    = "财富排行"
        sqladd = "emoney desc,id desc"
    Case Else
        tit    = "用户列表"
        sqladd = "id desc"
End Select

pageurl        = "?action=" & action & "&"

Call web_head(1,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Response.Write left_action("jt13",4)
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong

sql    = "select username,power,bbs_counter,sex,email,qq,url,tim,emoney from user_data order by " & sqladd
Set rs = Server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1

If Not(rs.eof And rs.bof) Then
    rssum = rs.recordcount
End If

Call format_pagecute()

If Int(viewpage) > 1 Then
    rs.move (viewpage - 1)*nummer
End If

For i = 1 To nummer
    If rs.eof Then Exit For
    user_temp = user_temp & user_type()
    rs.movenext
Next

rs.Close:Set rs = Nothing

Response.Write forum_table1 %>
<tr height=30<% Response.Write forum_table4 %> align=center>
<td><font class=red_3><b><% Response.Write tit %></b></font>&nbsp;&nbsp;&nbsp;
共&nbsp;<font class=red><% Response.Write rssum %></font>&nbsp;位&nbsp;┋&nbsp;
每&nbsp;<font class=red><% Response.Write nummer %></font>&nbsp;页&nbsp;┋&nbsp;
共&nbsp;<font class=red><% Response.Write thepages %></font>&nbsp;页&nbsp;┋&nbsp;
这是第&nbsp;<font class=red><% Response.Write viewpage %></font>&nbsp;页</td>
</tr>
</table>
<% Response.Write kong & forum_table1 %>
<tr align=center<% Response.Write forum_table2 %> height=25>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>排序</b></font></td>
<td width='27%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>用户名称</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>类型</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>性别</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>发贴</b></font></td>
<td width='6%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>Email</b></font></td>
<td width='6%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>QQ</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>主页</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>短信</b></font></td>
<td width='14%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>注册时间</b></font></td>
</tr>
<% Response.Write user_temp %>
</table>
<br>
<% Response.Write forum_table1 %>
<tr height=30<% Response.Write forum_table3 %>>
<td width='72%'>&nbsp;分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000") %></td>
<td width='28%' align=center><% Response.Write forum_go() %></td>
</tr>
<tr<% Response.Write forum_table4 %>><td align=center height=30 colspan=2><% Response.Write user_power_type(0) %></td></tr>
</table>
<br>
<%
'---------------------------------center end-------------------------------
Call web_end(0)

Function user_type()
    Dim tname,ttt
    tname     = rs("username")
    ttt       = rs("power")
    user_type = vbcrlf & "<tr align=center" & forum_table4 & "><td>" & i + (viewpage - 1)*nummer & ".</td>" & _
    vbcrlf & "<td align=left>" & format_user_view(tname,1,"") & "</td>" & _
    vbcrlf & "<td><img src='images/small/icon_" & ttt & ".gif' title='" & tname & " 是 " & format_power(ttt,1) & "' align=absmiddle border=0></td>"
    ttt       = rs("sex")

    If ttt = False Then
        ttt   = "<img src='images/small/forum_girl.gif' title='" & tname & " 是 青春女孩' align=absmiddle border=0>"
    Else
        ttt   = "<img src='images/small/forum_boy.gif' title='" & tname & " 是 阳光男孩' align=absmiddle border=0>"
    End If

    user_type = user_type & vbcrlf & "<td>" & ttt & "</td>" & _
    vbcrlf & "<td><font class=red>" & rs("bbs_counter") & "</font></td>" & _
    vbcrlf & "<td><a href='mailto:" & rs("email") & "'><img src='images/small/email.gif' title='给 " & tname & " 发电子邮件' align=absmiddle border=0></a></td>" & _
    vbcrlf & "<td>"
    ttt = rs("qq")

    If Not(IsNumeric(ttt)) Or Len(ttt) < 2 Then
        ttt = "<font class=gray>没有</font>"
    Else
        ttt = "<a href='http://search.tencent.com/cgi-bin/friend/user_show_info?ln=" & ttt & "' target=_blank><img src='images/small/qq.gif' title='查看 " & tname & " 的QQ信息' align=absmiddle border=0></a>"
    End If

    user_type = user_type & ttt & "</td>" & vbcrlf & "<td>"
    ttt = rs("url")

    If var_null(ttt) = "" Then
        ttt = "<font class=gray>没有</font>"
    Else
        ttt = "<a href='" & ttt & "' target=_blank><img src='images/small/url.gif' title='查看 " & tname & " 的个人主页' align=absmiddle border=0></a>"
    End If

    user_type = user_type & ttt & "</td><td><a href='user_message.asp?action=write&accept_uaername=" & Server.urlencode(tname) & "'><img src='images/mail/msg.gif' border=0 align=absmiddle title='给 " & tname & " 发送站内短信'></a></td>" & vbcrlf & "<td align=left>" & time_type(rs("tim"),3) & "</td></tr>"
End Function %>