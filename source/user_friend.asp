<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit = "我的好友（地址薄）"

Call web_head(2,0,0,0,0)

If Len(action) > 1 And Int(popedom_format(login_popedom,41)) Then Call close_conn():Call cookies_type("locked")
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong
Call user_mail_menu(0)
Response.Write table1 & vbcrlf & "<tr align=center" & table2 & " height=25>"

If action = "del" Then
    Response.Write del_select()
End If

Select Case action
    Case "add"
        Response.Write friend_add()
    Case Else
        Call friend_main()
End Select

Response.Write vbcrlf & "</table>"
'---------------------------------center end-------------------------------
Call web_end(0)

Sub friend_main() %>
<td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif width='7%'><font class=end><b>排序</b></font></td>
<td width='28%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>用户名称</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>类型</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>性别</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>发贴</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>Email</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>QQ</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>主页</b></font></td>
<td width='9%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>发短信</b></font></td>
<td width='8%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>操作</b></font></td>
</tr>
<script language=javascript src='STYLE/admin_del.js'></script>
<form name=del_form action='user_friend.asp?action=del' method=post>
<%
    Dim rs,sql,rssum,i,tname,ttt
    rssum    = 0
    sql      = "select user_data.username,user_data.power,user_data.sex,user_data.bbs_counter,user_data.email,user_data.qq,user_data.url,user_friend.id from user_data inner join user_friend on user_data.username=user_friend.username2 where user_friend.username1='" & login_username & "' order by user_friend.id desc"
    Set rs   = Server.CreateObject("adodb.recordset")
    rs.open sql,conn,1,1

    If Not(rs.eof And rs.bof) Then
        rssum = Int(rs.recordcount)
    End If

    For i = 1 To rssum
        tname = rs("username")
        ttt   = format_power(rs("power"),0)
        Response.Write vbcrlf & "<tr align=center" & table3 & "><td>" & i & ".</td>" & _
        vbcrlf & "<td>" & format_user_view(tname,1,1) & "</td>" & _
        vbcrlf & "<td><img src='images/small/icon_" & ttt & ".gif' title='" & tname & " 是 " & format_power(ttt,1) & "' align=absmiddle border=0></td>"
        ttt     = rs("sex")

        If ttt = False Then
            ttt = "<img src='images/small/forum_girl.gif' title='" & tname & " 是 青春女孩' align=absmiddle border=0>"
        Else
            ttt = "<img src='images/small/forum_boy.gif' title='" & tname & " 是 阳光男孩' align=absmiddle border=0>"
        End If

        Response.Write vbcrlf & "<td>" & ttt & "</td>" & _
        vbcrlf & "<td><font class=red_3>" & rs("bbs_counter") & "</font></td>" & _
        vbcrlf & "<td><a href='mailto:" & rs("email") & "'><img src='images/small/email.gif' title='给 " & tname & " 发电子邮件' align=absMiddle border=0></a></td>"
        ttt     = rs("qq")

        If var_null(ttt) = "" Or ttt = 0 Then
            ttt = "<font class=gray>没有</font>"
        Else
            ttt = "<a href='http://search.tencent.com/cgi-bin/friend/user_show_info?ln=" & ttt & "' target=_blank><img src='images/small/qq.gif' title='查看 " & tname & " 的QQ信息' align=absMiddle border=0></a>"
        End If

        Response.Write vbcrlf & "<td>" & ttt & "</td>"
        ttt     = rs("url")

        If var_null(ttt) = "" Then
            ttt = "<font class=gray>没有</font>"
        Else
            ttt = "<a href='" & ttt & "' target=_blank><img src='images/small/url.gif' title='查看 " & tname & " 的个人主页' align=absMiddle border=0></a>"
        End If

        Response.Write vbcrlf & "<td>" & ttt & "</td>" & _
        vbcrlf & "<td><a href='user_message.asp?action=write&accept_uaername=" & Server.urlencode(tname) & "'><img src='images/mail/msg.gif' border=0 align=absmiddle title='给 " & tname & " 发送站内短信'></a></td>" & _
        vbcrlf & "<td><input type=checkbox name=del_id value='" & rs("id") & "' class=bg_1></td></tr>"
        rs.movenext
    Next %>
<tr><td colspan=10 align=center height=30 bgcolor=<% Response.Write web_var(web_color,5) %>>
共有 <font class=red><% Response.Write rssum %></font> 位好友
　　<input type=button value='添加我的好友' onClick="document.location='user_friend.asp?action=add'">
　　<input type=checkbox name=del_all value=1 onClick="javascript:selectall('<% Response.Write rssum %>');" class=bg_3> 选中所有
　<input type=submit value='删除所选' onclick="return suredel('<% Response.Write rssum %>');">
</td></tr>
<%
End Sub

Function friend_add()
    friend_add = "<td><font class=end><b>添加我的好友</b></font></td></tr>" & _
    vbcrlf & "<tr" & table3 & "><td height=160 align=center>"

    If Trim(Request.form("add_ok")) = "ok" Then
        Dim username2,red,rs,sql
        red         = ""
        username2   = Trim(Request.form("username2"))

        If symbol_name(username2) <> "yes" Then
            red     = "<font class=red>好友名称</font> 为空或不符合相关规则！"
        Else
            sql     = "select username from user_data where username='" & username2 & "'"
            Set rs  = conn.execute(sql)

            If rs.eof And rs.bof Then
                red = "你填写的 <font class=red>好友名称</font> 好像不存在！"
            End If

            rs.Close:Set rs = Nothing
        End If

        If red = "" Then
            Set rs = Server.CreateObject("adodb.recordset")
            sql    = "select * from user_friend where username1='" & login_username & "' and username2='" & username2 & "'"
            rs.open sql,conn,1,3

            If rs.eof And rs.bof Then
                rs.addnew
                rs("username1") = login_username
                rs("username2") = username2
                rs.update
                friend_add     = friend_add & "<font class=red>您已成功的添加了好友（<font class=blue_1>" & username2 & "</font>）！</font>"
            Else
                friend_add     = friend_add & "<font class=red>您已经添加了好友（<font class=blue_1>" & username2 & "</font>）！</font>"
            End If

            rs.Close:Set rs = Nothing
            friend_add = friend_add & "<br><br><a href='user_friend.asp'>点击返回</a>"
        Else
            friend_add = friend_add & red & "<br><br>" & go_back
        End If

    Else
        friend_add = friend_add & "<form action='user_friend.asp?action=add' method=post><input type=hidden name=add_ok value='ok'>好友名称：<input type=text name=username2 value='" & Trim(Request.querystring("add_username")) & "' size=30 maxlength=20><br><br><input type=submit value='添加好友'></form>"
    End If

    friend_add     = friend_add & "</td></tr>"
End Function

Function del_select()
    Dim delid,del_i,del_num,del_dim,del_sql
    delid           = Trim(Request.form("del_id"))

    If var_null(delid) <> "" Then
        delid       = Replace(delid," ","")
        del_dim     = Split(delid,",")
        del_num     = UBound(del_dim)

        For del_i = 0 To del_num
            del_sql = "delete from user_friend where username1='" & login_username & "' and id=" & del_dim(del_i)
            conn.execute(del_sql)
        Next

        Erase del_dim
        del_select = vbcrlf & "<script language=javascript>alert(""好友删除成功！共删除了 " & del_num + 1 & " 位好友。"");</script>"
    End If

End Function %>