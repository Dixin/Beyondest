<!-- #include file="INCLUDE/config_forum.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim rssum,thepages,page,viewpage,sqladd,nummer,forum_temp,pageurl,usern
rssum      = 0:thepages = 0:viewpage = 1:nummer = web_var(web_num,1)
forum_temp = "":pageurl = ""

action     = Trim(Request.querystring("action"))

Select Case action
    Case "hot"
        tit   = "论坛热贴"
        sql   = "select * from bbs_topic where re_counter>10 order by re_counter desc,id desc"
    Case "top"
        tit   = "论坛置顶"
        sql   = "select * from bbs_topic where is" & action & "<>0 order by istop desc,id desc"
    Case "good"
        tit   = "论坛精华"
        sql   = "select * from bbs_topic where is" & action & "=1 order by id desc"
    Case "tim"
        tit   = "回复新贴"
        sql   = "select top 100 * from bbs_topic order by re_tim desc,id desc"
    Case "user"
        usern = Replace(Trim(Request.querystring("username")),"'","")

        If Len(usern) < 1 Then
            Call cookies_type("username")
        End If

        sql    = "select id from user_data where username='" & usern & "'"
        Set rs = conn.execute(sql)

        If rs.eof And rs.bof Then
            rs.Close:Set rs = Nothing
            close_conn
            Call cookies_type("username")
        End If

        rs.Close
        tit     = "查看 " & usern & " 参与过的主题"
        pageurl = "?action=" & action & "&username=" & usern & "&"
        sql     = "select bbs_topic.id,bbs_topic.forum_id,bbs_topic.username,bbs_topic.topic,bbs_topic.tim,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.re_username,bbs_topic.re_tim,bbs_topic.istop,bbs_topic.islock,bbs_topic.isgood " & _
        "from bbs_data inner join bbs_topic on bbs_data.reply_id=bbs_topic.id where bbs_data.username='" & usern & "' group by bbs_topic.id,bbs_topic.forum_id,bbs_topic.username,bbs_topic.topic,bbs_topic.tim,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.re_username,bbs_topic.re_tim,bbs_topic.istop,bbs_topic.islock,bbs_topic.isgood order by bbs_topic.id desc"
    Case "my"
        tit = "我所参与过的主题"
        sql = "select bbs_topic.id,bbs_topic.forum_id,bbs_topic.username,bbs_topic.topic,bbs_topic.tim,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.re_username,bbs_topic.re_tim,bbs_topic.istop,bbs_topic.islock,bbs_topic.isgood " & _
        "from bbs_data inner join bbs_topic on bbs_data.reply_id=bbs_topic.id where bbs_data.username='" & login_username & "' group by bbs_topic.id,bbs_topic.forum_id,bbs_topic.username,bbs_topic.topic,bbs_topic.tim,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.re_username,bbs_topic.re_tim,bbs_topic.istop,bbs_topic.islock,bbs_topic.isgood order by bbs_topic.id desc"
    Case Else
        tit = "论坛新贴"
        sql = "select top 100 * from bbs_topic order by id desc"
End Select

If pageurl = "" Then pageurl = "?action=" & action & "&"

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Response.Write left_action("jt13",4)
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong

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
    forum_temp = forum_temp & forum_view()
    rs.movenext
Next

rs.Close:Set rs = Nothing

Response.Write forum_table1 %>
<tr height=30 bgcolor=<% = web_var(web_color,6) %> align=center>
<td width='75%'><font class=red_3><b><% Response.Write tit %></b></font>&nbsp;&nbsp;&nbsp;
共&nbsp;<font class=red><% Response.Write rssum %></font>&nbsp;贴&nbsp;┋&nbsp;
每&nbsp;<font class=red><% Response.Write nummer %></font>&nbsp;页&nbsp;┋&nbsp;
共&nbsp;<font class=red><% Response.Write thepages %></font>&nbsp;页&nbsp;┋&nbsp;
这是第&nbsp;<font class=red><% Response.Write viewpage %></font>&nbsp;页</td>
</tr>
</table>
<% Response.Write kong & forum_table1 %>
<tr align=center<% Response.Write forum_table2 %> height=25>
<td width='5%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;</td>
<td width='58%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end>论坛主题</font></td>
<td width='14%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end>作者</font></td>
<td width='9%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end>人气</font></td>
<td width='14%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end>最后回复</font></td>
</tr>
<% Response.Write forum_temp %>
</table>
<br>
<% Response.Write forum_table1 %>
<tr height=30 bgcolor=<% = web_var(web_color,6) %>>
<td width='70%'>&nbsp;分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %></td>
<td width='30%' align=center><% Response.Write forum_go() %></td>
</tr>
<tr<% Response.Write forum_table4 %>><td align=center height=30 colspan=2>
<% Response.Write img_small("isok") %>&nbsp;开放的主题&nbsp;&nbsp;
<% Response.Write img_small("ishot") %>&nbsp;回复超过10贴&nbsp;&nbsp;
<% Response.Write img_small("islock") %>&nbsp;锁定的主题&nbsp;&nbsp;
<% Response.Write img_small("istop") %>&nbsp;固定顶端的主题&nbsp;&nbsp;
<% Response.Write img_small("isgood") %>&nbsp;精华帖子
</td></tr>
</table>
<br>
<%
'---------------------------------center end-------------------------------
Call web_end(0)

Function forum_view()
    Dim view_url,topic_head,forumid,id,username,topic,tim,counter,re_counter,re_username,re_tim,istop,islock,isgood,folder_type,reply_count
    folder_type = "isok"
    id          = rs("id")
    username    = rs("username")
    topic       = rs("topic")
    tim         = rs("tim")
    counter     = rs("counter")
    re_counter  = rs("re_counter")
    re_username = rs("re_username")
    re_tim      = rs("re_tim")
    istop       = rs("istop")
    islock      = rs("islock")
    isgood      = rs("isgood")

    Select Case Int(istop)
        Case 1
            folder_type = "istop"
        Case 2
            folder_type = "istops"
        Case Else

            If Int(isgood) = 1 Then
                folder_type = "isgood"
            Else

                If Int(islock) = 1 Then
                    folder_type = "islock"
                ElseIf Int(re_counter) >= 10 Then
                    folder_type = "ishot"
                End If

            End If

    End Select

    forumid = rs("forum_id")
    view_url = "forum_view.asp?forum_id=" & forumid & "&view_id=" & id

    If Int(re_counter) > 0 Then
        topic_head = "<img loaded=no src='images/small/fk_plus.gif' border=0>"
    Else
        topic_head = "<img src='images/small/fk_minus.gif' border=0>"
    End If

    forum_view = vbcrlf & "<tr align=center" & forum_table4 & "><td><img src='images/small/" & folder_type & ".gif' border=0></td>" & _
    vbcrlf & "<td align=left>" & topic_head & "<a href='" & view_url & "' title='主题：" & code_html(topic,1,0) & "<br>发贴时间：" & tim & "<br>最后回复：" & re_username & "<br>回复时间：" & re_tim & "'>" & code_html(topic,1,25) & "</a>&nbsp;" & index_pagecute(view_url,re_counter + 1,web_var(web_num,3),"#cc3300") & "</td>" & _
    vbcrlf & "<td>" & format_user_view(username,1,"") & "</td>" & _
    vbcrlf & "<td class=timtd>" & re_counter & "/" & counter & "</td>" & _
    vbcrlf & "<td>" & format_user_view(re_username,1,"") & "</td></tr>"
End Function %>