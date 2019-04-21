<!-- #include file="include/config.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim forumid
Dim viewid
Dim html_temp
Dim forumtype
html_temp     = ""
forumid       = Trim(Request.querystring("forum_id"))
viewid        = Trim(Request.querystring("view_id"))

If Not(IsNumeric(forumid)) Or Not(IsNumeric(viewid)) Then
    html_temp = "<tr><td><font class=red_2>您的操作有误：ID 出錯（1）！</font></td></tr>"
End If

If var_null(login_username) = "" Or var_null(login_password) = "" Then
    html_temp     = "<tr><td>&nbsp;&nbsp;" & web_var(web_error,2) & "</td></tr>"
Else
    sql           = "select forum_type from bbs_forum where forum_id=" & forumid & " and forum_hidden=0"
    Set rs        = conn.execute(sql)

    If rs.eof And rs.eof Then
        html_temp = html_temp & "<tr><td><font class=red_2>ForumID 出錯！</font></td></tr>"
    End If

    rs.Close:Set rs = Nothing

    If html_temp = "" Then
        sql           = "select username,word from bbs_data where forum_id=" & forumid & " and reply_id=" & viewid & " order by id desc"
        Set rs        = conn.execute(sql)

        If rs.eof And rs.bof Then
            html_temp = html_temp & "<tr><td><font class=red_2>您的操作有误：ID 出錯（2）！</font></td></tr>"
        Else

            Do While Not rs.eof
                html_temp = html_temp & "<tr><td><img src=""images/small/fk_minus.gif"" border=0> " & code_html(rs("word"),1,45) & "&nbsp;<font class=gray>-</font>&nbsp;" & Replace(format_user_view(rs("username"),1,0),"'","""") & "</td></tr>"
                rs.movenext
            Loop

        End If

        rs.Close:Set rs = Nothing
    End If

End If

Call close_conn()

html_temp = "<table border=0 width=99% align=right cellspacing=2 cellpadding=0>" & html_temp & "</table>" %>
<script language=javascript>
parent.followTd<% = viewid %>.innerHTML='<% = html_temp %>';
</script>