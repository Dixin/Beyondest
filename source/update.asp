<!-- #include file="include/config_other.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim id
id            = Trim(Request.querystring("id"))

If action = "forum" Then
    index_url = "forum"
    tit_fir   = format_menu(index_url)
    tit       = "论坛公告"
Else
    tit       = "网站更新"
    action    = "news"
End If

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Response.Write left_action("jt13",4)
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong

If IsNumeric(id) Then
    Call update_view()
Else
    Call update_main()
End If

Response.Write kong
'---------------------------------center end-------------------------------
Call web_end(0)

Sub update_main() %>
<table border=0 width='98%' cellspacing=2 cellpadding=2>
<tr><td colspan=2 height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr bgcolor=<% Response.Write web_var(web_color,5) %> valign=bottom height=20>
<% If action = "news" Then %>
<td width='70%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=news' class=red_3>网站更新</a></b></td>
<td width='30%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=forum' class=red_3>论坛公告</a></b></td>
<% Else %>
<td width='70%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=forum' class=red_3>论坛公告</a></b></td>
<td width='30%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=news' class=red_3>网站更新</a></b></td>
<% End If %>
</tr>
<tr><td colspan=2 height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr valign=top>
<td><% Response.Write update_top("jt0",action,20,15,1,2) %></td>
<td><%

    If action = "news" Then
        Response.Write update_top("jt0","forum",5,6,1,1)
    Else
        Response.Write update_top("jt0","news",5,6,1,1)
    End If %></td>
</tr>
<tr>
<td></td>
<td></td>
</tr>
</table>
<%
End Sub

Sub update_view()
    sql    = "select * from bbs_cast where id=" & id
    Set rs = conn.execute(sql)

    If rs.eof And rs.bof Then
        rs.Close
        Call update_main()

        Exit Sub
        End If %>
<table border=0 width='96%'>
<tr><td align=center height=40><font class=blue size=3><b><% Response.Write rs("topic") %></b></font></td></tr>
<tr><td align=center class=gray><% Response.Write web_var(web_config,1) %>&nbsp;&nbsp;发布人：<% Response.Write format_user_view(rs("username"),1,"") %>&nbsp;&nbsp;发布时间：<% Response.Write time_type(rs("tim"),88) %></td></tr>
<tr><td height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr><td align=center>
  <table border=0 width='96%'>
  <tr><td class=htd><% Response.Write code_jk(rs("word")) %></td></tr>
  </table>
</td></tr>
</table>
<%
        rs.Close %>
<br>
<table border=0 width='96%' cellspacing=0 cellpadding=2>
<tr><td colspan=2 height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr bgcolor=<% Response.Write web_var(web_color,5) %> valign=bottom height=20>
<td width='50%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=news' class=red_3>网站更新</a></b></td>
<td width='50%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=forum' class=red_3>论坛公告</a></b></td>
</tr>
<tr><td colspan=2 height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr valign=top>
<td><% Response.Write update_top("jt0","news",5,15,1,2) %></td>
<td><% Response.Write update_top("jt0","forum",5,15,1,2) %></td>
</tr>
<tr>
<td></td>
<td></td>
</tr>
</table>
<%
    End Sub

    Function update_top(u_jt,ut,u_num,c_num,et,timt)
        Dim temp1
        Dim topic
        If u_jt <> "" Then u_jt = img_small(u_jt)
        temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
        sql       = "select top " & u_num & " id,topic,tim from bbs_cast where sort='" & ut & "' order by id desc"
        Set rs    = conn.execute(sql)

        Do While Not rs.eof
            topic = rs("topic")
            temp1 = temp1 & vbcrlf & "<tr><td>" & u_jt & "<a href='update.asp?action=" & ut & "&id=" & rs("id") & "' title='" & code_html(topic,1,0) & "'>" & code_html(topic,1,c_num) & "</a>" & format_end(et,"<font class=gray>" & time_type(rs("tim"),timt) & "</font>") & "</td></tr>"
            rs.movenext
        Loop

        rs.Close
        temp1 = temp1 & "</table>"
        update_top = temp1
    End Function %>