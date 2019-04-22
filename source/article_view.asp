<!-- #include file="include/config_article.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim id:id = Trim(Request.querystring("id"))

If Not(IsNumeric(id)) Then
    Call format_redirect("article.asp")
    Response.End
End If %>
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/config_review.asp" -->
<!-- #include file="include/conn.asp" -->
<%
Dim username
Dim topic
Dim word
Dim tim
Dim counter
Dim cname
Dim sname
Dim power
Dim userp
Dim emoney
Dim author
Dim keyes
Dim sql2
sql    = "select * from article where hidden=1 and id=" & id
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing:Call close_conn()
    Call format_redirect("article.asp")
    Response.End
End If

cid      = Int(rs("c_id"))
sid      = Int(rs("s_id"))
username = rs("username")
topic    = rs("topic")
word     = code_jk(rs("word"))
tim      = rs("tim")
counter  = rs("counter")
power    = rs("power")
emoney   = rs("emoney")
author   = rs("author")
keyes    = rs("keyes")
rs.Close

cname = "文章浏览"
sname = ""

If cid > 0 Then

    If sid > 0 Then
        sql2  = "select jk_class.c_name,jk_sort.s_name from jk_sort inner join jk_class on jk_sort.c_id=jk_class.c_id where jk_sort.c_id=" & cid & " and jk_sort.s_id=" & sid
    Else
        sql2  = "select c_name from jk_class where c_id=" & cid
    End If

    Set rs    = conn.execute(sql2)

    If Not (rs.eof And rs.bof) Then
        cname = rs("c_name"):tit = cname
        If sid > 0 Then sname = rs("s_name"):tit = cname & "（" & sname & "）"
    End If

    rs.Close
End If

Call web_head(1,0,2,0,0)

Call emoney_notes(power,emoney,n_sort,id,"js",0,1,"article_list.asp?c_id=" & cid & "&s_id=" & sid)
sql = "update article set counter=counter+1 where id=" & id
conn.execute(sql)
'------------------------------------left----------------------------------
Call font_word_js() %>
<table border=0 width='96%' cellspacing=0 cellpadding=0>
<tr><td align=center height=50><font class=red_3 size=3><b><% Response.Write topic %></b></font></td></tr>
<tr><td align=center class=gray><% Response.Write time_type(tim,33) & "&nbsp;&nbsp;作者：" & author & "&nbsp;&nbsp;" & web_var(web_config,1) %>&nbsp;&nbsp;<% Call font_word_action() %>&nbsp;&nbsp;本文已被浏览&nbsp;<% Response.Write counter %>&nbsp;次</td></tr>
<tr><td height=10></td></tr>
<tr><td valign=top><% Call font_word_type(word) %></td></tr>
<tr><td height=10></td></tr>
<tr><td>
  <table border=0 width='100%'>
  <tr><td width='25%' class=htd>
&nbsp;发布人：<% Response.Write format_user_view(username,1,1) %><br>
&nbsp;<% Response.Write put_type("article") %>
  </td><td width='75%' class=htd>
<%
sql       = "select id,topic from article where hidden=1 and id=" & id - 1
Set rs    = conn.execute(sql)

If rs.eof And rs.bof Then
    topic = "<font class=gray>没有找到相关文章</font>"
Else
    topic = "<a href='article_view.asp?id=" & rs(0) & "'>" & code_html(rs(1),1,30) & "</a>"
End If

rs.Close
Response.Write "上篇文章：" & topic & "<br>"
sql = "select id,topic from article where hidden=1 and id=" & id + 1
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    topic = "<font class=gray>没有找到相关文章</font>"
Else
    topic = "<a href='article_view.asp?id=" & rs(0) & "'>" & code_html(rs(1),1,30) & "</a>"
End If

rs.Close
Response.Write "下篇文章：" & topic %>
  </td></tr></table>
</td></tr>
</table>
<% Call article_view_about() %>
<table border=0 width='96%' cellspacing=0 cellpadding=0 class=tf>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
<td><% Call article_sea() %></td>
<td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td></tr>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
<td><% Call article_view_review() %></td>
<td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
Call web_end(0) %>