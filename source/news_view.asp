<!-- #include file="INCLUDE/config_news.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim id:id = Trim(Request.querystring("id"))

If Not(IsNumeric(id)) Then
    Call format_redirect("news.asp")
    Response.End
End If %>
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="INCLUDE/config_review.asp" -->
<!-- #include file="include/conn.asp" -->
<%
Dim username,word,tim,counter,cname,sname,comto,pic,tt,sql2
sql    = "select * from news where hidden=1 and id=" & id
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing:close_conn
    Server.transfer "news.asp"
    Response.End
End If

cid      = rs("c_id")
sid      = rs("s_id")
username = rs("username")
topic    = rs("topic")
word     = code_jk(rs("word"))
tim      = rs("tim")
counter  = rs("counter")
ispic    = rs("ispic")
comto    = rs("comto")
pic      = rs("pic")
rs.Close

cname = "新闻浏览":sname = ""

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

Call web_head(0,0,1,0,0)

sql = "update news set counter=counter+1 where id=" & id
conn.execute(sql)
'-----------------------------------center---------------------------------
Call font_word_js() %>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center height=50><font class=red_3 size=3><b><% Response.Write topic %></b></font></td></tr>
<tr><td align=center class=gray><% Response.Write time_type(tim,33) & "&nbsp;" & web_var(web_config,1) %>&nbsp;<% Call font_word_action() %>&nbsp;出处：<% Response.Write comto %></td></tr>
<tr><td height=10></td></tr>
<tr><td valign=top><%

If ispic = True Then
    word = "<p align=center><img src='" & url_true(web_var(web_down,5) & "/",pic) & "' border=0 onload=""javascript:if(this.width>screen.width-430)this.width=screen.width-430""></p>" & word
End If

Call font_word_type(word & "&nbsp;<font class=gray>（本文已被浏览&nbsp;" & counter & "&nbsp;次）</font>") %></td></tr>
<tr><td height=10></td></tr>
<tr><td>
  <table border=0 width='100%'>
  <tr><td width='25%' class=htd>
&nbsp;发布人：<% Response.Write format_user_view(username,1,1) %><br>
&nbsp;<% Response.Write put_type("news") %>
  </td><td width='75%' class=htd>
<%
sql       = "select id,topic from news where hidden=1 and id=" & id - 1
Set rs    = conn.execute(sql)

If rs.eof And rs.bof Then
    topic = "<font class=gray>没有找到相关新闻</font>"
Else
    topic = "<a href='news_view.asp?id=" & rs(0) & "'>" & code_html(rs(1),1,25) & "</a>"
End If

rs.Close
Response.Write "上篇新闻：" & topic & "<br>"
sql = "select id,topic from news where hidden=1 and id=" & id + 1
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    topic = "<font class=gray>没有找到相关文章</font>"
Else
    topic = "<a href='news_view.asp?id=" & rs(0) & "'>" & code_html(rs(1),1,25) & "</a>"
End If

rs.Close
Response.Write "下篇新闻：" & topic %>
  </td></tr></table>
</td></tr>
</table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 class=tf>
<tr><td height=5></td></tr>
<tr><td><% Call review_type(n_sort,id,"news_view.asp?id=" & id,1) %></td></tr>
<tr><td height=5></td></tr>
<tr><td><% Call news_class_sort(cid,sid) %></td></tr>
<tr><td height=5></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
Call web_center(1)
'------------------------------------right---------------------------------
Call format_login()
Call news_sea()
Call news_scroll("jt0","",3,15,1)
Call news_new_hot("jt0","","new",10,12,1,6,0)
Call news_new_hot("jt0","","hot",10,12,1,6,0)
'----------------------------------right end-------------------------------
Call web_end(0) %>