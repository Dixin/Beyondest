<!-- #include file="INCLUDE/config_article.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim id:id=trim(request.querystring("id"))
if not(isnumeric(id)) then
  call format_redirect("article.asp")
  response.end
end if
%>
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="INCLUDE/config_review.asp" -->
<!-- #include file="include/conn.asp" -->
<%
dim username,topic,word,tim,counter,cname,sname,power,userp,emoney,author,keyes,sql2
sql="select * from article where hidden=1 and id="&id
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  rs.close:set rs=nothing:call close_conn()
  call format_redirect("article.asp")
  response.end
end if
cid=int(rs("c_id"))
sid=int(rs("s_id"))
username=rs("username")
topic=rs("topic")
word=code_jk(rs("word"))
tim=rs("tim")
counter=rs("counter")
power=rs("power")
emoney=rs("emoney")
author=rs("author")
keyes=rs("keyes")
rs.close

cname="文章浏览"
sname=""
if cid>0 then
  if sid>0 then
    sql2="select jk_class.c_name,jk_sort.s_name from jk_sort inner join jk_class on jk_sort.c_id=jk_class.c_id where jk_sort.c_id="&cid&" and jk_sort.s_id="&sid
  else
    sql2="select c_name from jk_class where c_id="&cid
  end if
  set rs=conn.execute(sql2)
  if not (rs.eof and rs.bof) then
    cname=rs("c_name"):tit=cname
    if sid>0 then sname=rs("s_name"):tit=cname&"（"&sname&"）"
  end if
  rs.close
end if

call web_head(1,0,2,0,0)

call emoney_notes(power,emoney,n_sort,id,"js",0,1,"article_list.asp?c_id="&cid&"&s_id="&sid)
sql="update article set counter=counter+1 where id="&id
conn.execute(sql)
'------------------------------------left----------------------------------
call font_word_js()
%>
<table border=0 width='96%' cellspacing=0 cellpadding=0>
<tr><td align=center height=50><font class=red_3 size=3><b><% response.write topic %></b></font></td></tr>
<tr><td align=center class=gray><%response.write time_type(tim,33)&"&nbsp;&nbsp;作者："&author&"&nbsp;&nbsp;"&web_var(web_config,1)%>&nbsp;&nbsp;<%call font_word_action()%>&nbsp;&nbsp;本文已被浏览&nbsp;<%response.write counter%>&nbsp;次</td></tr>
<tr><td height=10></td></tr>
<tr><td valign=top><%call font_word_type(word)%></td></tr>
<tr><td height=10></td></tr>
<tr><td>
  <table border=0 width='100%'>
  <tr><td width='25%' class=htd>
&nbsp;发布人：<%response.write format_user_view(username,1,1)%><br>
&nbsp;<%response.write put_type("article")%>
  </td><td width='75%' class=htd>
<%
sql="select id,topic from article where hidden=1 and id="&id-1
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  topic="<font class=gray>没有找到相关文章</font>"
else
  topic="<a href='article_view.asp?id="&rs(0)&"'>"&code_html(rs(1),1,30)&"</a>"
end if
rs.close
response.write "上篇文章："&topic&"<br>"
sql="select id,topic from article where hidden=1 and id="&id+1
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  topic="<font class=gray>没有找到相关文章</font>"
else
  topic="<a href='article_view.asp?id="&rs(0)&"'>"&code_html(rs(1),1,30)&"</a>"
end if
rs.close
response.write "下篇文章："&topic
%>
  </td></tr></table>
</td></tr>
</table>
<% call article_view_about() %>
<table border=0 width='96%' cellspacing=0 cellpadding=0 class=tf>
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td><% call article_sea() %></td>
<td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td></tr>
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td><% call article_view_review() %></td>
<td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)
%>