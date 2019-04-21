<!-- #include file="include/config_news.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim id:id=trim(request.querystring("id"))
if not(isnumeric(id)) then
  call format_redirect("news.asp")
  response.end
end if
%>
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/config_review.asp" -->
<!-- #include file="include/conn.asp" -->
<%
dim username,word,tim,counter,cname,sname,comto,pic,tt,sql2
sql="select * from news where hidden=1 and id="&id
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  rs.close:set rs=nothing:close_conn
  server.transfer "news.asp"
  response.end
end if
cid=rs("c_id")
sid=rs("s_id")
username=rs("username")
topic=rs("topic")
word=code_jk(rs("word"))
tim=rs("tim")
counter=rs("counter")
ispic=rs("ispic")
comto=rs("comto")
pic=rs("pic")
rs.close

cname="新闻浏览":sname=""
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

call web_head(0,0,1,0,0)

sql="update news set counter=counter+1 where id="&id
conn.execute(sql)
'-----------------------------------center---------------------------------
call font_word_js()
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center height=50><font class=red_3 size=3><b><% response.write topic %></b></font></td></tr>
<tr><td align=center class=gray><%response.write time_type(tim,33)&"&nbsp;"&web_var(web_config,1)%>&nbsp;<%call font_word_action()%>&nbsp;出处：<%response.write comto%></td></tr>
<tr><td height=10></td></tr>
<tr><td valign=top><%
if ispic=true then
  word="<p align=center><img src='"&url_true(web_var(web_down,5)&"/",pic)&"' border=0 onload=""javascript:if(this.width>screen.width-430)this.width=screen.width-430""></p>"&word
end if
call font_word_type(word&"&nbsp;<font class=gray>（本文已被浏览&nbsp;"&counter&"&nbsp;次）</font>")%></td></tr>
<tr><td height=10></td></tr>
<tr><td>
  <table border=0 width='100%'>
  <tr><td width='25%' class=htd>
&nbsp;发布人：<%response.write format_user_view(username,1,1)%><br>
&nbsp;<%response.write put_type("news")%>
  </td><td width='75%' class=htd>
<%
sql="select id,topic from news where hidden=1 and id="&id-1
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  topic="<font class=gray>没有找到相关新闻</font>"
else
  topic="<a href='news_view.asp?id="&rs(0)&"'>"&code_html(rs(1),1,25)&"</a>"
end if
rs.close
response.write "上篇新闻："&topic&"<br>"
sql="select id,topic from news where hidden=1 and id="&id+1
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  topic="<font class=gray>没有找到相关文章</font>"
else
  topic="<a href='news_view.asp?id="&rs(0)&"'>"&code_html(rs(1),1,25)&"</a>"
end if
rs.close
response.write "下篇新闻："&topic
%>
  </td></tr></table>
</td></tr>
</table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 class=tf>
<tr><td height=5></td></tr>
<tr><td><% call review_type(n_sort,id,"news_view.asp?id="&id,1) %></td></tr>
<tr><td height=5></td></tr>
<tr><td><% call news_class_sort(cid,sid) %></td></tr>
<tr><td height=5></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_center(1)
'------------------------------------right---------------------------------
call format_login()
call news_sea()
call news_scroll("jt0","",3,15,1)
call news_new_hot("jt0","","new",10,12,1,6,0)
call news_new_hot("jt0","","hot",10,12,1,6,0)
'----------------------------------right end-------------------------------
call web_end(0)
%>