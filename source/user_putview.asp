<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

select case action
case "article"
  tit="查看我发表的文章"
case "down"
  tit="查看我添加的软件"
case "gallery"
  tit="查看我上传的贴图"
case "website"
  tit="查看我推荐的网站"
case else
  action="news"
  tit="查看我发布的新闻"
end select

dim rssum,nummer,page,thepages,viewpage,pageurl,types,topic,tim
rssum=0:thepages=0:viewpage=1:nummer=web_var(web_num,1)
pageurl="?action="&action&"&"

call web_head(2,0,0,0,0)
'------------------------------------left----------------------------------
call left_user()
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong&table1
%>
<tr<%response.write table2%> height=25><td class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small(us)%>&nbsp;&nbsp;<b>查看我所发表的相关信息</b></td></tr>
<tr<%response.write table3%>><td align=center height=30>
<%response.write img_small("jt12")%><a href='?action=news'<%if action="news" then response.write "class=red_3"%>>查看我所发布的新闻</a>　
<%response.write img_small("jt12")%><a href='?action=article'<%if action="article" then response.write "class=red_3"%>>查看我所发表的文章</a>　
<%response.write img_small("jt12")%><a href='?action=down'<%if action="down" then response.write "class=red_3"%>>查看我所添加的软件</a>　
<%response.write img_small("jt12")%><a href='?action=gallery'<%if action="gallery" then response.write "class=red_3"%>>查看我所上传的图片</a>
</td></tr>
</table>
<%
select case action
case "article"
  sql="select id,topic,tim,counter from article where username='"&login_username&"' and hidden=1 order by id desc"
case "down"
  sql="select id,name,tim,counter from down where username='"&login_username&"' and hidden=1 order by id desc"
case "gallery"
  types=trim(request.querystring("types"))
  if types<>"logo" and types<>"baner" then types="paste"
  pageurl=pageurl&"types="&types&"&"
  select case types
  case "logo"
    nummer=nummer*2
  case "baner"
    nummer=web_var(web_num,3)
  end select
  sql="select * from gallery where hidden=1 and types='"&types&"' and username='"&login_username&"' order by id desc"
case else
  sql="select id,topic,tim,counter from news where username='"&login_username&"' and hidden=1 order by id desc"
end select

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
  rssum=rs.recordcount
end if
call format_pagecute()

response.write ukong&table1
%>

<%
if int(viewpage)>1 then
  rs.move (viewpage-1)*nummer
end if

select case action
case "article"
  call putview_article()
case "down"
  call putview_down()
case "gallery"
  call putview_gallery()
case else
  call putview_news()
end select

rs.close:set rs=nothing
%>
<tr><td align=center bgcolor=<%=web_var(web_color,6)%> height=30 colspan=2<%response.write table3%>>
  <table border=0 width='98%' cellspacing=0 cellpadding=0>
<tr align=center valign=bottom><td width='30%' >
现在有<font class=red><%response.write rssum%></font>条记录┋
每页<font class=red><%response.write nummer%></font>个
  </td><td width='70%' bgcolor=<%=web_var(web_color,6)%>>
页次：<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font> 分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%>
  </td></tr>
  </table>
</td></tr>  
</table>
<br>
<%
'---------------------------------center end-------------------------------
call web_end(0)


sub putview_news()
%>
<tr align=center<%response.write table2%> height=25>
<td width='6%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>序号</b></td>
<td width='84%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>新闻标题111</b></td>
</tr>
<%
  for i=1 to nummer
    if rs.eof then exit for
    topic=rs("topic"):tim=rs("tim")
    response.write vbcrlf&"<tr"&table3&"><td align=center>"&(viewpage-1)*nummer+i&".</td><td><a target=_blank href='news_view.asp?id="&rs("id")&"' title='新闻标题："&code_html(topic,1,0)&"<br>浏览次数："&rs("counter")&"<br>发布时间："&tim&"'>"&code_html(topic,1,35)&"</a>"&format_end(1,time_type(tim,3))&"</td></tr>"
    rs.movenext
  next
end sub

sub putview_article()
%>
<tr align=center<%response.write table2%> height=25>
<td width='6%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>序号</b></td>
<td width='84%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>文章标题</b></td>
</tr>
<%
  for i=1 to nummer
    if rs.eof then exit for
    topic=rs("topic"):tim=rs("tim")
    response.write vbcrlf&"<tr"&table3&"><td align=center>"&(viewpage-1)*nummer+i&".</td><td><a target=_blank href='article_view.asp?id="&rs("id")&"' title='文章标题："&code_html(topic,1,0)&"<br>发表时间："&tim&"'>"&code_html(topic,1,35)&"</a>"&format_end(1,time_type(tim,3)&",<font class=blue>"&rs("counter")&"</font>")&"</td></tr>"
    rs.movenext
  next
end sub

sub putview_down()
%>
<tr align=center<%response.write table2%> height=25>
<td width='6%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>序号</b></td>
<td width='84%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>软件名称</b></td>
</tr>
<%
  for i=1 to nummer
    if rs.eof then exit for
    topic=rs("name"):tim=rs("tim")
    response.write vbcrlf&"<tr"&table3&"><td align=center>"&(viewpage-1)*nummer+i&".</td><td><a target=_blank href='article_view.asp?id="&rs("id")&"' title='软件名称："&code_html(topic,1,0)&"<br>添加时间："&tim&"'>"&code_html(topic,1,35)&"</a>"&format_end(1,time_type(tim,3)&",<font class=blue>"&rs("counter")&"</font>")&"</td></tr>"
    rs.movenext
  next
end sub

sub putview_gallery()
  dim j,k,kn,pic,name,nnum:nnum=1
  response.write vbcrlf&"<tr"&table3&"><td align=center>" & _
		 vbcrlf&"<table border=0>" & _
		 vbcrlf&"<tr><td width=100>"&img_small("jt1")&"<a href='?action="&action&"&types=paste'"
  if types="paste" then response.write " class=red_3"
  response.write vbcrlf&">精彩贴图</a></td>" & _
		 vbcrlf&"<td width=100>"&img_small("jt1")&"<a href='?action="&action&"&types=logo'"
  if types="logo" then response.write " class=red_3"
  response.write vbcrlf&">精彩LOGO</a></td>" & _
		 vbcrlf&"<td width=100>"&img_small("jt1")&"<a href='?action="&action&"&types=baner'"
  if types="baner" then response.write " class=red_3"
  response.write vbcrlf&">精彩BANNER</a></td></tr>" & _
		 vbcrlf&"</table></td></tr><tr"&table3&"><td align=center><table border=0 width='100%'>"
select case types
case "logo"
  kn=5:nummer=30
  if nummer mod kn > 0 then
    k=nummer\kn+1
  else
    k=nummer\kn
  end if
  
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  
  for i=1 to k
    'if rs.eof then exit for
    response.write "<tr align=center>"
    for j=1 to kn
      if rs.eof or nnum>nummer then exit for
      pic=rs("pic"):name=rs("name")
      response.write "<td><table border=0><tr><td align=center><img src='"&web_var(web_upload,1)&pic&"' border=0 width=88 height=31></td></tr><tr><td align=center title='"&code_html(name,1,0)&"'>"&code_html(name,1,10)&"</td></tr></table></td>"
      rs.movenext
      nnum=nnum+1
    next
    response.write "</tr>"
  next
case "baner"
  nummer=web_var(web_num,3)
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  
  for i=1 to nummer
    if rs.eof then exit for
    pic=rs("pic"):name=rs("name")
    response.write "<tr><td><table border=0 align=center><tr><td align=center><img src='"&web_var(web_upload,1)&pic&"' border=0 width=468 height=60></td></tr><tr><td align=center>"&code_html(name,1,0)&"</td></tr></table></td></tr>"
    rs.movenext
  next
case else
  kn=3:nummer=12
  if nummer mod kn > 0 then
    k=nummer\kn+1
  else
    k=nummer\kn
  end if
  
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  
  for i=1 to k
    'if rs.eof then exit for
    response.write "<tr align=center>"
    for j=1 to kn
      if rs.eof or nnum>nummer then exit for
      pic=rs("pic"):name=rs("name")
      response.write "<td><table border=0><tr><td align=center><a href='gallery.asp?action=view&c_id="&rs("c_id")&"&s_id="&rs("s_id")&"&id="&rs("id")&"'><img src='"&web_var(web_down,5)&"/"&pic&"' border=0 width="&web_var(web_num,7)&" height="&web_var(web_num,7)&"></a></td></tr><tr><td align=center>"&rs("name")&"</td></tr></table></td>"
      rs.movenext
      nnum=nnum+1
    next
    response.write "</tr>"
  next
end select
  
  response.write "</table></td></tr>"
end sub
%>