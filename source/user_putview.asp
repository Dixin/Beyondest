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
  tit="�鿴�ҷ��������"
case "down"
  tit="�鿴����ӵ����"
case "gallery"
  tit="�鿴���ϴ�����ͼ"
case "website"
  tit="�鿴���Ƽ�����վ"
case else
  action="news"
  tit="�鿴�ҷ���������"
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
<tr<%response.write table2%> height=25><td class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small(us)%>&nbsp;&nbsp;<b>�鿴��������������Ϣ</b></td></tr>
<tr<%response.write table3%>><td align=center height=30>
<%response.write img_small("jt12")%><a href='?action=news'<%if action="news" then response.write "class=red_3"%>>�鿴��������������</a>��
<%response.write img_small("jt12")%><a href='?action=article'<%if action="article" then response.write "class=red_3"%>>�鿴�������������</a>��
<%response.write img_small("jt12")%><a href='?action=down'<%if action="down" then response.write "class=red_3"%>>�鿴������ӵ����</a>��
<%response.write img_small("jt12")%><a href='?action=gallery'<%if action="gallery" then response.write "class=red_3"%>>�鿴�����ϴ���ͼƬ</a>
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
������<font class=red><%response.write rssum%></font>����¼��
ÿҳ<font class=red><%response.write nummer%></font>��
  </td><td width='70%' bgcolor=<%=web_var(web_color,6)%>>
ҳ�Σ�<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font> ��ҳ��<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%>
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
<td width='6%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>���</b></td>
<td width='84%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>���ű���111</b></td>
</tr>
<%
  for i=1 to nummer
    if rs.eof then exit for
    topic=rs("topic"):tim=rs("tim")
    response.write vbcrlf&"<tr"&table3&"><td align=center>"&(viewpage-1)*nummer+i&".</td><td><a target=_blank href='news_view.asp?id="&rs("id")&"' title='���ű��⣺"&code_html(topic,1,0)&"<br>���������"&rs("counter")&"<br>����ʱ�䣺"&tim&"'>"&code_html(topic,1,35)&"</a>"&format_end(1,time_type(tim,3))&"</td></tr>"
    rs.movenext
  next
end sub

sub putview_article()
%>
<tr align=center<%response.write table2%> height=25>
<td width='6%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>���</b></td>
<td width='84%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>���±���</b></td>
</tr>
<%
  for i=1 to nummer
    if rs.eof then exit for
    topic=rs("topic"):tim=rs("tim")
    response.write vbcrlf&"<tr"&table3&"><td align=center>"&(viewpage-1)*nummer+i&".</td><td><a target=_blank href='article_view.asp?id="&rs("id")&"' title='���±��⣺"&code_html(topic,1,0)&"<br>����ʱ�䣺"&tim&"'>"&code_html(topic,1,35)&"</a>"&format_end(1,time_type(tim,3)&",<font class=blue>"&rs("counter")&"</font>")&"</td></tr>"
    rs.movenext
  next
end sub

sub putview_down()
%>
<tr align=center<%response.write table2%> height=25>
<td width='6%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>���</b></td>
<td width='84%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>�������</b></td>
</tr>
<%
  for i=1 to nummer
    if rs.eof then exit for
    topic=rs("name"):tim=rs("tim")
    response.write vbcrlf&"<tr"&table3&"><td align=center>"&(viewpage-1)*nummer+i&".</td><td><a target=_blank href='article_view.asp?id="&rs("id")&"' title='������ƣ�"&code_html(topic,1,0)&"<br>���ʱ�䣺"&tim&"'>"&code_html(topic,1,35)&"</a>"&format_end(1,time_type(tim,3)&",<font class=blue>"&rs("counter")&"</font>")&"</td></tr>"
    rs.movenext
  next
end sub

sub putview_gallery()
  dim j,k,kn,pic,name,nnum:nnum=1
  response.write vbcrlf&"<tr"&table3&"><td align=center>" & _
		 vbcrlf&"<table border=0>" & _
		 vbcrlf&"<tr><td width=100>"&img_small("jt1")&"<a href='?action="&action&"&types=paste'"
  if types="paste" then response.write " class=red_3"
  response.write vbcrlf&">������ͼ</a></td>" & _
		 vbcrlf&"<td width=100>"&img_small("jt1")&"<a href='?action="&action&"&types=logo'"
  if types="logo" then response.write " class=red_3"
  response.write vbcrlf&">����LOGO</a></td>" & _
		 vbcrlf&"<td width=100>"&img_small("jt1")&"<a href='?action="&action&"&types=baner'"
  if types="baner" then response.write " class=red_3"
  response.write vbcrlf&">����BANNER</a></td></tr>" & _
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