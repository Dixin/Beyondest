<!-- #include file="INCLUDE/config_other.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nummer,page,rssum,thepages,viewpage,pageurl,sqladd,keyword,sea_type,sea_name,topic,topic2,sql1,sql2,linkurl,keywords,tims
pageurl="?":sqladd="":topic="":sql1="":sql2="":linkurl="":keywords="":sea_name="����"
nummer=20:viewpage=1:thepages=0
tit="վ������"

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong
call web_search(0)
response.write ukong
call sea_types()
call sql_add()
if sqladd="" then
  call search_error()
else
  call search_main()
end if

response.write ukong
'---------------------------------center end-------------------------------
call web_end(0)

sub search_error()
%>
<table border=0 width='96%'>
<tr><td height=300 align=center>
  <table border=0>
  <tr><td colspan=2 height=30>������û����д�������ؼ��֡�����鿴���°���˵����</td></tr>
  <tr><td width=10></td><td><%response.write img_small("jt1")%>������ʱ������д�������ؼ��֡���</td></tr>
  <tr><td></td><td><%response.write img_small("jt12")%>��Ҫ��������ؼ�������<font class=red>�ո�</font>������ؼ��ָ������磺<font class=blue>V6&nbsp;���</font>��</td></tr>
  <tr><td></td><td><%response.write img_small("jt0")%>���ؼ��֡��в��ܺ��е����ţ�'����</td></tr>
  <tr><td></td><td><%response.write img_small("jt0")%>���ؼ��֡��к��еļӺţ�+��������Ϊ�ո���</td></tr>
  <tr><td></td><td><%response.write img_small("jt13")%>������������ֻ�ڣ����š���������������Ч��</td></tr>
  <tr><td></td><td><%response.write img_small("jt14")%>ף����ʹ�ñ�վ�ġ�վ��������ʱ������졣</td></tr>
  </table>
</td></tr>
</table>
<%
end sub

sub search_main()
  sql=sql1&sqladd&sql2
  tims=timer()
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
    rssum=0
  else
    rssum=rs.recordcount
  end if
  call format_pagecute()
%>
  <table border=0 width='96%' cellspacing=0 cellpadding=2>
  <tr><td height=1 colspan=4 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr align=center valign=bottom<%response.write table4%>>
  <td width='6%'>���</td>
  <td width='94%'>������ݣ�����ѯ�Ĺؼ����ǣ�<%response.write keywords%>ÿҳ <font class=red><%response.write nummer%></font> �� <font class=blue><%response.write sea_name%></font> ��ѯ�����</td>
  </tr>
  <tr><td height=1 colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr><td height=5></td></tr>
<%
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
%>
  <tr>
  <td align=center><%response.write (viewpage-1)*nummer+i%>.</td>
  <td><a href='<%
    response.write linkurl&rs(0)
    if sea_type="forum" then response.write "&forum_id="&rs(4)
%>' target=_blank><%response.write code_html(rs(1),1,32)%></a>&nbsp;<font class=gray size=1><%response.write time_type(rs(3),3)%></font>&nbsp;<%response.write format_user_view(rs(2),1,"blue")%></td>
  </tr>
<%
    rs.movenext
  next
  rs.close:set rs=nothing
%>
  <tr><td height=5></td></tr>
  <tr><td height=1 colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr><td colspan=2<%response.write table4%>>
    <table border=0 width='100%' cellspacing=0 cellpadding=0>
    <tr>
    <td>��&nbsp;<font class=red><%response.write rssum%></font>&nbsp;�����&nbsp;
ҳ�Σ�<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font>&nbsp;
��ҳ��<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,3,"#ff0000")%></td>
    <td align=right><font size=2 class=gray>��ѯ��ʱ��<font class=red_3><% response.write FormatNumber((timer()-tims)*1000,3) %></font> ����</font></td>
    </tr>
    </table>
  </td></tr>
  <tr><td height=1 colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  </table>
<%
end sub

sub sql_add()
  dim ddim,dnum,i
  keyword=code_form(request.querystring("keyword"))
  if len(keyword)<1 or len(topic)<1 then sqladd="":exit sub
  keyword=replace(keyword,"+"," ")
  pageurl=pageurl&"keyword="&server.urlencode(keyword)&"&"
  ddim=split(keyword," ")
  dnum=ubound(ddim)
  for i=0 to dnum
    keywords=keywords&"<font class=red_3><b>"&ddim(i)&"</b></font>&nbsp;&nbsp;"
    sqladd=sqladd&" and "&topic2&" like '%"&ddim(i)&"%'"
  next
  erase ddim
  if sea_type="forum" and sqladd<>"" then
    sqladd=right(sqladd,len(sqladd)-4)
  end if
end sub

sub sea_types()
  dim celerity
  celerity=trim(request.querystring("celerity"))
  sea_type=trim(request.querystring("sea_type"))
  select case sea_type
  case "news","article"
    topic="topic":topic2=topic
    if celerity="yes" then topic2="keyes"
    linkurl=sea_type&"_view.asp?id="
    sea_name="����"
    if sea_type="article" then sea_name="����"
    sql1="select id,"&topic&",username,tim from "&sea_type&" where hidden=1"
    sql2=" order by id desc"
  case "down"
    topic="name":topic2=topic
    if celerity="yes" then topic2="keyes"
    linkurl=sea_type&"_view.asp?id="
    sea_name="���"
    sql1="select id,"&topic&",username,tim from "&sea_type&" where hidden=1"
    sql2=" order by id desc"
  case "website"
    topic="name":topic2=topic
    linkurl=sea_type&".asp?action=view&id="
    sea_name="��վ"
    sql1="select id,"&topic&",username,tim from "&sea_type&" where hidden=1"
    sql2=" order by id desc"
  case "paste","flash"
    topic="name":topic2=topic
    linkurl="gallery.asp?action="&sea_type&"&types=view&id="
    sea_name="ͼƬ"
    if sea_type="flash" then sea_name="Flash"
    sql1="select id,"&topic&",username,tim from gallery where hidden=1 and types='"&sea_type&"'"
    sql2=" order by id desc"
  case else
    sea_type="forum"
    topic="topic":topic2=topic
    linkurl="forum_view.asp?view_id="
    sea_name="��̳"
    sql1="select id,"&topic&",username,tim,forum_id from bbs_topic where"
    sql2=" order by id desc"
  end select
  pageurl=pageurl&"sea_type="&sea_type&"&"
end sub
%>