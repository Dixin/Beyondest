<!-- #include file="INCLUDE/config_news.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nummer,page,rssum,thepages,viewpage,pageurl,keyword,sea_type,sqladd2
tit="新闻分类浏览"
call cid_sid()
nummer=int(web_var(web_num,2))

call web_head(0,0,1,0,0)
'-----------------------------------center---------------------------------
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center><%
response.write format_img("rnewslist.jpg")&gang
sqladd=""
if cid=0 and action<>"more" then
  call news_main("jt0",16,20,1,6,3,2,10,1)
else
  if cid>0 then sqladd=" and c_id="&cid
  if sid>0 then sqladd=sqladd&" and s_id="&sid
  sqladd2=sqladd
  if action="more" then
    call news_more("jt0",35,1,6,3,5,10,1)
  else
    if sid=0 then
      call news_list("jt0",10,20,1,6,3,1,10,1)
    else
      call news_list("jt0",30,20,1,6,3,5,10,1)
    end if
  end if
end if
%></td><td width=1 bgcolor=<%=web_var(web_color,3)%>></td></tr>
<tr><td align=center><%call news_class_sort(cid,sid)%></td><td width=1 bgcolor=<%=web_var(web_color,3)%>></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_center(1)
'------------------------------------right---------------------------------
call format_login()
call news_sea()
call news_scroll("jt0","",3,15,1)
call news_new_hot("jt0",sqladd2,"hot",10,12,1,6,0)
call news_picr(sqladd2,1,10,6)
'----------------------------------right end-------------------------------
call web_end(0)
%>