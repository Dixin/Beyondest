<!-- #include file="INCLUDE/config_article.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nummer,page,rssum,thepages,viewpage,pageurl,sqladd,keyword,sea_type
tit="ÎÄÀ¸·ÖÀàä¯ÀÀ"
call cid_sid()
nummer=int(web_var(web_num,2))

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
call article_left_hot("jt0",10,8,1,6)
response.write put_type("article")
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
%>
<table border=0 width='100%' align=center cellspacing=0 cellpadding=0>
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td><%response.write format_img("rart.jpg")%></td></tr>
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td><%=gang%></td></tr>
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td align=center><%call article_list(0,20,3,1)%></td></tr>
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td align=center><%call article_sea()%></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)
%>