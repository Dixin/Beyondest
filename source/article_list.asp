<!-- #include file="include/config_article.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer
Dim page
Dim rssum
Dim thepages
Dim viewpage
Dim pageurl
Dim sqladd
Dim keyword
Dim sea_type
tit = "ÎÄÀ¸·ÖÀàä¯ÀÀ"
Call cid_sid()
nummer = Int(web_var(web_num,2))

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Call article_left_hot("jt0",10,8,1,6)
Response.Write put_type("article")
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center--------------------------------- %>
<table border=0 width='100%' align=center cellspacing=0 cellpadding=0>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
<td><% Response.Write format_img("rart.jpg") %></td></tr>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
<td><% = gang %></td></tr>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
<td align=center><% Call article_list(0,20,3,1) %></td></tr>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
<td align=center><% Call article_sea() %></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
Call web_end(0) %>