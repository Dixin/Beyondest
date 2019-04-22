<!-- #include file="include/config_article.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit = tit_fir
tit_fir = ""

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Call article_left_hot("jt0",10,8,1,6)
Response.Write put_type("article")
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center--------------------------------- %>
<table border=0 width='100%' align=center cellspacing=0 cellpadding=0>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center><% Response.Write format_img("rart.jpg") %></td></tr>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center><% Response.Write format_barc("<font class=end><b>资料分类</b></font>",class_sort(n_sort,"article",0,0),3,0,14) %></td></tr>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center><% Call article_sea() %></td></tr>
<tr><td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center><% Call article_main("jt0",12,14,1,1) %></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
Call web_end(0) %>