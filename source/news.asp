<!-- #include file="include/config_news.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit = tit_fir
tit_fir = ""

Call web_head(0,0,1,0,0)
'-----------------------------------center--------------------------------- %>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=right><% Response.Write format_img("rnews.jpg") & gang %></td><td width=1 bgcolor=<% = web_var(web_color,3) %>></td></tr>
<tr><td align=center><% Call news_fpic("",12,180,88) %></td><td width=1 bgcolor=<% = web_var(web_color,3) %>></td></tr>
<tr><td align=center><% = gang %></td><td width=1 bgcolor=<% = web_var(web_color,3) %>></td></tr>
<tr><td align=center><% Call news_main("jt0",12,20,1,6,3,1,10,1) %></td><td width=1 bgcolor=<% = web_var(web_color,3) %>></td></tr>
<tr><td align=center><% Call news_class_sort(cid,sid) %></td><td width=1 bgcolor=<% = web_var(web_color,3) %>></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
Call web_center(1)
'------------------------------------right---------------------------------
Call format_login()
Call news_sea()
Call news_scroll("jt0","",3,15,1)
Call news_new_hot("jt0","","new",10,12,1,6,0)
Call news_picr("",3,10,6)
'----------------------------------right end-------------------------------
Call web_end(0) %>