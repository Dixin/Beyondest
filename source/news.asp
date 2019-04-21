<!-- #include file="include/config_news.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

tit=tit_fir
tit_fir=""

call web_head(0,0,1,0,0)
'-----------------------------------center---------------------------------
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=right><%response.write format_img("rnews.jpg")&gang%></td><td width=1 bgcolor=<%=web_var(web_color,3)%>></td></tr>
<tr><td align=center><%call news_fpic("",12,180,88)%></td><td width=1 bgcolor=<%=web_var(web_color,3)%>></td></tr>
<tr><td align=center><%=gang%></td><td width=1 bgcolor=<%=web_var(web_color,3)%>></td></tr>
<tr><td align=center><%call news_main("jt0",12,20,1,6,3,1,10,1)%></td><td width=1 bgcolor=<%=web_var(web_color,3)%>></td></tr>
<tr><td align=center><%call news_class_sort(cid,sid)%></td><td width=1 bgcolor=<%=web_var(web_color,3)%>></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_center(1)
'------------------------------------right---------------------------------
call format_login()
call news_sea()
call news_scroll("jt0","",3,15,1)
call news_new_hot("jt0","","new",10,12,1,6,0)
call news_picr("",3,10,6)
'----------------------------------right end-------------------------------
call web_end(0)
%>