<!-- #include file="INCLUDE/config_article.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************

'

'                     Beyondest.Com V3.6 Demo版

' 




'           网址：http://www.beyondest.com

' 

'*******************************************************************

tit=tit_fir
tit_fir=""

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
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%response.write format_img("rart.jpg")%></td></tr>
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%response.write format_barc("<font class=end><b>资料分类</b></font>",class_sort(n_sort,"article",0,0),3,0,14)%></td></tr>
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%call article_sea()%></td></tr>
<tr><td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%call article_main("jt0",12,14,1,1)%></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)
%>