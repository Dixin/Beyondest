<!-- #include file="INCLUDE/config_down.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit=tit_fir
tit_fir=""

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center><%call format_login()%></td></tr>
<tr><td align=center><%call down_atat()%></td></tr>
<tr><td align=center><%call down_sea()%></td></tr>
<tr><td align=center><%call down_new_hot("jt0","","","","good",20,0,13,1,0)%></td></tr>
<tr><td align=center><%call down_new_hot("jt0","","","","hot",20,0,13,1,0)%></td></tr>
</table>
<%
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%response.write format_img("rdown.jpg")%></td></tr>
</table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr align=center valign=top>
<td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td width='60%'><%call down_new_hotr("jt0","","","","new",21,4,30,1,7)%></td>
<td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td><%call down_tool()%></td>
</tr>
</table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%call down_class_sort(0,0)%></td></tr>
</table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%call down_remark("jt0")%></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)
%>