<!-- #include file="INCLUDE/config_down.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer,page,rssum,thepages,viewpage,pageurl,keyword,sea_type,sqladd2
tit    = "作品分类浏览"
Call cid_sid()
nummer = Int(web_var(web_num,2))

Call web_head(0,0,0,0,0)
'------------------------------------left---------------------------------- %>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center><% Call format_login() %></td></tr>
<tr><td align=center><% Call down_sea() %></td></tr>
<tr><td align=center><% Call down_new_hot("jt0","","","","good",15,0,13,1,0) %></td></tr>
<tr><td align=center><% Call down_new_hot("jt0","","","","hot",15,0,13,1,0) %></td></tr>
</table>
<%
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center--------------------------------- %>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td width=1 bgcolor='<% = web_var(web_color,3) %>'></td><td>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<%
sqladd = ""

If action = "more" Then
    Call down_more(30,8)
Else
    Call down_main()
End If %>
</table>
</td></tr></table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center><% Call down_class_sortt(cid,sid) %></td></tr>
</table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center><% Call down_remark("jt0") %></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
Call web_end(0) %>