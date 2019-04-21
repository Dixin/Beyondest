<!-- #include file="include/config_down.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim nummer,page,rssum,thepages,viewpage,pageurl,keyword,sea_type,sqladd2
tit="作品分类浏览"
call cid_sid()
nummer=int(web_var(web_num,2))

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center><%call format_login()%></td></tr>
<tr><td align=center><%call down_sea()%></td></tr>
<tr><td align=center><%call down_new_hot("jt0","","","","good",15,0,13,1,0)%></td></tr>
<tr><td align=center><%call down_new_hot("jt0","","","","hot",15,0,13,1,0)%></td></tr>
</table>
<%
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td width=1 bgcolor='<%=web_var(web_color,3)%>'></td><td>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<%
sqladd=""
if action="more" then
  call down_more(30,8)
else
  call down_main()
end if
%>
</table>
</td></tr></table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%call down_class_sortt(cid,sid)%></td></tr>
</table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%call down_remark("jt0")%></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)
%>