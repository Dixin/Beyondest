<!-- #include file="INCLUDE/config_other.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim atb:atb=" target=_blank"
index_url="main"
tit=format_menu(index_url)
tit_fir=""

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
response.write format_img("llogin.jpg")
call format_login()
call main_update_view()
call main_stat("","jt1",1,1,1)
call user_data_top("bbs_counter","jt12",1,10)
call vote_type(1,1,"","f7f7f7")
response.write kong
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
%>


<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
    <td><%response.write format_img("rmain.jpg")%></td>
  </tr>
</table>
<%=gang%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
    <td><%call main_news("jt0",10,20,1,1)%></td>
  </tr>
</table>
<%=gang%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
    <td><%call main_down("jt0","","","","new",10,0,25,1,1)%></td>
    <td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
    <td><%call main_video("jt0",10,20,1,1)%></td>
  </tr>
</table>
<%=gang%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
    <td><%call web_search(0)%></td>
  </tr>
</table>
<%=gang%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
    <td><%call main_article("jt0",10,20,1,1)%></td>
    <td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
    <td><%call main_forum("jt0",10,17,1,1)%></td>
  </tr>
</table>
<%=gang%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
    <td><%call main_best()%></td>
  </tr>
</table>
<%=gang%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<%response.write web_var(web_color,3)%>"></td>
    <td><%call links_maini("fir",5)%></td>
  </tr>
</table>

<%
'---------------------------------center end-------------------------------
call web_end(0)
%>