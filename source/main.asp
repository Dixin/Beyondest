<!-- #include file="INCLUDE/config_other.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim atb:atb = " target=_blank"
index_url = "main"
tit = format_menu(index_url)
tit_fir = ""

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Response.Write format_img("llogin.jpg")
Call format_login()
Call main_update_view()
Call main_stat("","jt1",1,1,1)
Call user_data_top("bbs_counter","jt12",1,10)
Call vote_type(1,1,"","f7f7f7")
Response.Write kong
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center--------------------------------- %>


<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td><% Response.Write format_img("rmain.jpg") %></td>
  </tr>
</table>
<% = gang %>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td><% Call main_news("jt0",10,20,1,1) %></td>
  </tr>
</table>
<% = gang %>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td><% Call main_down("jt0","","","","new",10,0,25,1,1) %></td>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td><% Call main_video("jt0",10,20,1,1) %></td>
  </tr>
</table>
<% = gang %>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td><% Call web_search(0) %></td>
  </tr>
</table>
<% = gang %>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td><% Call main_article("jt0",10,20,1,1) %></td>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td><% Call main_forum("jt0",10,17,1,1) %></td>
  </tr>
</table>
<% = gang %>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td><% Call main_best() %></td>
  </tr>
</table>
<% = gang %>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td><% Call links_maini("fir",5) %></td>
  </tr>
</table>

<%
'---------------------------------center end-------------------------------
Call web_end(0) %>