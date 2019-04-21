<!-- #include file="include/config_forum.asp" -->
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

forum_mode="half"	'forum_mode设为"full"时为全屏显示,"half"
if len(trim(request.querystring("mode")))>0 then
  forum_mode=trim(request.querystring("mode"))
end if

if forum_mode="full" then
  call mode_full()
else
  call mode_half()
end if

sub mode_half()
  call web_head(0,0,1,0,0)
  '-----------------------------------center---------------------------------
  response.write kong&format_img("rbbs.jpg")&kong&gang
  response.write ukong
  call forum_main("fk4")
  call forum_down(int(mid(web_setup,3,1)))
  '---------------------------------center end-------------------------------
  call web_center(1)
  '------------------------------------right---------------------------------
  call format_login()
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center><%call main_stat("fk4","jt1",0,5,1)%></td></tr>
<tr><td height=10></td></tr>
<tr><td align=center><%call forum_cast("fk4","jt0",10,11)%></td></tr>
<tr><td height=10></td></tr>
<tr><td align=center><%call forum_new("fk4","jt0",0,15,11,1)%></td></tr>
<tr><td height=5></td></tr>
<tr><td align=center><%response.write left_action("jt13",1)%></td></tr>
<tr><td height=10></td></tr>
</table>
<%
  '----------------------------------right end-------------------------------
  call web_end(0)
end sub

sub mode_full()
  call web_head(0,0,2,0,0)
  '-----------------------------------center---------------------------------
  response.write ukong
%>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<tr align=center valign=top>
<td width='24%'><%call format_login()%></td>
<td width='24%'><%call main_stat("fk4","jt1",0,5,1)%></td>
<td width='1%'></td>
<td width='25%'><%call forum_new("fk4","jt0",0,5,13,1)%></td>
<td width='1%'></td>
<td width='25%'><%call forum_cast("fk4","jt0",5,13)%></td>
</tr>
</table>
<%
  call forum_main("fk4")
  call forum_down(int(mid(web_setup,3,1)))
  '---------------------------------center end-------------------------------
  call web_end(0)
end sub
%>