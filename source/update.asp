<!-- #include file="INCLUDE/config_other.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim id
id=trim(request.querystring("id"))
if action="forum" then
  index_url="forum"
  tit_fir=format_menu(index_url)
  tit="论坛公告"
else
  tit="网站更新"
  action="news"
end if

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong

if isnumeric(id) then
  call update_view()
else
  call update_main()
end if

response.write kong
'---------------------------------center end-------------------------------
call web_end(0)

sub update_main()
%>
<table border=0 width='98%' cellspacing=2 cellpadding=2>
<tr><td colspan=2 height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr bgcolor=<%response.write web_var(web_color,5)%> valign=bottom height=20>
<%if action="news" then%>
<td width='70%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=news' class=red_3>网站更新</a></b></td>
<td width='30%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=forum' class=red_3>论坛公告</a></b></td>
<%else%>
<td width='70%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=forum' class=red_3>论坛公告</a></b></td>
<td width='30%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=news' class=red_3>网站更新</a></b></td>
<%end if%>
</tr>
<tr><td colspan=2 height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr valign=top>
<td><%response.write update_top("jt0",action,20,15,1,2)%></td>
<td><%
if action="news" then
  response.write update_top("jt0","forum",5,6,1,1)
else
  response.write update_top("jt0","news",5,6,1,1)
end if
%></td>
</tr>
<tr>
<td></td>
<td></td>
</tr>
</table>
<%
end sub

sub update_view()
  sql="select * from bbs_cast where id="&id
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close
    call update_main()
    exit sub
  end if
%>
<table border=0 width='96%'>
<tr><td align=center height=40><font class=blue size=3><b><%response.write rs("topic")%></b></font></td></tr>
<tr><td align=center class=gray><%response.write web_var(web_config,1)%>&nbsp;&nbsp;发布人：<%response.write format_user_view(rs("username"),1,"")%>&nbsp;&nbsp;发布时间：<%response.write time_type(rs("tim"),88)%></td></tr>
<tr><td height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr><td align=center>
  <table border=0 width='96%'>
  <tr><td class=htd><%response.write code_jk(rs("word"))%></td></tr>
  </table>
</td></tr>
</table>
<%
  rs.close
%>
<br>
<table border=0 width='96%' cellspacing=0 cellpadding=2>
<tr><td colspan=2 height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr bgcolor=<%response.write web_var(web_color,5)%> valign=bottom height=20>
<td width='50%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=news' class=red_3>网站更新</a></b></td>
<td width='50%' class=red_3><b>&nbsp;→&nbsp;<a href='update.asp?action=forum' class=red_3>论坛公告</a></b></td>
</tr>
<tr><td colspan=2 height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr valign=top>
<td><%response.write update_top("jt0","news",5,15,1,2)%></td>
<td><%response.write update_top("jt0","forum",5,15,1,2)%></td>
</tr>
<tr>
<td></td>
<td></td>
</tr>
</table>
<%
end sub

function update_top(u_jt,ut,u_num,c_num,et,timt)
  dim temp1,topic
  if u_jt<>"" then u_jt=img_small(u_jt)
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&u_num&" id,topic,tim from bbs_cast where sort='"&ut&"' order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("topic")
    temp1=temp1&vbcrlf&"<tr><td>"&u_jt&"<a href='update.asp?action="&ut&"&id="&rs("id")&"' title='"&code_html(topic,1,0)&"'>"&code_html(topic,1,c_num)&"</a>"&format_end(et,"<font class=gray>"&time_type(rs("tim"),timt)&"</font>")&"</td></tr>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&"</table>"
  update_top=temp1
end function
%>