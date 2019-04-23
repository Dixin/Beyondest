<!-- #include file="INCLUDE/config_forum.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim sqladd,nummer,user_temp,rssum,viewpage,thepages,page,pageurl
rssum=0:thepages=0:viewpage=1:nummer=web_var(web_num,1)
sqladd="":user_temp=""

select case action
case "top"
  tit="发贴排行"
  sqladd="bbs_counter desc,id desc"
case "emoney"
  tit="财富排行"
  sqladd="emoney desc,id desc"
case else
  tit="用户列表"
  sqladd="id desc"
end select
pageurl="?action="&action&"&"

call web_head(1,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong

sql="select username,power,bbs_counter,sex,email,qq,url,tim,emoney from user_data order by "&sqladd
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
  rssum=rs.recordcount
end if

call format_pagecute()

if int(viewpage)>1 then
  rs.move (viewpage-1)*nummer
end if
for i=1 to nummer
  if rs.eof then exit for
  user_temp=user_temp&user_type()
  rs.movenext
next
rs.close:set rs=nothing

response.write forum_table1
%>
<tr height=30<%response.write forum_table4%> align=center>
<td><font class=red_3><b><%response.write tit%></b></font>&nbsp;&nbsp;&nbsp;
共&nbsp;<font class=red><%response.write rssum%></font>&nbsp;位&nbsp;┋&nbsp;
每&nbsp;<font class=red><%response.write nummer%></font>&nbsp;页&nbsp;┋&nbsp;
共&nbsp;<font class=red><%response.write thepages%></font>&nbsp;页&nbsp;┋&nbsp;
这是第&nbsp;<font class=red><%response.write viewpage%></font>&nbsp;页</td>
</tr>
</table>
<% response.write kong & forum_table1 %>
<tr align=center<%response.write forum_table2%> height=25>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>排序</b></font></td>
<td width='27%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>用户名称</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>类型</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>性别</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>发贴</b></font></td>
<td width='6%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>Email</b></font></td>
<td width='6%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>QQ</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>主页</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>短信</b></font></td>
<td width='14%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>注册时间</b></font></td>
</tr>
<% response.write user_temp %>
</table>
<br>
<% response.write forum_table1 %>
<tr height=30<%response.write forum_table3%>>
<td width='72%'>&nbsp;分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000")%></td>
<td width='28%' align=center><% response.write forum_go() %></td>
</tr>
<tr<%response.write forum_table4%>><td align=center height=30 colspan=2><%response.write user_power_type(0)%></td></tr>
</table>
<br>
<%
'---------------------------------center end-------------------------------
call web_end(0)

function user_type()
  dim tname,ttt
  tname=rs("username")
  ttt=rs("power")
  user_type=vbcrlf&"<tr align=center"&forum_table4&"><td>"&i+(viewpage-1)*nummer&".</td>" & _
	    vbcrlf&"<td align=left>"&format_user_view(tname,1,"")&"</td>" & _
	    vbcrlf&"<td><img src='images/small/icon_"&ttt&".gif' title='"&tname&" 是 "&format_power(ttt,1)&"' align=absmiddle border=0></td>"
  ttt=rs("sex")
  if ttt=false then
    ttt="<img src='images/small/forum_girl.gif' title='"&tname&" 是 青春女孩' align=absmiddle border=0>"
  else
    ttt="<img src='images/small/forum_boy.gif' title='"&tname&" 是 阳光男孩' align=absmiddle border=0>"
  end if
  user_type=user_type&vbcrlf&"<td>"&ttt&"</td>" & _
	    vbcrlf&"<td><font class=red>"&rs("bbs_counter")&"</font></td>" & _
	    vbcrlf&"<td><a href='mailto:"&rs("email")&"'><img src='images/small/email.gif' title='给 "&tname&" 发电子邮件' align=absmiddle border=0></a></td>" & _
	    vbcrlf&"<td>"
  ttt=rs("qq")
  if not(isnumeric(ttt)) or len(ttt)<2 then
    ttt="<font class=gray>没有</font>"
  else
    ttt="<a href='http://search.tencent.com/cgi-bin/friend/user_show_info?ln="&ttt&"' target=_blank><img src='images/small/qq.gif' title='查看 "&tname&" 的QQ信息' align=absmiddle border=0></a>"
  end if
  user_type=user_type&ttt&"</td>"&vbcrlf&"<td>"
  ttt=rs("url")
  if var_null(ttt)="" then
    ttt="<font class=gray>没有</font>"
  else
    ttt="<a href='"&ttt&"' target=_blank><img src='images/small/url.gif' title='查看 "&tname&" 的个人主页' align=absmiddle border=0></a>"
  end if
  user_type=user_type&ttt&"</td><td><a href='user_message.asp?action=write&accept_uaername="&server.urlencode(tname)&"'><img src='images/mail/msg.gif' border=0 align=absmiddle title='给 "&tname&" 发送站内短信'></a></td>"&vbcrlf&"<td align=left>"&time_type(rs("tim"),3)&"</td></tr>"
end function
%>