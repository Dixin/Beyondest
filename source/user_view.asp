<!-- #include file="include/config_user.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim username,view_username,userp
dim login1,login2
username=code_form(trim(request.querystring("username")))
tit="查看用户信息（"&username&"）"

call web_head(2,0,0,0,0)
userp=int(format_power(login_mode,2))
'------------------------------------left----------------------------------
call left_user()
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
call user_view()
'---------------------------------center end-------------------------------
call web_end(0)

sub user_view()
  dim tim_login,user_popedom,user_p
  sql="select l_where,l_tim_login from user_login where l_username='"&username&"'"
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    login1="<font class=gray>该用户现在没有登陆，处于离线状态</font>"
    login2=login1
  else
    login1="在线时间 <font class=red>"&datediff("n",rs(1),now())&"</font> 分钟"
    login2="当前位置：<font class=blue>"&rs(0)&"</font>"
  end if
  rs.close
  
  sql="select * from user_data where username='"&username&"'"
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call close_conn()
    format_redirect("user_main.asp")
    response.end
  end if
  user_popedom=rs("popedom")
  user_p=int(format_power(rs("power"),2))
  if user_p=3 then
    if int(userp)>int(user_p) then
      rs.close:set rs=nothing
      call close_conn()
      call cookies_type("power")
      response.end
    end if
  end if
  
  response.write ukong&vbcrlf&table1
%>
<tr<%response.write table2%> height=25>
<td colspan=3 background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small(us)%>&nbsp;&nbsp;<font class=end><b>查看用户信息（<%response.write username%>）</b></font></td>
</tr>
<tr<%response.write table3%> height=30>
<td width='20%' align=center bgcolor=<%=web_var(web_color,6)%>>用户名称：</td>
<td width='40%'>&nbsp;<font class=blue><b><%response.write username%></b></font>&nbsp;&nbsp;<a href='user_message.asp?action=write&accept_uaername=<%response.write server.urlencode(username)%>'><img src='IMAGES/MAIL/MSG.GIF' border=0 align=absmiddle title='给 <%response.write username%> 发送站内短信'></a></td>
<td width='40%' align=center bgcolor=<%=web_var(web_color,6)%>><%response.write login1%></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>用户类型：</td>
<td>&nbsp;<font class=red_3><%response.write format_power(rs("power"),1)%></font></td>
<td rowspan=8 align=center><img src='images/face/<%response.write rs("face")%>.gif' border=0></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>用户头衔：</td>
<td>&nbsp;<%
  tit=rs("nname")
  if var_null(tit)="" then
    response.write "<font class=gray>没有</font>"
  else
    response.write ""&code_html(tit,1,0)
  end if
%></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>来自哪里：</td>
<td>&nbsp;<%response.write code_html(rs("whe"),1,0)%></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>论坛发贴：</td>
<td>&nbsp;<font class=red><%response.write rs("bbs_counter")%></font></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>社区积分：</td>
<td>&nbsp;<font class=red_4><%response.write rs("integral")%></font></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>用户性别：</td>
<td>&nbsp;<%
  tit=rs("sex")
  if tit=false then
    response.write "<img src='images/small/forum_girl.gif' align=absmiddle border=0>&nbsp;&nbsp;青春女孩"
  else
    response.write "<img src='images/small/forum_boy.gif' align=absmiddle border=0>&nbsp;&nbsp;阳光男孩"
  end if
%></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>出生年月：</td>
<td>&nbsp;<%response.write rs("birthday")%></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>用户ＱＱ：</td>
<td>&nbsp;<%
  tit=rs("qq")
  if not(isnumeric(tit)) or len(tit)<2 then
    response.write "<font class=gray>没有</font>"
  else
    response.write "<img src='images/small/qq.gif' align=absmiddle border=0>&nbsp;<a href='http://search.tencent.com/cgi-bin/friend/user_show_info?ln="&tit&"' target=_blank>"&tit&"</a>"
  end if
%></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>最后登陆：</td>
<td>&nbsp;<%response.write time_type(rs("last_tim"),88)%></td>
<td align=center bgcolor=<%=web_var(web_color,6)%>><%response.write login2%></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>E - mail：</td>
<td colspan=2>&nbsp;<%
  tit=code_html(rs("email"),1,0)
  response.write "<img src='images/small/email.gif' align=absmiddle border=0>&nbsp;<a href='mailto:"&tit&"' title=''>"&tit&"</a>"
%></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>个人主页：</td>
<td colspan=2>&nbsp;<%
  tit=code_html(rs("url"),1,0)
  if var_null(tit)="" then
    response.write "<font class=gray>没有</font>"
  else
    response.write "<img src='images/small/url.gif' align=absmiddle border=0>&nbsp;<a href='"&tit&"' target=_blank>"&tit&"</a>"
  end if
%></td>
</tr>
<tr<%response.write table3%> height=25>
<td align=center bgcolor=<%=web_var(web_color,6)%>>个人介绍：</td>
<td colspan=2 align=center><table border=0 width='100%' class=tf><tr><td class=bw><%response.write code_jk2(rs("remark"))%></td></tr></table></td>
</tr>
<% rs.close:set rs=nothing %>
<tr<%response.write table2%> height=25>
<td colspan=3  background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small(us)%>&nbsp;&nbsp;<font class=end><b>论坛最新发贴</b>&nbsp;&nbsp;-</font>&nbsp;&nbsp;<a href='forum_action.asp?action=user&username=<% response.write server.urlencode(username) %>' class=menu>查看 <%response.write username%> 参与过的主题</a></td>
</tr>
<tr<%response.write table3%>>
<td colspan=3 align=center>
  <table border=0 width='96%'>
<%
  sql="select top 10 id,forum_id,topic,tim from bbs_topic where username='"&username&"' order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    response.write "  <tr><td>"&img_small("jt0")&"<a href='forum_view.asp?forum_id="&rs("forum_id")&"&view_id="&rs("id")&"' target=_blank>"&code_html(rs("topic"),1,30)&"</a>"&format_end(1,time_type(rs("tim"),8))&"</td></tr>"
    rs.movenext
  loop
  rs.close:set rs=nothing
%>
  </table>
</td>
</tr>
<tr<%response.write table3%>><td colspan=3 height=30 bgcolor=<%=web_var(web_color,6)%>>
&nbsp;&nbsp;用户管理操作：&nbsp;&nbsp;<font class=gray>[<a href='user_isaction.asp?username=<%response.write server.urlencode(username)%>&action=locked<%
  if int(popedom_format(user_popedom,41))=0 then
    response.write "'>锁定"
  else
    response.write "&cancel=yes' class=red_3>解除锁定"
  end if
%></a>]&nbsp;&nbsp;[<a href='user_isaction.asp?username=<%response.write server.urlencode(username)%>&action=shield<%
  if int(popedom_format(user_popedom,42))=0 then
    response.write "'>屏蔽"
  else
    response.write "&cancel=yes' class=red_3>解除屏蔽"
  end if
%></a>]</font>
</td></tr>
</table>
<br>
<%
end sub
%>