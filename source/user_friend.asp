<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit="我的好友（地址薄）"

call web_head(2,0,0,0,0)

if len(action)>1 and int(popedom_format(login_popedom,41)) then call close_conn():call cookies_type("locked")
'------------------------------------left----------------------------------
call left_user()
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong
call user_mail_menu(0)
response.write table1&vbcrlf&"<tr align=center"&table2&" height=25>"

if action="del" then
  response.write del_select()
end if

select case action
case "add"
  response.write friend_add()
case else
  call friend_main()
end select

response.write vbcrlf&"</table>"
'---------------------------------center end-------------------------------
call web_end(0)

sub friend_main()
%>
<td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif width='7%'><font class=end><b>排序</b></font></td>
<td width='28%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>用户名称</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>类型</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>性别</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>发贴</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>Email</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>QQ</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>主页</b></font></td>
<td width='9%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>发短信</b></font></td>
<td width='8%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>操作</b></font></td>
</tr>
<script language=javascript src='STYLE/admin_del.js'></script>
<form name=del_form action='user_friend.asp?action=del' method=post>
<%
  dim rs,sql,rssum,i,tname,ttt
  rssum=0
  sql="select user_data.username,user_data.power,user_data.sex,user_data.bbs_counter,user_data.email,user_data.qq,user_data.url,user_friend.id from user_data inner join user_friend on user_data.username=user_friend.username2 where user_friend.username1='"&login_username&"' order by user_friend.id desc"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if not(rs.eof and rs.bof) then
    rssum=int(rs.recordcount)
  end if
  for i=1 to rssum
    tname=rs("username")
    ttt=format_power(rs("power"),0)
    response.write vbcrlf&"<tr align=center"&table3&"><td>"&i&".</td>" & _
		   vbcrlf&"<td>"&format_user_view(tname,1,1)&"</td>" & _
		   vbcrlf&"<td><img src='images/small/icon_"&ttt&".gif' title='"&tname&" 是 "&format_power(ttt,1)&"' align=absmiddle border=0></td>"
    ttt=rs("sex")
    if ttt=false then
      ttt="<img src='images/small/forum_girl.gif' title='"&tname&" 是 青春女孩' align=absmiddle border=0>"
    else
      ttt="<img src='images/small/forum_boy.gif' title='"&tname&" 是 阳光男孩' align=absmiddle border=0>"
    end if
    response.write vbcrlf&"<td>"&ttt&"</td>" & _
		   vbcrlf&"<td><font class=red_3>"&rs("bbs_counter")&"</font></td>" & _
		   vbcrlf&"<td><a href='mailto:"&rs("email")&"'><img src='images/small/email.gif' title='给 "&tname&" 发电子邮件' align=absMiddle border=0></a></td>"
    ttt=rs("qq")
    if var_null(ttt)="" or ttt=0 then
      ttt="<font class=gray>没有</font>"
    else
      ttt="<a href='http://search.tencent.com/cgi-bin/friend/user_show_info?ln="&ttt&"' target=_blank><img src='images/small/qq.gif' title='查看 "&tname&" 的QQ信息' align=absMiddle border=0></a>"
    end if
    response.write vbcrlf&"<td>"&ttt&"</td>"
    ttt=rs("url")
    if var_null(ttt)="" then
      ttt="<font class=gray>没有</font>"
    else
      ttt="<a href='"&ttt&"' target=_blank><img src='images/small/url.gif' title='查看 "&tname&" 的个人主页' align=absMiddle border=0></a>"
    end if
    response.write vbcrlf&"<td>"&ttt&"</td>" & _
		   vbcrlf&"<td><a href='user_message.asp?action=write&accept_uaername="&server.urlencode(tname)&"'><img src='images/mail/msg.gif' border=0 align=absmiddle title='给 "&tname&" 发送站内短信'></a></td>" & _
		   vbcrlf&"<td><input type=checkbox name=del_id value='"&rs("id")&"' class=bg_1></td></tr>"
    rs.movenext
  next
%>
<tr><td colspan=10 align=center height=30 bgcolor=<%response.write web_var(web_color,5)%>>
共有 <font class=red><%response.write rssum%></font> 位好友
　　<input type=button value='添加我的好友' onClick="document.location='user_friend.asp?action=add'">
　　<input type=checkbox name=del_all value=1 onClick="javascript:selectall('<%response.write rssum%>');" class=bg_3> 选中所有
　<input type=submit value='删除所选' onclick="return suredel('<%response.write rssum%>');">
</td></tr>
<%
end sub

function friend_add()
  friend_add="<td><font class=end><b>添加我的好友</b></font></td></tr>" & _
	     vbcrlf&"<tr"&table3&"><td height=160 align=center>"
  if trim(request.form("add_ok"))="ok" then
    dim username2,red,rs,sql
    red=""
    username2=trim(request.form("username2"))
    if symbol_name(username2)<>"yes" then
      red="<font class=red>好友名称</font> 为空或不符合相关规则！"
    else
      sql="select username from user_data where username='"&username2&"'"
      set rs=conn.execute(sql)
      if rs.eof and rs.bof then
        red="你填写的 <font class=red>好友名称</font> 好像不存在！"
      end if
      rs.close:set rs=nothing
    end if
    if red="" then
      set rs=server.createobject("adodb.recordset")
      sql="select * from user_friend where username1='"&login_username&"' and username2='"&username2&"'"
      rs.open sql,conn,1,3
      if rs.eof and rs.bof then
        rs.addnew
        rs("username1")=login_username
        rs("username2")=username2
        rs.update
        friend_add=friend_add&"<font class=red>您已成功的添加了好友（<font class=blue_1>"&username2&"</font>）！</font>"
      else
        friend_add=friend_add&"<font class=red>您已经添加了好友（<font class=blue_1>"&username2&"</font>）！</font>"
      end if
      rs.close:set rs=nothing
      friend_add=friend_add&"<br><br><a href='user_friend.asp'>点击返回</a>"
    else
      friend_add=friend_add&red&"<br><br>"&go_back
    end if
  else
    friend_add=friend_add&"<form action='user_friend.asp?action=add' method=post><input type=hidden name=add_ok value='ok'>好友名称：<input type=text name=username2 value='"&trim(request.querystring("add_username"))&"' size=30 maxlength=20><br><br><input type=submit value='添加好友'></form>"
  end if
  friend_add=friend_add&"</td></tr>"
end function

function del_select()
  dim delid,del_i,del_num,del_dim,del_sql
  delid=trim(request.form("del_id"))
  if var_null(delid)<>"" then
    delid=replace(delid," ","")
    del_dim=split(delid,",")
    del_num=UBound(del_dim)
    for del_i=0 to del_num
      del_sql="delete from user_friend where username1='"&login_username&"' and id="&del_dim(del_i)
      conn.execute(del_sql)
    next
    Erase del_dim
    del_select=vbcrlf&"<script language=javascript>alert(""好友删除成功！共删除了 "&del_num+1&" 位好友。"");</script>"
  end if
end function
%>