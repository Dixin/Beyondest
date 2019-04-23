<!-- #include file="config.asp" -->
<!-- #include file="skin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim table1,table2,table3,us
us="fk2"
table1=format_table(1,3)
table2=format_table(3,2)
table3=format_table(3,1)
index_url="user_main"
tit_fir=format_menu(index_url)

sub user_mail_menu(mmt)
%>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<tr><td align=center height=50>
<a href='user_mail.asp?action=inbox'><img src='images/mail/inbox.gif' border=0></a>&nbsp;
<a href='user_mail.asp?action=outbox'><img src='images/mail/outbox.gif' border=0></a>&nbsp;
<a href='user_mail.asp?action=issend'><img src='images/mail/issend.gif' border=0></a>&nbsp;
<a href='user_mail.asp?action=recycle'><img src='images/mail/recycle.gif' border=0></a>&nbsp;
<a href='user_friend.asp'><img src='images/mail/address.gif' border=0></a>&nbsp;
<a href='user_message.asp?action=write'><img src='images/mail/write.gif' border=0></a><%
if action="view" or action="reply" or action="fw" then
  response.write vbcrlf&"<a href='user_message.asp?action=reply&id="&id&"'><img src='images/mail/reply.gif' border=0></a>" & _
		 vbcrlf&"<a href='user_message.asp?action=fw&id="&id&"'><img src='images/mail/fw.gif' border=0></a>" & _
		 vbcrlf&"<a href='user_message.asp?action=del&id="&id&"'><img src='images/mail/delete.gif' border=0></a>"
end if
%></td></tr>
</table>
<%
end sub

sub left_user()
  call format_login()
  dim usql,urs,uface,temp1,jtn:jtn=img_small("jt13")
  usql="select face from user_data where username='"&login_username&"'"
  set urs=conn.execute(usql)
  uface=urs(0)
  urs.close:set urs=nothing
  temp1=vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=4 align=center>" & _
	vbcrlf&"<tr><td height=5 width='50%'></td><td width='50%'></td></tr>" & _
	vbcrlf&"<tr><td height=80 colspan=2 align=center><img src='images/face/"&uface&".gif' border=0></td></tr>" & _
	vbcrlf&"<tr><td height=25 colspan=2 align=center><font class=red>"&format_power(login_mode,1)&"</font>：<b><font class=blue>"&login_username&"</font></b></td></tr>" & _
	vbcrlf&"<tr><td>"&jtn&"<a href='user_mail.asp'>站内短信</a></td><td>"&jtn&"<a href='user_put.asp?action=website'>推荐网站</a></td></tr>" & _
	vbcrlf&"<tr><td>"&jtn&"<a href='user_bookmark.asp'>网络书签</a></td><td>"&jtn&"<a href='user_friend.asp'>我的好友</a></td></tr>" & _
	vbcrlf&"<tr><td>"&jtn&"<a href='user_edit.asp'>修改资料</a></td><td>"&jtn&"<a href='user_edit.asp#pass'>修改密码</a></td></tr>" & _
	vbcrlf&"<tr><td colspan=2>"&jtn&"<a href='user_putview.asp'>查看我所发表的相关信息</a></td></tr>" & _
	vbcrlf&"<tr><td>"&jtn&"<a href='user_put.asp?action=news'>发布新闻</a></td><td>"&jtn&"<a href='user_put.asp?action=article'>发表文章</a></td></tr>" & _
	vbcrlf&"<tr><td>"&jtn&"<a href='user_put.asp?action=down'>上传音乐</a></td><td>"&jtn&"<a href='user_put.asp?action=gallery'>上传文件</a></td></tr>" & _
	vbcrlf&"<tr><td></td><td></td></tr>" & _
	vbcrlf&"</table>"
  response.write vbcrlf&"<table border=0 width='96%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>" & _
		 kong&format_barc("<img src='images/"&web_var(web_config,5)&"/left_user.gif' border=0>",temp1,2,0,4) & _
		 "</td></tr><tr><td align=center>" & _
		 left_action("jt13",2) & _
		 "</td></tr></table>"
end sub
%>