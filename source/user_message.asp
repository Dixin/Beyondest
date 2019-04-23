<!-- #include file="INCLUDE/config_user.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim id:id=trim(request.querystring("id"))
if not(isnumeric(id)) and action<>"write" then call cookies_type("mail_id")
%>
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
dim send_u,accept_u,topic,word,types,isread,red_3
tit="站内短信"

call web_head(2,0,0,0,0)

if action<>"view" and int(popedom_format(login_popedom,41)) then call close_conn():call cookies_type("locked")
'------------------------------------left----------------------------------
call left_user()
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong
call user_mail_menu(0)
response.write ukong&table1

select case action
case "reply"
  call mail_reply()
case "fw"
  call mail_fw()
case "edit"
  call mail_edit()
case "view"
  response.write mail_view()
case "del"
  response.write mail_del()
case else
  call mail_write()
end select

response.write vbcrlf&"</table>"
'---------------------------------center end-------------------------------
call web_end(0)

function mail_del()
  mail_del=vbcrlf&"<tr"&table2&"><td align=center><font class=end><b>删除短信</b></font></td></tr>"
  dim rs,sql,html_temp
  html_temp=""
  sql="select id from user_mail where (send_u='"&login_username&"' or accept_u='"&login_username&"') and id="&id
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    html_temp="<font class=red_2>您所要删除的短信ID不存在或出错！</font><br><br>"&go_back
  end if
  rs.close:set rs=nothing
  if html_temp="" then
    sql="update user_mail set types=4 where id="&id
    conn.execute(sql)
    html_temp="<font class=red>短信删除成功！删除的短信将置于您的回收站内。</font><br><br><a href='user_mail.asp?action=recycle'>点击返回</a>"
  end if
  mail_del=mail_del&"<tr"&table3&"><td height=150 align=center>"&html_temp&"</td></tr>"
end function

sub mail_write()
  response.write vbcrlf&"<tr"&table2&" height=25><td colspan=2 align=center background=images/"&web_var(web_config,5)&"/bar_3_bg.gif><font class=end><b>撰写短信</b></font></td></tr>"
  if trim(request.form("write_ok"))="ok" then
    response.write vbcrlf&"<tr"&table3&"><td colspan=2 align=center height=150>"
    if post_chk()="no" then
      response.write web_var(web_error,1)
    else
      red_3=""
      accept_u=trim(request.form("accept_u"))
      topic=trim(request.form("topic"))
      word=request.form("word")
      if symbol_name(accept_u)<>"yes" then
        red_3=red_3 & "<br><li><font class=red_3>收 信 人</font> 为空或不符合相关规则！"
      else
        sql="select username from user_data where username='"&accept_u&"'"
        set rs=conn.execute(sql)
        if rs.eof and rs.bof then
          red_3=red_3 & "<br><li>你填写的 <font class=red_3>收 信 人</font> 好像不存在！"
        end if
        rs.close:set rs=nothing
      end if
      if var_null(topic)="" or len(topic)>20 then
        red_3=red_3 & "<br><li><font class=red_3>短信主题</font> 不能为空且长度不能大于20！"
      end if
      if var_null(word)="" or len(word)>250 then
        red_3=red_3 & "<br><li><font class=red_3>短信内容</font> 不能为空且长度不能大于250！"
      end if
      if red_3="" then
        set rs=server.createobject("adodb.recordset")
        sql="select * from user_mail"
        rs.open sql,conn,1,3
        rs.addnew
        rs("send_u")=login_username
        rs("accept_u")=accept_u
        rs("topic")=topic
        rs("word")=word
        rs("tim")=now_time
        if trim(request.form("send_later"))="yes" then
          rs("types")=2
        else
          rs("types")=1
        end if
        rs("isread")=false
        rs.update
        rs.close
        if trim(request.form("send_later"))="yes" then
          response.write "<font class=red>您已成功的保存了一条短信！</font><br><br><a href='user_mail.asp?action=outbox'>点击返回</a>"
        else
          response.write "<font class=red>您已成功的给 <font class=blue><b>"&accept_u&"</b></font> 发送了一条短信！</font><br><br><a href='user_mail.asp'>点击返回</a>"
        end if
      else
        response.write found_error(red_3,"250")
      end if
    end if
    response.write vbcrlf&"</td></tr>"
  else
    response.write vbcrlf&"<form name=mail_frm action='user_message.asp?action=write' method=post onsubmit=""javascript:frm_submitonce(this);""><input type=hidden name=write_ok value='ok'><input type=hidden name=send_later value=''>" & _
		   vbcrlf&"<tr height=30"&table3&"><td width='15%' align=center bgcolor="&web_var(web_color,6)&">收 信 人：</td><td width='85%'>&nbsp;<input type=text name=accept_u value='"&trim(request.querystring("accept_uaername"))&"' size=30 maxlength=20>"&redx&"&nbsp;　&nbsp;"&friend_select()&"</td></tr>" & _
		   vbcrlf&"<tr height=30"&table3&"><td align=center bgcolor="&web_var(web_color,6)&">短信主题：</td><td>&nbsp;<input type=text name=topic size=60 maxlength=20></td></tr>" & _
		   vbcrlf&"<tr height=100"&table3&"><td align=center class=htd bgcolor="&web_var(web_color,6)&">短信内容：<br>"&web_var(web_error,3)&"</td><td>&nbsp;<textarea cols=64 rows=6 name=word title='短信内容最多250个字符<br>按 Ctrl+Enter 可直接发送' onkeydown=""javascript:frm_quicksubmit();""></textarea></td></tr>" & _
		   vbcrlf&"<tr"&table3&"><td colspan=2 height=40 align=center><input type=Submit name=wsubmit value='发送短信'>&nbsp;　&nbsp;<input type=submit name=send value='保存短信' onclick=""javascript:mail_send_later();"">&nbsp;　&nbsp;<input type=reset value='清除重写'></td></tr></form>"
  end if
end sub

sub mail_reply()
  response.write vbcrlf&"<tr"&table2&"><td colspan=2 align=center><font class=end><b>回复短信</b></font></td></tr>"
  if trim(request.form("reply_ok"))="ok" then
    response.write vbcrlf&"<tr"&table3&"><td colspan=2 align=center height=150>"
    if post_chk()="no" then
      response.write web_var(web_error,1)
    else
      red_3=""
      accept_u=trim(request.form("accept_u"))
      topic=trim(request.form("topic"))
      word=request.form("word")
      if symbol_name(accept_u)<>"yes" then
        red_3=red_3 & "<br><li><font class=red_3>收 信 人</font> 为空或不符合相关规则！"
      else
        sql="select username from user_data where username='"&accept_u&"'"
        set rs=conn.execute(sql)
        if rs.eof and rs.bof then
          red_3=red_3 & "<br><li>你填写的 <font class=red_3>收 信 人</font> 好像不存在！"
        end if
        rs.close
      end if
      if var_null(topic)="" or len(topic)>20 then
        red_3=red_3 & "<br><li><font class=red_3>短信主题</font> 不能为空且长度不能大于20！"
      end if
      if var_null(word)="" or len(word)>250 then
        red_3=red_3 & "<br><li><font class=red_3>短信内容</font> 不能为空且长度不能大于250！"
      end if
      if red_3="" then
        set rs=server.createobject("adodb.recordset")
        sql="select * from user_mail"
        rs.open sql,conn,1,3
        rs.addnew
        rs("send_u")=login_username
        rs("accept_u")=accept_u
        rs("topic")=topic
        rs("word")=word
        rs("tim")=now_time
        if trim(request.form("send_later"))="yes" then
          rs("types")=2
        else
          rs("types")=1
        end if
        rs("isread")=false
        rs.update
        rs.close
        if trim(request.form("send_later"))="yes" then
          response.write "<font class=red>您已成功的保存了一条短信的内容！</font><br><br><a href='user_mail.asp?action=outbox'>点击返回</a>"
        else
          response.write "<font class=red>您已成功的给 <font class=blue_1><b>"&accept_u&"</b></font> 回复了一条短信！</font><br><br><a href='user_mail.asp'>点击返回</a>"
        end if
      else
        response.write found_error(red_3,"250")
      end if
    end if
    response.write vbcrlf&"</td></tr>"
  else
    sql="select send_u,topic from user_mail where (send_u='"&login_username&"' or accept_u='"&login_username&"') and id="&id
    set rs=conn.execute(sql)
    if rs.eof and rs.bof then
      rs.close
      red_3="<br><li>您所回复的 <font class=red_3>短信ID</font> 不存在或有错误！"
      red_3=found_error(red_3,"240")
      response.write vbcrlf&"<tr"&table3&"><td align=center height=150 colspan=2>"&red_3&"</td></tr>"
      exit sub
    else
      response.write vbcrlf&"<form name=mail_frm action='user_message.asp?action=reply&id="&id&"' method=post onsubmit=""javascript:frm_submitonce(this);""><input type=hidden name=reply_ok value='ok'><input type=hidden name=send_later value=''>" & _
		     vbcrlf&"<tr height=30"&table3&"><td width='15%' align=center>收 信 人：</td><td width='85%'>&nbsp;<input type=text name=accept_u value='"&rs("send_u")&"' size=30 maxlength=20>"&redx&"&nbsp;　&nbsp;"&friend_select()&"</td></tr>" & _
		     vbcrlf&"<tr height=30"&table3&"><td align=center>短信主题：</td><td>&nbsp;<input type=text name=topic value='RE:"&rs("topic")&"' size=60 maxlength=20></td></tr>" & _
		     vbcrlf&"<tr height=100"&table3&"><td align=center class=htd>短信内容：<br>"&web_var(web_error,3)&"</td><td>&nbsp;<textarea cols=64 rows=6 name=word title='短信内容最多250个字符<br>按 Ctrl+Enter 可直接发送' onkeydown=""javascript:frm_quicksubmit();""></textarea></td></tr>" & _
		     vbcrlf&"<tr"&table3&"><td colspan=2 height=40 align=center><input type=Submit name=wsubmit value='发送短信'>&nbsp;　&nbsp;<input type=Submit name=send value='保存短信' onclick=""javascript:mail_send_later();"">&nbsp;　&nbsp;<input type=reset value='清除重写'></td></tr></form>"
    end if
    rs.close
  end if
end sub

sub mail_fw()
  response.write vbcrlf&"<tr"&table2&"><td colspan=2 align=center><font class=end><b>转发短信</b></font></td></tr>"
  if trim(request.form("fw_ok"))="ok" then
    response.write vbcrlf&"<tr"&table3&"><td colspan=2 align=center height=150>"
    if post_chk()="no" then
      response.write web_var(web_error,1)
    else
      red_3=""
      accept_u=trim(request.form("accept_u"))
      topic=trim(request.form("topic"))
      word=request.form("word")
      if symbol_name(accept_u)<>"yes" then
        red_3=red_3 & "<br><li><font class=red_3>收 信 人</font> 为空或不符合相关规则！"
      else
        sql="select username from user_data where username='"&accept_u&"'"
        set rs=conn.execute(sql)
        if rs.eof and rs.bof then
          red_3=red_3 & "<br><li>你填写的 <font class=red_3>收 信 人</font> 好像不存在！"
        end if
        rs.close
      end if
      if var_null(topic)="" or len(topic)>20 then
        red_3=red_3 & "<br><li><font class=red_3>短信主题</font> 不能为空且长度不能大于20！"
      end if
      if var_null(word)="" or len(word)>250 then
        red_3=red_3 & "<br><li><font class=red_3>短信内容</font> 不能为空且长度不能大于250！"
      end if
      if red_3="" then
        set rs=server.createobject("adodb.recordset")
        sql="select * from user_mail"
        rs.open sql,conn,1,3
        rs.addnew
        rs("send_u")=login_username
        rs("accept_u")=accept_u
        rs("topic")=topic
        rs("word")=word
        rs("tim")=now_time
        if trim(request.form("send_later"))="yes" then
          rs("types")=2
        else
          rs("types")=1
        end if
        rs("isread")=false
        rs.update
        rs.close
        if trim(request.form("send_later"))="yes" then
          response.write "<font class=red>您已成功的保存了一条短信的内容！</font><br><br><a href='user_mail.asp?action=outbox'>点击返回</a>"
        else
          response.write "<font class=red>您已成功的给 <font class=blue_1><b>"&accept_u&"</b></font> 转发了一条短信！</font><br><br><a href='user_mail.asp'>点击返回</a>"
        end if
      else
        response.write found_error(red_3,"250")
      end if
    end if
    response.write vbcrlf&"</td></tr>"
  else
    sql="select send_u,topic,word,tim from user_mail where (send_u='"&login_username&"' or accept_u='"&login_username&"') and id="&id
    set rs=conn.execute(sql)
    if rs.eof and rs.bof then
      rs.close
      red_3="<br><li>您所转发的 <font class=red_3>短信ID</font> 不存在或有错误！"
      red_3=found_error(red_3,"240")
      response.write vbcrlf&"<tr"&table3&"><td align=center height=150 colspan=2>"&red_3&"</td></tr>"
      exit sub
    else
      response.write vbcrlf&"<form name=mail_frm action='user_message.asp?action=fw&id="&id&"' method=post onsubmit=""frm_submitonce(this);""><input type=hidden name=fw_ok value='ok'><input type=hidden name=send_later value=''>" & _
		     vbcrlf&"<tr height=30"&table3&"><td width='15%' align=center>收 信 人：</td><td width='85%'>&nbsp;<input type=text name=accept_u size=30 maxlength=20>"&redx&"&nbsp;　&nbsp;"&friend_select()&"</td></tr>" & _
		     vbcrlf&"<tr height=30"&table3&"><td align=center>短信主题：</td><td>&nbsp;<input type=text name=topic value='FW:"&rs("topic")&"' size=60 maxlength=20></td></tr>" & _
		     vbcrlf&"<tr height=100"&table3&"><td align=center class=htd>短信内容：<br>"&web_var(web_error,3)&"</td><td>&nbsp;<textarea cols=64 rows=6 name=word title='短信内容最多250个字符<br>按 Ctrl+Enter 可直接发送' onkeydown=""javascript:frm_quicksubmit();"">以下为 "&login_username&" 转发 "&rs("send_u")&" 于 "&rs("tim")&" 写的短信"&vbcrlf&"——————————————————————————————"&vbcrlf&rs("word")&"</textarea></td></tr>" & _
		     vbcrlf&"<tr"&table3&"><td colspan=2 height=40 align=center><input type=Submit name=wsubmit value='发送短信'>&nbsp;　&nbsp;<input type=Submit name=send value='保存短信' onclick=""javascript:mail_send_later();"">&nbsp;　&nbsp;<input type=reset value='清除重写'></td></tr></form>"
    end if
    rs.close
  end if
end sub

sub mail_edit()
  response.write vbcrlf&"<tr"&table2&"><td colspan=2 align=center><font class=end><b>编缉短信</b></font></td></tr>"
  if trim(request.form("edit_ok"))="ok" then
    response.write vbcrlf&"<tr"&table3&"><td colspan=2 align=center height=150>"
    if post_chk()="no" then
      response.write web_var(web_error,1)
    else
      red_3=""
      accept_u=trim(request.form("accept_u"))
      topic=trim(request.form("topic"))
      word=request.form("word")
      if symbol_name(accept_u)<>"yes" then
        red_3=red_3 & "<br><li><font class=red_3>收 信 人</font> 为空或不符合相关规则！"
      else
        sql="select username from user_data where username='"&accept_u&"'"
        set rs=conn.execute(sql)
        if rs.eof and rs.bof then
          red_3=red_3 & "<br><li>你填写的 <font class=red_3>收 信 人</font> 好像不存在！"
        end if
        rs.close
      end if
      if var_null(topic)="" or len(topic)>20 then
        red_3=red_3 & "<br><li><font class=red_3>短信主题</font> 不能为空且长度不能大于20！"
      end if
      if var_null(word)="" or len(word)>250 then
        red_3=red_3 & "<br><li><font class=red_3>短信内容</font> 不能为空且长度不能大于250！"
      end if
      if red_3="" then
        set rs=server.createobject("adodb.recordset")
        sql="select * from user_mail where id="&id
        rs.open sql,conn,1,3
        if rs.eof and rs.bof then
          rs.close:set rs=nothing
          call close_conn()
          call cookies_type(mail_id)
	  response.end
        end if
        
        rs("send_u")=login_username
        rs("accept_u")=accept_u
        rs("topic")=topic
        rs("word")=word
        rs("tim")=now_time
        rs("types")=1
        if trim(request.form("send_later"))="yes" then
          rs("types")=2
        else
          rs("types")=1
        end if
        rs("isread")=false
        rs.update
        rs.close
        if trim(request.form("send_later"))="yes" then
          response.write "<font class=red>您已成功的保存了短信的内容！</font><br><br><a href='user_mail.asp?action=outbox'>点击返回</a>"
        else
          response.write "<font class=red>您已成功的给 <font class=blue_1><b>"&accept_u&"</b></font> 发送了一条短信！</font><br><br><a href='user_mail.asp'>点击返回</a>"
        end if
      else
        response.write found_error(red_3,"250")
      end if
    end if
    response.write vbcrlf&"</td></tr>"
  else
    sql="select accept_u,topic,word,tim from user_mail where (send_u='"&login_username&"' or accept_u='"&login_username&"') and id="&id
    set rs=conn.execute(sql)
    if rs.eof and rs.bof then
      rs.close
      red_3="<br><li>您所编缉的 <font class=red_3>短信ID</font> 不存在或有错误！"
      red_3=found_error(red_3,"240")
      response.write vbcrlf&"<tr><td align=center height=150 colspan=2>"&red_3&"</td></tr>"
      exit sub
    else
      response.write vbcrlf&"<form name=mail_frm action=user_message.asp?action=edit&id="&id&" method=post onsubmit=""frm_submitonce(this);""><input type=hidden name=edit_ok value='ok'><input type=hidden name=send_later value=''>" & _
		     vbcrlf&"<tr height=30"&table3&"><td width='15%' align=center>收 信 人：</td><td width='85%'>&nbsp;<input type=text name=accept_u value='"&rs("accept_u")&"' size=30 maxlength=20>"&redx&"&nbsp;　&nbsp;"&friend_select()&"</td></tr>" & _
		     vbcrlf&"<tr height=30"&table3&"><td align=center>短信主题：</td><td>&nbsp;<input type=text name=topic value='"&rs("topic")&"' size=60 maxlength=20></td></tr>" & _
		     vbcrlf&"<tr height=100"&table3&"><td align=center class=htd>短信内容：<br>"&web_var(web_error,3)&"</td><td>&nbsp;<textarea cols=64 rows=6 name=word title='短信内容最多250个字符<br>按 Ctrl+Enter 可直接发送' onkeydown=""javascript:frm_quicksubmit();"">"&rs("word")&"</textarea></td></tr>" & _
		     vbcrlf&"<tr"&table3&"><td colspan=2 height=40 align=center><input type=Submit name=wsubmit value='发送短信'>&nbsp;　&nbsp;<input type=Submit name=send value='保存短信' onclick=""javascript:mail_send_later();"">&nbsp;　&nbsp;<input type=reset value='清除重写'></td></tr></form>"
    end if
    rs.close
  end if
end sub

function mail_view()
  mail_view=vbcrlf&"<tr"&table2&" height=25><td align=center background=images/"&web_var(web_config,5)&"/bar_3_bg.gif><font class=end><b>查看短信</b></font></td></tr>"
  red_3=""
  sql="select * from user_mail where (send_u='"&login_username&"' or accept_u='"&login_username&"') and id="&id
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    red_3="<br><li>您所查看的 <font class=red_3>短信ID</font> 不存在或有错误！"
    red_3=found_error(red_3,"240")
    mail_view=mail_view&"<tr"&table3&"><td align=center height=150>"&red_3&"</td></tr>"
    exit function
  end if
  send_u=rs("send_u")
  accept_u=rs("accept_u")
  types=int(rs("types"))
  isread=rs("isread")
  mail_view=mail_view&vbcrlf&"<tr"&table3&"><td height=50>&nbsp;&nbsp;短信主题：<font class=red_3>"&code_html(rs("topic"),1,0)&"</font></td></tr>" & _
	    vbcrlf&"<tr"&table3&"><td height=80 align=center><table border=0 width='96%' class=tf><tr><td height=8></td></tr><tr><td class=bw>"&code_jk(rs("word"))&"</td></tr><tr><td height=8></td></tr></table></td></tr>" & _
	    vbcrlf&"<tr"&table3&"><td align=center height=30>以上是 "&format_user_view(send_u,1,1)&" 于 "&time_type(rs("tim"),88)&" 给您发送的短信</td></tr>"
  rs.close:set rs=nothing
  if not(send_u=login_username and accept_u<>login_username) and isread=false then
    sql="update user_mail set isread=1 where types<>2 and id="&id
    conn.execute(sql)
    if login_message>0 then login_message=login_message-1
  end if
end function

function friend_select()
  dim sql,rs,ttt
  friend_select=vbcrlf&"<script language=javascript>" & _
		vbcrlf&"function Do_accept(addaccept) {" & _
		vbcrlf&"  if (addaccept!=0) { document.mail_frm.accept_u.value=addaccept; }" & _
		vbcrlf&"  return;" & _
		vbcrlf&"}</script>" & _
		vbcrlf&"<select name=friend_select size=1 onchange=Do_accept(this.options[this.selectedIndex].value)>" & _
		vbcrlf&"<option value='0'>选择我的好友</option>"
  sql="select username2 from user_friend where username1='"&login_username&"' order by id"
  set rs=conn.execute(sql)
  do while not rs.eof
    ttt=rs(0)
    friend_select=friend_select&vbcrlf&"<option value='"&ttt&"'>"&ttt&"</option>"
    rs.movenext
  loop
  rs.close
  friend_select=friend_select&vbcrlf&"</select>"
end function
%>
<script language=javascript>
<!--
//调用方法:onsubmit="frm_submitonce(this);"
function frm_submitonce(theform)
{
  if (document.all||document.getElementById)
  {
    for (i=0;i<theform.length;i++)
    {
      var tempobj=theform.elements[i]
      if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
      tempobj.disabled=true
    }
  }
}

function frm_quicksubmit(eventobject)
{
  if (event.keyCode==13 && event.ctrlKey)
  mail_frm.wsubmit.click();
}

function mail_send_later()
{
  this.document.mail_frm.send_later.value='yes';
}
-->
</script>