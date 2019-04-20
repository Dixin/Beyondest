<!-- #include file="INCLUDE/config_forum.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim editid
editid=trim(request.querystring("edit_id"))
if not(isnumeric(forumid)) or not(isnumeric(editid)) then call cookies_typ("view_id")
%>
<!-- #include file="INCLUDE/config_upload.asp" -->
<!-- #include file="INCLUDE/config_frm.asp" -->
<!-- #include file="include/conn.asp" -->
<%
dim topic_yes,iid
call forum_first()
call forum_word()
tit=forumname&"（编辑贴子）"

call web_head(2,0,2,0,0)

if int(popedom_format(login_popedom,41)) then call close_conn():call cookies_type("locked")

sql="select bbs_data.username as data_username,bbs_data.topic as data_topic,bbs_data.word as data_word,bbs_data.tim as data_tim,bbs_topic.id,bbs_topic.username as topic_username,bbs_topic.topic as topic_topic,bbs_topic.tim as topic_tim " & _
    "from bbs_topic inner join bbs_data on bbs_topic.id=bbs_data.reply_id where bbs_data.id="&editid&" and bbs_data.forum_id="&forumid&" and bbs_topic.islock<>1"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  rs.close:set rs=nothing:call close_conn()
  call cookies_type("edit_id")
end if

if format_user_power(login_username,login_mode,forumpower)<>"yes" then
  if login_username<>rs("data_username") then
    set rs=nothing:call close_conn()
    call cookies_type("edit_id")
  end if
end if

iid=rs("id")
if rs("data_tim")=rs("topic_tim") and rs("data_username")=rs("topic_username") then
  topic_yes="yes"
end if
'-----------------------------------center---------------------------------
response.write forum_top("编辑贴子") & kong

if trim(request.form("edit"))="ok" then
  response.write "<table border=0><tr><td align=center height=200>"
  if post_chk()="no" then
    response.write web_var(web_error,1)
  else
    response.write edit_chk()
  end if
  response.write "</td></tr></table>"
else
  response.write edit_type()
end if

rs.close:set rs=nothing
'---------------------------------center end-------------------------------
call web_end(0)

function edit_type()
%>
<script language=javascript><!--
function check(write_frm)
{
<% if topic_yes="yes" then %>
  if(write_frm.topic.value=="" || write_frm.topic.value.length>50)
  {
   alert("你还没完全留下所需信息！\r\n\n贴子的 主题 是必须要的；\n且长度不能超过50个字符。");
   return false;
  }
<% else %>
  if(write_frm.topic.value.length>50)
  {
   alert("你还没完全留下所需信息！\r\n\n回贴的 主题 长度不能超过50个字符。");
   return false;
  }
<% end if %>
  if(write_frm.jk_word.value=="" || write_frm.jk_word.value.length><%response.write word_size*1024%>)
  {
   alert("你还没完全留下所需信息！\r\n\n贴子的 内容 是必须要的；\n且大小不能超过<%response.write word_size%>KB。");
   return false;
  }
}
-->
</script>
<script language=javascript src='style/em_type.js'></script>
<script language=javascript src='style/forum_ok.js'></script>
<%response.write forum_table1%>
<form name=write_frm action='forum_edit.asp?forum_id=<%response.write forumid%>&edit_id=<%response.write editid%>' method=post onsubmit="frm_submitonce(this);">
<input type=hidden name=edit value='ok'>
<input type=hidden name=upid value=''>
<tr height=30<%response.write forum_table3%>>
<td width='20%' align=center>用户信息：</td>
<td width='80%'>&nbsp;&nbsp;用户名：<input type=username name=username value='<%response.write login_username%>' size=18 maxlength=20>&nbsp;&nbsp;
密码：<input type=password name=password value='<%response.write login_password%>' size=18 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='user_main.asp?username=<%response.write server.htmlencode(login_username)%>'>用户中心</a> ]</font>&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='login.asp?action=logout'>退出登陆</a> ]</font></td>
</tr>
<tr height=30<%response.write format_table(3,1)%>>
<td align=center>贴子主题：</td>
<td>
  <table border=0 cellspacing=0 cellpadding=0><tr>
  <td>&nbsp;&nbsp;<%call frm_topic("write_frm","topic")%></td>
  <td>&nbsp;<input type=text name=topic value='<%response.write rs("data_topic")%>' size=60 maxlength=50><%response.write redx%>长度不能超过50</td>
  </tr></table>
</td>
</tr>
<tr height=30<%response.write forum_table3%>>
<td align=center>当前心情：</td>
<td>&nbsp;&nbsp;<% response.write icon_type(9,3) %>
</td>
</tr>
<tr height=35<%response.write format_table(3,1)%>>
<td align=center><%call frm_ubb_type()%></td>
<td><%call frm_ubb("write_frm","jk_word","&nbsp;&nbsp;")%></td>
</tr>
<tr align=center<%response.write forum_table3%>>
<td><table border=0><tr><td class=htd>贴子内容：<br><br><%response.write word_remark%><br><br><br></td></tr></table></td>
<td><table border=0><tr><td><textarea name=jk_word rows=12 cols=95 title='按 Ctrl+Enter 可直接发送' onkeydown="javascript:frm_quicksubmit();"><%response.write rs("data_word")%></textarea></td></tr></table></td>
</tr>
<tr<%response.write format_table(3,1)%>><td align=center>上传文件：</td><td>&nbsp;<iframe frameborder=0 name=upload_frame width='99%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=&uptext=jk_word'></iframe></td></tr>
<tr height=30<%response.write forum_table3%>><td align=center>E M 贴图：</td><td>&nbsp;&nbsp;<script language=javascript>jk_em_type('s');</script></td></tr>
<tr<%response.write format_table(3,1)%>><td colspan=2 align=center height=60><script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<%response.write forum_table3%>>
<td>发表新贴：</td>
<td><input type=submit name=wsubmit value='可以发表啦' onclick="return check(write_frm)">　&nbsp;<input type=button value='预览内容'>　&nbsp;<input type=reset value='清除重写'>　&nbsp;（按 Ctrl + Enter 可快速发表）</td></tr>
</form></table><br>
<%
end function

function edit_chk()
  call time_load(1,0,1)
  dim topic,word,founderr,rs,sql,now_tim,new_id
  topic=trim(request.form("topic"))
  word=code_word(request.form("jk_word"))
  founderr=""
  if topic_yes="yes" then
    if topic="" or len(topic)>50 then
      founderr=founderr & "<br><li>贴子的 <font class=founderr>主题</font> 是必须要的；且长度不能超过50个字符！"
    end if
  else
    if len(topic)>50 then
      founderr=founderr & "<br><li>贴子的 <font class=founderr>主题</font> 长度不能超过50个字符！"
    end if
  end if
  if word="" or isnull(word) or len(word)>word_size*1024 then
    founderr=founderr & "<br><li>贴子的 <font class=founderr>内容</font> 是必须要的；且大小不能超过"&word_size&"KB！"
  end if
  if founderr="" then
    word=word&vbcrlf&vbcrlf&"[ALIGN=right][COLOR=#000066][本贴已被 "&login_username&" 于 "&now()&" 修改过][/COLOR][/ALIGN]"
    'rs.close
    set rs=nothing
    sql="select topic,word from bbs_data where id="&editid
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,3
    if not(rs.eof and rs.bof) then
      rs("topic")=topic
      rs("word")=word
      rs.update
    end if
    
    if topic_yes="yes" then conn.execute("update bbs_topic set topic='"&topic&"' where id="&iid)
    call upload_note(index_url,iid)
    
    call time_load(0,0,1)
    
    response.write VbCrLf & "<table border=0 width=300>" & _
        	   VbCrLf & "<tr><td align=center height=30><font class=red>您已成功编辑了贴子！</font></td></tr>" & _
        	   VbCrLf & "<tr><td height=30>您现在可以选择以下操作：</td></tr>" & _
        	   VbCrLf & "<tr><td>　　1、<a href='forum_view.asp?forum_id=" & forumid & "&view_id=" & iid & "'>您所编辑的帖子</a>" & _
        	   VbCrLf & "<tr><td>　　2、<a href='forum_list.asp?forum_id=" & forumid & "'>返回 <b>" & forumname & "</b></a></td></tr>" & _
        	   VbCrLf & "<tr><td>　　3、<a href='forum.asp'>返回论坛首页</a></td></tr>" & _
        	   VbCrLf & "<tr><td height=30>系统将在 " & web_var(web_num,5) & " 秒钟后自动返回 <b>" & forumname & "</b> 。</td></tr>" & _
		   VbCrLf & "</table><meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=forum_list.asp?forum_id=" & forumid & "'>"
  else
    response.write found_error(founderr,"300")
  end if
end function
%>