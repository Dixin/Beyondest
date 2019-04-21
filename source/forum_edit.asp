<!-- #include file="include/config_forum.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim editid
editid = Trim(Request.querystring("edit_id"))
If Not(IsNumeric(forumid)) Or Not(IsNumeric(editid)) Then Call cookies_typ("view_id") %>
<!-- #include file="include/config_upload.asp" -->
<!-- #include file="include/config_frm.asp" -->
<!-- #include file="include/conn.asp" -->
<%
Dim topic_yes
Dim iid
Call forum_first()
Call forum_word()
tit = forumname & "（编辑贴子）"

Call web_head(2,0,2,0,0)

If Int(popedom_format(login_popedom,41)) Then Call close_conn():Call cookies_type("locked")

sql    = "select bbs_data.username as data_username,bbs_data.topic as data_topic,bbs_data.word as data_word,bbs_data.tim as data_tim,bbs_topic.id,bbs_topic.username as topic_username,bbs_topic.topic as topic_topic,bbs_topic.tim as topic_tim " & _
"from bbs_topic inner join bbs_data on bbs_topic.id=bbs_data.reply_id where bbs_data.id=" & editid & " and bbs_data.forum_id=" & forumid & " and bbs_topic.islock<>1"
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing:Call close_conn()
    Call cookies_type("edit_id")
End If

If format_user_power(login_username,login_mode,forumpower) <> "yes" Then

    If login_username <> rs("data_username") Then
        Set rs = Nothing:Call close_conn()
        Call cookies_type("edit_id")
    End If

End If

iid           = rs("id")

If rs("data_tim") = rs("topic_tim") And rs("data_username") = rs("topic_username") Then
    topic_yes = "yes"
End If

'-----------------------------------center---------------------------------
Response.Write forum_top("编辑贴子") & kong

If Trim(Request.form("edit")) = "ok" Then
    Response.Write "<table border=0><tr><td align=center height=200>"

    If post_chk() = "no" Then
        Response.Write web_var(web_error,1)
    Else
        Response.Write edit_chk()
    End If

    Response.Write "</td></tr></table>"
Else
    Response.Write edit_type()
End If

rs.Close:Set rs = Nothing
'---------------------------------center end-------------------------------
Call web_end(0)

Function edit_type() %>
<script language=javascript><!--
function check(write_frm)
{
<% If topic_yes = "yes" Then %>
  if(write_frm.topic.value=="" || write_frm.topic.value.length>50)
  {
   alert("你还没完全留下所需信息！\r\n\n贴子的 主题 是必须要的；\n且长度不能超过50个字符。");
   return false;
  }
<% Else %>
  if(write_frm.topic.value.length>50)
  {
   alert("你还没完全留下所需信息！\r\n\n回贴的 主题 长度不能超过50个字符。");
   return false;
  }
<% End If %>
  if(write_frm.jk_word.value=="" || write_frm.jk_word.value.length><% Response.Write word_size*1024 %>)
  {
   alert("你还没完全留下所需信息！\r\n\n贴子的 内容 是必须要的；\n且大小不能超过<% Response.Write word_size %>KB。");
   return false;
  }
}
-->
</script>
<script language=javascript src='style/em_type.js'></script>
<script language=javascript src='style/forum_ok.js'></script>
<% Response.Write forum_table1 %>
<form name=write_frm action='forum_edit.asp?forum_id=<% Response.Write forumid %>&edit_id=<% Response.Write editid %>' method=post onsubmit="frm_submitonce(this);">
<input type=hidden name=edit value='ok'>
<input type=hidden name=upid value=''>
<tr height=30<% Response.Write forum_table3 %>>
<td width='20%' align=center>用户信息：</td>
<td width='80%'>&nbsp;&nbsp;用户名：<input type=username name=username value='<% Response.Write login_username %>' size=18 maxlength=20>&nbsp;&nbsp;
密码：<input type=password name=password value='<% Response.Write login_password %>' size=18 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='user_main.asp?username=<% Response.Write Server.htmlencode(login_username) %>'>用户中心</a> ]</font>&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='login.asp?action=logout'>退出登陆</a> ]</font></td>
</tr>
<tr height=30<% Response.Write format_table(3,1) %>>
<td align=center>贴子主题：</td>
<td>
  <table border=0 cellspacing=0 cellpadding=0><tr>
  <td>&nbsp;&nbsp;<% Call frm_topic("write_frm","topic") %></td>
  <td>&nbsp;<input type=text name=topic value='<% Response.Write rs("data_topic") %>' size=60 maxlength=50><% Response.Write redx %>长度不能超过50</td>
  </tr></table>
</td>
</tr>
<tr height=30<% Response.Write forum_table3 %>>
<td align=center>当前心情：</td>
<td>&nbsp;&nbsp;<% Response.Write icon_type(9,3) %>
</td>
</tr>
<tr height=35<% Response.Write format_table(3,1) %>>
<td align=center><% Call frm_ubb_type() %></td>
<td><% Call frm_ubb("write_frm","jk_word","&nbsp;&nbsp;") %></td>
</tr>
<tr align=center<% Response.Write forum_table3 %>>
<td><table border=0><tr><td class=htd>贴子内容：<br><br><% Response.Write word_remark %><br><br><br></td></tr></table></td>
<td><table border=0><tr><td><textarea name=jk_word rows=12 cols=95 title='按 Ctrl+Enter 可直接发送' onkeydown="javascript:frm_quicksubmit();"><% Response.Write rs("data_word") %></textarea></td></tr></table></td>
</tr>
<tr<% Response.Write format_table(3,1) %>><td align=center>上传文件：</td><td>&nbsp;<iframe frameborder=0 name=upload_frame width='99%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=&uptext=jk_word'></iframe></td></tr>
<tr height=30<% Response.Write forum_table3 %>><td align=center>E M 贴图：</td><td>&nbsp;&nbsp;<script language=javascript>jk_em_type('s');</script></td></tr>
<tr<% Response.Write format_table(3,1) %>><td colspan=2 align=center height=60><script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<% Response.Write forum_table3 %>>
<td>发表新贴：</td>
<td><input type=submit name=wsubmit value='可以发表啦' onclick="return check(write_frm)">　&nbsp;<input type=button value='预览内容'>　&nbsp;<input type=reset value='清除重写'>　&nbsp;（按 Ctrl + Enter 可快速发表）</td></tr>
</form></table><br>
<%

End Function

Function edit_chk()
Call time_load(1,0,1)
Dim topic
Dim word
Dim founderr
Dim rs
Dim sql
Dim now_tim
Dim new_id
topic    = Trim(Request.form("topic"))
word     = code_word(Request.form("jk_word"))
founderr = ""

If topic_yes = "yes" Then

    If topic = "" Or Len(topic) > 50 Then
        founderr = founderr & "<br><li>贴子的 <font class=founderr>主题</font> 是必须要的；且长度不能超过50个字符！"
    End If

Else

    If Len(topic) > 50 Then
        founderr = founderr & "<br><li>贴子的 <font class=founderr>主题</font> 长度不能超过50个字符！"
    End If

End If

If word = "" Or IsNull(word) Or Len(word) > word_size*1024 Then
    founderr = founderr & "<br><li>贴子的 <font class=founderr>内容</font> 是必须要的；且大小不能超过" & word_size & "KB！"
End If

If founderr = "" Then
    word   = word & vbcrlf & vbcrlf & "[ALIGN=right][COLOR=#000066][本贴已被 " & login_username & " 于 " & Now() & " 修改过][/COLOR][/ALIGN]"
    'rs.close
    Set rs = Nothing
    sql    = "select topic,word from bbs_data where id=" & editid
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql,conn,1,3

    If Not(rs.eof And rs.bof) Then
        rs("topic") = topic
        rs("word") = word
        rs.update
    End If

    If topic_yes = "yes" Then conn.execute("update bbs_topic set topic='" & topic & "' where id=" & iid)
    Call upload_note(index_url,iid)

    Call time_load(0,0,1)

    Response.Write VbCrLf & "<table border=0 width=300>" & _
    VbCrLf & "<tr><td align=center height=30><font class=red>您已成功编辑了贴子！</font></td></tr>" & _
    VbCrLf & "<tr><td height=30>您现在可以选择以下操作：</td></tr>" & _
    VbCrLf & "<tr><td>　　1、<a href='forum_view.asp?forum_id=" & forumid & "&view_id=" & iid & "'>您所编辑的帖子</a>" & _
    VbCrLf & "<tr><td>　　2、<a href='forum_list.asp?forum_id=" & forumid & "'>返回 <b>" & forumname & "</b></a></td></tr>" & _
    VbCrLf & "<tr><td>　　3、<a href='forum.asp'>返回论坛首页</a></td></tr>" & _
    VbCrLf & "<tr><td height=30>系统将在 " & web_var(web_num,5) & " 秒钟后自动返回 <b>" & forumname & "</b> 。</td></tr>" & _
    VbCrLf & "</table><meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=forum_list.asp?forum_id=" & forumid & "'>"
Else
    Response.Write found_error(founderr,"300")
End If

End Function %>