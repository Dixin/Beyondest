<!-- #include file="include/config_forum.asp" -->
<% If Not(IsNumeric(forumid)) Or Not(IsNumeric(viewid)) Then Call cookies_type("view_id") %>
<!-- #include file="include/config_upload.asp" -->
<!-- #include file="include/config_frm.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim re_topic
Dim re_word
Dim reid
Call forum_first()
Call forum_word()
tit = forumname & "（回复贴子）"

Call web_head(2,0,2,0,0)

If Int(popedom_format(login_popedom,41)) Then Call close_conn():Call cookies_type("locked")

sql    = "select bbs_topic.id,bbs_topic.topic,bbs_topic.islock,bbs_data.username,bbs_data.tim,bbs_data.word from bbs_topic inner join bbs_data on bbs_topic.id=bbs_data.reply_id where bbs_data.forum_id=" & forumid & " and bbs_data.id=" & viewid
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing:close_conn
    Call cookies_type("view_id")
End If

If Int(rs("islock")) = 1 Then
    rs.Close:Set rs = Nothing
    close_conn
    Call cookies_type("islock")
End If

re_topic    = rs("topic")
re_word     = ""
reid        = rs("id")

If Trim(Request.querystring("quote")) = "yes" Then
    re_word = "[QUOTE][b]以下是引用[i]" & rs("username") & "在" & rs("tim") & "[/i]的发言：[/b][br]" & Replace(rs("word"),vbcrlf,"[br]") & "[/QUOTE]" & vbcrlf
End If

rs.Close:Set rs = Nothing
'-----------------------------------center---------------------------------
Response.Write forum_top("回复贴子") & kong

If Trim(Request.form("reply")) = "ok" Then
    Response.Write "<table border=0><tr><td align=center height=200>"

    If post_chk() = "no" Then
        Response.Write web_var(web_error,1)
    Else
        Response.Write reply_chk()
    End If

    Response.Write "</td></tr></table>"
Else
    Response.Write reply_type()
End If

'---------------------------------center end-------------------------------
Call web_end(0)

Function reply_type() %>
<script language=javascript><!--
function check(write_frm)
{
  if(write_frm.topic.value.length>50)
  {
   alert("你还没完全留下所需信息！\r\n\n回贴的 主题 长度不能超过50个字符。");
   return false;
  }
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
<form name=write_frm action='forum_reply.asp?forum_id=<% = forumid %>&view_id=<% = viewid %>' method=post onsubmit="frm_submitonce(this);">
<input type=hidden name=reply value='ok'>
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
  <td>&nbsp;<input type=text name=topic value='回复：<% Response.Write re_topic %>' size=60 maxlength=50><% Response.Write redx %>长度不能超过50</td>
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
<td><table border=0><tr><td><textarea name=jk_word rows=12 cols=95 title='按 Ctrl+Enter 可直接发送' onkeydown="javascript:frm_quicksubmit();"><% If re_word <> "" Then Response.Write re_word %></textarea></td></tr></table></td>
</tr>
<tr<% Response.Write format_table(3,1) %>><td align=center>上传文件：</td><td>&nbsp;<iframe frameborder=0 name=upload_frame width='99%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=&uptext=jk_word'></iframe></td></tr>
<tr height=30<% Response.Write forum_table3 %>><td align=center>E M 贴图：</td><td>&nbsp;&nbsp;<script language=javascript>jk_em_type('s');</script></td></tr>
<tr<% Response.Write format_table(3,1) %>><td colspan=2 align=center height=60>&nbsp;&nbsp;<script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<% Response.Write forum_table3 %>>
<td>回复贴子：</td>
<td><input type=submit name=wsubmit value='可以回复啦' onclick="return check(write_frm)">　&nbsp;<input type=button value='预览回复'>　&nbsp;<input type=reset value='清除重写'>　&nbsp;（按 Ctrl + Enter 可快速发表）</td></tr>
</form></table><br>
<%
End Function

Function reply_chk()
    Call time_load(1,0,1)
    Dim topic
    Dim word
    Dim founderr
    Dim rs
    Dim sql
    Dim now_tim
    Dim icon
    topic        = code_form(Request.form("topic"))
    word         = code_word(Request.form("jk_word"))
    founderr     = ""

    If login_username = "" Or IsNull(login_username) Then
        founderr = founderr & "<br><li><font class=red_2>您还没有登陆本站！因此不能发表留言。</font>"
    End If

    If Len(topic) > 50 Then
        founderr = founderr & "<br><li>回贴的 <font class=founderr>主题</font> 长度不能超过50个字符！"
    End If

    If word = "" Or IsNull(word) Or Len(word) > word_size*1024 Then
        founderr = founderr & "<br><li>回贴的 <font class=founderr>内容</font> 是必须要的；且大小不能超过" & word_size & "KB！"
    End If

    If founderr = "" Then
        icon = Trim(Request.form("icon"))
        If var_null(icon) = "" Then icon = "0"

        sql = "insert into bbs_data (forum_id,reply_id,username,topic,icon,word,tim,ip,sys) " & _
        "values(" & forumid & "," & reid & ",'" & login_username & "','" & topic & "','" & icon & "','" & word & "','" & now_time & "','" & ip_sys(1,1) & "','" & ip_sys(3,0) & "')"
        conn.execute(sql)

        sql = "update bbs_topic set re_counter=re_counter+1,re_username='" & login_username & "',re_tim='" & now_time & "' where forum_id=" & forumid & " and id=" & reid
        conn.execute(sql)

        sql = "update user_data set bbs_counter=bbs_counter+1,integral=integral+1 where username='" & login_username & "'"
        conn.execute(sql)
        sql = "update configs set num_data=num_data+1 where id=1"
        conn.execute(sql)
        'sql="update bbs_forum set forum_data_num=forum_data_num+1 where forum_id="&forumid
        sql = "update bbs_forum set forum_data_num=forum_data_num+1,forum_new_info='"

        If Len(topic) > 0 Then
            sql = sql & code_form(Left(login_username & "|" & now_time & "|" & reid & "|" & topic,60))
        Else
            sql = sql & code_form(Left(login_username & "|" & now_time & "|" & reid & "|" & Replace(word,vbcrlf,""),60))
        End If

        sql = sql & "' where forum_id=" & forumid
        conn.execute(sql)

        Call upload_note(index_url,reid)
        Call time_load(0,0,1)

        Response.Write VbCrLf & "<table border=0 width=300>" & _
        VbCrLf & "<tr><td align=center height=30><font class=red>贴子回复成功！谢谢您的发贴。</font></td></tr>" & _
        VbCrLf & "<tr><td height=30>您现在可以选择以下操作：</td></tr>" & _
        VbCrLf & "<tr><td>　　1、<a href='forum_view.asp?forum_id=" & forumid & "&view_id=" & reid & "'>您所回复的帖子</a>" & _
        VbCrLf & "<tr><td>　　2、<a href='forum_list.asp?forum_id=" & forumid & "'>返回 <b>" & forumname & "</b></a></td></tr>" & _
        VbCrLf & "<tr><td>　　3、<a href='forum.asp'>返回论坛首页</a></td></tr>" & _
        VbCrLf & "<tr><td height=30>系统将在 " & web_var(web_num,5) & " 秒钟后自动返回 <b>" & forumname & "</b> 。</td></tr>" & _
        VbCrLf & "</table><meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=forum_list.asp?forum_id=" & forumid & "'>"
    Else
        Response.Write found_error(founderr,"350")
    End If

End Function %>