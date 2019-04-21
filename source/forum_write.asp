<!-- #include file="include/config_forum.asp" -->
<% If Not(IsNumeric(forumid)) Then Call cookies_type("forum_id") %>
<!-- #include file="include/config_frm.asp" -->
<!-- #include file="include/config_upload.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Call forum_first()
Call forum_word()
tit = forumname & "（发表新贴）"

Call web_head(2,0,2,0,0)

If Int(popedom_format(login_popedom,41)) Then Call close_conn():Call cookies_type("locked")
'-----------------------------------center---------------------------------
Response.Write forum_top("发表新贴") & kong

If Trim(Request.form("write")) = "ok" Then
    Response.Write "<table border=0><tr><td align=center height=200>"

    If post_chk() = "no" Then
        Response.Write web_var(web_error,1)
    Else
        Response.Write write_chk()
    End If

    Response.Write "</td></tr></table>"
Else
    Response.Write write_type()
End If

'---------------------------------center end-------------------------------
Call web_end(0)

Function write_type() %>
<script language=javascript><!--
function check(write_frm)
{
  if(write_frm.topic.value=="" || write_frm.topic.value.length>50)
  {
   alert("你还没完全留下所需信息！\r\n\n贴子的 主题 是必须要的；\n且长度不能超过50个字符。");
   return false;
  }
  if(write_frm.jk_word.value=="" || write_frm.jk_word.value.length><% Response.Write word_size*1024 %>)
  {
   alert("你还没完全留下所需信息！\r\n\n贴子的 内容 是必须要的；\n且大小不能超过<% Response.Write word_size %>KB。");
   return false;
  }
}
--></script>
<script language=javascript src='style/em_type.js'></script>
<script language=javascript src='style/forum_ok.js'></script>
<% Response.Write forum_table1 %><tr height=26><td colspan=2 background='images/<% = web_var(web_config,5) %>/bar_1_bg.gif'><img src=IMAGES/SMALL/FK4.GIF> <font class=end><b>撰写话题</b><font></td></tr>
<form name=write_frm action='forum_write.asp?forum_id=<% Response.Write forumid %>' method=post onsubmit="frm_submitonce(this);">
<input type=hidden name=write value='ok'>
<input type=hidden name=upid value=''>
<tr height=30<% Response.Write forum_table3 %>>
<td width='20%' align=center bgcolor='<% = web_var(web_color,6) %>'>用户信息：</td>
<td width='80%'>&nbsp;&nbsp;用户名：<input type=username name=username value='<% Response.Write login_username %>' size=18 maxlength=20>&nbsp;&nbsp;
密码：<input type=password name=password value='<% Response.Write login_password %>' size=18 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='user_main.asp?username=<% Response.Write Server.htmlencode(login_username) %>'>用户中心</a> ]</font>&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='login.asp?action=logout'>退出登陆</a> ]</font></td>
</tr>
<tr height=30<% Response.Write format_table(3,1) %>>
<td align=center bgcolor='<% = web_var(web_color,6) %>'>贴子主题：</td>
<td>
  <table border=0 cellspacing=0 cellpadding=0><tr>
  <td>&nbsp;&nbsp;<% Call frm_topic("write_frm","topic") %></td>
  <td>&nbsp;<input type=text name=topic size=60 maxlength=50><% Response.Write redx %>长度不能超过50</td>
  </tr></table>
</td>
</tr>
<tr height=30<% Response.Write forum_table3 %>>
<td align=center bgcolor='<% = web_var(web_color,6) %>'>当前心情：</td>
<td>&nbsp;&nbsp;<% Response.Write icon_type(9,3) %>
</td>
</tr>
<tr height=35<% Response.Write format_table(3,1) %>>
<td align=center bgcolor='<% = web_var(web_color,6) %>'><% Call frm_ubb_type() %></td>
<td><% Call frm_ubb("write_frm","jk_word","&nbsp;&nbsp;") %></td>
</tr>
<tr align=center<% Response.Write forum_table3 %>>
<td bgcolor='<% = web_var(web_color,6) %>'><table border=0><tr><td class=htd>贴子内容：<br><% Call frm_word_size("write_frm","jk_word",word_size,"贴子内容") %><br><br><% Response.Write word_remark %><br><br><br></td></tr></table></td>
<td><table border=0><tr><td><textarea name=jk_word rows=12 cols=93 title='按 Ctrl+Enter 可直接发送' onkeydown="javascript:frm_quicksubmit();"></textarea></td></tr></table></td>
</tr>
<tr<% Response.Write format_table(3,1) %>><td align=center bgcolor='<% = web_var(web_color,6) %>'>上传文件：</td><td>&nbsp;<iframe frameborder=0 name=upload_frame width='99%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=&uptext=jk_word'></iframe></td></tr>
<tr height=30<% Response.Write forum_table3 %>><td align=center bgcolor='<% = web_var(web_color,6) %>'>E M 贴图：</td><td>&nbsp;&nbsp;<script language=javascript>jk_em_type('s');</script></td></tr>
<tr<% Response.Write format_table(3,1) %>><td colspan=2 align=center height=60>&nbsp;&nbsp;<script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<% Response.Write forum_table3 %>>
<td bgcolor='<% = web_var(web_color,6) %>'>发表新贴：</td>
<td><input type=submit name=wsubmit value='可以发表啦' onclick="return check(write_frm)">　&nbsp;<input type=button value='预览内容'>　&nbsp;<input type=reset value='清除重写'>　&nbsp;（按 Ctrl + Enter 可快速发表）</td></tr>
</form></table><br>
<%
End Function

Function write_chk()
    Call time_load(1,0,1)
    Dim topic
    Dim word
    Dim founderr
    Dim rs
    Dim sql
    Dim now_tim
    Dim new_id
    topic        = code_form(Trim(Request.form("topic")))
    word         = code_word(Request.form("jk_word"))
    founderr     = ""

    If login_username = "" Or IsNull(login_username) Then
        founderr = founderr & "<br><li><font class=red_2>您还没有登陆本站！因此不能发表留言。</font>"
    End If

    If topic = "" Or Len(topic) > 50 Then
        founderr = founderr & "<br><li>贴子的 <font class=founderr>主题</font> 是必须要的；且长度不能超过50个字符！"
    End If

    If word = "" Or IsNull(word) Or Len(word) > word_size*1024 Then
        founderr = founderr & "<br><li>贴子的 <font class=founderr>内容</font> 是必须要的；且大小不能超过" & word_size & "KB！"
    End If

    If founderr = "" Then
        sql = "insert into bbs_topic (forum_id,username,topic,icon,counter,tim,re_username,re_counter,re_tim,istop,islock,isgood) " & _
        "values (" & forumid & ",'" & login_username & "','" & topic & "','" & Trim(Request.form("icon")) & "',0,'" & now_time & "','" & login_username & "',0,'" & now_time & "',0,0,0)"
        conn.execute(sql)

        new_id = first_id("bbs_topic")

        sql = "insert into bbs_data (forum_id,reply_id,username,topic,icon,word,tim,ip,sys) " & _
        "values (" & forumid & "," & new_id & ",'" & login_username & "','" & topic & "','" & Trim(Request.form("icon")) & "','" & word & "','" & now_time & "','" & ip_sys(1,1) & "','" & ip_sys(3,0) & "')"
        Set rs = conn.execute(sql)

        sql = "update user_data set bbs_counter=bbs_counter+1,integral=integral+2 where username='" & login_username & "'"
        conn.execute(sql)
        sql = "update configs set num_topic=num_topic+1,num_data=num_data+1 where id=1"
        conn.execute(sql)
        sql = "update bbs_forum set forum_topic_num=forum_topic_num+1,forum_data_num=forum_data_num+1,forum_new_info='" & login_username & "|" & now_time & "|" & new_id & "|" & topic & "' where forum_id=" & forumid
        conn.execute(sql)

        Call upload_note(index_url,new_id)

        Call time_load(0,0,1)

        Response.Write VbCrLf & "<table border=0 width=300>" & _
        VbCrLf & "<tr><td align=center height=30><font class=red>贴子发表成功！谢谢您的发贴。</font></td></tr>" & _
        VbCrLf & "<tr><td height=30>您现在可以选择以下操作：</td></tr>" & _
        VbCrLf & "<tr><td>　　1、<a href='forum_view.asp?forum_id=" & forumid & "&view_id=" & new_id & "'>您所发表的帖子</a>" & _
        VbCrLf & "<tr><td>　　2、<a href='forum_list.asp?forum_id=" & forumid & "'>返回 <b>" & forumname & "</b></a></td></tr>" & _
        VbCrLf & "<tr><td>　　3、<a href='forum.asp'>返回论坛首页</a></td></tr>" & _
        VbCrLf & "<tr><td height=30>系统将在 " & web_var(web_num,5) & " 秒钟后自动返回 <b>" & forumname & "</b> 。</td></tr>" & _
        VbCrLf & "</table><meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=forum_list.asp?forum_id=" & forumid & "'>"
    Else
        Response.Write found_error(founderr,"350")
    End If

End Function %>