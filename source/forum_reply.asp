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
tit = forumname & "���ظ����ӣ�"

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
    re_word = "[QUOTE][b]����������[i]" & rs("username") & "��" & rs("tim") & "[/i]�ķ��ԣ�[/b][br]" & Replace(rs("word"),vbcrlf,"[br]") & "[/QUOTE]" & vbcrlf
End If

rs.Close:Set rs = Nothing
'-----------------------------------center---------------------------------
Response.Write forum_top("�ظ�����") & kong

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
   alert("�㻹û��ȫ����������Ϣ��\r\n\n������ ���� ���Ȳ��ܳ���50���ַ���");
   return false;
  }
  if(write_frm.jk_word.value=="" || write_frm.jk_word.value.length><% Response.Write word_size*1024 %>)
  {
   alert("�㻹û��ȫ����������Ϣ��\r\n\n���ӵ� ���� �Ǳ���Ҫ�ģ�\n�Ҵ�С���ܳ���<% Response.Write word_size %>KB��");
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
<td width='20%' align=center>�û���Ϣ��</td>
<td width='80%'>&nbsp;&nbsp;�û�����<input type=username name=username value='<% Response.Write login_username %>' size=18 maxlength=20>&nbsp;&nbsp;
���룺<input type=password name=password value='<% Response.Write login_password %>' size=18 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='user_main.asp?username=<% Response.Write Server.htmlencode(login_username) %>'>�û�����</a> ]</font>&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='login.asp?action=logout'>�˳���½</a> ]</font></td>
</tr>
<tr height=30<% Response.Write format_table(3,1) %>>
<td align=center>�������⣺</td>
<td>
  <table border=0 cellspacing=0 cellpadding=0><tr>
  <td>&nbsp;&nbsp;<% Call frm_topic("write_frm","topic") %></td>
  <td>&nbsp;<input type=text name=topic value='�ظ���<% Response.Write re_topic %>' size=60 maxlength=50><% Response.Write redx %>���Ȳ��ܳ���50</td>
  </tr></table>
</td>
</tr>
<tr height=30<% Response.Write forum_table3 %>>
<td align=center>��ǰ���飺</td>
<td>&nbsp;&nbsp;<% Response.Write icon_type(9,3) %>
</td>
</tr>
<tr height=35<% Response.Write format_table(3,1) %>>
<td align=center><% Call frm_ubb_type() %></td>
<td><% Call frm_ubb("write_frm","jk_word","&nbsp;&nbsp;") %></td>
</tr>
<tr align=center<% Response.Write forum_table3 %>>
<td><table border=0><tr><td class=htd>�������ݣ�<br><br><% Response.Write word_remark %><br><br><br></td></tr></table></td>
<td><table border=0><tr><td><textarea name=jk_word rows=12 cols=95 title='�� Ctrl+Enter ��ֱ�ӷ���' onkeydown="javascript:frm_quicksubmit();"><% If re_word <> "" Then Response.Write re_word %></textarea></td></tr></table></td>
</tr>
<tr<% Response.Write format_table(3,1) %>><td align=center>�ϴ��ļ���</td><td>&nbsp;<iframe frameborder=0 name=upload_frame width='99%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=&uptext=jk_word'></iframe></td></tr>
<tr height=30<% Response.Write forum_table3 %>><td align=center>E M ��ͼ��</td><td>&nbsp;&nbsp;<script language=javascript>jk_em_type('s');</script></td></tr>
<tr<% Response.Write format_table(3,1) %>><td colspan=2 align=center height=60>&nbsp;&nbsp;<script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<% Response.Write forum_table3 %>>
<td>�ظ����ӣ�</td>
<td><input type=submit name=wsubmit value='���Իظ���' onclick="return check(write_frm)">��&nbsp;<input type=button value='Ԥ���ظ�'>��&nbsp;<input type=reset value='�����д'>��&nbsp;���� Ctrl + Enter �ɿ��ٷ���</td></tr>
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
        founderr = founderr & "<br><li><font class=red_2>����û�е�½��վ����˲��ܷ������ԡ�</font>"
    End If

    If Len(topic) > 50 Then
        founderr = founderr & "<br><li>������ <font class=founderr>����</font> ���Ȳ��ܳ���50���ַ���"
    End If

    If word = "" Or IsNull(word) Or Len(word) > word_size*1024 Then
        founderr = founderr & "<br><li>������ <font class=founderr>����</font> �Ǳ���Ҫ�ģ��Ҵ�С���ܳ���" & word_size & "KB��"
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
        VbCrLf & "<tr><td align=center height=30><font class=red>���ӻظ��ɹ���лл���ķ�����</font></td></tr>" & _
        VbCrLf & "<tr><td height=30>�����ڿ���ѡ�����²�����</td></tr>" & _
        VbCrLf & "<tr><td>����1��<a href='forum_view.asp?forum_id=" & forumid & "&view_id=" & reid & "'>�����ظ�������</a>" & _
        VbCrLf & "<tr><td>����2��<a href='forum_list.asp?forum_id=" & forumid & "'>���� <b>" & forumname & "</b></a></td></tr>" & _
        VbCrLf & "<tr><td>����3��<a href='forum.asp'>������̳��ҳ</a></td></tr>" & _
        VbCrLf & "<tr><td height=30>ϵͳ���� " & web_var(web_num,5) & " ���Ӻ��Զ����� <b>" & forumname & "</b> ��</td></tr>" & _
        VbCrLf & "</table><meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=forum_list.asp?forum_id=" & forumid & "'>"
    Else
        Response.Write found_error(founderr,"350")
    End If

End Function %>