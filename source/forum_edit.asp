<!-- #include file="INCLUDE/config_forum.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

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
tit=forumname&"���༭���ӣ�"

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
response.write forum_top("�༭����") & kong

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
   alert("�㻹û��ȫ����������Ϣ��\r\n\n���ӵ� ���� �Ǳ���Ҫ�ģ�\n�ҳ��Ȳ��ܳ���50���ַ���");
   return false;
  }
<% else %>
  if(write_frm.topic.value.length>50)
  {
   alert("�㻹û��ȫ����������Ϣ��\r\n\n������ ���� ���Ȳ��ܳ���50���ַ���");
   return false;
  }
<% end if %>
  if(write_frm.jk_word.value=="" || write_frm.jk_word.value.length><%response.write word_size*1024%>)
  {
   alert("�㻹û��ȫ����������Ϣ��\r\n\n���ӵ� ���� �Ǳ���Ҫ�ģ�\n�Ҵ�С���ܳ���<%response.write word_size%>KB��");
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
<td width='20%' align=center>�û���Ϣ��</td>
<td width='80%'>&nbsp;&nbsp;�û�����<input type=username name=username value='<%response.write login_username%>' size=18 maxlength=20>&nbsp;&nbsp;
���룺<input type=password name=password value='<%response.write login_password%>' size=18 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='user_main.asp?username=<%response.write server.htmlencode(login_username)%>'>�û�����</a> ]</font>&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='login.asp?action=logout'>�˳���½</a> ]</font></td>
</tr>
<tr height=30<%response.write format_table(3,1)%>>
<td align=center>�������⣺</td>
<td>
  <table border=0 cellspacing=0 cellpadding=0><tr>
  <td>&nbsp;&nbsp;<%call frm_topic("write_frm","topic")%></td>
  <td>&nbsp;<input type=text name=topic value='<%response.write rs("data_topic")%>' size=60 maxlength=50><%response.write redx%>���Ȳ��ܳ���50</td>
  </tr></table>
</td>
</tr>
<tr height=30<%response.write forum_table3%>>
<td align=center>��ǰ���飺</td>
<td>&nbsp;&nbsp;<% response.write icon_type(9,3) %>
</td>
</tr>
<tr height=35<%response.write format_table(3,1)%>>
<td align=center><%call frm_ubb_type()%></td>
<td><%call frm_ubb("write_frm","jk_word","&nbsp;&nbsp;")%></td>
</tr>
<tr align=center<%response.write forum_table3%>>
<td><table border=0><tr><td class=htd>�������ݣ�<br><br><%response.write word_remark%><br><br><br></td></tr></table></td>
<td><table border=0><tr><td><textarea name=jk_word rows=12 cols=95 title='�� Ctrl+Enter ��ֱ�ӷ���' onkeydown="javascript:frm_quicksubmit();"><%response.write rs("data_word")%></textarea></td></tr></table></td>
</tr>
<tr<%response.write format_table(3,1)%>><td align=center>�ϴ��ļ���</td><td>&nbsp;<iframe frameborder=0 name=upload_frame width='99%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=&uptext=jk_word'></iframe></td></tr>
<tr height=30<%response.write forum_table3%>><td align=center>E M ��ͼ��</td><td>&nbsp;&nbsp;<script language=javascript>jk_em_type('s');</script></td></tr>
<tr<%response.write format_table(3,1)%>><td colspan=2 align=center height=60><script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<%response.write forum_table3%>>
<td>����������</td>
<td><input type=submit name=wsubmit value='���Է�����' onclick="return check(write_frm)">��&nbsp;<input type=button value='Ԥ������'>��&nbsp;<input type=reset value='�����д'>��&nbsp;���� Ctrl + Enter �ɿ��ٷ���</td></tr>
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
      founderr=founderr & "<br><li>���ӵ� <font class=founderr>����</font> �Ǳ���Ҫ�ģ��ҳ��Ȳ��ܳ���50���ַ���"
    end if
  else
    if len(topic)>50 then
      founderr=founderr & "<br><li>���ӵ� <font class=founderr>����</font> ���Ȳ��ܳ���50���ַ���"
    end if
  end if
  if word="" or isnull(word) or len(word)>word_size*1024 then
    founderr=founderr & "<br><li>���ӵ� <font class=founderr>����</font> �Ǳ���Ҫ�ģ��Ҵ�С���ܳ���"&word_size&"KB��"
  end if
  if founderr="" then
    word=word&vbcrlf&vbcrlf&"[ALIGN=right][COLOR=#000066][�����ѱ� "&login_username&" �� "&now()&" �޸Ĺ�][/COLOR][/ALIGN]"
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
        	   VbCrLf & "<tr><td align=center height=30><font class=red>���ѳɹ��༭�����ӣ�</font></td></tr>" & _
        	   VbCrLf & "<tr><td height=30>�����ڿ���ѡ�����²�����</td></tr>" & _
        	   VbCrLf & "<tr><td>����1��<a href='forum_view.asp?forum_id=" & forumid & "&view_id=" & iid & "'>�����༭������</a>" & _
        	   VbCrLf & "<tr><td>����2��<a href='forum_list.asp?forum_id=" & forumid & "'>���� <b>" & forumname & "</b></a></td></tr>" & _
        	   VbCrLf & "<tr><td>����3��<a href='forum.asp'>������̳��ҳ</a></td></tr>" & _
        	   VbCrLf & "<tr><td height=30>ϵͳ���� " & web_var(web_num,5) & " ���Ӻ��Զ����� <b>" & forumname & "</b> ��</td></tr>" & _
		   VbCrLf & "</table><meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=forum_list.asp?forum_id=" & forumid & "'>"
  else
    response.write found_error(founderr,"300")
  end if
end function
%>