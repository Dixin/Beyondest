<!-- #include file="INCLUDE/config_forum.asp" -->
<% if not(isnumeric(forumid)) or not(isnumeric(viewid)) then call cookies_type("view_id") %>
<!-- #include file="INCLUDE/config_upload.asp" -->
<!-- #include file="INCLUDE/config_frm.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim re_topic,re_word,reid
call forum_first()
call forum_word()
tit=forumname&"���ظ����ӣ�"

call web_head(2,0,2,0,0)

if int(popedom_format(login_popedom,41)) then call close_conn():call cookies_type("locked")

sql="select bbs_topic.id,bbs_topic.topic,bbs_topic.islock,bbs_data.username,bbs_data.tim,bbs_data.word from bbs_topic inner join bbs_data on bbs_topic.id=bbs_data.reply_id where bbs_data.forum_id="&forumid&" and bbs_data.id="&viewid
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  rs.close:set rs=nothing:close_conn
  call cookies_type("view_id")
end if
if int(rs("islock"))=1 then
  rs.close:set rs=nothing
  close_conn
  call cookies_type("islock")
end if
re_topic=rs("topic")
re_word=""
reid=rs("id")
if trim(request.querystring("quote"))="yes" then
  re_word="[QUOTE][b]����������[i]"&rs("username")&"��"&rs("tim")&"[/i]�ķ��ԣ�[/b][br]"&replace(rs("word"),vbcrlf,"[br]")&"[/QUOTE]"&vbcrlf
end if
rs.close:set rs=nothing
'-----------------------------------center---------------------------------
response.write forum_top("�ظ�����") & kong

if trim(request.form("reply"))="ok" then
  response.write "<table border=0><tr><td align=center height=200>"
  if post_chk()="no" then
    response.write web_var(web_error,1)
  else
    response.write reply_chk()
  end if
  response.write "</td></tr></table>"
else
  response.write reply_type()
end if
'---------------------------------center end-------------------------------
call web_end(0)

function reply_type()
%>
<script language=javascript><!--
function check(write_frm)
{
  if(write_frm.topic.value.length>50)
  {
   alert("�㻹û��ȫ����������Ϣ��\r\n\n������ ���� ���Ȳ��ܳ���50���ַ���");
   return false;
  }
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
<form name=write_frm action='forum_reply.asp?forum_id=<%=forumid%>&view_id=<%=viewid%>' method=post onsubmit="frm_submitonce(this);">
<input type=hidden name=reply value='ok'>
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
  <td>&nbsp;<input type=text name=topic value='�ظ���<%response.write re_topic%>' size=60 maxlength=50><%response.write redx%>���Ȳ��ܳ���50</td>
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
<td><table border=0><tr><td><textarea name=jk_word rows=12 cols=95 title='�� Ctrl+Enter ��ֱ�ӷ���' onkeydown="javascript:frm_quicksubmit();"><% if re_word<>"" then response.write re_word %></textarea></td></tr></table></td>
</tr>
<tr<%response.write format_table(3,1)%>><td align=center>�ϴ��ļ���</td><td>&nbsp;<iframe frameborder=0 name=upload_frame width='99%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=&uptext=jk_word'></iframe></td></tr>
<tr height=30<%response.write forum_table3%>><td align=center>E M ��ͼ��</td><td>&nbsp;&nbsp;<script language=javascript>jk_em_type('s');</script></td></tr>
<tr<%response.write format_table(3,1)%>><td colspan=2 align=center height=60>&nbsp;&nbsp;<script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<%response.write forum_table3%>>
<td>�ظ����ӣ�</td>
<td><input type=submit name=wsubmit value='���Իظ���' onclick="return check(write_frm)">��&nbsp;<input type=button value='Ԥ���ظ�'>��&nbsp;<input type=reset value='�����д'>��&nbsp;���� Ctrl + Enter �ɿ��ٷ���</td></tr>
</form></table><br>
<%
end function

function reply_chk()
  call time_load(1,0,1)
  dim topic,word,founderr,rs,sql,now_tim,icon
  topic=code_form(request.form("topic"))
  word=code_word(request.form("jk_word"))
  founderr=""
  if login_username="" or isnull(login_username) then
    founderr=founderr & "<br><li><font class=red_2>����û�е�½��վ����˲��ܷ������ԡ�</font>"
  end if
  if len(topic)>50 then
    founderr=founderr & "<br><li>������ <font class=founderr>����</font> ���Ȳ��ܳ���50���ַ���"
  end if
  if word="" or isnull(word) or len(word)>word_size*1024 then
    founderr=founderr & "<br><li>������ <font class=founderr>����</font> �Ǳ���Ҫ�ģ��Ҵ�С���ܳ���"&word_size&"KB��"
  end if
  if founderr="" then
    icon=trim(request.form("icon"))
    if var_null(icon)="" then icon="0"
    
    sql="insert into bbs_data (forum_id,reply_id,username,topic,icon,word,tim,ip,sys) " & _
	"values("&forumid&","&reid&",'"&login_username&"','"&topic&"','"&icon&"','"&word&"','"&now_time&"','"&ip_sys(1,1)&"','"&ip_sys(3,0)&"')"
    conn.execute(sql)
    
    sql="update bbs_topic set re_counter=re_counter+1,re_username='"&login_username&"',re_tim='"&now_time&"' where forum_id="&forumid&" and id="&reid
    conn.execute(sql)
    
    sql="update user_data set bbs_counter=bbs_counter+1,integral=integral+1 where username='"&login_username&"'"
    conn.execute(sql)
    sql="update configs set num_data=num_data+1 where id=1"
    conn.execute(sql)
    'sql="update bbs_forum set forum_data_num=forum_data_num+1 where forum_id="&forumid
    sql="update bbs_forum set forum_data_num=forum_data_num+1,forum_new_info='"
    if len(topic)>0 then
      sql=sql&code_form(left(login_username&"|"&now_time&"|"&reid&"|"&topic,60))
    else
      sql=sql&code_form(left(login_username&"|"&now_time&"|"&reid&"|"&replace(word,vbcrlf,""),60))
    end if
    sql=sql&"' where forum_id="&forumid
    conn.execute(sql)
    
    call upload_note(index_url,reid)
    call time_load(0,0,1)
    
    response.write VbCrLf & "<table border=0 width=300>" & _
        	   VbCrLf & "<tr><td align=center height=30><font class=red>���ӻظ��ɹ���лл���ķ�����</font></td></tr>" & _
        	   VbCrLf & "<tr><td height=30>�����ڿ���ѡ�����²�����</td></tr>" & _
        	   VbCrLf & "<tr><td>����1��<a href='forum_view.asp?forum_id=" & forumid & "&view_id=" & reid & "'>�����ظ�������</a>" & _
        	   VbCrLf & "<tr><td>����2��<a href='forum_list.asp?forum_id=" & forumid & "'>���� <b>" & forumname & "</b></a></td></tr>" & _
        	   VbCrLf & "<tr><td>����3��<a href='forum.asp'>������̳��ҳ</a></td></tr>" & _
        	   VbCrLf & "<tr><td height=30>ϵͳ���� " & web_var(web_num,5) & " ���Ӻ��Զ����� <b>" & forumname & "</b> ��</td></tr>" & _
		   VbCrLf & "</table><meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=forum_list.asp?forum_id=" & forumid & "'>"
  else
    response.write found_error(founderr,"350")
  end if
end function
%>