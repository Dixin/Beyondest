<!-- #include file="INCLUDE/config_forum.asp" -->
<% if not(isnumeric(forumid)) then call cookies_type("forum_id") %>
<!-- #include file="INCLUDE/config_frm.asp" -->
<!-- #include file="INCLUDE/config_upload.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

call forum_first()
call forum_word()
tit=forumname&"������������"

call web_head(2,0,2,0,0)

if int(popedom_format(login_popedom,41)) then call close_conn():call cookies_type("locked")
'-----------------------------------center---------------------------------
response.write forum_top("��������") & kong

if trim(request.form("write"))="ok" then
  response.write "<table border=0><tr><td align=center height=200>"
  if post_chk()="no" then
    response.write web_var(web_error,1)
  else
    response.write write_chk()
  end if
  response.write "</td></tr></table>"
else
  response.write write_type()
end if

'---------------------------------center end-------------------------------
call web_end(0)

function write_type()
%>
<script language=javascript><!--
function check(write_frm)
{
  if(write_frm.topic.value=="" || write_frm.topic.value.length>50)
  {
   alert("�㻹û��ȫ����������Ϣ��\r\n\n���ӵ� ���� �Ǳ���Ҫ�ģ�\n�ҳ��Ȳ��ܳ���50���ַ���");
   return false;
  }
  if(write_frm.jk_word.value=="" || write_frm.jk_word.value.length><%response.write word_size*1024%>)
  {
   alert("�㻹û��ȫ����������Ϣ��\r\n\n���ӵ� ���� �Ǳ���Ҫ�ģ�\n�Ҵ�С���ܳ���<%response.write word_size%>KB��");
   return false;
  }
}
--></script>
<script language=javascript src='style/em_type.js'></script>
<script language=javascript src='style/forum_ok.js'></script>
<%response.write forum_table1%><tr height=26><td colspan=2 background='images/<%=web_var(web_config,5)%>/bar_1_bg.gif'><img src=IMAGES/SMALL/FK4.GIF> <font class=end><b>׫д����</b><font></td></tr>
<form name=write_frm action='forum_write.asp?forum_id=<%response.write forumid%>' method=post onsubmit="frm_submitonce(this);">
<input type=hidden name=write value='ok'>
<input type=hidden name=upid value=''>
<tr height=30<%response.write forum_table3%>>
<td width='20%' align=center bgcolor='<%=web_var(web_color,6)%>'>�û���Ϣ��</td>
<td width='80%'>&nbsp;&nbsp;�û�����<input type=username name=username value='<%response.write login_username%>' size=18 maxlength=20>&nbsp;&nbsp;
���룺<input type=password name=password value='<%response.write login_password%>' size=18 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='user_main.asp?username=<%response.write server.htmlencode(login_username)%>'>�û�����</a> ]</font>&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='login.asp?action=logout'>�˳���½</a> ]</font></td>
</tr>
<tr height=30<%response.write format_table(3,1)%>>
<td align=center bgcolor='<%=web_var(web_color,6)%>'>�������⣺</td>
<td>
  <table border=0 cellspacing=0 cellpadding=0><tr>
  <td>&nbsp;&nbsp;<%call frm_topic("write_frm","topic")%></td>
  <td>&nbsp;<input type=text name=topic size=60 maxlength=50><%response.write redx%>���Ȳ��ܳ���50</td>
  </tr></table>
</td>
</tr>
<tr height=30<%response.write forum_table3%>>
<td align=center bgcolor='<%=web_var(web_color,6)%>'>��ǰ���飺</td>
<td>&nbsp;&nbsp;<% response.write icon_type(9,3) %>
</td>
</tr>
<tr height=35<%response.write format_table(3,1)%>>
<td align=center bgcolor='<%=web_var(web_color,6)%>'><%call frm_ubb_type()%></td>
<td><%call frm_ubb("write_frm","jk_word","&nbsp;&nbsp;")%></td>
</tr>
<tr align=center<%response.write forum_table3%>>
<td bgcolor='<%=web_var(web_color,6)%>'><table border=0><tr><td class=htd>�������ݣ�<br><%call frm_word_size("write_frm","jk_word",word_size,"��������")%><br><br><%response.write word_remark%><br><br><br></td></tr></table></td>
<td><table border=0><tr><td><textarea name=jk_word rows=12 cols=93 title='�� Ctrl+Enter ��ֱ�ӷ���' onkeydown="javascript:frm_quicksubmit();"></textarea></td></tr></table></td>
</tr>
<tr<%response.write format_table(3,1)%>><td align=center bgcolor='<%=web_var(web_color,6)%>'>�ϴ��ļ���</td><td>&nbsp;<iframe frameborder=0 name=upload_frame width='99%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=&uptext=jk_word'></iframe></td></tr>
<tr height=30<%response.write forum_table3%>><td align=center bgcolor='<%=web_var(web_color,6)%>'>E M ��ͼ��</td><td>&nbsp;&nbsp;<script language=javascript>jk_em_type('s');</script></td></tr>
<tr<%response.write format_table(3,1)%>><td colspan=2 align=center height=60>&nbsp;&nbsp;<script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<%response.write forum_table3%>>
<td bgcolor='<%=web_var(web_color,6)%>'>����������</td>
<td><input type=submit name=wsubmit value='���Է�����' onclick="return check(write_frm)">��&nbsp;<input type=button value='Ԥ������'>��&nbsp;<input type=reset value='�����д'>��&nbsp;���� Ctrl + Enter �ɿ��ٷ���</td></tr>
</form></table><br>
<%
end function

function write_chk()
  call time_load(1,0,1)
  dim topic,word,founderr,rs,sql,now_tim,new_id
  topic=code_form(trim(request.form("topic")))
  word=code_word(request.form("jk_word"))
  founderr=""
  if login_username="" or isnull(login_username) then
    founderr=founderr & "<br><li><font class=red_2>����û�е�½��վ����˲��ܷ������ԡ�</font>"
  end if
  if topic="" or len(topic)>50 then
    founderr=founderr & "<br><li>���ӵ� <font class=founderr>����</font> �Ǳ���Ҫ�ģ��ҳ��Ȳ��ܳ���50���ַ���"
  end if
  if word="" or isnull(word) or len(word)>word_size*1024 then
    founderr=founderr & "<br><li>���ӵ� <font class=founderr>����</font> �Ǳ���Ҫ�ģ��Ҵ�С���ܳ���"&word_size&"KB��"
  end if
  if founderr="" then
    sql="insert into bbs_topic (forum_id,username,topic,icon,counter,tim,re_username,re_counter,re_tim,istop,islock,isgood) " & _
	"values ("&forumid&",'"&login_username&"','"&topic&"','"&trim(request.form("icon"))&"',0,'"&now_time&"','"&login_username&"',0,'"&now_time&"',0,0,0)"
    conn.execute(sql)
    
    new_id=first_id("bbs_topic")
    
    sql="insert into bbs_data (forum_id,reply_id,username,topic,icon,word,tim,ip,sys) " & _
	"values ("&forumid&","&new_id&",'"&login_username&"','"&topic&"','"&trim(request.form("icon"))&"','"&word&"','"&now_time&"','"&ip_sys(1,1)&"','"&ip_sys(3,0)&"')"
    set rs=conn.execute(sql)
    
    sql="update user_data set bbs_counter=bbs_counter+1,integral=integral+2 where username='"&login_username&"'"
    conn.execute(sql)
    sql="update configs set num_topic=num_topic+1,num_data=num_data+1 where id=1"
    conn.execute(sql)
    sql="update bbs_forum set forum_topic_num=forum_topic_num+1,forum_data_num=forum_data_num+1,forum_new_info='"&login_username&"|"&now_time&"|"&new_id&"|"&topic&"' where forum_id="&forumid
    conn.execute(sql)
    
    call upload_note(index_url,new_id)
    
    call time_load(0,0,1)
    
    response.write VbCrLf & "<table border=0 width=300>" & _
        	   VbCrLf & "<tr><td align=center height=30><font class=red>���ӷ���ɹ���лл���ķ�����</font></td></tr>" & _
        	   VbCrLf & "<tr><td height=30>�����ڿ���ѡ�����²�����</td></tr>" & _
        	   VbCrLf & "<tr><td>����1��<a href='forum_view.asp?forum_id=" & forumid & "&view_id=" & new_id & "'>�������������</a>" & _
        	   VbCrLf & "<tr><td>����2��<a href='forum_list.asp?forum_id=" & forumid & "'>���� <b>" & forumname & "</b></a></td></tr>" & _
        	   VbCrLf & "<tr><td>����3��<a href='forum.asp'>������̳��ҳ</a></td></tr>" & _
        	   VbCrLf & "<tr><td height=30>ϵͳ���� " & web_var(web_num,5) & " ���Ӻ��Զ����� <b>" & forumname & "</b> ��</td></tr>" & _
		   VbCrLf & "</table><meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=forum_list.asp?forum_id=" & forumid & "'>"
  else
    response.write found_error(founderr,"350")
  end if
end function
%>