<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nummer,rssum,action_temp
tit="վ�ڶ���"
nummer=0

call web_head(2,0,0,0,0)
'------------------------------------left----------------------------------
call left_user()
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong
call user_mail_menu(0)
response.write table1
%>
<tr align=center<%response.write table2%> height=25>
<td width='6%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>�Ѷ�</b></font></td>
<td width='20%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b><%
if action="outbox" or action="issend" then
  response.write "��"
else
  response.write "��"
end if
%>����</b></font></td>
<td width='38%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>��������</b></font></td>
<td width='20%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>��������</b></font></td>
<td width='10%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>��С</b></font></td>
<td width='6%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>����</b></font></td>
</tr>
<script language=javascript src='STYLE/admin_del.js'></script>
<form name=del_form action='user_mail.asp?action=<%response.write action%>' method=post>
<input type=hidden name=action2 value='delete'>
<input type=hidden name=del_type value='<%response.write action%>'>
<%
if trim(request.form("action2"))="delete" and len(trim(request.form("del_sel"))) then
  response.write del_select()
end if
function del_select()
  dim delid,del_i,del_num,del_dim,del_sql,del_type
  del_type=trim(request.form("del_type"))
  delid=trim(request.form("del_id"))
  select case del_type
  case "outbox","issend"
    del_sql="update user_mail set types=4 where send_u='"&login_username&"' and id="
  case "recycle"
    del_sql="delete from user_mail where (send_u='"&login_username&"' or accept_u='"&login_username&"') and id="
  case else
    del_sql="update user_mail set types=4 where accept_u='"&login_username&"' and id="
  end select

  if var_null(delid)<>"" then
    delid=replace(delid," ","")
    del_dim=split(delid,",")
    del_num=UBound(del_dim)
    for del_i=0 to del_num
      conn.execute(del_sql&del_dim(del_i))
    next
    Erase del_dim
    if del_type="recycle" then
      del_select="����ɾ���ɹ�����ɾ���� "&del_num+1&" �����š�\n\n�����ѳ���ɾ����"
    else
      del_select="����ɾ���ɹ�����ɾ���� "&del_num+1&" �����š�\n\nɾ���Ķ��Ž��������Ļ���վ�ڡ�"
    end if
    del_select=vbcrlf&"<script language=javascript>alert("""&del_select&""");</script>"
  end if
end function

if len(trim(request.form("clear")))>0 then
  response.write mail_clear()
end if

sql="select * from user_mail where "
select case action
case "outbox"
  sql=sql&"send_u='"&login_username&"' and types=2"
  action_temp="�ݸ���"
case "issend"
  sql=sql&"send_u='"&login_username&"' and types=1"
  action_temp="�ѷ�����"
case "recycle"
  sql=sql&"(accept_u='"&login_username&"' or send_u='"&login_username&"') and types=4"
  action_temp="������"
case else
  action="inbox"
  sql=sql&"accept_u='"&login_username&"' and types=1"
  action_temp="������"
end select

sql=sql&" order by id desc"
login_message=0
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
  rssum=rs.recordcount
  nummer=rssum
  for i=1 to rssum
    response.write mail_type(rs)
    rs.movenext
  next
end if
%>
<tr><td colspan=6 bgcolor=<%response.write web_var(web_color,5)%> height=30 align=center class=htd>��<font class=red><%response.write nummer%></font>������<font class=gray>��Ϊ�˽�ʡ�ռ䣬�뼰ʱɾ��������Ϣ��</font>
<input type=checkbox name=del_all value=1 onClick=selectall('<%response.write nummer%>') class=bg_3> ѡ������
<input type=submit name=del_sel value='ɾ����ѡ' onclick="return suredel('<%response.write nummer%>');">
<input type=submit name=clear onclick="{if(confirm('ȷ�����<%response.write action_temp%>���еļ�¼��?\n\n��պ��޷��ָ���')){this.document.del_form.submit();return true;}return false;}" value="���<%response.write action_temp%>" style='width:90px'></td></tr>
</table>
<%
response.write ukong
'---------------------------------center end-------------------------------
call web_end(0)

function mail_clear()
  dim clear_type
  select case trim(request.form("del_type"))
  case "inbox"
    conn.execute("delete from user_mail where accept_u='"&login_username&"' and types=1")
    clear_type="������"
  case "outbox"
    conn.execute("delete from user_mail where send_u='"&login_username&"' and types=2")
    clear_type="�ݸ���"
  case "issend"
    conn.execute("delete from user_mail where send_u='"&login_username&"' and types=1")
    clear_type="�ѷ�����"
  case "recycle"
    conn.execute("delete from user_mail where (accept_u='"&login_username&"' or send_u='"&login_username&"') and types=4")
    clear_type="������"
  end select
end function

function mail_type(rs)
  dim ttim,isread,td_temp,read_pic,iid,link_temp,name_temp
  td_temp=""
  read_pic="olds"
  link_temp="view"
  iid=rs("id"):isread=rs("isread"):ttim=rs("tim")
  if isread=false then
    td_temp=" class=btd"
    read_pic="news"
    if action="inbox" then
      login_message=login_message+1
    end if
  end if
  if action="outbox" then
    td_temp=" class=btd"
    read_pic="sends"
    link_temp="edit"
  end if
  if action="outbox" or action="issend" then
    name_temp=format_user_view(rs("accept_u"),1,1)
  else
    name_temp=format_user_view(rs("send_u"),1,1)
  end if
  
  ttim=time_type(ttim,8)
  mail_type=vbcrlf&"<tr align=center"&td_temp&table3&"><td><img src='images/mail/"&read_pic&".gif' border=0></td><td>"&name_temp&"</td><td align=left><a href='user_message.asp?action="&link_temp&"&id="&iid&"'>"&cuted(rs("topic"),15)&"</a></td><td class=timtd>"&ttim&"</td><td>"&len(rs("word"))&"Byte</td><td><input type=checkbox name=del_id value='"&iid&"' class=bg_1></td></tr>"
end function
%>