<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim nummer,rssum,action_temp
tit="站内短信"
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
<td width='6%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>已读</b></font></td>
<td width='20%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b><%
if action="outbox" or action="issend" then
  response.write "收"
else
  response.write "发"
end if
%>信人</b></font></td>
<td width='38%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>短信主题</b></font></td>
<td width='20%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>发送日期</b></font></td>
<td width='10%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>大小</b></font></td>
<td width='6%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end><b>操作</b></font></td>
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
      del_select="短信删除成功！共删除了 "&del_num+1&" 条短信。\n\n短信已彻底删除！"
    else
      del_select="短信删除成功！共删除了 "&del_num+1&" 条短信。\n\n删除的短信将置于您的回收站内。"
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
  action_temp="草稿箱"
case "issend"
  sql=sql&"send_u='"&login_username&"' and types=1"
  action_temp="已发短信"
case "recycle"
  sql=sql&"(accept_u='"&login_username&"' or send_u='"&login_username&"') and types=4"
  action_temp="废信箱"
case else
  action="inbox"
  sql=sql&"accept_u='"&login_username&"' and types=1"
  action_temp="收信箱"
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
<tr><td colspan=6 bgcolor=<%response.write web_var(web_color,5)%> height=30 align=center class=htd>共<font class=red><%response.write nummer%></font>条短信<font class=gray>（为了节省空间，请及时删除无用信息）</font>
<input type=checkbox name=del_all value=1 onClick=selectall('<%response.write nummer%>') class=bg_3> 选中所有
<input type=submit name=del_sel value='删除所选' onclick="return suredel('<%response.write nummer%>');">
<input type=submit name=clear onclick="{if(confirm('确定清空<%response.write action_temp%>所有的纪录吗?\n\n清空后将无法恢复！')){this.document.del_form.submit();return true;}return false;}" value="清空<%response.write action_temp%>" style='width:90px'></td></tr>
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
    clear_type="收信箱"
  case "outbox"
    conn.execute("delete from user_mail where send_u='"&login_username&"' and types=2")
    clear_type="草稿箱"
  case "issend"
    conn.execute("delete from user_mail where send_u='"&login_username&"' and types=1")
    clear_type="已发短信"
  case "recycle"
    conn.execute("delete from user_mail where (accept_u='"&login_username&"' or send_u='"&login_username&"') and types=4")
    clear_type="废信箱"
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