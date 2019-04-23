<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit="<a href='admin_update.asp?'>��վ����</a> �� " & _
    "<a href='admin_data.asp'>���ݸ���</a> �� " & _
    "<a href='admin_update.asp?nsort=news'>���¹���</a> �� " & _
    "<a href='admin_update.asp?nsort=forum'>��̳����</a> �� " & _
    "<a href='admin_update.asp?action=add'>��������</a>"
response.write header(7,tit)

dim actions:actions=trim(request.querystring("actions"))

select case action
case "update_config"
  call update_config()
case "update_forum"
  call update_forum()
case "clear_notes"
  call clear_notes()
case "clear_message"
  call clear_message()
end select

sub update_config()
  dim rs,sql,num_topic,num_data,num_reg,new_username,num_news,num_article,num_down,num_flash,num_film,num_photo,num_desktop
  num_reg=0:num_topic=0:num_data=0:num_news=0:num_article=0:num_down=0:num_flash=0:num_film=0:num_photo=0:num_desktop=0
  set rs=server.createobject("adodb.recordset")
  sql="select username from user_data order by id desc"
  rs.open sql,conn,1,1
  if not(rs.eof and rs.bof) then
    num_reg=int(rs.recordcount)
    new_username=rs("username")
  end if
  rs.close
  
  sql="select count(id) from bbs_topic"
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then num_topic=int(rs(0))
  rs.close
  
  sql="select count(id) from bbs_data"
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then num_data=int(rs(0))
  rs.close
  
  sql="select count(id) from news where hidden=1"
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then num_news=int(rs(0))
  rs.close
  
  sql="select count(id) from article where hidden=1"
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then num_article=int(rs(0))
  rs.close
  
  sql="select count(id) from down where hidden=1"
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then num_down=int(rs(0))
  rs.close
  
  sql="select count(id) from gallery where hidden=1 and types='flash'"
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then num_flash=int(rs(0))
  rs.close
  
  sql="select count(id) from gallery where hidden=1 and types='film'"
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then num_film=int(rs(0))
  rs.close
  
  sql="select count(id) from gallery where hidden=1 and types='baner'"
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then num_photo=int(rs(0))
  rs.close
  
  sql="select count(id) from gallery where hidden=1 and types='paste'"
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then num_desktop=int(rs(0))
  rs.close
  
  sql="update configs set num_topic="&num_topic&",num_data="&num_data&",num_reg="&num_reg&",new_username='"&new_username&"',num_news="&num_news&",num_article="&num_article&",num_down="&num_down&",num_flash="&num_flash&",num_film="&num_film&",num_photo="&num_photo&",num_desktop="&num_desktop&" where id=1"
  conn.execute(sql)
  
  response.write "<script language=javascript>alert(""�ɹ���������վͳ�����ݣ�"");</script>"
end sub

sub update_forum()
  dim rsf,sqlf,rssum,i,rs,sql,forumid,t1,t2,t3
  sqlf="select * from bbs_forum order by forum_id"
  set rsf=conn.execute(sqlf)
  do while not rsf.eof
    forumid=rsf("forum_id")
    set rs=server.createobject("adodb.recordset")
    sql="select * from bbs_topic where forum_id="&forumid&" order by id desc"
    rs.open sql,conn,1,1
    if rs.eof and rs.bof then
      t1=0
      t2="|||"
    else
      t1=rs.recordcount
      t2=rs("username") &"|"& rs("tim") &"|"& rs("id") &"|"& rs("topic")
      t2=replace(t2,"'","")
    end if
    rs.close:set rs=nothing
    
    sql="select count(*) from bbs_data where forum_id="&forumid
    set rs=conn.execute(sql)
    t3=rs(0)
    rs.close:set rs=nothing
    if int(t3)<1 then t3=0

    sql="update bbs_forum set forum_topic_num="&t1&",forum_new_info='"&t2&"',forum_data_num="&t3&" where forum_id="&forumid
    conn.execute(sql)
    rsf.movenext
  loop
  rsf.close:set rsf=nothing
  
  response.write "<script language=javascript>alert(""�ɹ������˷���̳���ݣ�"");</script>"
end sub

sub clear_notes()
  dim clear_msg
  select case actions
  case "week"
    sql="delete from notes where DateDiff('d',tim,'"&now_time&"')>7"
    clear_msg="һ��ǰ"
  case "all"
    sql="delete from notes"
    clear_msg="����"
  case else
    sql="delete from notes where DateDiff('d',tim,'"&now_time&"')>30"
    clear_msg="һ����ǰ"
  end select
  conn.execute(sql)
  response.write "<script language=javascript>alert(""�ɹ�������"&clear_msg&"���������ݣ�"");</script>"
end sub

sub clear_message()
  dim clear_msg
  select case actions
  case "week"
    sql="delete from user_mail where DateDiff('d',tim,'"&now_time&"')>7"
    clear_msg="һ��ǰ"
  case "all"
    sql="delete from user_mail"
    clear_msg="����"
  case else
    sql="delete from user_mail where DateDiff('d',tim,'"&now_time&"')>30"
    clear_msg="һ����ǰ"
  end select
  conn.execute(sql)
  response.write "<script language=javascript>alert(""�ɹ�������"&clear_msg&"���û�����Ϣ��"");</script>"
end sub
%>
<table border=1 cellspacing=0 cellpadding=2 width=500 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<tr height=50 align=center>
<td width='20%'><font class=red_2>ע������</font></td>
<td width='80%'>�����еĲ������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�У�</td>
</tr>
<tr align=center height=80>
<td><font class=red_3>������վ������</font></td>
<td class=htd>�������İ�ť�����¼�����������վ�Ļ�����Ϣ���������š����¡��������̳�ȣ�����ÿ��һ��ʱ������һ�Ρ�<br>
<input type=button value='����������վͳ������' onclick=update_config() class=red></td>
</tr>
<tr align=center height=80>
<td><font class=red_3>���·���̳����</font></td>
<td class=htd>�������İ�ť�����¼���ÿ����̳���������⡢�ظ��������������⡢�ظ���ʱ�����Ϣ������ÿ��һ��ʱ������һ�Ρ�<br>
<input type=button value='�������·���̳����' onclick=update_forum() class=red></td>
</tr>
<tr align=center height=80>
<td><font class=red_3>������������</font></td>
<td class=htd>�������İ�ť�����������û����ѻ��ҵ�������Ϣ������ÿ��һ��ʱ������һ�Ρ�<br>
<input type=button value='���һ����ǰ������' onclick="javascript:clear_notes('month');" class=red style="width:140px;">&nbsp;&nbsp;
<input type=button value='���һ��ǰ������' onclick="javascript:clear_notes('week');" class=red style="width:120px;">&nbsp;&nbsp;
<input type=button value='ȫ�����' onclick="javascript:clear_notes('all');" class=red>
</td>
</tr>
<tr align=center height=80>
<td><font class=red_3>�����û�����Ϣ</font></td>
<td class=htd>�������İ�ť�����������û�����Ϣ��������Ϣ������ÿ��һ��ʱ������һ�Ρ�<br>
<input type=button value='���һ����ǰ�Ķ���' onclick="javascript:clear_message('month');" class=red style="width:140px;">&nbsp;&nbsp;
<input type=button value='���һ��ǰ�Ķ���' onclick="javascript:clear_message('week');" class=red style="width:120px;">&nbsp;&nbsp;
<input type=button value='ȫ�����' onclick="javascript:clear_message('all');" class=red>
</td>
</tr>
</table>
<script language=JavaScript>
<!--
function update_config()
{
  if (confirm("�˲����� ���·���̳���ݣ�\n\n���Ҫ������\n\n���º��޷��ָ���"))
    window.location="?action=update_config"
}

function update_forum()
{
  if (confirm("�˲����� ������վͳ�����ݣ�\n\n���Ҫ������\n\n���º��޷��ָ���"))
    window.location="?action=update_forum"
}

function clear_notes(cv)
{
  if (confirm("�˲����� �����������ݣ�\n\n���Ҫ������\n\n���º��޷��ָ���"))
    window.location="?action=clear_notes&actions="+cv
}

function clear_message(cv)
{
  if (confirm("�˲����� �����û�����Ϣ��\n\n���Ҫ������\n\n���º��޷��ָ���"))
    window.location="?action=clear_message&actions="+cv
}
//-->
</script>
<%
call close_conn()
response.write ender()
%>