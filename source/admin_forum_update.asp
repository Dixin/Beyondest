<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim admin_menu
admin_menu="<a href='admin_forum.asp'>论坛管理</a> ┋ " & _
	   "<a href='admin_forum_update.asp'>更新论坛数据</a> ┋ " & _
	   "<a href='admin_forum.asp?action=mod'>合并论坛</a> ┋ " & _
	   "<a href='admin_forum.asp?action=order'>重新排序</a>"
response.write header(11,admin_menu)

select case action
case "update_config"
  call update_config()
case "update_forum"
  call update_forum()
end select

sub update_config()
  dim rs,sql,num_topic,num_data,num_reg,new_username,num_news,num_article,num_down
  num_reg=0:num_topic=0:num_data=0:num_news=0:num_article=0:num_down=0
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
  
  sql="update configs set num_topic="&num_topic&",num_data="&num_data&",num_reg="&num_reg&",new_username='"&new_username&"',num_news="&num_news&",num_article="&num_article&",num_down="&num_down&" where id=1"
  conn.execute(sql)
  
  response.write "<script language=javascript>alert(""成功更新了网站统计数据！"");</script>"
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
  
  response.write "<script language=javascript>alert(""成功更新了分论坛数据！"");</script>"
end sub
%>
<table border=1 cellspacing=0 cellpadding=2 width=500 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<tr height=50 align=center>
<td width='20%'><font class=red_2>注意事项</font></td>
<td width='80%'>下面有的操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行！</td>
</tr>
<tr align=center height=80>
<td><font class=red_3>更新论坛总数据</font></td>
<td class=htd>点击下面的按钮将重新计算整个论坛的帖子主题、回复数量、最新加入用户等信息，建议每隔一段时间运行一次。<br>
<input type=button value='立即更新网站统计数据' onclick=update_config() class=red></td>
</tr>
<tr align=center height=80>
<td><font class=red_3>更新分论坛数据</font></td>
<td class=htd>点击下面的按钮将重新计算每个论坛的帖子主题、回复数量，最新主题、回复、时间等信息，建议每隔一段时间运行一次。<br>
<input type=button value='立即更新分论坛数据' onclick=update_forum() class=red></td>
</tr>
<tr align=center>
<td></td>
<td></td>
</tr>
</table>
<script language=JavaScript>
<!--
function update_config()
{
if (confirm("此操作将 更新分论坛数据！\n\n真的要进行吗？\n\n更新后将无法恢复！"))
  window.location="admin_forum_update.asp?action=update_config"
}

function update_forum()
{
if (confirm("此操作将 更新网站统计数据！\n\n真的要进行吗？\n\n更新后将无法恢复！"))
  window.location="admin_forum_update.asp?action=update_forum"
}
//-->
</script>
<%
close_conn
response.write ender()
%>