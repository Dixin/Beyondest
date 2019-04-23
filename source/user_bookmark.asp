<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim n_sort,gid,g_id,gname,select_add,id,name,url,rssum
tit="网络书签":n_sort="book":rssum=0
g_id=trim(request.querystring("g_id"))
if not(isnumeric(g_id)) then g_id=0

call web_head(2,0,0,0,0)
'------------------------------------left----------------------------------
call left_user()
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
select case action
case "groupedit"
  call group_edit()
case "bookmarkedit"
  call bookmark_edit()
case "groupdel"
  call group_del()
case "bookmarkdel"
  call bookmark_del()
case "groupadd"
  call group_add()
case "bookmarkadd"
  call bookmark_add()
end select
%>
<%response.write ukong&table1%>
<tr<%response.write table2%> height=25>
<td class=end width='90%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small(us)%>&nbsp;<b>我的书签组</b></td>
<td class=end width='10%' align=center background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>组操作</b></td>
</tr>
<form name=groupedit_frm action='?action=groupedit' method=post>
<input type=hidden name=g_id value=''>
<input type=hidden name=g_name value=''>
</form>
<tr<%response.write table3%>>
<td><%response.write img_small("jt0")%><a href='?g_id=0' class=gray>[ 根书签组 ]</a>&nbsp;&nbsp;
<%response.write img_small("jt0")%><a href='?action=all' class=gray>[ 浏览所有书签 ]</a></td>
<td align=center class=gray>无</td>
</tr>
<%
sql="select g_id,g_name from jk_group where g_sort='"&n_sort&"' and username='"&login_username&"' order by g_id"
set rs=conn.execute(sql)
do while not rs.eof
  gid=rs("g_id"):gname=rs("g_name")
  select_add=select_add&vbcrlf&"<option value='"&gid&"'"
  if int(gid)=int(g_id) then select_add=select_add&" selected"
  select_add=select_add&">"&gname&"</option>"
%>
<tr<%response.write table3%>>
<td><%response.write img_small("jt0")%><a href='?g_id=<%response.write gid%>'<%if int(g_id)=int(gid) then response.write " class=red_3"%>><%response.write gname%></a></td>
<td align=center><a href="javascript:group_edit(<%response.write gid%>,'<%response.write gname%>');"><img src='IMAGES/SMALL/EDIT.GIF' border=0 title='修改'></a>&nbsp;<a href="javascript:group_del(<%response.write gid%>);"><img src='IMAGES/SMALL/DEL.GIF' border=0 title='删除'></a></td>
</tr>
<%
  rs.movenext
loop
rs.close
%>
</table>
<%response.write kong&table1%>
<tr<%response.write table2%> align=center height=25>
<td class=end width='6%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>序号</b></td>
<td class=end width='34%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>我的个人书签名称</b></td>
<td class=end width='50%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>书签地址</b></td>
<td class=end width='10%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><b>操 作</b></td>
</tr>
<form name=bookmarkedit_frm action='?action=bookmarkedit' method=post>
<input type=hidden name=id value=''>
<input type=hidden name=name value=''>
<input type=hidden name=url value=''>
</form>
<%
sql="select id,name,url from user_bookmark where"
if action<>"all" then
  sql=sql&" g_id="&g_id&" and"
end if
sql=sql&" username='"&login_username&"' order by id desc"
set rs=conn.execute(sql)
do while not rs.eof
  id=rs("id"):name=rs("name"):url=rs("url")
  rssum=rssum+1
%>
<tr<%response.write table3%>>
<td align=center><%response.write rssum%></td>
<td><a href='<%response.write url%>' target=_blank title='<%response.write code_html(name,1,0)%>'><%response.write code_html(name,1,15)%></a></td>
<td><a href='<%response.write url%>' target=_blank title='<%response.write code_html(url,1,0)%>'><%response.write code_html(url,1,25)%></a></td>
<td align=center><a href="javascript:bookmark_edit(<%response.write id%>,'<%response.write name%>','<%response.write url%>');"><img src='IMAGES/SMALL/EDIT.GIF' border=0></a>&nbsp;<a href="javascript:bookmark_del(<%response.write id%>);"><img src='IMAGES/SMALL/DEL.GIF' border=0></a></td>
</tr>
<%
  rs.movenext
loop
rs.close
%>
</table>
<%response.write kong&table1%>
<tr<%response.write table2%> height=25>
<td class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small(us)%>&nbsp;<b>添加新的书签组</b></td>
</tr>
<tr<%response.write table3%>><td>
  <table border=0 cellpadding=5>
  <form action='?action=groupadd' method=post>
  <tr>
  <td>　组名称：</td>
  <td><input type=text name=g_name size=20 maxlength=20></td>
  <td><input type=submit value='添加书签组'></td>
  </tr>
  </form>
  </table>
</td></tr>
</table>
<%response.write kong&table1%>
<tr<%response.write table2%> height=25>
<td class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small(us)%>&nbsp;<b>添加新的个人书签</b></td>
</tr>
<tr<%response.write table3%>><td>
  <table border=0 cellpadding=2>
  <form action='?action=bookmarkadd' method=post>
  <tr>
  <td>　书签名称：</td>
  <td>
    <table border=0>
    <tr>
    <td><input type=text name=name size=30 maxlength=50></td>
    <td>　书签组：</td>
    <td><select name=g_id>
    <option value='0'>[ 根书签组 ]</option>
<%response.write select_add%>
</select></td>
    </tr>
    </table>
  </td></tr>
  <tr>
  <td>　书签地址：</td>
  <td>
    <table border=0>
    <tr>
    <td><input type=text name=url size=50 value='http://' maxlength=100></td>
    <td>　<input type=submit value='添加书签'></td>
    </tr>
    </table>
  </td></tr>
  </form>
  </table>
</td></tr>
</table>
<br>
<%
'---------------------------------center end-------------------------------
call web_end(0)

sub group_del()
  gid=trim(request.querystring("g_id"))
  if not(isnumeric(gid)) then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""您在删除书签组时输入的书签组ID出错！\n\n请返回重新输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
    close_conn
    exit sub
  end if
  sql="delete from jk_group where g_id="&gid&" and g_sort='"&n_sort&"' and username='"&login_username&"'"
  conn.execute(sql)
  sql="delete from user_bookmark where g_id="&gid&" and username='"&login_username&"'"
  conn.execute(sql)
  response.write("<script language=javascript>" & _
		 vbcrlf & "alert(""成功的删除了一书签组！"");" & _
		 vbcrlf & "</script>")
end sub

sub group_edit()
  gname=code_form(request.form("g_name"))
  gid=trim(request.form("g_id"))
  if len(gname)<1 or len(gname)>20 or not(isnumeric(gid)) then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""您在修改书签组的 组名称 时输入的数据有误！\n\n请返回重新输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
    close_conn
    exit sub
  end if
  sql="update jk_group set g_name='"&gname&"' where g_id="&gid&" and g_sort='"&n_sort&"' and username='"&login_username&"'"
  conn.execute(sql)
  g_id=gid
  response.write("<script language=javascript>" & _
		 vbcrlf & "alert(""成功的修改了书签组的名称："&gname&""");" & _
		 vbcrlf & "</script>")
end sub

sub bookmark_del()
  id=trim(request.querystring("id"))
  if not(isnumeric(id)) then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""您在删除个人书签时输入的书签ID出错！\n\n请返回重新输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
    close_conn
    exit sub
  end if
  sql="delete from user_bookmark where id="&id&" and username='"&login_username&"'"
  conn.execute(sql)
  response.write("<script language=javascript>" & _
		 vbcrlf & "alert(""成功的删除了一书签组！"");" & _
		 vbcrlf & "</script>")
end sub

sub bookmark_edit()
  name=code_form(request.form("name"))
  url=code_form(request.form("url"))
  id=trim(request.form("id"))
  if len(name)<1 or len(name)>50 or len(url)<1 or len(url)>100 or not(isnumeric(id)) then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""您在修改个人书签时输入的数据有误！\n\n请返回重新输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
    close_conn
    exit sub
  end if
  sql="update user_bookmark set name='"&name&"',url='"&url&"' where id="&id&" and username='"&login_username&"'"
  conn.execute(sql)
  response.write("<script language=javascript>" & _
		 vbcrlf & "alert(""成功的修改了个人书签（名称："&name&"）！"");" & _
		 vbcrlf & "</script>")
end sub

sub group_add()
  gname=code_form(request.form("g_name"))
  if len(gname)<1 or len(gname)>20 then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""添加书签组的 组名称 是必须要的！\n\n请返回新输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
    close_conn
    exit sub
  end if
  sql="insert into jk_group(g_sort,g_name,username) values('"&n_sort&"','"&gname&"','"&login_username&"')"
  conn.execute(sql)
  response.write("<script language=javascript>" & _
		 vbcrlf & "alert(""成功的添加了一书签组："&gname&""");" & _
		 vbcrlf & "</script>")
end sub

sub bookmark_add()
  dim gg
  gg=trim(request.form("g_id"))
  if not(isnumeric(gg)) then gg=0
  name=code_form(request.form("name"))
  url=code_form(request.form("url"))
  if len(name)<1 or len(name)>50 or len(url)<8 or len(url)>100 then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""添加新书签的 书签名称 和 书签地址 是必须要的！\n\n请返回新输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
    close_conn
    exit sub
  end if
  sql="insert into user_bookmark(g_id,username,name,url) values("&gg&",'"&login_username&"','"&name&"','"&url&"')"
  conn.execute(sql)
  response.write("<script language=javascript>" & _
		 vbcrlf & "alert(""成功的添加了一个我的个人书签："&name&""");" & _
		 vbcrlf & "</script>")
end sub
%>
<script language=javascript>
<!--
function group_edit(geid,gename)
{
  var gevar='请输入要修改的书签组（ID：'+geid+'）的新名称，长度不能超过20位';
  this.document.groupedit_frm.g_id.value=geid;
  var gename=prompt(gevar+'：',gename);
  if (gename == null || gename == '' || gename.length>20)
  { alert(gevar+"！");return; }
  else
  { this.document.groupedit_frm.g_name.value=gename; }
  this.document.groupedit_frm.submit();
}

function group_del(gdid)
{
  if (confirm("此操作将删除ID为 "+gdid+" 的书签组！\n真的要删除吗？\n删除后将无法恢复！"))
  { window.location="?action=groupdel&g_id="+gdid; }
}

function bookmark_edit(bid,bname,burl)
{
  var var1='请输入要修改的个人书签（ID：'+bid+'）的名称，长度不能超过50位';
  var var2='请输入要修改的个人书签（ID：'+bid+'）的地址，长度不能超过100位';
  this.document.bookmarkedit_frm.id.value=bid;
  var bename=prompt(var1+'：',bname);
  if (bename == null || bename == '' || bename.length>50)
  { alert(var1+"！");return; }
  else
  {
    this.document.bookmarkedit_frm.name.value=bename;
    var beurl=prompt(var2+'：',burl);
    if (beurl == null || beurl == '' || beurl.length>100)
    { alert(var2+"！");return; }
    else
    {this.document.bookmarkedit_frm.url.value=beurl;}
  }
  this.document.bookmarkedit_frm.submit();
}

function bookmark_del(bdid)
{
  if (confirm("此操作将删除ID为 "+bdid+" 的个人书签！\n真的要删除吗？\n删除后将无法恢复！"))
  { window.location="?action=bookmarkdel&id="+bdid; }
}
-->
</script>