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

dim rssnum,j,id,vid,vname,nid
tit="<a href='?'>查看现有调查列表</a> ┋ <a href='?action=add'>添加新调查列表</a>"
response.write header(8,tit)
id=trim(request.querystring("id"))
vid=trim(request.querystring("vid"))

select case action
case "add"
  call vote_add()
case "edit"
  call vote_edit()
case "edit2"
  call vote_edit2()
case "view"
  call vote_view()
case "del"
  call vote_del()
case "delete"
  call vote_delete()
case else
  call vote_main()
end select

call close_conn()
response.write ender()

sub vote_del()
  if not(isnumeric(id)) then call vote_main():exit sub
  conn.execute("delete from vote where vtype=1 and id="&id)
  response.write "<script language=javascript>alert(""成功删除了调查项目（"&id&"）！\n\n点击返回……"");location.href='?action=view&vid="&vid&"';</script>"
end sub

sub vote_delete()
  if not(isnumeric(vid)) then call vote_main():exit sub
  conn.execute("delete from vote where vid="&vid)
  response.write "<script language=javascript>alert(""成功删除了调查列表（"&vid&"）！\n\n点击返回……"");location.href='?';</script>"
end sub

sub vote_edit2()
  if not(isnumeric(id)) then call vote_main():exit sub
  sql="select vid,vname,counter from vote where vtype=1 and id="&id
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    response.write "<script language=javascript>alert(""调查项目不存在！\n\n点击返回……"");location.href='?';</script>"
    exit sub
  end if
  dim counter
  vid=rs("vid"):vname=rs("vname"):counter=rs("counter")
  rs.close:set rs=nothing
  if trim(request.querystring("chk"))="yes" then
    counter=code_admin(request.form("counter"))
    if not(isnumeric(counter)) then counter=-1
    vname=code_admin(request.form("vname"))
    if int(counter)<0 or instr(1,counter,".")>0 then
      response.write "<font class=red_2>投票计数只能为正整数且不能为空！</font><br><br>"&go_back:exit sub
    end if
    if len(vname)<1 then
      response.write "<font class=red_2>项目名称不能为空！</font><br><br>"&go_back:exit sub
    end if
    sql="update vote set vname='"&vname&"',counter="&counter&" where vtype=1 and id="&id
    conn.execute(sql)
    response.write "<script language=javascript>alert(""成功修改了一个调查项目名称！\n\n点击返回……"");location.href='?action=view&vid="&vid&"';</script>"
    exit sub
  end if
%>
<table border=0>
<form action='?action=edit2&id=<%response.write id%>&chk=yes' method=post>
<tr><td colspan=2 align=center height=50><a href='?action=view&vid=<%response.write vid%>' class=red>修改现有调查项目</a></td></tr>
<tr><td>项目名称：</td><td><input type=text name=vname value='<%response.write vname%>' size=30 maxlength=20></td></tr>
<tr><td height=30>投票计数：</td><td><input type=text name=counter value='<%response.write counter%>' size=10 maxlength=10><%response.write redx%>只能为0或正整数</td></tr>
<tr><td colspan=2 align=center><input type=submit value='提 交 修 改'>　　<input type=reset value='重新填写'></td></tr>
</form>
</table>
<%
end sub

sub vote_edit()
  if not(isnumeric(vid)) then call vote_main():exit sub
  sql="select id,vname from vote where vtype=0 and vid="&vid
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    response.write "<script language=javascript>alert(""调查列表（"&vid&"）不存在！\n\n点击返回……"");location.href='?';</script>"
    exit sub
  end if
  vname=rs("vname")
  rs.close:set rs=nothing
  if trim(request.querystring("chk"))="yes" then
    vname=code_admin(request.form("vname"))
    if len(vname)<1 then
      response.write "<font class=red_2>调查名称不能为空！</font><br><br>"&go_back:exit sub
    end if
    sql="update vote set vname='"&vname&"' where vtype=0 and vid="&vid
    conn.execute(sql)
    response.write "<script language=javascript>alert(""成功修改了调查列表（"&vid&"）的名称！\n\n点击返回……"");location.href='?action=view&vid="&vid&"';</script>"
    exit sub
  end if
%>
<table border=0>
<form action='?action=edit&vid=<%response.write vid%>&chk=yes' method=post>
<tr><td colspan=2 align=center height=50 class=red>修改调查列表名称</td></tr>
<tr><td>调查 ID：</td><td><input type=text name=vid value='<%response.write vid%>' size=10 maxlength=10 disabled><%response.write redx%>只能为正整数</td></tr>
<tr><td height=50>调查名称：</td><td><input type=text name=vname value='<%response.write vname%>' size=30 maxlength=20><%response.write redx%></td></tr>
<tr><td colspan=2 align=center><input type=submit value='提 交 修 改'>　　<input type=reset value='重新填写'></td></tr>
</form>
</table>
<%
end sub

sub vote_add()
  if trim(request.querystring("chk"))="yes" then
    vid=code_admin(request.form("vid"))
    if not(isnumeric(vid)) then vid=0
    vname=code_admin(request.form("vname"))
    if int(vid)<1 or instr(1,vid,".")>0 then
      response.write "<font class=red_2>调查列表 ID 只能为正整数且不能为空！</font><br><br>"&go_back:exit sub
    end if
    if len(vname)<1 then
      response.write "<font class=red_2>调查名称不能为空！</font><br><br>"&go_back:exit sub
    end if
    sql="select id from vote where vtype=0 and vid="&vid
    set rs=conn.execute(sql)
    if not(rs.eof and rs.bof) then
      rs.close:set rs=nothing
      response.write "<font class=red_2>调查列表 ID（"&vid&"）已存在！请重新输入。</font><br><br>"&go_back:exit sub
    end if
    rs.close:set rs=nothing
    sql="insert into vote(vid,vtype,vname,counter) values("&vid&",0,'"&vname&"',0)"
    conn.execute(sql)
    response.write "<script language=javascript>alert(""成功添加了一个新的调查列表！\n\n点击返回……"");location.href='?';</script>"
    exit sub
  end if
%>
<table border=0>
<form action='?action=add&chk=yes' method=post>
<tr><td colspan=2 align=center height=50 class=red>添加新的调查列表</td></tr>
<tr><td>调查 ID：</td><td><input type=text name=vid size=10 maxlength=10><%response.write redx%>只能为正整数</td></tr>
<tr><td height=50>调查名称：</td><td><input type=text name=vname size=30 maxlength=20><%response.write redx%></td></tr>
<tr><td colspan=2 align=center><input type=submit value='提 交 添 加'>　　<input type=reset value='重新填写'></td></tr>
</form>
</table>
<%
end sub

sub vote_view()
  if not(isnumeric(vid)) then call vote_main():exit sub
  if trim(request.querystring("chk"))="yes" then
    vname=code_admin(request.form("vname"))
    if len(vname)<1 then
      response.write "<font class=red_2>调查项目不能为空！</font><br><br>"&go_back:exit sub
    end if
    sql="insert into vote(vid,vtype,vname,counter) values("&vid&",1,'"&vname&"',0)"
    conn.execute(sql)
    response.write "<script language=javascript>alert(""成功添加了一条新调查项目！\n\n点击返回……"");location.href='?action=view&vid="&vid&"';</script>"
    exit sub
  end if
%>
<table border=1 width=400 cellspacing=0 cellpadding=2<%response.write table1%>>
<%
  sql="select id,vid,vname,counter from vote where vid="&vid&" order by id"
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    response.write "<script language=javascript>alert(""调查列表（"&vid&"）不存在！\n\n点击返回……"");location.href='?';</script>"
    exit sub
  end if
  j=0
  do while not rs.eof
    nid=rs("id")
    if j=0 then
%>
<tr>
<td colspan=2 height=25 bgcolor=<%response.write color3%> class=red_3>&nbsp;&nbsp;<b><%response.write code_html(rs("vname"),1,0)%></b>（ID：<%response.write vid%>）</td>
<td align=center><a href='?action=edit&vid=<%response.write vid%>'>编辑标题</a></td>
</td></tr>
<%  else %>
<tr align=center<%response.write mtr%>>
<td width='8%'><%response.write j%></td>
<td width='76%' align=left><%response.write rs("vname")%> <font class=blue><%response.write rs("counter")%></font></td>
<td width='16%'><a href='?action=edit2&id=<%response.write nid%>'>编辑</a> <a href="javascript:do_del(<%response.write vid%>,<%response.write nid%>);">删除</a></td>
</tr>
<%
    end if
    j=j+1
    rs.movenext
  loop
  rs.close:set rs=nothing
%>
<tr><td colspan=3 height=25 align=center>
  <table border=0>
  <form action='?action=view&vid=<%response.write vid%>&chk=yes' method=post>
  <tr>
  <td>新的项目名称：</td>
  <td><input type=text name=vname size=20 maxlength=20></td>
  <td>&nbsp;&nbsp;<input type=submit value='点击添加'></td>
  </tr>
  </form>
  </table>
</td></tr>
</table>
<%
end sub

sub vote_main()
%>
<table border=1 width=400 cellspacing=0 cellpadding=2<%response.write table1%>>
<tr align=center height=20 bgcolor=<%response.write color3%>>
<td width='8%'>ID</td>
<td width='76%'>调查列表名称</td>
<td width='16%'>操作</td>
</tr>
<%
  sql="select id,vid,vname from vote where vtype=0 order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    nid=rs("id"):vid=rs("vid")
%>
<tr align=center<%response.write mtr%>>
<td class=blue><b><%response.write vid%></b></td>
<td align=left><a href='?action=view&vid=<%response.write vid%>'><%response.write code_html(rs("vname"),1,0)%></a></td>
<td><a href='?action=edit&vid=<%response.write vid%>'>编辑</a> <a href="javascript:do_delete(<%response.write vid%>);">删除</a></td>
</tr>
<%
    rs.movenext
  loop
  rs.close:set rs=nothing
%>
</table>
<br>
<table border=0 width=450>
<tr><td colspan=2>调用方法：</td></tr>
<tr><td colspan=2 height=40>&lt;script language=javascript src='vote.asp?id=<font class=red>1</font>&types=<font class=red>1</font>&mcolor=<font class=red>ff0000</font>&bgcolor=<font class=red>ededed</font>'&gt;&lt;/script&gt;</td></tr>
<tr><td>使用说明：</td><td>1、第一个参数是要调用的调查ID；</td></tr>
<tr><td></td><td>2、第二个参数是调查显示的类型：“1”为单选，“2”为多选；</td></tr>
<tr><td></td><td>3、第三个参数是调查标题显示颜色；（不要加“#”）</td></tr>
<tr><td></td><td>4、第四个参数是调查选择框背景色；（不要加“#”）</td></tr>
</table>
<%
end sub
%>
<script language=JavaScript><!--
function do_del(data1,data2)
{
  if (confirm("此操作将删除ID为 "+data2+" 的调查项目！\n\n真的要删除吗？\n删除后将无法恢复！"))
    window.location="?action=del&vid="+data1+"&id="+data2
}
function do_delete(data1)
{
  if (confirm("此操作将删除ID为 "+data1+" 的调查列表！\n\n真的要删除吗？\n删除后将无法恢复！"))
    window.location="?action=delete&vid="+data1
}
//--></script>