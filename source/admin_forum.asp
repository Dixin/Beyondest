<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim admin_menu
admin_menu="<a href='admin_forum.asp'>��̳����</a> �� " & _
	   "<a href='admin_forum.asp?action=mod'>�ϲ���̳</a> �� " & _
	   "<a href='admin_forum.asp?action=order'>��������</a>"
response.write header(11,admin_menu)

select case action
case "mod"
  call forum_mod()
case "order"
  call forum_order()
case "forum_add"
  call forum_add()
case "forum_edit"
  call forum_edit()
case "del_forum"
  call del_forum()
case "class_add"
  call class_add()
case "class_edit"
  call class_edit()
case "del_class"
  call del_class()
case else
  call forum_main()
end select

close_conn
response.write ender()

sub forum_order()
  dim rs,sql,rsf,sqlf,i,j,cid,fid
  i=1
  sql="select class_id from bbs_class order by class_order,class_id"
  set rs=conn.execute(sql)
  do while not rs.eof
    j=1:cid=rs(0)
    conn.execute("update bbs_class set class_order="&i&" where class_id="&cid)
    sqlf="select forum_id from bbs_forum where class_id="&cid&" order by forum_order,forum_id"
    set rsf=conn.execute(sqlf)
    do while not rsf.eof
      fid=rsf(0)
      conn.execute("update bbs_forum set forum_order="&j&" where forum_id="&fid)
      rsf.movenext
      j=j+1
    loop
    rsf.close:set rsf=nothing
    rs.movenext
    i=i+1
  loop
  rs.close:set rs=nothing
  call forum_main()
end sub

sub class_edit()
  dim classid,rs,strsql,class_name,class_order
  classid=trim(request("class_id"))
  if not(isnumeric(classid)) then
    response.redirect "admin_forum.asp"
    response.end
  end if
  set rs=server.createobject("adodb.recordset")
  strsql="Select * from bbs_class where class_id="&classid
  rs.open strsql,conn,1,3
%><font class=red>�޸���̳����</font><br><br><br>
<table border=0 width=300><%
if trim(request("edit"))="ok" then
  class_name=code_form(request.form("class_name"))
  if class_name="" then
    response.write( VbCrLf & "<tr><td height=80 align=center><font class=red_2>��̳�������Ʋ���Ϊ�գ�</font><br><br>"&go_back&"</td></tr>")
  else
    rs("class_name")=class_name
    rs.update
    response.write( VbCrLf & "<tr><td height=80 align=center>�ɹ����޸�����̳���ࣺ<font class=red>" & class_name & "</font></td></tr>")
  end if
else
%>
<tr>
<form method=post action='admin_forum.asp?action=class_edit&class_id=<%=classid%>&edit=ok'>
<td width='40%' align=center></td><td width='60%'></td>
</tr>
<tr height=30>
<td align=center>��̳�������ƣ�</td> 
<td><input type=text name=class_name value='<%=rs("class_name")%>' size=20 maxlength=20></td> 
</tr>
<tr height=30> 
<td colspan=2 align=center height=30><input type=submit value=' �� �� �� �� '></td>
</form>
</tr><%
  end if
  rs.close:set rs=nothing
%></table><%
end sub

sub class_add()
%><font class=red>�����̳����</font><br><br><br>
<table border=0 width=300>
<%
if trim(request.querystring("add"))="ok" then
  dim rs,strsql,class_name,class_order
  class_name=code_form(request.form("class_name"))
  if class_name="" then
    response.write( VbCrLf & "<tr><td height=80 align=center><font class=red_2>��̳�������Ʋ���Ϊ�գ�</font><br><br>"&go_back&"</td></tr>")
  else
    set rs=server.createobject("adodb.recordset")
    strsql="Select top 1 * from bbs_class order by class_order desc"
    rs.open strsql,conn,1,1
    if rs.eof and rs.bof then
      class_order=0
    else
      class_order=rs("class_order")
    end if
    class_order=class_order+1
    rs.close
    strsql="Select * from bbs_class"
    rs.open strsql,conn,1,3
    rs.addnew
    rs("class_order")=class_order
    rs("class_name")=class_name
    rs.update
    response.write( VbCrLf & "<tr><td height=80 align=center>�ɹ����������̳���ࣺ<font class=red>" & class_name & "</font></td></tr>")
    rs.close:set rs=nothing
  end if
else
%>
<tr>
<form method=post action='admin_forum.asp?action=class_add&add=ok'>
<td width='40%' align=center></td><td width='60%'></td>
</tr>
<tr height=30>
<td align=center>��̳�������ƣ�</td> 
<td><input type=text name=class_name size=20 maxlength=20></td> 
</tr>
<tr height=30> 
<td colspan=2 align=center height=30><input type=submit value=' �� �� �� �� '></td>
</form>
</tr><%
end if
%></table><%
end sub

sub forum_edit()
  dim classid,forumid,rs,strsql,classname,forum_name
  classid=trim(request("class_id"))
  forumid=trim(request("forum_id"))
  if not(isnumeric(classid)) or not(isnumeric(forumid)) then
    call forum_main():exit sub
  end if
  strsql="select class_name from bbs_class where class_id="&classid
  set rs=conn.execute(strsql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call forum_main():exit sub
  end if
  classname=rs("class_name")
  rs.close:set rs=nothing
%><font class=red>�޸���̳</font>��<font class=blue_1><%=classname%></font>��<br><br><br>
<table border=0 width=400><%
  set rs=server.createobject("adodb.recordset")
  strsql="Select * from bbs_forum where forum_id="&forumid
  rs.open strsql,conn,1,3
  if trim(request.querystring("edit"))="ok" then
    forum_name=code_form(request.form("forum_name"))
    if forum_name="" then
      response.write( VbCrLf & "<tr><td height=80 align=center><font class=red_2>��̳���Ʋ���Ϊ�գ�</font><br><br>"&go_back&"</td></tr>")
    else
      rs("class_id")=classid
      rs("forum_name")=forum_name
      rs("forum_pic")=trim(request.form("forum_pic"))
      if request.form("forum_hidden")="no" then
        rs("forum_hidden")=false
      else
        rs("forum_hidden")=true
      end if
      rs("forum_type")=request.form("forum_type")
      rs("forum_remark")=request.form("forum_remark")
      rs("forum_power")=code_form(request.form("forum_power"))
      rs.update
      response.write( VbCrLf & "<tr><td height=80 align=center>�ɹ����޸�����̳��<font class=red>" & forum_name & "</font></td></tr>")
    end if
  else
%><form method=post action='admin_forum.asp?action=forum_edit&forum_id=<%=forumid%>&edit=ok'>
<tr><td width='20%' align=center></td><td width='80%'></td></tr>
<tr height=30>
<td align=center>��̳���ƣ�</td> 
<td><input type=text name=forum_name value='<%=rs("forum_name")%>' size=30 maxlength=20></td> 
</tr>
<tr height=30>
<td align=center>�������ࣺ</td> 
<td><select name=class_id size=1>
<%
dim crs,csql,cid,ctype
csql="select * from bbs_class order by class_order"
set crs=conn.execute(csql)
do while not crs.eof
  cid=crs("class_id")
  response.write vbcrlf & "<option value='"&cid&"'"
  if int(classid)=int(cid) then
    response.write " selected class=bg_1"
  end if
  response.write ">"&crs("class_name")&"</option>"
  crs.movenext
loop

ctype=int(rs("forum_type"))
%>
</select></td> 
</tr>
<tr>
<td align=center>��̳˵����</td> 
<td><textarea name=forum_remark rows=5 cols=50><%=rs("forum_remark")%></textarea></td> 
</tr>
<tr>
<td align=center>��̳ͼƬ��</td> 
<td><input type=text name=forum_pic value='<%=rs("forum_pic")%>' size=30 maxlength=50></td> 
</tr>
<tr>
<td align=center>��̳���ͣ�</td> 
<td><select name=forum_type size=1>
<%
  dim tdim,t2
  tdim=split(forum_type,"|")
  for i=0 to ubound(tdim)
    response.write vbcrlf&"<option value='"&i+1&"'"
    if ctype=i+1 then response.write " selected"
    response.write ">"&right(tdim(i),len(tdim(i))-instr(tdim(i),":"))&"</option>"
  next
  erase tdim
%>
</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�Ƿ񿪷ţ�<input type=checkbox name=forum_hidden value='no'<% if rs("forum_hidden")=false then response.write " checked" %>>&nbsp;��ѡ��Ϊ���ţ�</td> 
</tr>
<tr height=50>
<td align=center>��̳������<br><br></td> 
<td><input type=text name=forum_power value='<%=rs("forum_power")%>' size=50 maxlength=50><br>������á�|���ֿ����磺������|apple|5271��</td> 
</tr>
<tr height=30><td colspan=2 align=center height=30><input type=submit value=' �� �� �� �� '></td></tr>
</form><%
  end if
%></table><%
end sub

sub forum_add()
  dim rs,strsql,classname,classid,forum_name,forum_order
  classid=trim(request("class_id"))
  if not(isnumeric(classid)) then
    call forum_main():exit sub
  end if
  strsql="select class_name from bbs_class where class_id="&classid
  set rs=conn.execute(strsql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call forum_main():exit sub
  end if
  classname=rs("class_name")
  rs.close:set rs=nothing
%><font class=red>�����̳</font>��<font class=blue_1><%=classname%></font>��<br><br><br>
<table border=0 width=400>
<%
if trim(request("add"))="ok" then
  forum_name=code_form(request.form("forum_name"))
  if forum_name="" then
    response.write( VbCrLf & "<tr><td height=80 align=center><font class=red_2>��̳���Ʋ���Ϊ�գ�</font><br><br>"&go_back&"</td></tr>")
  else
    set rs=server.createobject("adodb.recordset")
    strsql="Select top 1 * from bbs_forum where class_id="&classid&" order by forum_order desc"
    rs.open strsql,conn,1,1
    if rs.eof and rs.bof then
      forum_order=0
    else
      forum_order=rs("forum_order")
    end if
    forum_order=forum_order+1
    rs.close
    strsql="Select * from bbs_forum"
    rs.open strsql,conn,1,3
    rs.addnew
    rs("class_id")=classid
    rs("forum_order")=forum_order
    rs("forum_name")=forum_name
    rs("forum_remark")=request.form("forum_remark")
    rs("forum_power")=code_form(request.form("forum_power"))
    rs("forum_hidden")=false
    rs("forum_type")=1
    rs("forum_topic_num")=0
    rs("forum_data_num")=0
    rs("forum_new_info")="|||"
    rs.update
    response.write( VbCrLf & "<tr><td height=80 align=center>�ɹ����������̳��<font class=red>" & forum_name & "</font></td></tr>")
    rs.close:set rs=nothing
  end if
else
%>
<form method=post action='admin_forum.asp?action=forum_add&add=ok&class_id=<%=classid%>'>
<tr><td width='20%' align=center></td><td width='80%'></td></tr>
<tr height=30>
<td align=center>��̳���ƣ�</td> 
<td><input type=text name=forum_name size=30 maxlength=20></td> 
</tr>
<tr>
<td align=center>��̳˵����</td> 
<td><textarea name=forum_remark rows=5 cols=50></textarea></td> 
</tr>
<tr height=50>
<td align=center>��̳������<br><br></td> 
<td><input type=text name=forum_power size=50 maxlength=50><br>������á�|���ֿ����磺������|apple|5271��</td> 
</tr>
<tr height=30><td colspan=2 align=center height=30><input type=submit value=' �� �� �� �� '></td></tr>
</form><%
  end if
  response.write "</table>"
end sub

sub forum_mod()
%>
<table border=0>
<form action='admin_forum.asp?action=mod' method=post>
<input type=hidden name=modok value='ok'>
<tr><td align=center height=50 colspan=4><font class=red>�ϲ���̳</font></td></tr>
<%
  if trim(request.form("modok"))="ok" then
    response.write "<tr><td align=center height=50 colspan=4>"
    dim sel1,sel2,rs,sql
    sel1=trim(request.form("sel_1"))
    sel2=trim(request.form("sel_2"))
    if not(isnumeric(sel1)) or not(isnumeric(sel2)) then
      response.write "<font class=red_2>��û��ѡ��Ҫ�ϲ�����̳��</font>"
    else
      sql="update bbs_topic set forum_id="&int(sel2)&" where forum_id="&int(sel1)
      conn.execute(sql)
      sql="update bbs_data set forum_id="&int(sel2)&" where forum_id="&int(sel1)
      conn.execute(sql)
      response.write "<font class=red_3>��̳�ϲ��ɹ���</font>"
    end if
    response.write "</td></tr>"
  end if
%>
<tr height=50>
<td>��</td>
<td><select name=sel_1><% call forum_list() %></select></td>
<td>�ϲ���</td>
<td><select name=sel_2><% call forum_list() %></select></td>
</tr>
<tr><td align=center height=50 colspan=4><input type=submit value='��ʼ�ϲ�'></td></tr>
</form>
</table>
<%
end sub

sub forum_list()
  dim strsqlclass,rsclass,strsqlboard,rsboard
  strsqlclass="select class_id,class_name from bbs_class order by class_order"
  set rsclass=conn.execute(strsqlclass)
  if not(rsclass.bof and rsclass.eof) then
    do while not rsclass.eof
      response.write vbcrlf & "<option class=bg_2>�� "& rsclass("class_name") &"</option>"
      strsqlboard="select forum_id,forum_name from bbs_forum where class_id=" & rsclass("class_id") & " order by forum_order"
      set rsboard=conn.execute(strsqlboard)
      if rsboard.eof and rsboard.bof then
        response.write vbcrlf & "<option>û����̳</option>"
      else
        do while not rsboard.eof
          response.write vbcrlf & "<option value='" &rsboard("forum_id")& "'>����" & rsboard("forum_name") & "</option>"
	  rsboard.movenext
        loop
      end if
      rsclass.movenext
    loop
  end if
  set rsclass=nothing:set rsboard=nothing
end sub

sub forum_main()
%><table border=1 cellspacing=0 cellpadding=2 width=500 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<%
dim rsclass,strsqlclass,rsboard,strsqlboard,classid,forumid,forumname
strsqlclass="select * from bbs_class order by class_order"
set rsclass=conn.execute(strsqlclass)
if rsclass.bof and rsclass.eof then
  response.write vbcrlf & "<tr><td align=center height=200><font class=red_2>���ں���û����̳���࣡</font></td></tr>"
else
  do while not rsclass.eof
    classid=rsclass("class_id")
    response.write vbcrlf & "<tr height=20 bgcolor=#ffffff align=center><td align=left>" & img_small("fk2") & vbcrlf & "<font class=red_3><b>" & rsclass("class_name") & "</b></font></td><td><a href='admin_forum.asp?action=forum_add&class_id=" & classid & "'>�����̳</a></td><td><a href='admin_forum.asp?action=class_edit&class_id=" & classid & "'>�޸�</a></td><td><a href=""javascript:Do_del_class('" & classid & "');"">ɾ��</a></td><td>����<a href='admin_class_order.asp?class_id="&classid&"&action=up'>����</a> <a href='admin_class_order.asp?class_id="&classid&"&action=down'>����</a></td></tr>"
    strsqlboard="select forum_id,forum_name,forum_power,forum_hidden from bbs_forum where class_id=" & classid & " order by forum_order"
    set rsboard=conn.execute(strsqlboard)
    if rsboard.eof and rsboard.bof then
      response.write vbcrlf & "<tr><td colspan=5><font class=gray>���������໹û����̳</font></td></tr>"
    else
      do while not rsboard.eof
        forumid=rsboard("forum_id"):forumname=rsboard("forum_name")
        response.write vbcrlf&"<tr align=center><td align=left>����<font class=blue><b>" & forumname & "</b></font>"
        if rsboard("forum_hidden")=true then response.write " <font class=gray>����</font>"
        response.write "</td><td align=left>��������" & rsboard("forum_power") & "��</td><td><a href='admin_forum.asp?action=forum_edit&class_id="&classid&"&forum_id=" & forumid & "'>�༭</a></td><td><a href=""javascript:Do_del_forum(" & forumid & ");"">ɾ��</a></td><td>����<a href='admin_forum_order.asp?forum_id="&forumid&"&class_id="&classid&"&action=up'>����</a> <a href='admin_forum_order.asp?forum_id="&forumid&"&class_id="&classid&"&action=down'>����</a></td></tr>"
	rsboard.movenext
      loop
    end if
    rsclass.movenext
  loop
end if
set rsclass=nothing:set rsboard=nothing
%>
<tr><td align=center height=30 colspan=5><a href='admin_forum.asp?action=class_add'>�����̳����</a></td></tr>
</table>
<script language=JavaScript>
<!--
function Do_del_class(data1)
{
if (confirm("�˲�����ɾ��idΪ "+data1+" ����̳���࣡\n\n���Ҫɾ����\n\nɾ�����޷��ָ���"))
  window.location="admin_forum.asp?action=del_class&class_id="+data1
}

function Do_del_forum(data1)
{
if (confirm("�˲�����ɾ��idΪ "+data1+" ����̳��\n\n���Ҫɾ����\n\nɾ�����޷��ָ���"))
  window.location="admin_forum.asp?action=del_forum&forum_id="+data1
}
//-->
</script><%
end sub

sub del_class()
  dim classid,sql,rs,forumid
  classid=trim(request.querystring("class_id"))
  if not(isnumeric(classid)) then
    call forum_main():exit sub
  end if
  sql="delete from bbs_class where class_id="&classid
  conn.execute(sql)
  sql="select forum_id from bbs_forum where class_id="&classid
  set rs=conn.execute(sql)
  do while not rs.eof
    forumid=rs("forum_id")
    sql="delete from bbs_topic where forum_id="&forumid
    conn.execute(sql)
    sql="delete from bbs_data where forum_id="&forumid
    conn.execute(sql)
    rs.movenext
  loop
  sql="delete from bbs_forum where class_id="&classid
  conn.execute(sql)
  response.write "<script language=javascript>alert(""�ѳɹ���ɾ����һ����̳���࣡\n\n����������������̳�����ӣ�"");</script>"
  call forum_main()
end sub

sub del_forum()
  dim classid,forumid,sql
  forumid=trim(request.querystring("forum_id"))
  if not(isnumeric(forumid)) then
    call forum_main():exit sub
  end if
  sql="delete from bbs_forum where forum_id="&forumid
  conn.execute(sql)
  sql="delete from bbs_topic where forum_id="&forumid
  conn.execute(sql)
  sql="delete from bbs_data where forum_id="&forumid
  conn.execute(sql)
  response.write "<script language=javascript>alert(""�ѳɹ���ɾ����һ����̳��\n\n�����������������ӣ�"");</script>"
  call forum_main()
end sub
%>