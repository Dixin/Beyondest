<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim id,c_name,c_pass,c_emoney,c_hidden,rssum,nummer,thepages,viewpage,pageurl,page
tit="<a href='?'>会 员 卡</a>"
response.write header(1,tit)
id=trim(request.querystring("id"))

if trim(request("del_ok"))="ok" then
  response.write del_select(trim(request.form("del_id")))
end if

function del_select(delid)
  dim del_i,del_num,del_dim,del_sql,del_rs,del_username,fobj,picc
  if delid<>"" and not isnull(delid) then
    delid=replace(delid," ","")
    del_dim=split(delid,",")
    del_num=UBound(del_dim)
    for del_i=0 to del_num
      del_sql="delete from cards where c_id="&del_dim(del_i)
      conn.execute(del_sql)
    next
    Erase del_dim
    del_select=vbcrlf&"<script language=javascript>alert(""共删除了 "&del_num+1&" 条记录！"");</script>"
  end if
end function

if (action="hidden") and isnumeric(id) then
  sql="select c_hidden from cards where c_id="&id
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then
    if int(rs("c_hidden"))=0 then
      sql="update cards set c_hidden=1 where c_id="&id
    else
      sql="update cards set c_hidden=0 where c_id="&id
    end if
    conn.execute(sql)
  end if
  rs.close
  action=""
end if

select case action
case "del"
  if isnumeric(id) then
    call cards_del()
  else
    call cards_main()
  end if
case "add"
  call cards_add()
case "edit"
  if isnumeric(id) then
    call cards_edit()
  else
    call cards_main()
  end if
case else
  call cards_main()
end select

close_conn
response.write ender()

sub cards_edit()
  dim sql2,rs2
  set rs=server.createobject("adodb.recordset")
  sql="select * from cards where c_id="&id
  rs.open sql,conn,1,3
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""您的操作有错误（error in edit）存在！\n\n点击返回。"");" & _
		   vbcrlf & "location='?'" & _
		   vbcrlf & "</script>")
    exit sub
  end if
  if trim(request.querystring("chk"))="ok" then
    c_name=code_admin(request.form("c_name"))
    c_pass=code_admin(request.form("c_pass"))
    c_emoney=code_admin(request.form("c_emoney"))
    if len(c_name)<1 or len(c_pass)<1 or not(isnumeric(c_emoney)) then
      response.write "会员卡号、密码和分值不能为空！<br><br>"&go_back:exit sub
    end if
    
  if c_name<>code_admin(request.form("c_name2")) then
    sql2="select * from cards where c_name='"&c_name&"'"
    set rs2=conn.execute(sql2)
    if not(rs2.eof and rs2.bof) then
      rs2.close:set rs2=nothing
      response.write "会员卡号："&c_name&" 已存在！请选用其它的代码。<br><br>"&go_back:exit sub
    end if
    rs2.close:set rs2=nothing
  end if
    
    rs("c_name")=c_name
    rs("c_pass")=c_pass
    rs("c_emoney")=c_emoney
    if isnumeric(trim(request.form("c_hidden"))) then
      if int(trim(request.form("c_hidden")))=0 then
        rs("c_hidden")=0
      else
        rs("c_hidden")=1
      end if
    else
      rs("c_hidden")=0
    end if
    rs.update
    rs.close:set rs=nothing
    response.write "<script lanuage=javascrip>alert(""修改会员卡号成功！"");location.href='?page="&trim(request.querystring("page"))&"';</script>"
    exit sub
  end if
%>
<table border=0 align=center>
<form action='?action=edit&chk=ok&page=<%response.write trim(request.querystring("page"))%>&id=<%response.write id%>' method=post>
<tr><td>卡号：&nbsp;<input type=text name=c_name value='<%response.write rs("c_name")%>' size=20 maxlength=20></td></tr>
<input type=hidden name=c_name2 value='<%response.write rs("c_name")%>'>
<tr><td>密码：&nbsp;<input type=text name=c_pass value='<%response.write rs("c_pass")%>' size=20 maxlength=20></td></tr>
<tr><td>分值：&nbsp;<input type=text name=c_emoney value='<%response.write rs("c_emoney")%>' size=20 maxlength=20></td></tr>
<tr><td>是否使用：<input type=radio name=c_hidden value='1'<%if int(rs("c_hidden"))=1 then response.write " checked"%>>&nbsp;已使用&nbsp;
<input type=radio name=c_hidden value='0'<%if int(rs("c_hidden"))=0 then response.write " checked"%>>&nbsp;未使用</td></tr>
<tr><td align=center height=30><input type=submit value='修改会员卡'></td></tr>
</form>
</table>
<%
end sub

sub cards_add()
    c_name=code_admin(request.form("c_name"))
    c_pass=code_admin(request.form("c_pass"))
    c_emoney=code_admin(request.form("c_emoney"))
    if len(c_name)<1 or len(c_pass)<1 or not(isnumeric(c_emoney)) then
      response.write "会员卡号、密码和分值不能为空！<br><br>"&go_back:exit sub
    end if
    
    set rs=server.createobject("adodb.recordset")
    sql="select * from cards where c_name='"&c_name&"'"
    rs.open sql,conn,1,3
    if not(rs.eof and rs.bof) then
      rs.close:set rs=nothing
      response.write "会员卡号："&c_name&" 已存在！请选用其它的代码。<br><br>"&go_back:exit sub
    end if
    rs.addnew
    rs("c_name")=c_name
    rs("c_pass")=c_pass
    rs("c_emoney")=c_emoney
    rs("c_hidden")=0
    rs.update
    rs.close:set rs=nothing
    response.write "<script lanuage=javascrip>alert(""添加会员卡号成功！"");location.href='?';</script>"
end sub

sub cards_main()
  dim i,hidden,sqladd,sname,iid,del_temp
  hidden=trim(request.querystring("hidden"))
  pageurl="?hidden="&hidden&"&"
%>
<script language=javascript src='STYLE/admin_del.js'></script>
<form name=del_form action='<%response.write pageurl%>del_ok=ok' method=post>
<%
  sql="select * from cards order by c_id desc"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
   rssum=0
  else
    rssum=rs.recordcount
  end if
  nummer=15
  call format_pagecute()
  del_temp=nummer
  if rssum=0 then del_temp=0
  if int(page)=int(thepages) then
    del_temp=rssum-nummer*(thepages-1)
  end if
%>
<table border=1 width='80%' cellspacing=0 cellpadding=1 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<tr><td colspan=5 align=center height=30>
  <table border=0 width='100%'cellspacing=0 cellpadding=0>
  <tr align=center>
  <td width='40%'>现在有 <font class=red><%response.write rssum%></font> 个 <font class=red_3><%response.write sname%></font> 会员卡号</td>
  <td width='60%'><input type=checkbox name=del_all value=1 onClick=selectall('<%response.write del_temp%>')> 选中所有　<input type=submit value='删除所选' onclick="return suredel('<%response.write del_temp%>');"></td>
  </tr>
  </table>
</td></tr>
<tr align=center bgcolor=#ededed>
<td width='8%'>序号</td>
<td width='26%'>会员卡号代码</td>
<td width='26%'>会员卡号类型</td>
<td width='18%'>分值</td>
<td width='24%'>操作</td>
</tr>
<%
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    iid=rs("c_id")
%>
<tr>
<td align=center><%response.write i+(viewpage-1)*nummer%>.</td>
<td><%response.write rs("c_name")%></td>
<td><%response.write rs("c_pass")%></td>
<td><%response.write rs("c_emoney")%></td>
<td align=center><a href='?action=hidden&page=<%response.write viewpage%>&id=<%response.write iid%>'>
<%
if int(rs("c_hidden"))=0 then
  response.write "未使用"
else
  response.write "<font class=red>已使用</font>"
end if
%></a>&nbsp;
<a href='?action=edit&id=<%response.write iid%>'>修改</a>&nbsp;
<input type=checkbox name=del_id value='<%response.write iid%>' class=bg_1></td></tr>
<%
    rs.movenext
  next
  rs.close:set rs=nothing
%>
<tr><td colspan=5>页次：<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font>
分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%></td></tr>
</form>
</table>
<br>
<table border=0 align=center>
<form name=add_frm action='?action=add' method=post>
<tr>
<td>卡号：</td>
<td><input type=text name=c_name size=12 maxlength=20></td>
<td>密码：</td>
<td><input type=text name=c_pass size=12 maxlength=20></td>
<td>分值：</td>
<td><input type=text name=c_emoney size=10 maxlength=20></td>
<td>&nbsp;<input type=submit value='添加会员卡'></td>
</tr>
</form>
</table>
<%
end sub
%>