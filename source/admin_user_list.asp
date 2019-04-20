<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim admin_menu,udim,unum
admin_menu="<a href='admin_user_list.asp'>用户管理</a>　　┋"
udim=split(user_power,"|"):unum=ubound(udim)+1
for i=0 to unum-1
  admin_menu=admin_menu&"<a href='admin_user_list.asp?power="&left(udim(i),instr(udim(i),":")-1)&"'>"&right(udim(i),len(udim(i))-instr(udim(i),":"))&"</a>┋"
next
admin_menu=admin_menu&"　　<a href='admin_user_list.asp?hidden=true'>正常用户</a>┋" & _
	   "<a href='admin_user_list.asp?hidden=false'>锁定用户</a>"

response.write header(1,admin_menu)
%>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='98%' cellspacing=0 cellpadding=2 align=center>
<tr><td align=center valign=top height=350>
<%
dim id,rssum,thepages,viewpage,page,nummer,pageurl,sqladd,user_tit,now_username,now_id,now_power,now_hidden,del_temp,checkbox_val,power,hidden,keyword
pageurl="admin_user_list.asp?"
sqladd=""
user_tit=format_power2(unum,2)
power=format_power(trim(request.querystring("power")),0)
if power<>"" and not isnull(power) then
  sqladd="where power='"&power&"' "
  pageurl=pageurl&"power="&power&"&"
  user_tit=format_power(power,1)
end if

id=trim(request.querystring("id"))
hidden=trim(request.querystring("hidden"))
if hidden="true" then
  sqladd="where hidden=1 "
  pageurl=pageurl&"hidden=true&"
  user_tit="正常用户"
elseif hidden="false" then
  sqladd="where hidden=0 "
  pageurl=pageurl&"hidden=false&"
  user_tit="锁定用户"
end if

keyword=trim(request.querystring("keyword"))
if keyword<>"" and not isnull(keyword) then
  if sqladd<>"" then
    sqladd=sqladd&"and username like '%"&keyword&"%' "
  else
  sqladd=sqladd&"where username like '%"&keyword&"%' "
  end if
  pageurl=pageurl&"keyword="&server.urlencode(keyword)&"&"
end if

if action="hidden" and isnumeric(id) then call user_hidden()
sub user_hidden()
  dim rs,sql,hid:hid=""
  sql="select username,hidden from user_data where id="&id
  set rs=conn.execute(sql)
  if not (rs.eof and rs.bof) then
    if rs(0)=web_var(web_config,3) then exit sub

    if rs("hidden")=true then
      hid=" hidden=0"
    else
      hid=" hidden=1"
    end if
  end if
  rs.close:set rs=nothing

  if hid<>"" then conn.execute("update user_data set"&hid&" where id="&id)
end sub

if trim(request("del_ok"))="ok" then response.write del_select()
function del_select()
  dim delid,del_i,del_num,del_dim,del_sql
  delid=request("del_id")
  if delid<>"" and not isnull(delid) then
    delid=replace(delid," ","")
    del_dim=split(delid,",")
    del_num=UBound(del_dim)
    for del_i=0 to del_num
      del_sql="delete from user_data where id="&del_dim(del_i)
      conn.execute(del_sql)
    next
    Erase del_dim
    del_select=vbcrlf&"<script language=javascript>alert(""共删除了 "&del_num+1&" 条记录！"");</script>"
  else
    del_select=vbcrlf&"<script language=javascript>alert(""没有删除记录！"");</script>"
  end if
end function

set rs=server.createobject("adodb.recordset")
sql="select id,username,tim,power,hidden from user_data "&sqladd&" order by id desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
  response.write "<br><br><br><br><br><br><br>还没有"&user_tit
else
  rssum=rs.recordcount
  nummer=15
  call format_pagecute()
  del_temp=nummer
%>
<table border=1 width='98%' cellspacing=0 cellpadding=0 align=center bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<form name=del_form action='<%=pageurl%>del_ok=ok' method=post>
  <tr align=center height=30>
  <td colspan=2>现在有 <font class=red><%=rssum%></font> 位<%=user_tit%></td>
  <td colspan=4><%=jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000")%></td>
  </tr>
  <tr align=center bgcolor=#ededed height=25>
    <td width='6%'>序号</td>
    <td width='30%'>会员名称</td>
    <td width='28%'>注册时间</td>
    <td width='14%'>类型</td>
    <td width='12%'>状态</td>
    <td width='10%'>操作</td>
  </tr>
<%
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    checkbox_val=""
    now_id=rs("id")
    now_username=rs("username")
    now_power=rs("power")
    now_hidden=rs("hidden")
%>
  <tr align=center>
    <td align=left><a href='user_view.asp?username=<%response.write server.urlencode(now_username)%>' target=_blank><font color=#000000><%=i+(viewpage-1)*nummer%>.</font></a></td>
    <td><a href='admin_user_edit.asp?id=<%=now_id%>'><font class=blue_1><%=now_username%></font></a></td>
    <td><%=rs("tim")%></td>
    <td><%
select case now_power
case format_power2(1,1)
  response.write "<font class=red>"&format_power2(1,2)&"</font>"
  checkbox_val="no"
  del_temp=del_temp-1
case format_power2(2,1)
  response.write "<font class=red_2>"&format_power2(2,2)&"</font>"
  checkbox_val="no"
  del_temp=del_temp-1
case format_power2(3,1)
  response.write "<font class=red_3>"&format_power2(3,2)&"</font>"
case format_power2(4,1)
  response.write "<font class=red_4>"&format_power2(4,2)&"</font>"
case else
  response.write format_power2(5,2)
end select
%></td>
    <td><a href='admin_user_list.asp?power=<%response.write power%>&hidden=<%response.write hidden%>&action=hidden&id=<%response.write now_id%>'><%
select case now_hidden
case true
  response.write "正常"
case else
  response.write "<font class=red_2>锁定</font>"
end select
%></a></td>
    <td><%
if checkbox_val<>"no" then
  response.write "<input type=checkbox name=del_id value='"&now_id&"'>"
else
  response.write "&nbsp;"
end if
%></td>
  </tr>
<%
    rs.movenext
  next
%>
  <tr align=center height=30>
  <td colspan=2><input type=submit value='删除所选' onclick="return suredel('<%=del_temp%>');"> &nbsp;<input type=checkbox name=del_all value=1 onClick=selectall('<%=del_temp%>')>&nbsp;选择所有</td>
</form>
  <td colspan=4>
<table border=0>
<form name=sea_frm action='<%=pageurl%>'>
<tr>
<td>关键字：</td>
<td><input type=text name=keyword value='<%=keyword%>' size=20 maxlength=20>&nbsp;</td>
<td>&nbsp;<input type=submit value=' 搜 索 '>&nbsp;</td>
</tr>
</form>
</table>
  </td>
  </tr>
</table>
<%
end if
rs.close:set rs=nothing
%>
</td></tr></table>
<%
call close_conn()
response.write ender()
%>