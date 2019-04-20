<!-- #include file="include/onlogin.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim id:id=trim(request.querystring("id"))
if not(isnumeric(id)) then
  response.redirect "admin_user_list.asp"
  response.end
end if
%>
<!-- #include file="include/conn.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/jk_md5.asp" -->
<%
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
<table border=0 width='98%' cellspacing=0 cellpadding=2 align=center>
<tr><td align=center height=350>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from user_data where id="&id
rs.open sql,conn,1,3
if rs.eof and rs.bof then
  rs.close:set rs=nothing
  call close_conn()
  response.redirect "admin_user_list.asp"
  response.end
end if

if rs("username")="笼民" then
  rs.close:set rs=nothing
  call close_conn()
  response.redirect "admin_user_list.asp"
  response.end
end if

if trim(request("edit"))="ok" then
  response.write user_chk()
else
  response.write user_type()
end if

rs.close:set rs=nothing
call close_conn()
%>
</td></tr></table>
<%
response.write ender()

function user_chk()
  dim password,password2,passwd,passwd2,bbs_counter,counter,integral,emoney,power,hidden
  password=trim(request.form("password"))
  password2=trim(request.form("password2"))
  passwd=trim(request.form("passwd"))
  passwd2=trim(request.form("passwd2"))
  power=trim(request.form("power"))
  hidden=trim(request.form("hidden"))
  if password<>password2 then
    rs("password")=jk_md5(password,"short")
  end if
  if passwd<>passwd2 then
    rs("passwd")=jk_md5(passwd,"short")
  end if
  bbs_counter=trim(request.form("bbs_counter"))
  counter=trim(request.form("counter"))
  integral=trim(request.form("integral"))
  emoney=trim(request.form("emoney"))
  '-2147483648 +2147483647
  if isnumeric(bbs_counter) then
    bbs_counter=int(bbs_counter)
    if bbs_counter<>int(request.form("bbs_counter2")) and bbs_counter>0 and bbs_counter<=2147483647 then
      rs("bbs_counter")=bbs_counter
    end if
  end if
  if isnumeric(counter) then
    counter=int(counter)
    if counter<>int(request.form("counter2")) and counter>0 and counter<=2147483647 then
      rs("counter")=counter
    end if
  end if
  if isnumeric(integral) then
    integral=int(integral)
    if integral<>int(request.form("integral2")) and integral>0 and integral<=2147483647 then
      rs("integral")=integral
    end if
  end if
  if isnumeric(emoney) then
    emoney=int(emoney)
    if emoney<>int(request.form("emoney2")) and emoney>0 and emoney<=2147483647 then
      rs("emoney")=emoney
    end if
  end if
  rs("power")=power
  rs("hidden")=hidden
  rs.update
  response.write "<font class=red>用户信息修改成功！</font><br><br><a href='admin_user_list.asp'>点击返回</a>"
end function

function user_type()
%>
<table border=0 width=300>
<form action='admin_user_edit.asp?edit=ok&id=<%response.write id%>' method=post>
  <tr>
    <td colspan=2 align=center height=50><font class=red>用户管理修改</font></td>
  </tr>
  <tr>
    <td width='30%'>用户名称：</td>
    <td width='70%'><input type=text value='<%response.write rs("username")%>' readonly size=25></td>
  </tr>
  <tr>
    <td>用户密码：</td>
    <td><input type=text name=password value='<%response.write rs("password")%>' size=25 maxlength=20><input type=hidden name=password2 value='<%response.write rs("password")%>'></td>
  </tr>
  <tr>
    <td>密码钥匙：</td>
    <td><input type=text name=passwd value='<%response.write rs("passwd")%>' size=25 maxlength=20><input type=hidden name=passwd2 value='<%response.write rs("passwd")%>'></td>
  </tr>
  <tr>
    <td>论坛发贴：</td>
    <td><input type=text name=bbs_counter value='<%response.write rs("bbs_counter")%>' size=15 maxlength=10></td>
  </tr><input type=hidden name=bbs_counter2 value='<%response.write rs("bbs_counter")%>'>
  <tr>
    <td>文栏发贴：</td>
    <td><input type=text name=counter value='<%response.write rs("counter")%>' size=15 maxlength=10></td>
  </tr><input type=hidden name=counter2 value='<%response.write rs("counter")%>'>
  <tr>
    <td>用户积分：</td>
    <td><input type=text name=integral value='<%response.write rs("integral")%>' size=15 maxlength=10></td>
  </tr><input type=hidden name=integral2 value='<%response.write rs("integral")%>'>
  <tr>
    <td>用户金钱：</td>
    <td><input type=text name=emoney value='<%response.write rs("emoney")%>' size=15 maxlength=10></td>
  </tr><input type=hidden name=emoney2 value='<%response.write rs("emoney")%>'>
  <tr>
    <td>用户类型：</td>
    <td><select name=power size=1><%
dim power,pi,hidden,h1,h2
power=rs("power")
for pi=1 to unum
  response.write vbcrlf & "<option value='"&format_power2(pi,1)&"'"
  if power=format_power2(pi,1) then response.write " selected"
  response.write ">"&format_power2(pi,2)&"</option>"
next
%></select>（<%response.write power%>）</td>
  </tr>
  <tr>
    <td>类型状态：</td>
    <td><%
hidden=rs("hidden")
if hidden=true then
  h1=" checked"
  h2=""
else
  h1=""
  h2=" checked"
end if
%><input type=radio name=hidden value=true<%response.write h1%>>正常<input type=radio name=hidden value=false<%response.write h2%>>锁定</td>
  </tr>
  <tr>
    <td colspan=2 align=center height=30><input type=submit value=' 提 交 修 改 '></td>
  </tr>
</form>
</table>
<%
end function
%>