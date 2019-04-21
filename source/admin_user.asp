<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/jk_md5.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim udim,unum,id,rssum,thepages,viewpage,page,nummer,pageurl,j,sqladd,admin_user,user_tit,now_username,now_id,now_power,now_hidden,del_temp,checkbox_val,power,hidden,keyword
tit="<a href='?'>用户管理</a>　┋"
udim=split(user_power,"|"):unum=ubound(udim)+1
for i=0 to unum-1
  tit=tit&"<a href='?power="&left(udim(i),instr(udim(i),":")-1)&"'>"&right(udim(i),len(udim(i))-instr(udim(i),":"))&"</a>┋"
next
erase udim
tit=tit&"　<a href='?hidden=true'>正常用户</a>┋" & _
    "<a href='?hidden=false'>未审核用户</a>"
response.write header(1,tit)
%>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='98%' cellspacing=0 cellpadding=2 align=center>
<tr><td align=center valign=top height=350>
<%
pageurl="?":sqladd="":user_tit=format_power2(unum,2)
admin_user=web_var(web_config,3)
id=trim(request.querystring("id"))
power=format_power(trim(request.querystring("power")),0)
if power<>"" and not isnull(power) then
  sqladd="where power='"&power&"' "
  pageurl=pageurl&"power="&power&"&"
  user_tit=format_power(power,1)
end if

if isnumeric(id) then
  select case action
  case "hidden"
    call user_hidden()
  case "locked"
    call useres_popedom(41)
  case "shield"
    call useres_popedom(42)
  end select
end if

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
  if hid<>"" then conn.execute("update user_data set"&hid&" where username<>'"&admin_user&"' and id="&id)
end sub

sub useres_popedom(pn)
  dim sql,rs,temp1,temp2,temp3,user_popedom
  sql="select popedom from user_data where id="&id
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then rs.close:set rs=nothing:exit sub
  user_popedom=rs("popedom")
  rs.close:set rs=nothing
  if len(user_popedom)<>50 then
    user_popedom="00000000000000000000000000000000000000000000000000"
  else
    if pn>len(user_popedom) then exit sub
    temp1=left(user_popedom,pn-1)
    temp2=popedom_format(user_popedom,pn)
    temp3=right(user_popedom,len(user_popedom)-pn)
    if int(temp2)=0 then
      temp2=1
    else
      temp2=0
    end if
  end if
  sql="update user_data set popedom='"&temp1&temp2&temp3&"' where id="&id
  conn.execute(sql)
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
      call delete_userdata(del_dim(del_i))
      del_sql="delete from user_data where username='"&del_dim(del_i)&"'"
      conn.execute(del_sql)
    next
    Erase del_dim
    del_select=vbcrlf&"<script language=javascript>alert(""共删除了 "&del_num+1&" 条记录！"");</script>"
  else
    del_select=vbcrlf&"<script language=javascript>alert(""没有删除记录！"");</script>"
  end if
end function

sub delete_userdata(username)
  if len(username)<1 then response.write 1:exit sub
  dim sql,rs,nn,tnum,dnum:tnum=0:dnum=0
  sql="select id,forum_id,re_counter from bbs_topic where username='"&username&"' order by id"
  set rs=conn.execute(sql)
  do while not rs.eof
    nn=rs("re_counter")+1
    tnum=tnum+1:dnum=dnum+nn
    sql="delete from bbs_data where reply_id="&rs("id")
    conn.execute(sql)
    sql="update bbs_forum set forum_topic_num=forum_topic_num-1,forum_data_num=forum_data_num-"&nn&" where forum_id="&rs("forum_id")
    conn.execute(sql)
    rs.movenext
  loop
  rs.close
  sql="delete from bbs_topic where username='"&username&"'"
  conn.execute(sql)
  
  sql="select forum_id,reply_id from bbs_data where username='"&username&"' order by id"
  set rs=conn.execute(sql)
  do while not rs.eof
    dnum=dnum+1
    sql="update bbs_topic set re_counter=re_counter-1 where id="&rs("reply_id")
    conn.execute(sql)
    sql="update bbs_forum set forum_data_num=forum_data_num-1 where forum_id="&rs("forum_id")
    conn.execute(sql)
    rs.movenext
  loop
  rs.close:set rs=nothing
  sql="delete from bbs_data where username='"&username&"'"
  conn.execute(sql)
  sql="update configs set num_topic=num_topic-"&tnum&",num_data=num_data-"&dnum&" where id=1"
  conn.execute(sql)
end sub

select case action
case "edit"
  if isnumeric(id) then
    call user_edit()
  else
    call user_main()
  end if
case else
  call user_main()
end select

call close_conn()
response.write ender()

sub user_edit()
  dim hidden,h1,h2,password,password2,passwd,passwd2,bbs_counter,counter,integral,emoney,u_popedom
  set rs=server.createobject("adodb.recordset")
  sql="select * from user_data where id="&id
  rs.open sql,conn,1,3
  if rs.eof and rs.bof then rs.close:set rs=nothing:call user_main():exit sub
  if rs("username")=web_var(web_config,3) then rs.close:set rs=nothing:call user_main():exit sub
  u_popedom=rs("popedom")
  if trim(request.querystring("edit"))="ok" then
    dim temp1,temp2,temp3
    if len(u_popedom)<>50 then
      u_popedom="00000000000000000000000000000000000000000000000000"
    else
      temp1=left(u_popedom,40)
      temp2=trim(request.form("locked"))&trim(request.form("shield"))
      temp3=right(u_popedom,8)
      u_popedom=temp1&temp2&temp3
    end if
    password=trim(request.form("password"))
    password2=trim(request.form("password2"))
    passwd=trim(request.form("passwd"))
    passwd2=trim(request.form("passwd2"))
    power=trim(request.form("power"))
    hidden=trim(request.form("hidden"))
    if password<>password2 then rs("password")=jk_md5(password,"short")
    if passwd<>passwd2 then rs("passwd")=jk_md5(passwd,"short")
    bbs_counter=trim(request.form("bbs_counter"))
    counter=trim(request.form("counter"))
    integral=trim(request.form("integral"))
    emoney=trim(request.form("emoney"))
    '-2147483648 +2147483647
    if isnumeric(bbs_counter) then
      bbs_counter=int(bbs_counter)
      if bbs_counter<>int(request.form("bbs_counter2")) and bbs_counter>0 and bbs_counter<=2147483647 then rs("bbs_counter")=bbs_counter
    end if
    if isnumeric(counter) then
      counter=int(counter)
      if counter<>int(request.form("counter2")) and counter>0 and counter<=2147483647 then rs("counter")=counter
    end if
    if isnumeric(integral) then
      integral=int(integral)
      if integral<>int(request.form("integral2")) and integral>0 and integral<=2147483647 then rs("integral")=integral
    end if
    if isnumeric(emoney) then
      emoney=int(emoney)
      if emoney<>int(request.form("emoney2")) and emoney>0 and emoney<=2147483647 then rs("emoney")=emoney
    end if
    rs("power")=power
    rs("hidden")=hidden
    rs("popedom")=u_popedom
    rs.update
    response.write "<br><br><br><br><br><br><font class=red>用户信息修改成功！</font><br><br><a href='?power="&power&"'>点击返回</a>"
  else
    power=rs("power"):hidden=rs("hidden")
%>
<table border=0 width=300>
  <form action='?action=edit&edit=ok&power=<%response.write power%>&id=<%response.write id%>' method=post>
  <tr><td colspan=2 align=center height=50><font class=red>用户管理修改</font></td></tr>
  <tr><td width='30%'>用户名称：</td><td width='70%'><input type=text value='<%response.write rs("username")%>' readonly size=25></td></tr>
  <tr><td>用户密码：</td><td><input type=text name=password value='<%response.write rs("password")%>' size=25 maxlength=20><input type=hidden name=password2 value='<%response.write rs("password")%>'></td></tr>
  <tr><td>密码钥匙：</td><td><input type=text name=passwd value='<%response.write rs("passwd")%>' size=25 maxlength=20><input type=hidden name=passwd2 value='<%response.write rs("passwd")%>'></td></tr>
  <tr><td>论坛发贴：</td><td><input type=text name=bbs_counter value='<%response.write rs("bbs_counter")%>' size=15 maxlength=10></td></tr><input type=hidden name=bbs_counter2 value='<%response.write rs("bbs_counter")%>'>
  <tr><td>文栏发贴：</td><td><input type=text name=counter value='<%response.write rs("counter")%>' size=15 maxlength=10></td></tr><input type=hidden name=counter2 value='<%response.write rs("counter")%>'>
  <tr><td>用户积分：</td><td><input type=text name=integral value='<%response.write rs("integral")%>' size=15 maxlength=10></td></tr><input type=hidden name=integral2 value='<%response.write rs("integral")%>'>
  <tr><td>用户金钱：</td><td><input type=text name=emoney value='<%response.write rs("emoney")%>' size=15 maxlength=10></td></tr><input type=hidden name=emoney2 value='<%response.write rs("emoney")%>'>
  <tr><td>用户类型：</td><td><select name=power size=1><%
    for i=1 to unum
      response.write vbcrlf & "<option value='"&format_power2(i,1)&"'"
      if power=format_power2(i,1) then response.write " selected"
      response.write ">"&format_power2(i,2)&"</option>"
    next
%></select>（<%response.write power%>）</td></tr>
  <tr><td>注册审核：</td><td><%
    if hidden=true then
      h1=" checked":h2=""
    else
      h1="":h2=" checked"
    end if
%><input type=radio name=hidden value=true<%response.write h1%>>正常<input type=radio name=hidden value=false<%response.write h2%>>未审核</td></tr>
  <tr><td>是否锁定：</td><td><%
    if int(popedom_format(u_popedom,41))=0 then
      h1=" checked":h2=""
    else
      h1="":h2=" checked"
    end if
%><input type=radio name=locked value='0'<%response.write h1%>>正常<input type=radio name=locked value='1'<%response.write h2%>>锁定</td></tr>
  <tr><td>论坛屏蔽：</td><td><%
    if int(popedom_format(u_popedom,42))=0 then
      h1=" checked":h2=""
    else
      h1="":h2=" checked"
    end if
%><input type=radio name=shield value='0'<%response.write h1%>>正常<input type=radio name=shield value='1'<%response.write h2%>>屏蔽</td></tr>
  <tr><td colspan=2 align=center height=30><input type=submit value=' 提 交 修 改 '></td></tr>
  </form>
</table>
<%
  end if
end sub

sub user_main()
  dim u_popedom
  hidden=trim(request.querystring("hidden"))
  if hidden="true" then
    sqladd="where hidden=1 "
    pageurl=pageurl&"hidden=true&"
    user_tit="正常用户"
  elseif hidden="false" then
    sqladd="where hidden=0 "
    pageurl=pageurl&"hidden=false&"
    user_tit="未审核用户"
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
  
  set rs=server.createobject("adodb.recordset")
  sql="select id,username,tim,power,hidden,popedom from user_data "&sqladd&" order by id desc"
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
    response.write "<br><br><br><br><br><br><br>还没有"&user_tit
    rs.close:set rs=nothing:exit sub
  end if
  
  rssum=rs.recordcount
  nummer=15
  call format_pagecute()
  del_temp=nummer
%>
<table border=1 width='98%' cellspacing=0 cellpadding=2 align=center bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<form name=del_form action='<%=pageurl%>del_ok=ok' method=post>
  <tr align=center height=30>
  <td colspan=2>现在有 <font class=red><%=rssum%></font> 位<%=user_tit%></td>
  <td colspan=6><%=jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000")%></td>
  </tr>
  <tr align=center bgcolor=#ededed height=20>
    <td width='6%'>序号</td>
    <td width='30%'>会员名称</td>
    <td width='22%'>注册时间</td>
    <td width='12%'>类型</td>
    <td width='8%'>审核</td>
    <td width='7%'>锁定</td>
    <td width='7%'>屏蔽</td>
    <td width='8%'>操作</td>
  </tr>
<%
  if int(viewpage)>1 then rs.move (viewpage-1)*nummer

  for i=1 to nummer
    if rs.eof then exit for
    checkbox_val=""
    now_id=rs("id"):now_username=rs("username")
    now_power=rs("power"):now_hidden=rs("hidden"):u_popedom=rs("popedom")
%>
  <tr align=center<%response.write mtr%>>
    <td align=left><a href='user_view.asp?username=<%response.write server.urlencode(now_username)%>' target=_blank><font color=#000000><%response.write i+(viewpage-1)*nummer%>.</font></a></td>
    <td align=left><a href='?power=<%response.write power%>&action=edit&id=<%response.write now_id%>'><font class=blue_1><%response.write now_username%></font></a></td>
    <td align=left><%response.write time_type(rs("tim"),8)%></td>
    <td><%
    for j=1 to unum
      if now_power=format_power2(j,1) then
        select case j
        case 1
          response.write "<font class=red>"&format_power2(j,2)&"</font>"
          checkbox_val="no":del_temp=del_temp-1:exit for
        case 2
          response.write "<font class=red_2>"&format_power2(j,2)&"</font>"
          checkbox_val="no":del_temp=del_temp-1:exit for
        case 3
          response.write "<font class=red_3>"&format_power2(j,2)&"</font>":exit for
        case 4
          response.write "<font class=blue>"&format_power2(j,2)&"</font>":exit for
        case else
          response.write format_power2(j,2):exit for
        end select
      end if
    next
%></td>
    <td><a href='?power=<%response.write power%>&hidden=<%response.write hidden%>&action=hidden&id=<%response.write now_id%>'><%
    if now_hidden=true then
      response.write "正常"
    else
      response.write "<font class=red_2>未审核</font>"
    end if
%></a></td>
    <td><a href='?power=<%response.write power%>&hidden=<%response.write hidden%>&action=locked&id=<%response.write now_id%>'><%
    if int(popedom_format(u_popedom,41))=0 then
      response.write "正常"
    else
      response.write "<font class=red_2>锁定</font>"
    end if
%></a></td>
    <td><a href='?power=<%response.write power%>&hidden=<%response.write hidden%>&action=shield&id=<%response.write now_id%>'><%
    if int(popedom_format(u_popedom,42))=0 then
      response.write "正常"
    else
      response.write "<font class=red_2>屏蔽</font>"
    end if
%></a></td>
    <td><%
    if checkbox_val<>"no" then
      response.write "<input type=checkbox name=del_id value='"&now_username&"'>"
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
  <td colspan=2><input type=submit value='删除所选' onclick="return suredel('<%response.write del_temp%>');"> &nbsp;<input type=checkbox name=del_all value=1 onClick=selectall('<%response.write del_temp%>')>&nbsp;选择所有</td>
</form>
  <td colspan=6>
    <table border=0>
    <form name=sea_frm action='<%response.write pageurl%>'>
    <tr>
    <td>关键字：</td>
    <td><input type=text name=keyword value='<%response.write keyword%>' size=20 maxlength=20>&nbsp;</td>
    <td>&nbsp;<input type=submit value=' 搜 索 '>&nbsp;</td>
    </tr>
    </form>
    </table>
  </td>
  </tr>
</table>
<%
  rs.close:set rs=nothing
%>
</td></tr></table>
<%
end sub
%>