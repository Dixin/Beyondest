<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

function code_admin(strers)
  dim strer:strer=trim(strers)
  if isnull(strer) or strer="" then code_admin="":exit function
  strer=replace(strer,"'","""")
  code_admin=strer
end function

function code_config(strers)
  dim strer:strer=trim(strers)
  if isnull(strer) or strer="" then code_config="":exit function
  strer=replace(strer,"'","")
  strer=replace(strer,":","")
  strer=replace(strer,"|","")
  strer=replace(strer,".","")
  code_config=strer
end function

function ender()
  ender=vbcrlf & "</td></tr><tr><td height=30 align=center>"&web_var(web_error,4)&"</td></tr>" & _
	vbcrlf & "</center></body></html>"
end function

sub sql_cid_sid()
  if cid<>0 then
    sqladd=" where c_id="&cid
    pageurl=pageurl&"c_id="&cid&"&"
    if sid<>0 then
      sqladd=sqladd&" and s_id="&sid
      pageurl=pageurl&"s_id="&sid&"&"
    end if
  end if
end sub

sub chk_cid_sid()
  if isnumeric(csid) then
    cid=csid:sid=0
  else
    cid=mid(csid,1,instr(csid,"-")-1)
    sid=mid(csid,instr(csid,"-")+1,len(csid))
  end if
end sub

sub admin_cid_sid()
  cid=trim(request.querystring("c_id"))
  sid=trim(request.querystring("s_id"))
  if not(isnumeric(cid)) then cid=0
  if not(isnumeric(sid)) then sid=0
  cid=int(cid):sid=int(sid)
end sub

sub chk_csid(cid,sid)
  dim sql3,rs3
  response.write "<select name=csid size=1>"
  sql3="select c_id,c_name from jk_class where nsort='"&nsort&"' order by c_order,c_id"
  set rs3=conn.execute(sql3)
  do while not rs3.eof
    nid=int(rs3(0))
    response.write vbcrlf&"<option value='"&nid&"' class=bg_2"
    if cid=nid then response.write " selected"
    response.write ">"&rs3(1)&"</option>"
    sql2="select s_id,s_name from jk_sort where c_id="&nid&" order by s_order,s_id"
    set rs2=conn.execute(sql2)
    do while not rs2.eof
      now_id=int(rs2(0))
      response.write vbcrlf&"<option value='"&nid&"-"&now_id&"'"
      if sid=now_id then response.write " selected"
      response.write ">　"&rs2(1)&"</option>"
      rs2.movenext
    loop
    rs2.close:set rs2=nothing
    rs3.movenext
  loop
  rs3.close:set rs3=nothing
  response.write "</select>"&redx
end sub

sub left_sort()
  dim rs,sql
  sql="select c_id,c_name from jk_class where nsort='"&nsort&"' order by c_order,c_id"
  set rs=conn.execute(sql)
  do while not rs.eof
    nid=int(rs(0))
    if cid=nid then
      response.write vbcrlf&img_small("jt1")&"<a href='?c_id="&nid&"'><b><font class=red_3>"&rs(1)&"</b></font></a><br>"
    else
      response.write vbcrlf&img_small("jt0")&"<a href='?c_id="&nid&"'><font class=red_3>"&rs(1)&"</font></a><br>"
    end if
    sql2="select s_id,s_name from jk_sort where c_id="&nid&" order by s_order,s_id"
    set rs2=conn.execute(sql2)
    do while not rs2.eof
      now_id=int(rs2(0))
      if sid=now_id then
        response.write vbcrlf&"　<a href='?c_id="&nid&"&s_id="&now_id&"'><font class=blue>"&rs2(1)&"</a></a><br>"
      else
        response.write vbcrlf&"　<a href='?c_id="&nid&"&s_id="&now_id&"'>"&rs2(1)&"</a><br>"
      end if
      rs2.movenext
    loop
    rs2.close:set rs2=nothing
    rs.movenext
  loop
  rs.close
end sub

sub left_sort2()
  dim rs,sql
  sql="select c_id,c_name from jk_class where nsort='"&nsort&"' order by c_order,c_id"
  set rs=conn.execute(sql)
  do while not rs.eof
    nid=int(rs(0))
    if cid=nid then
      response.write vbcrlf&img_small("jt1")&"<a href='?types="&types&"&c_id="&nid&"'><b><font class=red_3>"&rs(1)&"</b></font></a><br>"
    else
      response.write vbcrlf&img_small("jt0")&"<a href='?types="&types&"&c_id="&nid&"'><font class=red_3>"&rs(1)&"</font></a><br>"
    end if
    sql2="select s_id,s_name from jk_sort where c_id="&nid&" order by s_order,s_id"
    set rs2=conn.execute(sql2)
    do while not rs2.eof
      now_id=int(rs2(0))
      if sid=now_id then
        response.write vbcrlf&"　<a href='?types="&types&"&c_id="&nid&"&s_id="&now_id&"'><font class=blue>"&rs2(1)&"</a></a><br>"
      else
        response.write vbcrlf&"　<a href='?types="&types&"&c_id="&nid&"&s_id="&now_id&"'>"&rs2(1)&"</a><br>"
      end if
      rs2.movenext
    loop
    rs2.close:set rs2=nothing
    rs.movenext
  loop
  rs.close
end sub


sub chk_power(power,pt)
  dim ddim:ddim=split(user_power,"|")
  for i=0 to ubound(ddim)
    response.write vbcrlf&"<input type=checkbox name=power value='"&i+1&"' class=bg_1"
    if instr(1,"."&power&".","."&i+1&".")>0 or pt=1 then response.write " checked"
    response.write ">"&right(ddim(i),len(ddim(i))-instr(ddim(i),":"))
  next
  erase ddim
%><input type=checkbox name=power value='0' class=bg_1<%if instr(1,"."&power&".",".0.")>0 then response.write " checked"%>>游客<%
end sub

sub chk_emoney(ee)
  response.write "&nbsp;货币：<input type=text name=emoney value='"&ee&"' size=6 maxlength=10>"
end sub

sub chk_h_u()
%>&nbsp;&nbsp;<input type=checkbox name=hidden<%if rs("hidden")=false then response.write " checked"%> value='yes'>&nbsp;隐藏
&nbsp;<input type=checkbox name=username_my value='yes'>&nbsp;<font alt='发布人：<%response.write rs("username")%>'>修改发布人为我</font><%
end sub
%>