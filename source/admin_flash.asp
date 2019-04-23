<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nsort,sql2,rs2,del_temp,data_name,cid,sid,ncid,nsid,nid,id,left_type,now_id,nummer,sqladd,page,rssum,thepages,viewpage,pageurl,pic,ispic,types,csid
types=trim(request.querystring("types"))
types="flash"
pageurl="?action="&action&"&types="&types&"&":nsort=types:data_name="gallery":sqladd="":nummer=30
tit=vbcrlf & "<a href='?'>Flash管理</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='?action=add&types="&nsort&"'>添加Flash</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='admin_nsort.asp?nsort=flash'>Flash分类</a>"
response.write header(16,tit)
call admin_cid_sid()

if trim(request("del_ok"))="ok" then
  response.write del_select(trim(request.form("del_id")))
end if

id=trim(request.querystring("id"))
if (action="hidden" or action="istop") and isnumeric(id) then
  sql="select "&action&" from "&data_name&" where id="&id
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then
    if action="istop" then
      if int(rs(action))=1 then
        sql="update "&data_name&" set "&action&"=0 where id="&id
      else
        sql="update "&data_name&" set "&action&"=1 where id="&id
      end if
    else
      if rs(action)=true then
        sql="update "&data_name&" set "&action&"=0 where id="&id
      else
        sql="update "&data_name&" set "&action&"=1 where id="&id
      end if
    end if
    conn.execute(sql)
  end if
  rs.close:action=""
end if

select case action
case "add"
  call news_add()
case "edit"
  if not(isnumeric(id)) then
    call news_main()
  else
    set rs=server.createobject("adodb.recordset")
    sql="select * from "&data_name&" where id="&id
    rs.open sql,conn,1,3
    call news_edit()
  end if
case else
  call news_main()
end select

call close_conn()
response.write ender()

function select_type(st1,st2)
  select_type=vbcrlf&"<option"
  if st1=st2 then select_type=select_type&" selected"
  select_type=select_type&">"&st1&"</option>"
end function

sub news_edit()
  dim rs3,sql3,name
  if trim(request.querystring("edit"))="chk" then
    name=code_admin(request.form("name"))
    csid=trim(request.form("csid"))
    pic=code_admin(request.form("pic"))
    types=trim(request.form("types"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择文件分类！</font><br><br>"&go_back
    elseif len(name)<1 then
      response.write "<font class=red_2>文件名称说明不能为空！</font><br><br>"&go_back
    elseif len(pic)<3 then
      response.write "<font class=red_2>请上传文件或输入文件的地址！</font><br><br>"&go_back
    else
      call chk_cid_sid()
      rs("c_id")=cid
      rs("s_id")=sid
      if trim(request.form("username_my"))="yes" then rs("username")=login_username
      rs("types")=types
      rs("name")=name
      if len(code_admin(request.form("spic")))<3 then
        rs("spic")="no_pic.gif"
      else
        rs("spic")=code_admin(request.form("spic"))
      end if
      rs("pic")=pic
      rs("remark")=left(request.form("remark"),250)
      rs("power")=replace(replace(trim(request.form("power"))," ",""),",",".")
      if isnumeric(trim(request.form("emoney"))) then
        rs("emoney")=trim(request.form("emoney"))
      else
        rs("emoney")=0
      end if
      if trim(request.form("istop"))="yes" then
        rs("istop")=1
      else
        rs("istop")=0
      end if
      if isnumeric(trim(request.form("counter"))) then rs("counter")=trim(request.form("counter"))
      if trim(request.form("hidden"))="yes" then
        rs("hidden")=false
      else
        rs("hidden")=true
      end if
      rs.update
      rs.close:set rs=nothing
      call upload_note(data_name,id)
      response.write "<font class=red>已成功修改了一张文件！</font><br><br><a href='?c_id="&cid&"&s_id="&sid&"&types="&types&"'>点击返回</a><br><br>"
    end if
  else
    types=rs("types")
%><table border=0 cellspacing=0 cellpadding=3>
<form action='<%response.write pageurl%>c_id=<%response.write cid%>&s_id=<%response.write sid%>&id=<%response.write id%>&edit=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='12%'>文件名称：</td><td width='88%'><input type=text size=40 name=name value='<%response.write rs("name")%>' maxlength=50><%=redx%></td></tr>
  <tr><td>文件分类：</td><td><%call chk_csid(cid,sid)%>&nbsp;&nbsp;文件类型：<select name=types size=1>
<option value='flash'<%if types="flash" then response.write " selected"%>>Flash</option>
<option value='logo'<%if types="logo" then response.write " selected"%>>其他</option>
</select><%=redx%>&nbsp;&nbsp;<%call chk_emoney(rs("emoney"))%></td></tr>
  <tr><td align=center>浏览权限：</td><td><%call chk_power(rs("power"),0)%></td></tr>
  <tr><td align=center>浏览人气：</td><td><input type=text name=counter value='<%response.write rs("counter")%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name=istop value='yes'<%if int(rs("istop"))=1 then response.write " checked"%>>&nbsp;推荐&nbsp;&nbsp;<%call chk_h_u()%></td></tr>
<%
pic=rs("spic")
if pic="no_pic.gif" then pic=""
ispic=pic
if Instr(ispic,"/")>0 then ispic=right(ispic,len(ispic)-Instr(ispic,"/"))
if Instr(ispic,".")>0 then ispic=left(ispic,Instr(ispic,".")-1)
if len(ispic)<1 then ispic="n"&upload_time(now_time)
%>
  <tr><td>小 图 片：</td><td><input type=test name=spic value='<%response.write pic%>' size=70 maxlength=100></td></tr>
  <tr><td>上传图片：</td><td><iframe frameborder=0 name=upload_frames width='100%' height=30 scrolling=no src='upload.asp?uppath=gallery&upname=<%response.write ispic%>&uptext=spic'></iframe></td></tr>
<%
pic=rs("pic")
if pic="no_pic.gif" then pic=""
ispic=pic
if Instr(ispic,"/")>0 then ispic=right(ispic,len(ispic)-Instr(ispic,"/"))
if Instr(ispic,".")>0 then ispic=left(ispic,Instr(ispic,".")-1)
if len(ispic)<1 then ispic="n"&upload_time(now_time)
%>
  <tr><td>文件地址：</td><td><input type=test name=pic value='<%response.write pic%>' size=70 maxlength=100><%response.write redx%></td></tr>
  <tr><td>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=gallery&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr><td>文件说明：<br><br><=250字符</td><td><textarea name=remark maxlength=250 rows=5 cols=70><%response.write rs("remark")%></textarea></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' 提 交 修 改 '></td></tr>
</form></table><%
  end if
end sub

sub news_add()
  dim name,csid
  types=trim(request.querystring("types"))
  if types<>"flash" and types<>"logo" and types<>"baner" then types="paste"
  if trim(request.querystring("add"))="chk" then
    name=code_admin(request.form("name"))
    csid=trim(request.form("csid"))
    pic=code_admin(request.form("pic"))
    types=trim(request.form("types"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择文件分类！</font><br><br>"&go_back
    elseif len(name)<1 then
      response.write "<font class=red_2>文件名称说明不能为空！</font><br><br>"&go_back
    elseif len(pic)<3 then
      response.write "<font class=red_2>请上传文件或输入文件的地址！</font><br><br>"&go_back
    else
      call chk_cid_sid()
      set rs=server.createobject("adodb.recordset")
      sql="select * from "&data_name
      rs.open sql,conn,1,3
      rs.addnew
      rs("c_id")=cid
      rs("s_id")=sid
      rs("username")=login_username
      rs("types")=types
      rs("name")=name
      if len(code_admin(request.form("spic")))<3 then
        rs("spic")="no_pic.gif"
      else
        rs("spic")=code_admin(request.form("spic"))
      end if
      rs("pic")=pic
      rs("remark")=left(request.form("remark"),250)
      rs("power")=replace(replace(trim(request.form("power"))," ",""),",",".")
      if isnumeric(trim(request.form("emoney"))) then
        rs("emoney")=trim(request.form("emoney"))
      else
        rs("emoney")=0
      end if
      if trim(request.form("istop"))="yes" then
        rs("istop")=1
      else
        rs("istop")=0
      end if
      rs("counter")=0
      rs("tim")=now_time
      rs("hidden")=true
      rs.update
      rs.close:set rs=nothing
      call upload_note(data_name,first_id(data_name))
      response.write "<font class=red>已成功添加了一个文件！</font><br><br><a href='?c_id="&cid&"&s_id="&sid&"&types="&types&"'>点击返回</a><br><br>"
    end if
  else
%><table border=0 cellspacing=0 cellpadding=3>
<form action='<%response.write pageurl%>add=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='12%' align=center>文件名称：</td><td width='88%'><input type=text size=70 name=name maxlength=50><%=redx%></td></tr>
  <tr><td align=center>文件分类：</td><td><%call chk_csid(cid,sid)%>&nbsp;&nbsp;文件类型：<select name=types size=1>
<option value='flash'<%if types="flash" then response.write " selected"%>>Flash</option>
<option value='logo'<%if types="logo" then response.write " selected"%>>其他</option>
</select><%response.write redx%>&nbsp;&nbsp;<%call chk_emoney(0)%></td></tr>
  <tr><td align=center>浏览权限：</td><td><%call chk_power("",1)%></td></tr>
<%ispic="gs"&upload_time(now_time)%>
  <tr><td align=center>小 图 片：</td><td><input type=test name=spic size=70 maxlength=100></td></tr>
  <tr><td align=center>上传图片：</td><td><iframe frameborder=0 name=upload_frames width='100%' height=28 scrolling=no src='upload.asp?uppath=gallery&upname=<%response.write ispic%>&uptext=spic'></iframe></td></tr>
<%ispic="g"&upload_time(now_time)%>
  <tr><td align=center>文件地址：</td><td><input type=test name=pic size=70 maxlength=100><%response.write redx%></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=28 scrolling=no src='upload.asp?uppath=gallery&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr><td align=center>文件说明：<br><br><=250字符</td><td><textarea name=remark rows=5 cols=70></textarea></td></tr>
  <tr><td colspan=2 align=center height=30><input type=submit value=' 提 交 添 加 '></td></tr>
</form></table><%
  end if
end sub

sub news_main()
%>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='100%' cellpadding=2>
  <tr valign=top height=350>
    <td width='25%' class=htd><br>文件类型：<br>

<a href='?types=flash'<%if types="flash" then response.write " class=red_3"%>>Flash</a><br>
<a href='?types=logo'<%if types="logo" then response.write " class=red_3"%>>其他</a><br>
<br>文件分类：<br><%call left_sort2()%></td>
    <td width='75%' align=center>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<form name=del_form action='<%=pageurl%>del_ok=ok' method=post>
<tr><td width='6%'></td><td width='80%'></td><td width='14%'></td></tr>
<%
  call sql_cid_sid()
  if len(sqladd)<1 then
    sqladd=" where types='"&types&"'"
  else
    sqladd=sqladd&" and types='"&types&"'"
  end if
  sql="select id,c_id,s_id,name,pic,hidden,istop from "&data_name&sqladd&" order by id desc"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
    rssum=0
  else
    rssum=rs.recordcount
  end if
  call format_pagecute()
  del_temp=nummer
  if rssum=0 then del_temp=0
  if int(page)=int(thepages) then
    del_temp=rssum-nummer*(thepages-1)
  end if
%>
<tr><td colspan=3 align=center height=25>
现有<font class=red><%response.write rssum%></font>个文件　<%response.write "<a href='?action=add&c_id="&cid&"&s_id="&sid&"'>添加文件</a>"%>
　<input type=checkbox name=del_all value=1 onClick=selectall('<%response.write del_temp%>')> 选中所有　<input type=submit value='删除所选' onclick=""return suredel('<%response.write del_temp%>');"">
</td></tr>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<%
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    now_id=rs("id"):ncid=rs("c_id"):nsid=rs("s_id")
    response.write gallery_center()
    rs.movenext
  next
  rs.close:set rs=nothing
%></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>页次：<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font>
分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%>
</td></tr></table>
</td></tr></table>
<%
end sub

function gallery_center()
  gallery_center=vbcrlf&"<tr"&mtr&">" & _
		 vbcrlf&"<td><a href='"&url_true(web_var(web_upload,1),rs("pic"))&"' target=_blank title='浏览该文件'>" & i+(viewpage-1)*nummer & ".</a> </td><td>" & _
		 vbcrlf&"<a href='?action=edit&c_id="&rs(1)&"&s_id="&rs(2)&"&id=" & now_id & "'>" & rs("name") & "</a></td><td align=center><a href='?action=hidden&c_id="&cid&"&s_id="&sid&"&id="&now_id&"&types="&types&"&page="&viewpage&"'>"
    if rs("hidden")=true then
      gallery_center=gallery_center&"显"
    else
      gallery_center=gallery_center&"<font class=red_2>隐</font>"
    end if
  gallery_center=gallery_center&"</a> <a href='?action=istop&c_id="&cid&"&s_id="&sid&"&id="&now_id&"&types="&types&"&page="&viewpage&"'>"
  if int(rs("istop"))=1 then
    gallery_center=gallery_center&"<font class=red>是</font>"
  else
    gallery_center=gallery_center&"否"
  end if
  gallery_center=gallery_center&"</a> <input type=checkbox name=del_id value='"&now_id&"'></td></tr>"
end function
%>