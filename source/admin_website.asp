<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<!-- #INCLUDE file="INCLUDE/functions.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim website_menu,nsort,sql2,rs2,del_temp,data_name,cid,sid,ncid,nsid,nid,id,left_type,now_id,nummer,sqladd,page,rssum,thepages,viewpage,pageurl,pic,ispic,csid
website_menu=vbcrlf & "<a href='?'>网站推荐</a>&nbsp;┋&nbsp;" & _
	     vbcrlf & "<a href='?action=add'>添加网站</a>&nbsp;┋&nbsp;" & _
	     vbcrlf & "<a href='admin_nsort.asp?nsort=web'>网站分类</a>"
response.write header(15,website_menu)
pageurl="?action="&action&"&":nsort="web":data_name="website":sqladd="":nummer=15
call admin_cid_sid()

if trim(request("del_ok"))="ok" then
  response.write del_select(trim(request.form("del_id")))
end if

function del_select(delid)
  dim del_i,del_num,del_dim,del_sql,fobj,picc
  if delid<>"" and not isnull(delid) then
    delid=replace(delid," ","")
    del_dim=split(delid,",")
    del_num=UBound(del_dim)
    for del_i=0 to del_num
      call upload_del(data_name,del_dim(del_i))
      del_sql="delete from "&data_name&" where id="&del_dim(del_i)
      conn.execute(del_sql)
    next
    Erase del_dim
    del_select=vbcrlf&"<script language=javascript>alert(""共删除了 "&del_num+1&" 条记录！"");</script>"
  end if
end function

id=trim(request.querystring("id"))
if (action="hidden" or action="isgood") and isnumeric(id) then
  sql="select "&action&" from "&data_name&" where id="&id
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then
    if rs(action)=true then
      sql="update "&data_name&" set "&action&"=0 where id="&id
    else
      sql="update "&data_name&" set "&action&"=1 where id="&id
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

close_conn()
response.write ender()

function select_type(st1,st2)
  select_type=vbcrlf&"<option"
  if st1=st2 then select_type=select_type&" selected"
  select_type=select_type&">"&st1&"</option>"
end function

sub news_edit()
  dim rs3,sql3,name,url,isgood,country,lang,remark
  if trim(request.querystring("edit"))="chk" then
    name=code_admin(request.form("name"))
    csid=trim(request.form("csid"))
    url=code_admin(request.form("url"))
    isgood=trim(request.form("isgood"))
    remark=request.form("remark")
    country=trim(request.form("country"))
    lang=trim(request.form("lang"))
    pic=trim(request.form("pic"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择网站类型！</font><br><br>"&go_back
    elseif len(name)<1 or len(url)<1 then
      response.write "<font class=red_2>网站名称和地址不能为空！</font><br><br>"&go_back
    elseif len(remark)>250 then
      response.write "<font class=red_2>网站说明不能长于250个字符！</font><br><br>"&go_back
    else
      call chk_cid_sid()
      rs("c_id")=cid
      rs("s_id")=sid
      if trim(request.form("username_my"))="yes" then rs("username")=login_username
      rs("name")=name
      rs("url")=url
      rs("country")=country
      rs("lang")=lang
      rs("remark")=remark
      if isgood="yes" then
        rs("isgood")=true
      else
        rs("isgood")=false
      end if
      if trim(request.form("hidden"))="yes" then
        rs("hidden")=false
      else
        rs("hidden")=true
      end if
      if len(pic)<3 then
        rs("pic")="no_pic.gif"
      else
        rs("pic")=pic
      end if
      rs("tim")=now_time
      rs.update
      rs.close:set rs=nothing
      call upload_note(data_name,id)
      response.write "<font class=red>已成功修改了一个网站！</font><br><br><a href='?c_id="&cid&"&s_id="&sid&"'>点击返回</a><br><br>"
    end if
  else
%><table border=0 cellspacing=0 cellpadding=3>
<form action='<%response.write pageurl%>c_id=<%response.write cid%>&s_id=<%response.write sid%>&id=<%response.write id%>&edit=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='12%'>网站名称：</td><td width='88%'><input type=text size=70 name=name value='<%response.write rs("name")%>' maxlength=50><%=redx%></td></tr>
  <tr><td>网站类型：</td><td><%call chk_csid(cid,sid):call chk_h_u()%></td></tr>
  <tr><td>网站地址：</td><td><input type=text size=70 name=url value='<%response.write rs("url")%>' maxlength=100><%=redx%></td></tr>
  <tr><td>国家地区：</td><td><select name=country size=1>
<%
pic=rs("pic")
if pic="no_pic.gif" then pic=""
ispic=pic
if Instr(ispic,"/")>0 then
  ispic=right(ispic,len(ispic)-Instr(ispic,"/"))
end if
if Instr(ispic,".")>0 then
  ispic=left(ispic,Instr(ispic,".")-1)
end if
if len(ispic)<1 then ispic="n"&upload_time(now_time)
tit=rs("country")
response.write select_type("中国",tit)
response.write select_type("香港",tit)
response.write select_type("台湾",tit)
response.write select_type("美国",tit)
response.write select_type("英国",tit)
response.write select_type("日本",tit)
response.write select_type("韩国",tit)
response.write select_type("加拿大",tit)
response.write select_type(">澳大利亚",tit)
response.write select_type("新西兰",tit)
response.write select_type("俄罗斯",tit)
response.write select_type("意大利",tit)
response.write select_type("法国",tit)
response.write select_type("西班牙",tit)
response.write select_type("德国",tit)
response.write select_type("其它国家",tit)
%>
</select>&nbsp;&nbsp;&nbsp;&nbsp;站点语言：<select name=lang size=1>
<%
tit=rs("lang")
response.write select_type("简体中文",tit)
response.write select_type("繁体中文",tit)
response.write select_type("English",tit)
response.write select_type("其它语言",tit)
%>
</select>&nbsp;&nbsp;&nbsp;推荐：<input type=checkbox name=isgood<%if rs("isgood")=true then response.write " checked"%> value='yes'></td></tr>
  <tr><td>图片地址：</td><td><input type=test name=pic value='<%response.write pic%>' size=70 maxlength=100></td></tr>
  <tr><td>上传图片：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=website&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr><td valign=top class=htd><br>网站内容：<br><=250B</td><td><textarea name=remark rows=5 cols=70><%response.write rs("remark")%></textarea></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' 修 改 网 站 '></td></tr>
</form></table><%
  end if
end sub

sub news_add()
  if trim(request.querystring("add"))="chk" then
    dim name,url,isgood,country,lang,remark
    name=code_admin(request.form("name"))
    csid=trim(request.form("csid"))
    url=code_admin(request.form("url"))
    isgood=trim(request.form("isgood"))
    remark=request.form("remark")
    country=trim(request.form("country"))
    lang=trim(request.form("lang"))
    pic=trim(request.form("picg"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择网站类型！</font><br><br>"&go_back
    elseif len(name)<1 or len(url)<1 then
      response.write "<font class=red_2>网站名称和地址不能为空！</font><br><br>"&go_back
    elseif len(remark)>250 then
      response.write "<font class=red_2>网站说明不能长于250个字符！</font><br><br>"&go_back
    else
      call chk_cid_sid()
      set rs=server.createobject("adodb.recordset")
      sql="select * from "&data_name
      rs.open sql,conn,1,3
      rs.addnew
      rs("c_id")=cid
      rs("s_id")=sid
      rs("username")=login_username
      rs("hidden")=true
      rs("name")=name
      rs("url")=url
      rs("country")=country
      rs("lang")=lang
      rs("remark")=remark
      if isgood="yes" then
        rs("isgood")=true
      else
        rs("isgood")=false
      end if
      rs("username")=login_username
      if len(pic)<3 then
        rs("pic")="no_pic.gif"
      else
        rs("pic")=pic
      end if
      rs("tim")=now_time
      rs("counter")=0
      rs.update
      rs.close:set rs=nothing
      call upload_note(data_name,first_id(data_name))
      response.write "<font class=red>已成功添加了一个网站！</font><br><br><a href='?c_id="&cid&"&s_id="&sid&"'>点击返回</a><br><br>"
    end if
  else
%><table border=0 cellspacing=0 cellpadding=3>
<form action='<%response.write pageurl%>add=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='12%'>网站名称：</td><td width='88%'><input type=text size=70 name=name maxlength=50><%=redx%></td></tr>
  <tr><td>网站类型：</td><td><%call chk_csid(cid,sid)%></td></tr>
  <tr><td>网站地址：</td><td><input type=text size=70 name=url value='http://' maxlength=100><%=redx%></td></tr>
  <tr><td>国家地区：</td><td><select name=country size=1>
<option>中国</option>
<option>香港</option>
<option>台湾</option>
<option>美国</option>
<option>英国</option>
<option>日本</option>
<option>韩国</option>
<option>加拿大</option>
<option>澳大利亚</option>
<option>新西兰</option>
<option>俄罗斯</option>
<option>意大利</option>
<option>法国</option>
<option>西班牙</option>
<option>德国</option>
<option>其它国家</option>
</select>&nbsp;&nbsp;&nbsp;&nbsp;站点语言：<select name=lang size=1>
<option>简体中文</option>
<option>繁体中文</option>
<option>English</option>
<option>其它语言</option>
</select>&nbsp;&nbsp;&nbsp;推荐：<input type=checkbox name=isgood value='yes'></td></tr>
<% ispic="w"&upload_time(now_time) %>
  <tr><td>图片地址：</td><td><input type=test name=pic size=70 maxlength=100></td></tr>
  <tr><td>上传图片：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=website&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr><td valign=top class=htd><br>网站内容：<br><=250B</td><td><textarea name=remark rows=5 cols=70></textarea></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' 添 加 网 站 '></td></tr>
</form></table><%
  end if
end sub

sub news_main()
%>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='100%' cellpadding=2>
  <tr valign=top height=350>
    <td width='25%' class=htd><br><%call left_sort()%></td>
    <td width='75%' align=center>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<form name=del_form action='<%=pageurl%>del_ok=ok' method=post>
<tr><td width='6%'></td><td width='81%'></td><td width='13%'></td></tr>
<%
  call sql_cid_sid()
  sql="select id,c_id,s_id,name,url,isgood,hidden from "&data_name&sqladd&" order by id desc"
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
现有<font class=red><%response.write rssum%></font>个网站　<%response.write "<a href='?action=add&c_id="&cid&"&s_id="&sid&"'>添加网站</a>"%>
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
    response.write website_center()
    rs.movenext
  next
  rs.close:set rs=nothing
%></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>页次：<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font>
分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%>
</td></tr></table>
    </td>
  </tr>
</table>
<%
end sub

function website_center()
  website_center=VbCrLf & "<tr"&mtr&">" & _
		 VbCrLf & "<td><a href='"&rs("url")&"' target=_blank title='浏览该网站'>" & i+(viewpage-1)*nummer & ".</a> </td><td>" & _
		 VbCrLf & "<a href='?action=edit&c_id="&ncid&"&s_id="&nsid&"&id=" & now_id & "'>" & rs("name") & "</a></td><td align=right><a href='?action=hidden&c_id="&cid&"&s_id="&sid&"&id="&now_id&"&page="&viewpage&"'>"
  if rs("hidden")=true then
    website_center=website_center&"显"
  else
    website_center=website_center&"<font class=red_2>隐</font>"
  end if
  website_center=website_center&"</a> <a href='?action=isgood&c_id="&cid&"&s_id="&sid&"&id="&now_id&"&page="&viewpage&"'>"
  if rs("isgood")=true then
    website_center=website_center&"<font class=red>是</font>"
  else
    website_center=website_center&"否"
  end if
  website_center=website_center&"</a><input type=checkbox name=del_id value='"&now_id&"'></td></tr>"
end function
%>