<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim nsort,sql2,rs2,del_temp,data_name,cid,sid,nid,ncid,nsid,id,left_type,now_id,nummer,sqladd,page,rssum,thepages,viewpage,pageurl,topic,csid
tit=vbcrlf & "<a href='?'>文栏管理</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='?action=add'>添加文章</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='admin_nsort.asp?nsort=article'>文栏分类</a>"
response.write header(13,tit)
pageurl="?action="&action&"&":nsort="art":data_name="article":sqladd="":nummer=15
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

sub news_edit()
  if trim(request.querystring("edit"))="chk" then
    topic=code_admin(request.form("topic"))
    csid=trim(request.form("csid"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择文章类型！</font><br><br>"&go_back
    elseif topic="" then
      response.write "<font class=red_2>文章标题不能为空！</font><br><br>"&go_back
    else
      call chk_cid_sid()
      rs("c_id")=cid
      rs("s_id")=sid
      if trim(request.form("username_my"))="yes" then rs("username")=login_username
      rs("topic")=topic
      rs("word")=request.form("word")
      if isnumeric(trim(request.form("emoney"))) then
        rs("emoney")=trim(request.form("emoney"))
      else
        rs("emoney")=0
      end if
      rs("author")=code_admin(request.form("author"))
      rs("power")=replace(replace(trim(request.form("power"))," ",""),",",".")
      rs("keyes")=code_admin(request.form("keyes"))
      if trim(request.form("istop"))="yes" then
        rs("istop")=1
      else
        rs("istop")=0
      end if
      if trim(request.form("hidden"))="yes" then
        rs("hidden")=false
      else
        rs("hidden")=true
      end if
      if isnumeric(trim(request.form("counter"))) then rs("counter")=trim(request.form("counter"))
      rs.update
      rs.close:set rs=nothing
      call upload_note(data_name,id)
      response.write "<font class=red>已成功修改了一篇文章！</font><br><br><a href='?c_id="&cid&"&s_id="&sid&"'>点击返回</a><br><br>"
    end if
  else
    dim sql3,rs3
%><table border=0 width='98%' cellspacing=0 cellpadding=1>
<form name='add_frm' action='<%response.write pageurl%>c_id=<%response.write cid%>&s_id=<%response.write sid%>&id=<%response.write id%>&edit=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>文章标题：</td><td width='85%'><input type=text size=70 name=topic value='<%=rs("topic")%>' maxlength=40><%=redx%></td></tr>
  <tr><td align=center>文章类型：</td><td><%call chk_csid(cid,sid):call chk_emoney(rs("emoney")):call chk_h_u()%></td></tr>
  <tr><td align=center>浏览权限：</td><td><%call chk_power(rs("power"),0)%></td></tr>
  <tr><td align=center>文章作者：</td><td><input type=text size=12 name=author value='<%response.write rs("author")%>' maxlength=20>&nbsp;&nbsp;关键字：<input type=text name=keyes value='<%response.write rs("keyes")%>' size=12 maxlength=20>&nbsp;&nbsp;推荐：<input type=checkbox name=istop value='yes'<%if int(rs("istop"))=1 then response.write " checked"%>>&nbsp;&nbsp;人次：<input type=text name=counter value='<%response.write rs("counter")%>' size=10 maxlength=10></td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","word","&nbsp;&nbsp;")%></td></tr>
  <tr><td valign=top align=center><br>文章内容：</td><td><textarea name=word rows=15 cols=70><%=rs("word")%></textarea></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=article&upname=a&uptext=word'></iframe></td></tr>
  <tr height=25><td></td><td><input type=submit value=' 修 改 文 章 '></td></tr>
</form>
</table><%
  end if
end sub

sub news_add()
  if trim(request.querystring("add"))="chk" then
    topic=code_admin(request.form("topic"))
    csid=trim(request.form("csid"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择文章类型！</font><br><br>"&go_back
    elseif topic="" then
      response.write "<font class=red_2>文章标题不能为空！</font><br><br>"&go_back
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
      rs("topic")=topic
      rs("word")=request.form("word")
      if isnumeric(trim(request.form("emoney"))) then
        rs("emoney")=trim(request.form("emoney"))
      else
        rs("emoney")=0
      end if
      rs("author")=code_admin(request.form("author"))
      rs("power")=replace(replace(trim(request.form("power"))," ",""),",",".")
      rs("keyes")=code_admin(request.form("keyes"))
      if trim(request.form("istop"))="yes" then
        rs("istop")=1
      else
        rs("istop")=0
      end if
      rs("tim")=now_time
      rs("counter")=0
      rs.update
      rs.close:set rs=nothing
      call upload_note(data_name,first_id(data_name))
      response.write "<font class=red>已成功添加了一篇文章！</font><br><br><a href='?c_id="&cid&"&s_id="&sid&"'>点击返回</a><br><br>"
    end if
  else
%><table border=0 width='98%' cellspacing=0 cellpadding=1>
<form name='add_frm' action='<%response.write pageurl%>add=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>文章标题：</td><td width='85%'><input type=text size=70 name=topic maxlength=40><%=redx%></td></tr>
  <tr><td align=center>文章类型：</td><td><%call chk_csid(cid,sid):call chk_emoney(0)%></td></tr>
  <tr><td align=center>浏览权限：</td><td><%call chk_power("",1)%></td></tr>
  <tr><td align=center>文章作者：</td><td><input type=text size=12 name=author maxlength=20>&nbsp;&nbsp;关键字：<input type=text name=keyes size=12 maxlength=20>&nbsp;&nbsp;推荐：<input type=checkbox name=istop value='yes'></td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","word","&nbsp;&nbsp;")%></td></tr>
  <tr><td valign=top align=center><br>文章内容：</td><td><textarea name=word rows=15 cols=70></textarea></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=article&upname=a&uptext=word'></iframe></td></tr>
  <tr><td></td><td height=25><input type=submit value=' 添 加 文 章 '></td></tr>
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
  sql="select id,c_id,s_id,topic,hidden,istop from "&data_name&sqladd&" order by id desc"
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
现有<font class=red><%response.write rssum%></font>篇文章　<%response.write "<a href='?action=add&c_id="&cid&"&s_id="&sid&"'>添加文章</a>"%>
　<input type=checkbox name=del_all value=1 onClick=selectall('<%response.write del_temp%>')> 选中所有　<input type=submit value='删除所选' onclick=""return suredel('<%response.write del_temp%>');"">
</td></tr>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<%
  if int(viewpage)<>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    now_id=rs("id"):ncid=rs("c_id"):nsid=rs("s_id")
    response.write article_center()
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

function article_center()
  article_center=VbCrLf & "<tr"&mtr&">" & _
		 VbCrLf & "<td>" & i+(viewpage-1)*nummer & ". </td><td>" & _
		 VbCrLf & "<a href='?action=edit&c_id="&ncid&"&s_id="&nsid&"&id=" & now_id & "'>" & cuted(rs("topic"),30) & "</a>" & _
		 "</td><td align=right><a href='?action=hidden&c_id="&cid&"&s_id="&sid&"&id="&now_id&"&page="&viewpage&"'>"
  if rs("hidden")=true then
    article_center=article_center&"显"
  else
    article_center=article_center&"<font class=red_2>隐</font>"
  end if
  article_center=article_center&"</a> <a href='?action=istop&c_id="&cid&"&s_id="&sid&"&id="&now_id&"&page="&viewpage&"'>"
  if int(rs("istop"))=1 then
    article_center=article_center&"<font class=red>是</font>"
  else
    article_center=article_center&"否"
  end if
  article_center=article_center&"</a> <input type=checkbox name=del_id value='"&now_id&"' class=bg_1></td></tr>"
end function
%>