<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #INCLUDE file="INCLUDE/jk_page_cute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim id,nsort,rssum,nummer,thepages,viewpage,pageurl,page
nsort=trim(request("nsort"))
select case nsort
case "forum"
  nsort=nsort
case else
  nsort="news"
end select

sql="select * from bbs_cast"			' where sort='"&nsort&"'"

tit="<a href='admin_update.asp?'>网站更新</a> ┋ " & _
    "<a href='admin_data.asp'>数据更新</a> ┋ " & _
    "<a href='admin_update.asp?nsort=news'>更新公告</a> ┋ " & _
    "<a href='admin_update.asp?nsort=forum'>论坛公告</a> ┋ " & _
    "<a href='admin_update.asp?action=add'>添加更新</a>"
    
response.write header(7,tit)
id=trim(request.querystring("id"))

select case action
case "add"
  response.write news_add()
case "addchk"
  response.write news_addchk()
case "del"
  if isnumeric(id) then
    response.write news_del(id)
  else
    response.write news_main()
  end if
case "edit"
  if isnumeric(id) then
    response.write news_edit(id)
  else
    response.write news_main()
  end if
case "editchk"
  if isnumeric(id) then
    response.write news_editchk(id)
  else
    response.write news_main()
  end if
case else
  response.write news_main()
end select

response.write ender()

function news_del(id)
  on error resume next
  conn.execute("delete from bbs_cast where sort='"&nsort&"' and id="&id)
  call upload_del("update",id)
  if err then
    err.clear
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""您的操作有错误（error in del）存在！\n\n点击返回。"");" & _
		   vbcrlf & "location='?nsort="&nsort&"'" & _
		   vbcrlf & "</script>")
  else
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""成功删除了一条更新！\n\n点击返回。"");" & _
		   vbcrlf & "location='?nsort="&nsort&"'" & _
		   vbcrlf & "</script>")
  end if
end function

function news_main()
  pageurl="?nsort="&nsort&"&action=main&"
  set rs=server.createobject("adodb.recordset")
  sql=sql&" where sort='"&nsort&"' order by id desc"
  rs.open sql,conn,1,1
  if not(rs.eof and rs.bof) then
    rssum=rs.recordcount
    nummer=15
    call format_pagecute
    
    news_main=news_main&vbcrlf&"<script language=JavaScript><!--" & _
	      vbcrlf&"function Do_del_data(data1)" & _
	      vbcrlf&"{" & _
	      vbcrlf&"if (confirm(""此操作将删除id为 ""+data1+"" 的展会信息！\n真的要删除吗？\n删除后将无法恢复！""))" & _
	      vbcrlf&"  window.location=""?nsort="&nsort&"&action=del&id=""+data1" & _
	      vbcrlf&"}" & _
	      vbcrlf&"//--></script>" & _
	      vbcrlf&"<table border=1 width=500 cellspacing=0 cellpadding=1 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>" & _
	      vbcrlf&"<tr><td colspan=3 align=center height=30>现在有 <font class=red>"&rssum&"</font> 条新闻</td></tr>" & _
	      "<tr align=center><td width='8%'>序号</td><td width='75%'>标题</td><td width='17%'>操作</td></tr>"
    if int(viewpage)>1 then
      rs.move (viewpage-1)*nummer
    end if
    for i=1 to nummer
      if rs.eof then exit for
      news_main=news_main&vbcrlf&"<tr align=center><td>"&i+(viewpage-1)*nummer&".</td><td align=left>"&code_html(rs("topic"),1,28)&"</td><td><a href='?nsort="&nsort&"&action=edit&id="&rs("id")&"'>修改</a> ┋ <a href='javascript:Do_del_data("&rs("id")&")'>删除</a></td></tr>"
      rs.movenext
    next
    news_main=news_main&vbcrlf&"</table>"&kong&pagecute_fun(viewpage,thepages,pageurl)
  end if
  rs.close:set rs=nothing
end function

function news_add()
%><table border=0 width='98%' cellspacing=0 cellpadding=2>
<form name='add_frm' action='?action=addchk' method=post>
<input type=hidden name=upid value=''>
  <tr><td colspan=2 align=center height=50><font class=red>添加公告更新</font></td></tr>
  <tr><td width='15%' align=center>更新标题：</td><td width='85%'><input type=text name=topic size=65 maxlength=50></td></tr>
  <tr><td align=center height=30>新增类型：</td><td><input type=radio name=nsort value='news' checked>&nbsp;网站更新&nbsp;&nbsp;<input type=radio name=nsort value='forum'>&nbsp;论坛公告</td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","word","&nbsp;&nbsp;")%></td></tr>
  <tr><td align=center valign=top><br>更新内空：</td><td><textarea name=word rows=15 cols=65></textarea></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=f&uptext=word'></iframe></td></tr>
  <tr height=30 align=center><td colspan=2><input type=submit value='新 增 更 新'>　　　<input type=reset value='重新填写'></td></tr>
</form></table><%
end function

function news_addchk()
  dim topic
  topic=trim(request.form("topic"))
  if len(topic)<1 then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""更新标题 是必须要的！\n\n请返回输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
  else
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,3
    rs.addnew
    rs("sort")=nsort
    rs("username")=login_username
    rs("topic")=topic
    rs("word")=request.form("word")
    rs("tim")=now
    rs.update
    rs.close:set rs=nothing
    call upload_note("update",first_id("bbs_cast"))
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""成功新增了更新！\n\n点击返回。"");" & _
		   vbcrlf & "location='?nsort="&nsort&"'" & _
		   vbcrlf & "</script>")
  end if
end function

function news_edit(id)
  sql=sql&" where id="&id
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""您的操作有错误（error in edit）存在！\n\n点击返回。"");" & _
		   vbcrlf & "location='?nsort="&nsort&"'" & _
		   vbcrlf & "</script>")
  else
    dim msort:msort=rs("sort")
%><table border=0 width='98%' cellspacing=0 cellpadding=2>
<form name='add_frm' action='?action=editchk&id=<%response.write id%>' method=post>
<input type=hidden name=upid value=''>
  <tr><td colspan=2 align=center height=50><font class=red>修改更新</font></td></tr>
  <tr><td width='15%' align=center>更新标题：</td><td width='85%'><input type=text name=topic value='<%response.write rs("topic")%>' size=65 maxlength=50></td></tr>
  <tr><td height=30 align=center>更新类型：</td><td><input type=radio name=nsort value='news'<% if msort="news" then response.write "checked" %>>&nbsp;网站更新&nbsp;&nbsp;<input type=radio name=nsort value='forum'<% if msort="forum" then response.write "checked" %>>&nbsp;论坛公告</td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","word","&nbsp;&nbsp;")%></td></tr>
  <tr><td align=center>更新内空：</td><td><textarea name=word rows=15 cols=65><%response.write rs("word")%></textarea></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=f&uptext=word'></iframe></td></tr>
  <tr height=30 align=center><td colspan=2><input type=submit value='修 改 更 新'>　　　<input type=reset value='重新填写'></td></tr>
</form></table><%
  end if
  rs.close:set rs=nothing
end function

function news_editchk(id)
  dim topic:topic=trim(request.form("topic"))
  call upload_note("update",id)
  if len(topic)<1 then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""更新标题 是必须要的！\n\n请返回输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
  else
    set rs=server.createobject("adodb.recordset")
    sql=sql&" where id="&id
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
      response.write("<script language=javascript>" & _
		     vbcrlf & "alert(""您的操作有错误（error in editchk）存在！\n\n点击返回。"");" & _
		     vbcrlf & "location='?nsort="&nsort&"'" & _
		     vbcrlf & "</script>")
    else
      rs("sort")=nsort
      rs("topic")=topic
      rs("word")=request.form("word")
      rs.update
      rs.close:set rs=nothing
      response.write("<script language=javascript>" & _
		     vbcrlf & "alert(""成功修改了更新！\n\n点击返回。"");" & _
		     vbcrlf & "location='?nsort="&nsort&"'" & _
		     vbcrlf & "</script>")
    end if
  end if
end function
%>