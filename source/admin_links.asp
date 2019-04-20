<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #INCLUDE file="INCLUDE/jk_page_cute.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

'	fir	sec	txt
dim id,sort,rssum,nummer,thepages,viewpage,pageurl,page
id=trim(request.querystring("id"))
sort=trim(request.querystring("sort"))
tit="<a href='?'>友情链接</a>┋" & _
    "<a href='?action=main&sort=fir'>首页链接</a>┋" & _
    "<a href='?action=main&sort=sec'>内页链接</a>┋" & _
    "<a href='?action=main&sort=txt'>文字链接</a>┋" & _
    "<a href='?action=add'>新增链接</a>┋" & _
    "<a href='?action=list'>重新排序</a>"
response.write header(17,tit)

select case action
case "list"
  call links_list()
case "add"
  response.write links_add()
case "addchk"
  response.write links_addchk()
case "order"
  if isnumeric(id) and ( trim(request.querystring("actiones"))="up" or trim(request.querystring("actiones"))="down" ) then
    response.write links_order(id)
  else
    response.write links_main()
  end if
case "del"
  if isnumeric(id) then
    response.write links_del(id)
  else
    response.write links_main()
  end if
case "hidden"
  if isnumeric(id) then
    response.write links_hidden(id)
  else
    response.write links_main()
  end if
case "edit"
  if isnumeric(id) then
    response.write links_edit(id)
  else
    response.write links_main()
  end if
case "editchk"
  if isnumeric(id) then
    response.write links_editchk(id)
  else
    response.write links_main()
  end if
case else
  response.write links_main()
end select

response.write ender()

sub links_list()
  dim rssum,i
  set rs=server.createobject("adodb.recordset")
  sql="select * from links where sort='fir' order by orders,id"
  rs.open sql,conn,1,3
  if rs.eof and rs.bof then
    rssum=0
  else
    rssum=rs.recordcount
  end if
  for i=1 to rssum
    rs("orders")=i
    rs.update
    rs.movenext
  next
  rs.close
  rssum=0
  sql="select * from links where sort='sec' order by orders,id"
  rs.open sql,conn,1,3
  if rs.eof and rs.bof then
    rssum=0
  else
    rssum=rs.recordcount
  end if
  for i=1 to rssum
    rs("orders")=i
    rs.update
    rs.movenext
  next
  rs.close
  rssum=0
  sql="select * from links where sort='txt' order by orders,id"
  rs.open sql,conn,1,3
  if rs.eof and rs.bof then
    rssum=0
  else
    rssum=rs.recordcount
  end if
  for i=1 to rssum
    rs("orders")=i
    rs.update
    rs.movenext
  next
  rs.close:set rs=nothing
  response.write links_main()
end sub

function links_order(id)
  dim action,sort,tmp_id_1,tmp_id_2,tmp_order_1,tmp_order_2,sqladd,update_ok
  action=trim(request.querystring("actiones"))
  update_ok="no":sort="no"
  if action="up" then
    sqladd=" desc"
  else
    sqladd=""
  end if
  
  sql="select sort from links where id="&id
  set rs=conn.execute(sql)
  if not rs.eof or not rs.bof then
    sort=rs("sort")
  end if
  rs.close:set rs=nothing
  
  if sort<>"no" then
    sql="select * from links where sort='"&sort&"' order by orders"&sqladd
    set rs=conn.execute(sql)
    do while not rs.eof
      if int(rs("id"))=int(id) then
        tmp_id_1=id
        tmp_order_1=rs("orders")
        rs.movenext
        if not rs.eof then
          tmp_id_2=rs("id")
          tmp_order_2=rs("orders")
          update_ok="yes"
          exit do
        end if
        exit do
      end if
      rs.movenext
    loop
    rs.close:set rs=nothing
  end if
  
  if update_ok="yes" then
    sql="update links set orders="&tmp_order_2&" where id="&tmp_id_1
    conn.execute(sql)
    sql="update links set orders="&tmp_order_1&" where id="&tmp_id_2
    conn.execute(sql)
  end if
  
  response.redirect request.servervariables("http_referer")
end function

function links_del(id)
  on error resume next
  conn.execute("delete from links where id="&id)
  if err then
    err.clear
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""您的操作有错误（error in del）存在！\n\n点击返回。"");" & _
		   vbcrlf & "location='?action=main&sort="&sort&"'" & _
		   vbcrlf & "</script>")
  else
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""成功删除了一条友情链接！\n\n点击返回。"");" & _
		   vbcrlf & "location='?action=main&sort="&sort&"'" & _
		   vbcrlf & "</script>")
  end if
end function

function links_hidden(id)
  dim hid,hh:hh="no"
  set rs=conn.execute("select hidden from links where id="&id)
  if rs.eof and rs.bof then
    '
  else
    hid=rs("hidden")
    hh="yes"
  end if
  rs.close:set rs=nothing
  if hh="yes" then
    if hid=true then
      hid=0
    else
      hid=1
    end if
    conn.execute("update links set hidden="&hid&" where id="&id)
  end if
  
  response.redirect request.servervariables("http_referer")
end function

function links_main()
  dim i,sort,sqladd,sname,iid
  pageurl="?"
  sort=trim(request.querystring("sort"))
  if sort="fir" or sort="sec" or sort="txt" then
    sqladd=" where sort='"&sort&"'"
    pageurl=pageurl&"sort="&sort&"&"
    select case sort
    case "fir"
      sname="首页"
    case "sec"
      sname="内页"
    case "txt"
      sname="文字"
    end select
  end if
  sql="select * from links"&sqladd&" order by orders,id"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
    links_main="现在还没有！"
  else
    rssum=rs.recordcount
    nummer=8
    call format_pagecute()

    links_main=links_main&vbcrlf&"<script language=JavaScript><!--" & _
	      vbcrlf&"function Do_del_data(data1)" & _
	      vbcrlf&"{" & _
	      vbcrlf&"if (confirm(""此操作将删除id为 ""+data1+"" 的友情链接！\n真的要删除吗？\n删除后将无法恢复！""))" & _
	      vbcrlf&"  window.location="""&pageurl&"action=del&id=""+data1" & _
	      vbcrlf&"}" & _
	      vbcrlf&"//--></script>" & _
	      vbcrlf&"<table border=1 width=500 cellspacing=0 cellpadding=1 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>" & _
	      vbcrlf&"<tr><td colspan=4 align=center height=30><table border=0 width='100%'cellspacing=0 cellpadding=0>" & _
	      vbcrlf&"<tr align=center><td width='40%'>现在有 <font class=red>"&rssum&"</font> 个 <font class=red_4>"&sname&"</font> 链接</td>" & _
	      vbcrlf&"<td width='60%'>"&pagecute_fun(viewpage,thepages,pageurl)&"</td></tr></table></td></tr>" & _
	      "<tr align=center bgcolor=#ededed><td width='8%'>序号</td><td width='20%'>LOGO</td><td width='35%'>网站名称</td><td width='37%'>操作</td></tr>"
    if int(viewpage)>1 then
      rs.move (viewpage-1)*nummer
    end if
    pageurl=pageurl&"page="&viewpage&"&"
    for i=1 to nummer
      if rs.eof then exit for
      iid=rs("id")
      links_main=links_main&vbcrlf&"<tr align=center height=40><td>"&i+(viewpage-1)*nummer&".</td><td>"
      if rs("sort")="txt" then
        links_main=links_main&"txt"
      else
        links_main=links_main&"<img src='"&rs("pic")&"' width=88 height=31 border=0>"
      end if
      links_main=links_main&"</td><td><a href='"&rs("url")&"' target=_blank>"&code_html(rs("nname"),1,12)&"</a></td><td>"
      if rs("hidden")=true then
        links_main=links_main&"<a href='"&pageurl&"action=hidden&id="&iid&"'>显示</a>┋"
      else
        links_main=links_main&"<a href='"&pageurl&"action=hidden&id="&iid&"'><font class=red_2>隐藏</font></a>┋"
      end if
      links_main=links_main&"<a href='"&pageurl&"action=order&actiones=up&id="&iid&"'>向上</a>┋<a href='"&pageurl&"action=order&actiones=down&id="&iid&"'>向下</a>┋<a href='"&pageurl&"action=edit&id="&iid&"'>修改</a>┋<a href='javascript:Do_del_data("&iid&")'>删除</a></td></tr>"
      rs.movenext
    next
    links_main=links_main&vbcrlf&"</table>"
  end if
  rs.close:set rs=nothing
end function

function links_add()
%><table border=0 width=450 cellspacing=0 cellpadding=2>
<form action='admin_links.asp?action=addchk' method=post>
  <tr>
    <td colspan=2 align=center height=50><font class=red>新增链接</font></td>
  </tr>
  <tr height=30>
    <td width='20%'>链接类型：</td>
    <td width='80%'><input type=radio name=sort value='fir' checked>首页链接
    <input type=radio name=sort value='sec'>内页链接
    <input type=radio name=sort value='txt'>文字链接</td>
  </tr>
  <tr height=30>
    <td>网站名称：</td>
    <td><input type=text name=nname size=50 maxlength=20></td>
  </tr>
  <tr height=30>
    <td>链接地址：</td>
    <td><input type=text name=url value='http://' size=50 maxlength=100></td>
  </tr>
  <tr height=30>
    <td>链接LOGO：</td>
    <td><input type=text name=pic value='images/links/' size=60 maxlength=100></td>
  </tr>
  <tr height=30 align=center>
    <td colspan=2><input type=submit value='新 增 链 接'></td>
  </tr>
</form></table><%
end function

function links_addchk()
  dim nname,orders
  nname=trim(request.form("nname"))
  sort=trim(request.form("sort"))
  if len( nname)<1 or ( sort="fir" and sort="sec" and sort="txt" ) then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""网站名称 和 链接类型 是必须要的！\n\n请返回输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
  else
    set rs=server.createobject("adodb.recordset")
    sql="select top 1 orders from links where sort='"&sort&"' order by orders desc"
    rs.open sql,conn,1,1
    if rs.eof and rs.bof then
      orders=0
    else
      orders=int(rs("orders"))
    end if
    rs.close
    orders=int(orders)+1
    
    sql="select * from links"
    rs.open sql,conn,1,3
    rs.addnew
    rs("orders")=orders
    rs("sort")=sort
    rs("nname")=nname
    rs("url")=trim(request.form("url"))
    rs("pic")=trim(request.form("pic"))
    rs("hidden")=true
    rs.update
    rs.close:set rs=nothing
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""成功新增了链接！\n\n点击返回。"");" & _
		   vbcrlf & "location='?action=main&sort="&sort&"'" & _
		   vbcrlf & "</script>")
  end if
end function

function links_edit(id)
  dim sss
  sql="select * from links where id="&id
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""您的操作有错误（error in edit）存在！\n\n点击返回。"");" & _
		   vbcrlf & "location='?action=main&sort="&sort&"'" & _
		   vbcrlf & "</script>")
  else
    sss=rs("sort")
%><table border=0 width=450 cellspacing=0 cellpadding=2>
<form action='admin_links.asp?action=editchk&id=<%response.write id%>' method=post>
  <tr>
    <td colspan=2 align=center height=50><font class=red>修改链接</font></td>
  </tr>
  <tr height=30>
    <td width='20%'>链接类型：</td>
    <td width='80%'><input type=radio name=sort value='fir'<% if sss="fir" then response.write " checked" %>>首页链接
    <input type=radio name=sort value='sec'<% if sss="sec" then response.write " checked" %>>内页链接
    <input type=radio name=sort value='txt'<% if sss="txt" then response.write " checked" %>>文字链接</td>
  </tr>
  <tr height=30>
    <td>网站名称：</td>
    <td><input type=text name=nname value='<%response.write rs("nname")%>' size=50 maxlength=20></td>
  </tr>
  <tr height=30>
    <td>链接地址：</td>
    <td><input type=text name=url value='<%response.write rs("url")%>' size=50 maxlength=100></td>
  </tr>
  <tr height=30>
    <td>链接LOGO：</td>
    <td><input type=text name=pic value='<%response.write rs("pic")%>' size=60 maxlength=100></td>
  </tr>
  <tr height=30 align=center>
    <td colspan=2><input type=submit value='修 改 链 接'></td>
  </tr>
</form></table><%
  end if
  rs.close:set rs=nothing
end function

function links_editchk(id)
  dim nname
  nname=trim(request.form("nname"))
  sort=trim(request.form("sort"))
  if len( nname)<1 or ( sort="fir" and sort="sec" and sort="txt" ) then
    response.write("<script language=javascript>" & _
		   vbcrlf & "alert(""网站名称 和 链接类型 是必须要的！\n\n请返回输入。"");" & _
		   vbcrlf & "history.back(1)" & _
		   vbcrlf & "</script>")
  else
    set rs=server.createobject("adodb.recordset")
    sql="select * from links where id="&id
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
      response.write("<script language=javascript>" & _
		     vbcrlf & "alert(""您的操作有错误（error in editchk）存在！\n\n点击返回。"");" & _
		     vbcrlf & "location='?action=main&sort="&sort&"'" & _
		     vbcrlf & "</script>")
    else
      rs("sort")=sort
      rs("nname")=nname
      rs("url")=trim(request.form("url"))
      rs("pic")=trim(request.form("pic"))
      rs.update
      rs.close:set rs=nothing
      response.write("<script language=javascript>" & _
		     vbcrlf & "alert(""成功修改了链接！\n\n点击返回。"");" & _
		     vbcrlf & "location='?action=main&sort="&sort&"'" & _
		     vbcrlf & "</script>")
    end if
  end if
end function
%>