<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nsort,rs2,sql2,id,j,sqladd,cid,sid,ncid,nsid,nid,now_id,power,pageurl,ispic
dim data_name,nummer,ddim,genre,os,rssum,thepages,page,viewpage,del_temp,csid
tit="<a href='?action='>�����б�</a>&nbsp;��&nbsp;" & _
    "<a href='?action=add'>�������</a>&nbsp;��&nbsp;" & _
    "<a href='admin_nsort.asp?nsort=down'>���ط���</a>&nbsp;��&nbsp;" & _
    "<a href='?action=code'>ע�����б�</a>&nbsp;��&nbsp;" & _
    "<a href='?action=code_add'>���ע����</a>"
response.write header(14,tit)
pageurl="?action="&action&"&":nsort="down":sqladd="":data_name="down":sqladd="":nummer=15
call admin_cid_sid()

if trim(request.querystring("del_ok"))="ok" then
  response.write del_select(request.form("del_id"))
end if

id=trim(request.querystring("id"))
if action="hidden" and isnumeric(id) then
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
  rs.close
  action=""
end if

select case action
case "add"
  call down_add()
case "down_edit"
  call down_edit()
case "code"
  call code_main()
case "code_add"
  call code_add()
case "code_edit"
  call code_edit()
case "code_del"
  call code_del()
case else
  call down_main()
end select

close_conn
response.write ender()

sub down_edit()
  dim sql3,rs3,id,name,sizes,url,url2,homepage,remark,counter,types,keyes,pic
  id=request.querystring("id")
  if not(isnumeric(id)) then call down_main():exit sub
%><table border=0 width=600 cellspacing=0 cellpadding=0>
<tr><td align=center height=300>
<%
  sql="select * from "&data_name&" where id="&id
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,3
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call down_main():exit sub
  end if
  if trim(request.querystring("types"))="edit" then
    csid=trim(request.form("csid"))
    name=code_admin(request.form("name"))
    sizes=code_admin(request.form("sizes"))
    url=code_admin(request.form("url"))
    url2=code_admin(request.form("url2"))
    pic=request.form("pic")
    if len(pic)<3 then pic="no_pic.gif"
    homepage=code_admin(request.form("homepage"))
    keyes=code_admin(request.form("keyes"))
    remark=request.form("remark")
    counter=trim(request.form("counter"))
    types=request.form("types")
    if len(csid)<1 or var_null(name)="" or var_null(url)="" then
      response.write("<font class=red_2>��������͡����ƺ����ص�ַ����Ϊ�գ�</font><br><br>"&go_back)
    else
      call chk_cid_sid()
      rs("c_id")=cid
      rs("s_id")=sid
      if trim(request.form("username_my"))="yes" then rs("username")=login_username
      if trim(request.form("hidden"))="yes" then
        rs("hidden")=false
      else
        rs("hidden")=true
      end if
      rs("name")=name
      rs("sizes")=sizes
      if isnumeric(trim(request.form("emoney"))) then
        rs("emoney")=trim(request.form("emoney"))
      else
        rs("emoney")=0
      end if
      rs("genre")=trim(request.form("genre"))
      rs("os")=replace(trim(request.form("os"))," ","")
      rs("power")=replace(replace(trim(request.form("power"))," ",""),",",".")
      rs("url")=url
      rs("url2")=url2
      rs("homepage")=homepage
      rs("remark")=remark
      rs("keyes")=keyes
      rs("pic")=pic
      rs("tim")=now_time
      rs("types")=types
      if isnumeric(counter) then rs("counter")=counter
      rs.update
      call upload_note(data_name,id)
      response.write "<font class=red>����޸ĳɹ���</font><br><br><a href='?c_id="&cid&"&s_id="&sid&"'>�������</a>" & _
		     vbcrlf&"<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=?c_id="&cid&"&s_id="&sid&"'>"
    end if
  else
    cid=int(rs("c_id")):sid=int(rs("s_id")):types=int(rs("types"))
    ispic=rs("pic"):pic=ispic
    if Instr(ispic,"/")>0 then ispic=right(ispic,len(ispic)-Instr(ispic,"/"))
    if Instr(ispic,".")>0 then ispic=left(ispic,Instr(ispic,".")-1)
    if ispic="no_pic" then ispic="n_"&id:pic=""
%><table border=0 width='98%' cellspacing=0 cellpadding=2>
  <tr><td colspan=2 height=50 align=center><font class=red>������������޸�</font></td></tr>
<form name='add_frm' action="?action=down_edit&types=edit&id=<%response.write id%>" method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>�������ƣ�</td><td width='85%'><input type=text name=name value='<%response.write rs("name")%>' size=70 maxlength=40><% response.write redx %></td></tr>
  <tr><td align=center>�������</td><td><%call chk_csid(cid,sid):call chk_emoney(rs("emoney")):call chk_h_u()%></td></tr>
  <tr><td align=center>����Ȩ�ޣ�</td><td><%call chk_power(rs("power"),0)%></td></tr>
  <tr><td align=center>�ļ���С��</td><td><input type=text name=sizes value='<%response.write rs("sizes")%>' size=20 maxlength=10>&nbsp;&nbsp;&nbsp;�Ƽ��ȼ���<select name=types size=1>
<option value='0'<% if types=0 then response.write " selected" %>>û�еȼ�</option>
<option value='1'<% if types=1 then response.write " selected" %>>һ�Ǽ�</option>
<option value='2'<% if types=2 then response.write " selected" %>>���Ǽ�</option>
<option value='3'<% if types=3 then response.write " selected" %>>���Ǽ�</option>
<option value='4'<% if types=4 then response.write " selected" %>>���Ǽ�</option>
<option value='5'<% if types=5 then response.write " selected" %>>���Ǽ�</option>
</select>&nbsp;&nbsp;&nbsp;�������ͣ�<select name=genre size=1><%
  dim tt:tt=rs("genre"):ddim=split(web_var(web_down,4),":")
  for i=0 to ubound(ddim)
    response.write vbcrlf&"<option"
    if tt=ddim(i) then response.write " selected"
    response.write ">"&ddim(i)&"</option>"
  next
  erase ddim
%></select></td></tr>
  <tr><td align=center>���������</td><td><%
  tt=rs("os"):ddim=split(web_var(web_down,3),":")
  for i=0 to ubound(ddim)
    response.write "<input type=checkbox name=os value='"&ddim(i)&"'"
    if instr(1,tt,ddim(i))>0 then response.write " checked"
    response.write " class=bg_1>"&ddim(i)
  next
  erase ddim
%></td></tr>
  <tr><td align=center>���ص�ַ1��</td><td><input type=text name=url value='<%response.write rs("url")%>' size=70 maxlength=200><% response.write redx %></td></tr>
  <tr><td align=center>���ص�ַ2��</td><td><input type=text name=url2 value='<%response.write rs("url2")%>' size=70 maxlength=200></td></tr>
  <tr><td align=center>�ļ����ԣ�</td><td><input type=text name=homepage value='<%response.write rs("homepage")%>' size=50 maxlength=50>&nbsp;&nbsp;&nbsp;���ش�����<input type=text name=counter value='<%response.write rs("counter")%>' size=4 maxlength=10></td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","remark","&nbsp;&nbsp;")%></td></tr>
  <tr><td align=center valign=top><br>���ֱ�ע��</td><td><%response.write("<textarea rows=6 name=remark cols=70>"&rs("remark")&"</textarea>")%></td></tr>
  <tr><td align=center>�� �� �֣�</td><td><input type=text name=keyes value='<%response.write rs("keyes")%>' size=12 maxlength=20>&nbsp;ͼƬ��<input type=test name=pic value='<% if ispic<>"no_pic.gif" then response.write pic %>' size=30 maxlength=100>&nbsp;<a href='upload.asp?uppath=down&upname=<%response.write ispic%>&uptext=pic' target=upload_frame>�ϴ�ͼƬ</a>&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=d&uptext=remark' target=upload_frame>�ϴ�������</a></td></tr>
  <tr><td align=center>�ϴ�ͼƬ��</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=down&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr height=30><td></td><td><input type=submit value=' �� �� �� �� '></td></tr>
</form></table><%
  end if
  rs.close:set rs=nothing
%></td></tr></table><%
end sub

sub down_add()
%><table border="0" width="600" cellspacing="0" cellpadding="0">
<tr><td align=center height=300><%
if request.querystring("types")="add" then
  dim name,sizes,url,url2,homepage,remark,types,keyes,pic
  csid=trim(request.form("csid"))
  name=code_admin(request.form("name"))
  sizes=code_admin(request.form("sizes"))
  url=code_admin(request.form("url"))
  url2=code_admin(request.form("url2"))
  homepage=code_admin(request.form("homepage"))
  keyes=code_admin(request.form("keyes"))
  remark=request.form("remark")
  pic=request.form("pic")
  if len(pic)<3 then pic="no_pic.gif"
  types=request.form("types")
  if len(csid)<1 or var_null(name)="" or var_null(url)="" then
    response.write("<font class=red_2>�ļ������͡����ƺ����ص�ַ����Ϊ�գ�</font><br><br>"&go_back)
  else
    call chk_cid_sid()
    sql="select * from "&data_name
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,3
    rs.addnew
    rs("c_id")=cid
    rs("s_id")=sid
    rs("username")=login_username
    rs("hidden")=true
    rs("name")=name
    rs("sizes")=sizes
    if isnumeric(trim(request.form("emoney"))) then
      rs("emoney")=trim(request.form("emoney"))
    else
      rs("emoney")=0
    end if
    rs("genre")=trim(request.form("genre"))
    rs("os")=replace(trim(request.form("os"))," ","")
    rs("power")=replace(replace(trim(request.form("power"))," ",""),",",".")
    rs("url")=url
    rs("url2")=url2
    rs("homepage")=homepage
    rs("remark")=remark
    rs("keyes")=keyes
    rs("pic")=pic
    rs("tim")=now_time
    rs("counter")=0
    rs("types")=types
    rs.update
    rs.close:set rs=nothing
    call upload_note(data_name,first_id(data_name))
    response.write "<font class=red>������ӳɹ���</font>&nbsp;<a href='?c_id="&cid&"&s_id="&sid&"'>�������</a><br><br><a href='?c_id="&cid&"&s_id="&sid&"&action=down_add'>����������</a>" & _
		   VbCrLf&"<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=?c_id="&cid&"&s_id="&sid&"&action=down_add'>"
  end if
else
%>
<table border=0 width='98%' cellspacing=0 cellpadding=2>
  <tr><td colspan=2 height=50 align=center><font class=red>������������</font></td></tr>
<form name='add_frm' action='?action=add&types=add' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>�������ƣ�</td><td width='85%'><input type=text name=name size=70 maxlength=40><% response.write redx %></td></tr>
  <tr><td align=center>�������</td><td><%call chk_csid(cid,sid):call chk_emoney(0)%></td></tr>
  <tr><td align=center>����Ȩ�ޣ�</td><td><%call chk_power("",1)%></td></tr>
  <tr><td align=center>�ļ���С��</td><td><input type=text name=sizes value='KB' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;�Ƽ��ȼ���<select name=types size=1>
<option value='0'>û�еȼ�</option>
<option value='1'>һ�Ǽ�</option>
<option value='2'>���Ǽ�</option>
<option value='3'>���Ǽ�</option>
<option value='4'>���Ǽ�</option>
<option value='5'>���Ǽ�</option>
</select>&nbsp;&nbsp;&nbsp;�������ͣ�<select name=genre size=1><%
  ddim=split(web_var(web_down,4),":")
  for i=0 to ubound(ddim)
    response.write vbcrlf&"<option>"&ddim(i)&"</option>"
  next
  erase ddim
%></select></td></tr>
  <tr><td align=center>���������</td><td><%
  ddim=split(web_var(web_down,3),":")
  for i=0 to ubound(ddim)
    response.write "<input type=checkbox name=os value='"&ddim(i)&"' class=bg_1>"&ddim(i)
  next
  erase ddim
%></td></tr>
  <tr><td align=center>��վ���أ�</td><td><input type=text name=url size=70 maxlength=200><% response.write redx %></td></tr>
  <tr><td align=center>�������أ�</td><td><input type=text name=url2 value='http://' size=70 maxlength=200></td></tr>
  <tr><td align=center>�ļ����ԣ�</td><td><input type=text name=homepage value='http://' size=50 maxlength=50></td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","remark","&nbsp;&nbsp;")%></td></tr>
  <tr><td valign=top align=center><br>���ֱ�ע</td><td><textarea rows=6 name=remark cols=70></textarea></td></tr>
<%ispic="d"&upload_time(now_time)%>
  <tr><td align=center>�� �� �֣�</td><td><input type=text name=keyes size=12 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ��<input type=text name=pic size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=<%response.write ispic%>&uptext=pic' target=upload_frame>�ϴ�ͼƬ</a>&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=d&uptext=remark' target=upload_frame>�ϴ�������</a></td></tr>
  <tr><td align=center>�ϴ��ļ���</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=down&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr height=30><td></td><td><input type=submit value=' �� �� �� �� '></td></tr>
</form></table><%
  end if
%></td></tr></table><%
end sub

sub code_del()
  id=trim(request.querystring("id"))
  if not(isnumeric(id)) then call code_main():exit sub
  conn.execute("delete from down_code where id="&id)
  call code_main()
end sub

sub code_edit()
  dim titler,rs,strsql
  if id="" or isnull(id) then call code_main():exit sub
%><table border="0" width="600" cellspacing="0" cellpadding="0">
<tr><td align=center height=300><%
  strsql="select * from down_code where id="&id
  set rs=server.createobject("adodb.recordset")
  rs.open strsql,conn,1,3
  if request("types")="edit" then
    dim name,username,code,remark
    name=code_form(trim(request("name")))
    username=code_form(trim(request("username")))
    code=code_form(trim(request("code")))
    remark=request("remark")
    if name="" or isnull(name) or code="" or isnull(code) then
      response.write("�������ƺ�ע �� �벻��Ϊ�գ�<br><br>"&go_back)
    else
      rs("name")=name
      rs("username")=username
      rs("code")=code
      rs("remark")=remark
      rs.update
      rs.close:set rs=nothing
      response.write("ע�����޸ĳɹ���<br><br><a href='admin_down.asp?action=code'>�������</a>")
      response.write(VbCrLf & "<meta http-equiv='refresh' content='" & time_go & "; url=admin_down.asp?action=code'>")
    end if
else
%>
<table border="0" width="400" cellspacing="0" cellpadding="2">
  <tr><td colspan=2 height=50 align=center><font class=font_color1>ע���������޸�</font></td></tr>
  <form action="?action=code_edit&types=edit&id=<%=id%>" method=post>
  <tr><td>�ļ�����</td><td><input type=text name=name value='<%=rs("name")%>' size=50 maxlength=100></td></tr>
  <tr><td>ע������</td><td><input type=text name=username value='<%=rs("username")%>' size=50 maxlength=100></td></tr>
  <tr><td>ע �� ��</td><td><input type=text name=code value='<%=rs("code")%>' size=50 maxlength=100></td></tr>
  <tr><td>��ע˵��</td><td><%response.write("<textarea rows=6 name=remark cols=50>"&rs("remark")&"</textarea>")%></td></tr>
  <tr height=30><td></td><td><input type="submit" value=" �� �� "></td></tr>
</form></table>
<%  end if%></td></tr><tr></table><%
end sub

sub code_add()
%><table border="0" width="600" cellspacing="0" cellpadding="0">
<tr><td align=center height=300><%
if request("types")="add" then
  dim name,username,code,remark
  name=code_form(trim(request("name")))
  username=code_form(trim(request("username")))
  code=code_form(trim(request("code")))
  remark=request("remark")
  if name="" or isnull(name) or code="" or isnull(code) then
    response.write("�ļ����ƺ�ע �� �벻��Ϊ�գ�<br><br>"&go_back)
  else
    dim rs,strsql
    strsql="select * from down_code where (id is null)"
    set rs=server.createobject("adodb.recordset")
    rs.open strsql,conn,1,3
    rs.addnew
    rs("name")=name
    rs("username")=username
    rs("code")=code
    rs("remark")=remark
    rs.update
    rs.close
    set rs=nothing
    response.write("ע������ӳɹ���<br><br><a href='admin_down.asp?action=code_add'>����������</a>")
    response.write(VbCrLf & "<meta http-equiv='refresh' content='" & time_go & "; url=admin_down.asp?action=code_add'>")
  end if
else
%><table border="0" width="400" cellspacing="0" cellpadding="2">
<form action="?action=code_add&types=add" method=post>
  <tr><td colspan=2 height=50 align=center><font class=font_color1>�����ע����</font></td></tr>
  <tr><td>�ļ�����</td><td><input type=text name=name size=50 maxlength=100></td></tr>
  <tr><td>ע������</td><td><input type=text name=username size=50 maxlength=100></td></tr>
  <tr><td>ע �� ��</td><td><input type=text name=code size=50 maxlength=100></td></tr>
  <tr><td>��ע˵��</td><td><textarea rows="6" name=remark cols="50"></textarea></td></tr>
  <tr height=30><td></td><td><input type="submit" value=" �� �� "></td></tr>
</form></table>
<%  end if%></td></tr></table><%
end sub

sub code_main()
%><table border=0 width='95%' cellspacing=0 cellpadding=2><%
  dim rs,strsql
  sqladd=""
  strsql="select * from down_code " & sqladd & "order by id desc"
  set rs=server.createobject("adodb.recordset")
  rs.open strsql,conn,1,1
  if rs.eof and rs.bof then
    rssum=0
  else
    rssum=rs.recordcount
  end if
  call format_pagecute()
  response.write "<tr><td colspan=3 align=center height=30>���и� <font class=red>"&rssum&"</font> ע����  ��ҳ:"&jk_pagecute(nummer,thepages,viewpage,pageurl,10,"#ff0000")&"</td></tr>" & _
	         "<tr align=center><td width='10%'>���</td><td width='75%'>���ͺ�����</td><td width='15%'>����</td></tr>"
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    response.write "<tr align=center><td>"&(viewpage-1)*nummer+i&"</td><td align=left>"&rs("name")&"</td><td><a href='admin_down.asp?action=code_edit&id="&rs("id")&"'>�޸�</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href='admin_down.asp?action=code_del&id="&rs("id")&"'>ɾ��</a></td></tr>"
    rs.movenext
  next
  rs.close:set rs=nothing
%></table><%
end sub

sub down_main()
%><table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr align=center height=300 valign=top>
<td width='20%' class=htd><br><%call left_sort()%></td>
<td width='80%'>
<table border=1 width='100%' cellspacing=0 cellpadding=1 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<script language=javascript src='STYLE/admin_del.js'></script>
<form name=del_form action='<%response.write pageurl%>del_ok=ok' method=post><%
  call sql_cid_sid()
  sql="select id,c_id,s_id,name,types,hidden from down"&sqladd&" order by tim desc"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
    rssum=0
  else
    rssum=int(rs.recordcount)
  end if
  call format_pagecute()
  del_temp=nummer
  if rssum=0 then del_temp=0
  if int(page)=int(thepages) then
    del_temp=rssum-nummer*(thepages-1)
  end if
%>
<tr><td colspan=3 align=center height=25>
����<font class=red><%response.write rssum%></font>�������<%response.write "<a href='?action=add&c_id="&cid&"&s_id="&sid&"'>�������</a>"%>
��<input type=checkbox name=del_all value=1 onClick=selectall('<%response.write del_temp%>')> ѡ�����С�<input type=submit value='ɾ����ѡ' onclick=""return suredel('<%response.write del_temp%>');"">
</td></tr>
<tr align=center bgcolor=#ffffff><td width='8%'>���</td><td width='77%'>���ͺ�����</td><td width='15%'>����</td></tr>
<%
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    now_id=rs("id"):nid=int(rs("types")):ncid=rs("c_id"):nsid=rs("s_id")
    response.write vbcrlf&"<tr align=center><td>"&(viewpage-1)*nummer+i&"</td><td align=left><a href='?action=down_edit&id="&now_id&"'>"&rs("name")&"</a></td><td align=right><a href='?action=hidden&c_id="&cid&"&s_id="&sid&"&id="&now_id&"&page="&viewpage&"'>"
    if rs("hidden")=true then
      response.write "��</a> "
    else
      response.write "<font class=red_2>��</font></a> "
    end if
    response.write "<font class=red>"&nid&"</font>&nbsp;��&nbsp;<input type=checkbox name=del_id value='"&now_id&"'></td></tr>"
    rs.movenext
  next
  rs.close:set rs=nothing
%></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>ҳ�Σ�<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font>
��ҳ��<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%>
</td></tr>
</table></td></tr></table><%
end sub
%>