<!-- #include file="include/onlogin.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nummer,sqladd,page,rssum,thepages,viewpage,pageurl,del_temp,url,types,nname
tit="<a href='?'>�����ϴ��ļ�</a> �� " & _
    "<a href='?types=1'>��Ч�ϴ�</a> �� " & _
    "<a href='?types=0'>��Ч�ϴ�</a>"
response.write header(9,tit)
nummer=15:rssum=0:thepages=0:viewpage=1
types=trim(request.querystring("types"))
if not(isnumeric(types)) then types=-1
select case int(types)
case 0
  nname="��Ч�ϴ�"
  pageurl="?types=0&"
case 1
  nname="��Ч�ϴ�"
  pageurl="?types=1&"
case else
  types=-1
  nname="����"
  pageurl="?"
end select

if trim(request("del_ok"))="ok" then
  call del_select(trim(request.form("del_id")))
end if

call upload_main()
response.write ender()

sub del_select(delid)
  'on error resume next
  dim del_i,del_num,del_dim,del_sql
  if delid<>"" and not isnull(delid) then
    delid=replace(delid," ","")
    del_dim=split(delid,",")
    del_num=UBound(del_dim)
    for del_i=0 to del_num
      del_sql="select url from upload where id="&del_dim(del_i)
      set rs=conn.execute(del_sql)
      if not(rs.eof and rs.bof) then
        call del_file(rs("url"))
      end if
      rs.close
      del_sql="delete from upload where id="&del_dim(del_i)
      conn.execute(del_sql)
    next
    Erase del_dim
    response.write "<script language=javascript>alert(""��ɾ���� "&del_num+1&" ���ļ���"");</script>"
  end if
end sub

sub upload_main()
  dim ntypes,upload_path,nnsort,nsortn
  upload_path=web_var(web_upload,1)
%>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=1 width='100%' cellspacing=0 cellpadding=1<%response.write table1%>>
<form name=del_form action='<%response.write pageurl%>del_ok=ok' method=post>
<%
  if types>-1 then sqladd=" where types="&types
  sql="select * from upload"&sqladd&" order by id desc"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if not(rs.eof and rs.bof) then rssum=rs.recordcount
  call format_pagecute()
  del_temp=nummer
  if rssum=0 then del_temp=0
  if int(page)=int(thepages) then del_temp=rssum-nummer*(thepages-1)
%>
<tr bgcolor=<%response.write color1%>><td colspan=8 align=center height=30>
���� <font class=red><%response.write rssum%></font> �� <font class=red_3><%response.write nname%></font> �ļ� �� ҳ�Σ�<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font>
��<input type=checkbox name=del_all value=1 onClick="javascript:selectall('<%response.write del_temp%>');"> ѡ�����С�<input type=submit value='ɾ����ѡ' onclick="return suredel('<%response.write del_temp%>');"></td></tr>
<tr align=center height=18 bgcolor=<%response.write color3%>>
<td width='5%'>���</td>
<td width='31%'>�ļ���</td>
<td width='5%'>����</td>
<td width='10%'>��С(B)</td>
<td width='12%'>��Ŀ��ID</td>
<td width='14%'>����</td>
<td width='18%'>ʱ��</td>
<td width='5%'>����</td>
</tr>
<%
  if int(viewpage)>1 then rs.move (viewpage-1)*nummer
  for i=1 to nummer
    if rs.eof then exit for
    url=rs("url"):ntypes=rs("types"):nnsort=rs("nsort")
%>
<tr align=center<%response.write mtr%>>
<td><%response.write (viewpage-1)*nummer+i%>.</td>
<td align=left><a href='<%response.write url_true(upload_path,url)%>' target=_blank><%response.write url%></a></td>
<td><%response.write rs("genre")%></td>
<td align=left><%response.write rs("sizes")%></td>
<td><%
if int(ntypes)<>0 then
  nsortn=format_menu(nnsort)
  if len(nsortn)<1 then nsortn=nnsort
  response.write "<font alt='ID��"&rs("iid")&"'>"&nsortn&"</font>"
else
  response.write "<font class=red_2>��Ч</font>"
end if
%></td>
<td><%response.write format_user_view(rs("username"),1,"")%></td>
<td><%response.write time_type(rs("tim"),7)%></td>
<td><input type=checkbox name=del_id value='<%response.write rs("id")%>'></td>
</tr>
<%
    rs.movenext
  next
  rs.close:set rs=nothing
%></form>
<tr>
<td colspan=8>��ҳ��<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,6,"#ff0000")%></td>
</tr>
</table>
<%
end sub
%>