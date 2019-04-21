<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim nummer,sqladd,page,rssum,thepages,viewpage,pageurl,del_temp,url,types,nname
tit="<a href='?'>管理上传文件</a> ┋ " & _
    "<a href='?types=1'>有效上传</a> ┋ " & _
    "<a href='?types=0'>无效上传</a>"
response.write header(9,tit)
nummer=15:rssum=0:thepages=0:viewpage=1
types=trim(request.querystring("types"))
if not(isnumeric(types)) then types=-1
select case int(types)
case 0
  nname="无效上传"
  pageurl="?types=0&"
case 1
  nname="有效上传"
  pageurl="?types=1&"
case else
  types=-1
  nname="所有"
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
    response.write "<script language=javascript>alert(""共删除了 "&del_num+1&" 个文件！"");</script>"
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
现有 <font class=red><%response.write rssum%></font> 个 <font class=red_3><%response.write nname%></font> 文件 ┋ 页次：<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font>
　<input type=checkbox name=del_all value=1 onClick="javascript:selectall('<%response.write del_temp%>');"> 选中所有　<input type=submit value='删除所选' onclick="return suredel('<%response.write del_temp%>');"></td></tr>
<tr align=center height=18 bgcolor=<%response.write color3%>>
<td width='5%'>序号</td>
<td width='31%'>文件名</td>
<td width='5%'>类型</td>
<td width='10%'>大小(B)</td>
<td width='12%'>栏目、ID</td>
<td width='14%'>作者</td>
<td width='18%'>时间</td>
<td width='5%'>操作</td>
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
  response.write "<font alt='ID："&rs("iid")&"'>"&nsortn&"</font>"
else
  response.write "<font class=red_2>无效</font>"
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
<td colspan=8>分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,6,"#ff0000")%></td>
</tr>
</table>
<%
end sub
%>