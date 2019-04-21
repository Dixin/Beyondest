<!-- #INCLUDE file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim nsort,nsortn,jk_an
tit=vbcrlf & "<a href='?nsort=art'>文栏分类</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='?nsort=down'>下载分类</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='?nsort=news'>新闻分类</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='?nsort=paste'>壁纸分类</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='?nsort=film'>视频分类</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='?nsort=flash'>Flash分类</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='?nsort=web'>网站分类</a>"
response.write header(5,tit)
nsort=trim(request.querystring("nsort"))
action=trim(request.querystring("action"))
select case nsort
case "down"
  nsortn="下载分类"
case "news"
  nsortn="新闻分类"
case "web"
  nsortn="网站分类"
case "gall"
  nsortn="图库分类"
case "film"
  nsortn="视频分类"
case "flash"
  nsortn="Flash分类"
case "baner"
  nsortn="相册分类"
case "paste"
  nsortn="壁纸分类"
case else
  nsort="art"
  nsortn="文栏分类"
end select

select case action
case "up","down"
  jk_an="分类查看"
  call jk_order()
case "del"
  jk_an="分类查看"
  call jk_del()
case "list"
  jk_an="分类查看"
  call jk_list()
case "addc"
  jk_an="添加一级分类"
  call jk_addc()
case "adds"
  jk_an="添加二级分类"
  call jk_adds()
case "editc"
  jk_an="修改一级分类"
  call jk_editc()
case "edits"
  jk_an="修改二级分类"
  call jk_edits()
case else
  jk_an="分类查看"
  call jk_main()
end select

response.write ender()

sub jk_list()
  dim i,j,cid,sql2,rs2:i=1
  sql="select c_id from jk_class where nsort='"&nsort&"' order by c_order,c_id"
  set rs=conn.execute(sql)
  do while not rs.eof
    cid=rs(0):j=1
    conn.execute("update jk_class set c_order="&i&" where c_id="&cid)
    sql2="select s_id from jk_sort where c_id="&cid&" order by s_order,s_id"
    set rs2=conn.execute(sql2)
    do while not rs2.eof
      conn.execute("update jk_sort set s_order="&j&" where s_id="&rs2(0))
      rs2.movenext
      j=j+1
    loop
    rs2.close
    rs.movenext
    i=i+1
  loop
  rs.close:set rs=nothing:set rs2=nothing
  call jk_main()
end sub

sub jk_del()
  dim cid,sid
  cid=trim(request.querystring("c_id")):sid=trim(request.querystring("s_id"))
  if not(isnumeric(cid)) and not(isnumeric(sid)) then call jk_main():exit sub
  if isnumeric(cid) then sid=""
  if sid="" then
    sql="delete from jk_class where c_id="&cid
    conn.execute(sql)
    sql="delete from jk_sort where c_id="&cid
    conn.execute(sql)
  else
    sql="delete from jk_sort where s_id="&sid
    conn.execute(sql)
  end if
  call jk_main()
end sub

sub jk_order()
  dim cid,sid,nid,t1,t11,t2,t22,sqladd:sqladd=""
  cid=trim(request.querystring("c_id")):sid=trim(request.querystring("s_id"))
  if not(isnumeric(cid)) and not(isnumeric(sid)) then call jk_main():exit sub
  if isnumeric(cid) then sid=""
  if action="up" then sqladd=" desc"
  if sid="" then
    t1=int(cid)
    sql="select c_id,c_order from jk_class where nsort='"&nsort&"' order by c_order"&sqladd&",c_id"&sqladd
    set rs=conn.execute(sql)
    do while not rs.eof
      nid=int(rs(0))
      if int(cid)=nid then
        t22=rs(1)
        rs.movenext
        if rs.eof then exit do
        t2=rs(0):t11=rs(1)
        conn.execute("update jk_class set c_order="&t11&" where c_id="&t1)
        conn.execute("update jk_class set c_order="&t22&" where c_id="&t2)
        exit do
      end if
      rs.movenext
    loop
    
rs.close:set rs=nothing
  else
    t1=int(sid)
    sql="select jk_sort.c_id from jk_class inner join jk_sort on jk_class.c_id=jk_sort.c_id where jk_sort.s_id="&sid
    set rs=conn.execute(sql)
    if rs.eof and rs.bof then
      rs.close:set rs=nothing
      call jk_main():exit sub
    end if
    cid=int(rs(0))
    
rs.close
    sql="select s_id,s_order from jk_sort where c_id="&cid&" order by s_order"&sqladd&",s_id"&sqladd
    set rs=conn.execute(sql)
    do while not rs.eof
      nid=int(rs(0))
      if int(sid)=nid then
        t22=rs(1)
        rs.movenext
        if rs.eof then exit do
        t2=rs(0):t11=rs(1)
        conn.execute("update jk_sort set s_order="&t11&" where s_id="&t1)
        conn.execute("update jk_sort set s_order="&t22&" where s_id="&t2)
        exit do
      end if
      rs.movenext
    loop
    
rs.close:set rs=nothing
  end if
  call jk_main()
end sub

sub jk_editc()
  dim c_name,cid
  cid=trim(request.querystring("c_id"))
  if not(isnumeric(cid)) then call jk_main():exit sub
  sql="select c_name from jk_class where nsort='"&nsort&"' and c_id="&cid
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,3
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call jk_main():exit sub
  end if
  response.write jk_tit()&"<table border=1 width=350 cellspacing=0 cellpadding=2 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"
  if trim(request.querystring("edit"))="ok" then
    response.write vbcrlf&"<tr><td height=100 align=center>"
    c_name=replace(trim(request.form("c_name")),"'","")
    if var_null(c_name)="" or len(c_name)>16 then
      response.write "<font class=red_2>一级分类名称不能为空（长度不大于16）！</font><br><br>"&go_back
    else
      rs("c_name")=c_name
      rs.update
      response.write "<font class=red_3>修改一级分类成功！</font><br><br><a href='?nsort="&nsort&"'>点击返回</a>"
    end if
    response.write vbcrlf&"</td></tr>"
  else
%><form action='?nsort=<% response.write nsort %>&action=editc&c_id=<% response.write cid %>&edit=ok' method=post>
<tr height=50 align=center>
<td>一级分类名称：</td>
<td><input type=text name=c_name value='<% response.write rs(0) %>' size=30 maxlength=16></td>
</tr>
<tr><td colspan=2 height=50 align=center><input type=submit value='修改一级分类'></td></tr>
</form><%
  end if
  rs.close:set rs=nothing
  response.write "</table>"
end sub

sub jk_edits()
  dim s_name,pic,s_order,intro,sid,cid,ccid,ncid,sqladd
  sqladd=""
  sid=trim(request.querystring("s_id"))
  if not(isnumeric(sid)) then sid=0
  sql="select c_id,s_name,pic,intro from jk_sort where s_id="&sid
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call jk_main():exit sub
  end if
  cid=rs(0)
  response.write jk_tit()&"<table border=1 width=500 cellspacing=0 cellpadding=2 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"
  if trim(request.querystring("edit"))="ok" then
    response.write vbcrlf&"<tr><td height=100 align=center>"
    ccid=trim(request.form("c_id"))
    s_name=replace(trim(request.form("s_name")),"'","")
    pic=replace(trim(request.form("pic")),"'","")
    intro=replace(trim(request.form("intro")),"'","")
    if len(s_name)<1 or len(s_name)>16 then
      response.write "<font class=red_2>二级分类名称不能为空（长度不大于16）！</font><br><br>"&go_back
    else
      if int(ccid)<>int(cid) then
        rs.close
        sql="select top 1 s_order from jk_sort where c_id="&ccid&" order by s_order desc"
        set rs=conn.execute(sql)
        if rs.eof and rs.bof then
          s_order=1
        else
          s_order=int(rs(0))+1
        end if
        sqladd=",s_order="&s_order
      end if
      sql="update jk_sort set intro='"&intro&"',pic='"&pic&"',c_id="&ccid&",s_name='"&s_name&"'"&sqladd&" where s_id="&sid
      conn.execute(sql)
      response.write "<font class=red_3>修改二级分类成功！</font><br><br><a href='?nsort="&nsort&"'>点击返回</a>"
    end if
    response.write vbcrlf&"</td></tr>"
  else
%><form action='?nsort=<% response.write nsort %>&action=edits&s_id=<% response.write sid %>&edit=ok' method=post>
<tr height=30 align=center>
<td width=100>一级分类类型：</td>
<td><select name=c_id size=1><%
  pic=rs(2)
  intro=rs(3)
  s_name=rs(1):rs.close
  sql="select c_id,c_name from jk_class where nsort='"&nsort&"' order by c_order,c_id"
  set rs=conn.execute(sql)
  do while not rs.eof
    ncid=int(rs(0))
    response.write vbcrlf&"<option value='"&ncid&"'"
    if cid=ncid then response.write " selected"
    response.write ">"&rs(1)&"</option>"
    rs.movenext
  loop
  
%>
</select></td>
</tr>
<tr height=30 align=center>
<td>二级分类名称：</td>
<td><input type=text name=s_name value='<% response.write s_name %>' size=30 maxlength=16></td>
</tr>
<tr height=30 align=center>
<td>二级分类图片：</td>
<td><input type=text name=pic value='<% response.write pic %>' size=30 maxlength=16></td>
</tr>
<tr height=30 align=center>
<td>二级分类介绍：</td>
<td><textarea rows=6 name=intro cols=70 value=''><% response.write intro %></textarea></td>
</tr>
<tr><td colspan=2 height=50 align=center><input type=submit value='修改二级分类'></td></tr>
</form><%
  end if
  rs.close:set rs=nothing
  response.write "</table>"
end sub

sub jk_addc()
  dim c_name,c_order
  response.write jk_tit()&"<table border=1 width=350 cellspacing=0 cellpadding=2 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"
  if trim(request.querystring("add"))="ok" then
    response.write vbcrlf&"<tr><td height=100 align=center>"
    c_name=replace(trim(request.form("c_name")),"'","")
    if var_null(c_name)="" or len(c_name)>16 then
      response.write "<font class=red_2>一级分类名称不能为空（长度不大于16）！</font><br><br>"&go_back
    else
      sql="select top 1 c_order from jk_class where nsort='"&nsort&"' order by c_order desc"
      set rs=conn.execute(sql)
      if rs.eof and rs.bof then
        c_order=1
      else
        c_order=int(rs(0))+1
      end if
      rs.close:set rs=nothing
      sql="insert into jk_class(nsort,c_name,c_order) values('"&nsort&"','"&c_name&"',"&c_order&")"
      conn.execute(sql)
      response.write "<font class=red_3>添加一级分类成功！</font><br><br><a href='?nsort="&nsort&"'>点击返回</a>"
    end if
    response.write vbcrlf&"</td></tr>"
  else
%><form action='?nsort=<% response.write nsort %>&action=addc&add=ok' method=post>
<tr height=50 align=center>
<td>一级分类名称：</td>
<td><input type=text name=c_name size=30 maxlength=16></td>
</tr>
<tr><td colspan=2 height=50 align=center><input type=submit value='添加一级分类'></td></tr>
</form><%
  end if
  response.write "</table>"
end sub

sub jk_adds()
  dim s_name,s_order,cname,cid,ncid
  cid=trim(request.querystring("c_id"))
  if not(isnumeric(cid)) then cid=0
  cid=int(cid)
  response.write jk_tit()&"<table border=1 width=350 cellspacing=0 cellpadding=2 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"
  if trim(request.querystring("add"))="ok" then
    response.write vbcrlf&"<tr><td height=100 align=center>"
    s_name=replace(trim(request.form("s_name")),"'","")
    if len(s_name)<1 or len(s_name)>16 then
      response.write "<font class=red_2>二级分类名称不能为空（长度不大于16）！</font><br><br>"&go_back
    else
      cid=trim(request.form("c_id"))
      if not(isnumeric(cid)) then cid=0
      sql="select c_name from jk_class where nsort='"&nsort&"' and c_id="&cid
      set rs=conn.execute(sql)
      if rs.eof and rs.bof then
        rs.close:set rs=nothing
        call jk_main():exit sub
      end if
      cname=rs(0)
      rs.close
      
      sql="select top 1 s_order from jk_sort where c_id="&cid&" order by s_order desc"
      set rs=conn.execute(sql)
      if rs.eof and rs.bof then
        s_order=1
      else
        s_order=int(rs(0))+1
      end if
      rs.close:set rs=nothing
      
      sql="insert into jk_sort(c_id,s_name,s_order) values("&cid&",'"&s_name&"',"&s_order&")"
      conn.execute(sql)
      response.write "<font class=red_3>添加二级分类成功！</font><br><br><a href='?nsort="&nsort&"'>点击返回</a>"
    end if
    response.write vbcrlf&"</td></tr>"
  else
%><form action='?nsort=<% response.write nsort %>&action=adds&c_id=<% response.write cid %>&add=ok' method=post>
<tr height=30 align=center>
<td>一级分类类型：</td>
<td><select name=c_id size=1><%
  sql="select c_id,c_name from jk_class where nsort='"&nsort&"' order by c_order"
  set rs=conn.execute(sql)
  do while not rs.eof
    ncid=int(rs(0))
    response.write vbcrlf&"<option value='"&ncid&"'"
    if cid=ncid then response.write " selected"
    response.write ">"&rs(1)&"</option>"
    rs.movenext
  loop
  rs.close:set rs=nothing
%>
</select></td>
</tr>
<tr height=30 align=center>
<td>二级分类名称：</td>
<td><input type=text name=s_name size=30 maxlength=16></td>
</tr>
<tr><td colspan=2 height=50 align=center><input type=submit value='添加二级分类'></td></tr>
</form><%
  end if
  response.write "</table>"
end sub

sub jk_main()
  response.write jk_tit()
  dim sql2,rs2,cid,sid
  response.write vbcrlf&"<table border=1 cellspacing=0 cellpadding=2 width=400 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"
  sql="select c_id,c_name from jk_class where nsort='"&nsort&"' order by c_order,c_id"
  set rs=conn.execute(sql)
  do while not rs.eof
    cid=rs(0)
    response.write vbcrlf&"<tr bgcolor=#ffffff align=center><td align=left>&nbsp;<font class=red_3><b>"&img_small("jt1")&rs(1)&"</b></font>&nbsp;&nbsp;（<a href='?nsort="&nsort&"&action=adds&c_id="&cid&"'>添加二级分类</a>）</td><td><a href='?nsort="&nsort&"&action=editc&c_id="&cid&"'>修改</a>&nbsp;&nbsp;<a href=""javascript:Do_del_class('"&cid&"');"">删除</a></td><td>排序：<a href='?nsort="&nsort&"&action=up&c_id="&cid&"'>向上</a>&nbsp;&nbsp;<a href='?nsort="&nsort&"&action=down&c_id="&cid&"'>向下</a></td></tr>"
    sql2="select s_id,s_name from jk_sort where c_id="&cid&" order by s_order,s_id"
    set rs2=conn.execute(sql2)
    do while not rs2.eof
       sid=rs2(0)
       response.write vbcrlf&"<tr align=center><td align=left>　　<font class=blue>"&rs2(1)&"</font></td><td><a href='?nsort="&nsort&"&action=edits&s_id="&sid&"'>修改</a>&nbsp;&nbsp;<a href=""javascript:Do_del_sort('"&sid&"');"">删除</a></td><td>排序：<a href='?nsort="&nsort&"&action=up&s_id="&sid&"'>向上</a>&nbsp;&nbsp;<a href='?nsort="&nsort&"&action=down&s_id="&sid&"'>向下</a></td></tr>"
       rs2.movenext
    loop
    rs2.close:set rs2=nothing
    rs.movenext
  loop
  rs.close:set rs=nothing
  response.write vbcrlf&"<tr><td height=30 align=center colspan=3><a href='?nsort="&nsort&"&action=addc'>添加一级分类</a>&nbsp;&nbsp;-&nbsp;&nbsp;<a href='?nsort="&nsort&"&action=list'>重新排序</a></td></tr></table>"
%><script language=JavaScript>
<!--
function Do_del_class(data1)
{
if (confirm("此操作将删除id为 "+data1+" 的一级分类！\n\n真的要删除吗？\n\n删除后将无法恢复！"))
  window.location="?nsort=<% response.write nsort %>&action=del&c_id="+data1
}

function Do_del_sort(data1)
{
if (confirm("此操作将删除id为 "+data1+" 的二级分类！\n\n真的要删除吗？\n\n删除后将无法恢复！"))
  window.location="?nsort=<% response.write nsort %>&action=del&s_id="+data1
}
//-->
</script><%
end sub

function jk_tit()
  jk_tit=vbcrlf&"<table border=0><tr><td height=30><font class=red>"&nsortn&"</font>&nbsp;&nbsp;-&nbsp;&nbsp;<font class=blue>"&jk_an&"</font></td></tr></table>"&vbcrlf
end function
%>