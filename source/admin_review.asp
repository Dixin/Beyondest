<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim val_sort
val_sort="|news|art|down|gall|web|pro|"

dim sql2,rs2,del_temp,data_name,id,nummer,sqladd,page,rssum,thepages,viewpage,pageurl,nid,nsort
tit=vbcrlf & "<a href='?'>评论管理</a>&nbsp;┋&nbsp;" & _
    vbcrlf & "<a href='?action=delete'>批量删除</a>"
response.write header(7,tit)
pageurl="?":data_name="review":sqladd="":nummer=20
nsort=trim(request.querystring("nsort"))
if instr(1,val_sort,"|"&nsort&"|")<=0 then nsort=""

if trim(request("del_ok"))="ok" then
  response.write del_selects(trim(request.form("del_id")))
end if

function del_selects(delid)
  dim del_i,del_num,del_dim,del_sql,del_rs,del_username,picc,app,appn
  app=trim(request.form("app"))
  if delid<>"" and not isnull(delid) then
    delid=replace(delid," ","")
    del_dim=split(delid,",")
    del_num=UBound(del_dim)
    for del_i=0 to del_num
      appn="删除"
      del_sql="delete from "&data_name&" where rid="&del_dim(del_i)
      conn.execute(del_sql)
    next
    erase del_dim
    del_selects=vbcrlf&"<script language=javascript>alert(""共"&appn&"了 "&del_num+1&" 条记录！"");</script>"
  end if
end function

call review_main()

call close_conn()
response.write ender()

sub review_main()
  dim rword
  pageurl=pageurl&"nsort="&nsort&"&"
%>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='100%' cellpadding=2>
  <tr valign=top height=350>
    <td width='15%' class=htd><br><a href='?'<%if nsort="" then response.write " class=red_3"%>>全部评论</a><br>
<a href='?nsort=news'<%if nsort="news" then response.write " class=red_3"%>>新闻评论</a><br>
<a href='?nsort=art'<%if nsort="art" then response.write " class=red_3"%>>文栏评论</a><br>
<a href='?nsort=down'<%if nsort="down" then response.write " class=red_3"%>>下载评论</a><br>
<a href='?nsort=gall'<%if nsort="gall" then response.write " class=red_3"%>>贴图评论</a><br>
<a href='?nsort=web'<%if nsort="web" then response.write " class=red_3"%>>网站评论</a><br>
    </td>
    <td width='85%' align=center>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<form name=del_form action='<%response.write pageurl%>del_ok=ok' method=post>
<tr><td width='6%'></td><td width='88%'></td><td width='6%'></td></tr>
<%
  if nsort<>"" then sqladd=" where rsort='"&nsort&"'"
  rssum=0
  sql="select rid,rusername,rword,rtim from "&data_name&sqladd&" order by rid desc"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if not(rs.eof and rs.bof) then rssum=rs.recordcount
  call format_pagecute()
  del_temp=nummer
  if rssum=0 then del_temp=0
  if int(page)=int(thepages) then
    del_temp=rssum-nummer*(thepages-1)
  end if
%>
<tr><td colspan=3 align=center height=25>
现有<font class=red><%response.write rssum%></font>条信息　<input type=radio name=app value='del' checked> 删除
 <input type=checkbox name=del_all value=1 onClick=selectall('<%response.write del_temp%>')> 选中所有　<input type=submit value='操作所选' onclick=""return suredel('<%response.write del_temp%>');"">
</td></tr>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<%
  if int(viewpage)<>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    nid=rs("rid"):rword=rs("rword")
%>
<tr<%response.write mtr%>>
<td><%response.write i+(viewpage-1)*nummer%>. </td><td>
<a title='<%response.write nid%>：<%response.write code_html(rword,1,0)%>'><%response.write code_html(rword,1,35)%></a>
</td><td align=right>&nbsp;<input type=checkbox name=del_id value='<%response.write nid%>' class=bg_1></td></tr>
<%
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
%>