<!-- #include file="INCLUDE/config_vouch.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************
dim nid,url,name
nummer=web_var(web_num,4):n_sort="web"
tit="网站推荐"

if action="view" and isnumeric(id) then
  sql="select url from website where id="&id
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then
    tit=rs(0)
    rs.close:set rs=nothing
    sql="update website set counter=counter+1 where id="&id
    conn.execute(sql)
    call close_conn()
    response.redirect ""&tit&""
    response.end
  end if
  rs.close
end if

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
call vouch_skin("网站分类","<table border=0 width='60%' align=center><tr><td>"&nsort_left(n_sort,cid,sid,"?",1)&"</td></tr></table>","",1)
call vouch_left("jt12","jt1")
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong

pageurl="?c_id="&cid&"&s_id="&sid&"&action="&action&"&"
keyword=code_form(request.querystring("keyword"))
sea_type=trim(request.querystring("sea_type"))
call cid_sid_sql(2,sea_type)

sql="select * from website where hidden=1"&sqladd
select case action
case "counter"
  sql=sql&" order by counter desc,id desc"
case "tim"
  sql=sql&" order by tim desc"
case "id"
  sql=sql&" order by id"
case else
  sql=sql&" order by id desc"
end select

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
  rssum=rs.recordcount
end if
call format_pagecute()
%>
<table border=0 width='96%'>
<tr><td colspan=3 height=30 align=center>
  <table border=0 width='98%'>
  <tr>
  <td>选择排序方式：</td>
  <td><select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option value='?c_id=<%response.write cid%>&s_id=<%response.write sid%>&action=default'<%if action="" then response.write " selected"%>>&nbsp;&nbsp;按默认排序&nbsp;&nbsp;</option>
<option value='?c_id=<%response.write cid%>&s_id=<%response.write sid%>&action=counter'<%if action="counter" then response.write " selected"%>>&nbsp;&nbsp;按人气排序&nbsp;&nbsp;</option>
<option value='?c_id=<%response.write cid%>&s_id=<%response.write sid%>&action=tim'<%if action="tim" then response.write " selected"%>>&nbsp;&nbsp;按时间排序&nbsp;&nbsp;</option>
<option value='?c_id=<%response.write cid%>&s_id=<%response.write sid%>&action=id'<%if action="id" then response.write " selected"%>>&nbsp;&nbsp;按先后排序&nbsp;&nbsp;</option>
</select></td>
  <td align=right>
    <table border=0>
    <form action='?' method=get>
    <input type=hidden name=action value='<%response.write action%>'>
    <input type=hidden name=c_id value='<%response.write cid%>'>
    <input type=hidden name=s_id value='<%response.write sid%>'>
    <input type=hidden name=page value='<%response.write viewpage%>'>
    <tr>
    <td>网站搜索：</td>
    <td><select name=sea_type size=1>
<option value='name'<%if sea_type="'name" then response.write " selected"%>>按名称</option>
<option value='remark'<%if sea_type="remark" then response.write " selected"%>>按介绍</option>
<option value='username'<%if sea_type="username" then response.write " selected"%>>按推荐人</option>
</select></td>
    <td><input type=text name=keyword value='<%response.write keyword%>' size=15 maxlength=20></td>
    <td>&nbsp;<input type=submit value='搜 索'></td>
    </tr>
    </table>
  </td>
  </tr>
  </table>
</td></tr>
<tr><td colspan=3 height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<tr><td align=center height=30>
  <table border=0 width='98%' cellspacing=0 cellpadding=0>
  <tr align=center valign=bottom><td width='30%'>现在有<font class=red><%response.write rssum%></font>条记录┋每页<font class=red><%response.write nummer%></font>个</td>
  <td width='70%'>页次：<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font> 分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%></td></tr>
  </table>
</td></tr>
<tr><td height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<%
if int(viewpage)>1 then
  rs.move (viewpage-1)*nummer
end if
for i=1 to nummer
  if rs.eof then exit for
  nid=rs("id"):url=rs("url"):name=rs("name")
%>
<tr><td><%call web_site_type()%></td></tr>
<tr><td height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<%
  rs.movenext
next
rs.close:set rs=nothing
%>
</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)
%>