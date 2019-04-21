<!-- #include file="include/config_vouch.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
Dim nid
Dim url
Dim name
nummer = web_var(web_num,4):n_sort = "web"
tit    = "网站推荐"

If action = "view" And IsNumeric(id) Then
    sql     = "select url from website where id=" & id
    Set rs  = conn.execute(sql)

    If Not(rs.eof And rs.bof) Then
        tit = rs(0)
        rs.Close:Set rs = Nothing
        sql = "update website set counter=counter+1 where id=" & id
        conn.execute(sql)
        Call close_conn()
        Response.redirect "" & tit & ""
        Response.End
    End If

    rs.Close
End If

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Call vouch_skin("网站分类","<table border=0 width='60%' align=center><tr><td>" & nsort_left(n_sort,cid,sid,"?",1) & "</td></tr></table>","",1)
Call vouch_left("jt12","jt1")
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong

pageurl  = "?c_id=" & cid & "&s_id=" & sid & "&action=" & action & "&"
keyword  = code_form(Request.querystring("keyword"))
sea_type = Trim(Request.querystring("sea_type"))
Call cid_sid_sql(2,sea_type)

sql = "select * from website where hidden=1" & sqladd

Select Case action
    Case "counter"
        sql = sql & " order by counter desc,id desc"
    Case "tim"
        sql = sql & " order by tim desc"
    Case "id"
        sql = sql & " order by id"
    Case Else
        sql = sql & " order by id desc"
End Select

Set rs = Server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1

If Not(rs.eof And rs.bof) Then
    rssum = rs.recordcount
End If

Call format_pagecute() %>
<table border=0 width='96%'>
<tr><td colspan=3 height=30 align=center>
  <table border=0 width='98%'>
  <tr>
  <td>选择排序方式：</td>
  <td><select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option value='?c_id=<% Response.Write cid %>&s_id=<% Response.Write sid %>&action=default'<% If action = "" Then Response.Write " selected" %>>&nbsp;&nbsp;按默认排序&nbsp;&nbsp;</option>
<option value='?c_id=<% Response.Write cid %>&s_id=<% Response.Write sid %>&action=counter'<% If action = "counter" Then Response.Write " selected" %>>&nbsp;&nbsp;按人气排序&nbsp;&nbsp;</option>
<option value='?c_id=<% Response.Write cid %>&s_id=<% Response.Write sid %>&action=tim'<% If action = "tim" Then Response.Write " selected" %>>&nbsp;&nbsp;按时间排序&nbsp;&nbsp;</option>
<option value='?c_id=<% Response.Write cid %>&s_id=<% Response.Write sid %>&action=id'<% If action = "id" Then Response.Write " selected" %>>&nbsp;&nbsp;按先后排序&nbsp;&nbsp;</option>
</select></td>
  <td align=right>
    <table border=0>
    <form action='?' method=get>
    <input type=hidden name=action value='<% Response.Write action %>'>
    <input type=hidden name=c_id value='<% Response.Write cid %>'>
    <input type=hidden name=s_id value='<% Response.Write sid %>'>
    <input type=hidden name=page value='<% Response.Write viewpage %>'>
    <tr>
    <td>网站搜索：</td>
    <td><select name=sea_type size=1>
<option value='name'<% If sea_type = "'name" Then Response.Write " selected" %>>按名称</option>
<option value='remark'<% If sea_type = "remark" Then Response.Write " selected" %>>按介绍</option>
<option value='username'<% If sea_type = "username" Then Response.Write " selected" %>>按推荐人</option>
</select></td>
    <td><input type=text name=keyword value='<% Response.Write keyword %>' size=15 maxlength=20></td>
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
  <tr align=center valign=bottom><td width='30%'>现在有<font class=red><% Response.Write rssum %></font>条记录┋每页<font class=red><% Response.Write nummer %></font>个</td>
  <td width='70%'>页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font> 分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %></td></tr>
  </table>
</td></tr>
<tr><td height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<%

If Int(viewpage) > 1 Then
    rs.move (viewpage - 1)*nummer
End If

For i = 1 To nummer
    If rs.eof Then Exit For
    nid = rs("id"):url = rs("url"):name = rs("name") %>
<tr><td><% Call web_site_type() %></td></tr>
<tr><td height=1 background='IMAGES/BG_DIAN.GIF'></td></tr>
<%
    rs.movenext
Next

rs.Close:Set rs = Nothing %>
</table>
<%
'---------------------------------center end-------------------------------
Call web_end(0) %>