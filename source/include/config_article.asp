<!-- #include file="config.asp" -->
<!-- #include file="config_nsort.asp" -->
<!-- #include file="skin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim atb:atb = " target=_blank"
sk_bar    = 15
index_url = "article":n_sort = "art"
tit_fir   = format_menu(index_url)

Sub article_main(n_jt,l_num,c_num,timt,et)
    Response.Write vbcrlf & "<table border=0 width='100%' align=center cellspacing=0 cellpadding=0>"
    Dim cnum
    Dim snum
    Dim rssum
    Dim j
    Dim nummer
    Dim topic
    Dim sql2
    Dim rs2
    Dim nid
    Dim temp1
    If n_jt <> "" Then n_jt = img_small(n_jt)
    sql    = "select count(c_id) from jk_class where nsort='" & n_sort & "'"
    Set rs = conn.execute(sql)
    cnum   = rs(0):rs.Close
    sql    = "select count(jk_sort.s_id) from jk_class inner join jk_sort on jk_class.c_id=jk_sort.c_id where nsort='" & n_sort & "'"
    Set rs = conn.execute(sql)
    snum   = rs(0):rs.Close

    sql    = "select count(id) from article where hidden=1"
    Set rs = conn.execute(sql)
    rssum  = rs(0):rs.Close
    sql    = "select c_id,c_name from jk_class where nsort='" & n_sort & "' order by c_order,c_id"
    Set rs = conn.execute(sql)

    Do While Not rs.eof
        Response.Write vbcrlf & "<tr align=center valign=top>"

        For j = 1 To 2
            If j = 2 Then rs.movenext
            If rs.eof Then Exit For
            nid       = rs(0)
            temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
            sql2      = "select top " & l_num & " id,topic,tim,counter from article where hidden=1 and c_id=" & nid & " order by id desc"
            Set rs2   = conn.execute(sql2)

            Do While Not rs2.eof
                topic = rs2("topic")
                temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & ">" & n_jt & "<a href='article_view.asp?id=" & rs2("id") & "'" & atb & " title='" & code_html(topic,1,0) & "'>" & code_html(topic,1,c_num) & "</a>" & format_end(et,"<font class=gray>" & time_type(rs2("tim"),timt) & "</font>,<font class=blue>" & rs2("counter") & "</font>") & "</td></tr>"
                rs2.movenext
            Loop

            rs2.Close:Set rs2 = Nothing
            temp1 = temp1 & "</table>"
            Response.Write vbcrlf & "<td width=289>" & format_barc("<a href='article_list.asp?c_id=" & nid & "'><b><font class=end>" & rs(1) & "</font></b></a>",temp1 & kong,3,0,11) & "</td>"
            If j = 1 Then Response.Write "<td width=1 bgcolor=" & web_var(web_color,3) & "></td>"
        Next

        If Not rs.eof Then rs.movenext
        Response.Write vbcrlf & "</tr>"
    Loop

    rs.Close:Set rs = Nothing
    Response.Write vbcrlf & "</table>"
End Sub

Sub article_view_review() %>
<table border=0 width='96%' cellspacing=0 cellpadding=0 class=tf>
<tr><td><% Call review_type(n_sort,id,"article_view.asp?id=" & id,1) %></td></tr>
<tr><td height=5></td></tr>
</table>
<%
End Sub

Sub article_view_about() %>
<table border=0 width='96%' cellspacing=0 cellpadding=0 class=tf>
<tr><td height=5></td></tr>

<tr><td height=1 background='images/bg_dian.gif'></td></tr>

<tr><td height=30 align=center>
  <table border=0 width='98%'>
  <tr>
  <td class=red_3><b>→&nbsp;主题所属分类：</b>&nbsp;&nbsp;<a href='article_list.asp?c_id=<% Response.Write cid %>'><% Response.Write cname %></a>&nbsp;→&nbsp;<a href='article_list.asp?c_id=<% Response.Write cid %>&s_id=<% Response.Write sid %>'><% Response.Write sname %></a></td>
  <td class=red_3 align=right>→&nbsp;<% Response.Write closer %></td>
  </tr>
  </table>
</td></tr>

<tr><td height=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td></tr>

<tr><td>
  <table border=0 width='100%' cellspacing=0 cellpadding=0>
  <tr valign=top align=center>
  <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
  <td><% Call article_left_hot("jt0",10,24,1,6) %></td>
  <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
  <td><% Call article_left_new("jt0",10,24,1,6,11) %></td>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td></tr>
  </table>
</td></tr>

<tr><td height=1 bgcolor=<% Response.Write web_var(web_color,3) %>></td></tr>
</table>
<%
End Sub

Sub article_list(n_num,c_num,timt,et)
    Dim topic
    Dim tim
    Dim nid
    Dim theheight
    Dim ss:ss = 0:theheight = 26
    If n_num > 5 Then nummer = n_num
    pageurl  = "?"
    keyword  = code_form(Request.querystring("keyword"))
    sea_type = Trim(Request.querystring("sea_type"))
    If sea_type <> "username" Then sea_type = "topic"
    Call cid_sid_sql(2,sea_type) %>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr ><td bgcolor=<% = web_var(web_color,2) %>  width='100'><table border=0 width='100%' cellspacing=0 cellpadding=0><tr height=25><td align=center><font class=red><b>一级分类</b></font></td></tr><tr><td><% = gang %></td></tr></table></td><%
    sql      = "select c_name,c_id from jk_class where nsort='" & n_sort & "' order by c_order"
    Set rs   = conn.execute(sql)

    Do While Not rs.eof
        nid  = rs("c_id")
        Response.Write "<td width=1 bgcolor=" & web_var(web_color,3) & "></td><td><table border=0 width='100%' align=center cellspacing=0 cellpadding=0 "
        If nid = cid Then Response.Write "bgcolor=" & web_var(web_color,6)
        If nid <> cid Then theheight = 25
        If nid = cid Then theheight = 26
        Response.Write "><tr height='" & theheight & "'><td  align=center>"
        Response.Write "<a href='?c_id=" & nid & "'"
        If nid = cid Then ss = 1:Response.Write " class=red_3"
        Response.Write ">" & rs("c_name") & "</a>"
        If nid <> cid Then Response.Write "<tr><td>" & gang & "</td></tr>"
        Response.Write "</td></tr></table></td>"
        rs.movenext
    Loop

    rs.Close %></td></tr></table>
<% If ss = 1 Then %>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
  <tr height=25  align=center><td bgcolor=<% = web_var(web_color,2) %> width='100'><font class=red><b>二级分类</b></font></td><td width=1 bgcolor=<% = web_var(web_color,3) %>></td><%
    sql     = "select s_name,s_id from jk_sort where c_id=" & cid & " order by s_order"
    Set rs  = conn.execute(sql)

    Do While Not rs.eof
        nid = rs("s_id")
        Response.Write "<td bgcolor=" & web_var(web_color,6) & ">|<a href='?c_id=" & cid & "&s_id=" & nid & "'"
        If nid = sid Then Response.Write " class=red_3"
        Response.Write ">" & rs("s_name") & "</a>|</td>"
        rs.movenext
    Loop

    rs.Close %>
</tr></table>
<%
End If %>
<table border=0 width='100%'  cellspacing=0 cellpadding=0>
  <tr><td colspan=4><% = gang %></td></tr>
  <tr align=center height=27>
  <td width='7%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>序号</b></td>
  <td width='63%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>文章主题</b></td>
  <td width='10%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>点击次数</b></td>
  <td width='20%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>浏览权限</b></td>
  </tr>
<%
Set rs   = Server.CreateObject("adodb.recordset")
sql      = "select id,username,topic,tim,counter,emoney,power from article where hidden=1" & sqladd & " order by id desc"
rs.open sql,conn,1,1

If rs.eof And rs.bof Then
    rssum = 0
Else
    rssum = rs.recordcount
End If

Call format_pagecute()

If Int(viewpage) > 1 Then
    rs.move (viewpage - 1)*nummer
End If

For i = 1 To nummer
    If rs.eof Then Exit For
    topic = rs("topic"):tim = rs("tim")
    Response.Write vbcrlf & "  <tr>" & _
    vbcrlf & "  <td height=25 bgcolor=" & web_var(web_color,6) & ">&nbsp;&nbsp;" & i + (viewpage - 1)*nummer & ".</td>" & _
    vbcrlf & "  <td>&nbsp;&nbsp;<a href='article_view.asp?id=" & rs("id") & "'" & atb & " title='文章标题：" & code_html(topic,1,0) & "<br>发 布 人：" & rs("username") & "<br>整理时间：" & rs("tim") & "'>" & code_html(topic,1,c_num) & "</a>" & format_end(et,"<font class=gray>" & time_type(tim,timt) & "</font>,<font class=blue>" & rs("counter") & "</font>") & "</td>" & _
    vbcrlf & "  <td class=red_3 align=center  bgcolor=" & web_var(web_color,6) & ">" & rs("counter") & "</td>" & _
    vbcrlf & "  <td>" & power_pic(0,rs("power"),0) & "</td></tr><tr><td height=1 colspan=4 background='images/bg_dian.gif'></td></tr>"
    rs.movenext
Next

rs.Close:Set rs = Nothing %>
</table>
<table border=0 width='100%'  cellspacing=0 cellpadding=0>
<tr><td align=center colspan=2><% = gang %></td></tr>
<tr><td align=left height=30>&nbsp;
本分类共有&nbsp;<font class=red><% Response.Write rssum %></font>&nbsp;篇文章</td><td align=right>
页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>&nbsp;
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000") %>
</td></tr>
</table>
<%
End Sub

Sub article_sea()
Dim temp1
Dim nid
Dim nid2
Dim rs
Dim sql
Dim rs2
Dim sql2
temp1 = vbcrlf & "<table border=0 cellspacing=0 cellpadding=0 align=center>" & _
vbcrlf & "<script language=javascript><!--" & _
vbcrlf & "function article_sea()" & _
vbcrlf & "{" & _
vbcrlf & "  if (article_sea_frm.keyword.value==""请输入关键字"")" & _
vbcrlf & "  {" & _
vbcrlf & "    alert(""请在搜索新闻前先输入要查询的 关键字 ！"");" & _
vbcrlf & "    article_sea_frm.keyword.focus();" & _
vbcrlf & "    return false;" & _
vbcrlf & "  }" & _
vbcrlf & "}" & _
vbcrlf & "--></script>" & _
vbcrlf & "<form name=article_sea_frm action='article_list.asp' method=get onsubmit=""return article_sea()"">" & _
vbcrlf & "<tr><td height=5></td></tr><tr align=center>" & _
vbcrlf & "<td>搜索选项：</td>" & _
vbcrlf & "<td>&nbsp;&nbsp;<select name=sea_type sizs=1><option value='topic'>文章标题</option><option value='username'>发布人</option></seelct></td>" & _
vbcrlf & "<td>&nbsp;&nbsp;<select name=c_id sizs=1><option value=''>全部分类</option>"
sql           = "select c_id,c_name from jk_class where nsort='" & n_sort & "' order by c_order,c_id"
Set rs        = conn.execute(sql)

Do While Not rs.eof
    nid       = Int(rs(0))
    temp1     = temp1 & vbcrlf & "<option value='" & nid & "' class=bg_2"
    If cid = nid Then temp1 = temp1 & " selected"
    temp1     = temp1 & ">" & rs(1) & "</option>"
    sql2      = "select s_id,s_name from jk_sort where c_id=" & nid & " order by s_order,s_id"
    Set rs2   = conn.execute(sql2)

    Do While Not rs2.eof
        nid2  = rs2(0)
        temp1 = temp1 & vbcrlf & "<option value='" & nid & "&s_id=" & nid2 & "'"
        If sid = nid2 Then temp1 = temp1 & " selected"
        temp1 = temp1 & ">　" & rs2(1) & "</option>"
        rs2.movenext
    Loop

    rs2.Close:Set rs2 = Nothing
    rs.movenext
Loop

rs.Close:Set rs = Nothing
temp1 = temp1 & vbcrlf & "</select></td>" & _
vbcrlf & "<td>&nbsp;&nbsp;<input type=text name=keyword value='请输入关键字' onfocus=""if (value =='请输入关键字'){value =''}"" onblur=""if (value ==''){value='请输入关键字'}"" size=20 maxlength=20></td>" & _
vbcrlf & "<td>&nbsp;&nbsp;<input type=image src='images/small/search_go.gif' border=0></td></tr>" & _
vbcrlf & "</form><tr><td height=5></td></tr></table>"
Response.Write format_barc("<font class=end><b>文章搜索</b></font>",temp1,1,1,3)
End Sub

Sub article_left_hot(n_jt,n_num,c_num,et,ct)
Dim rs
Dim sql
Dim ltemp
Dim topic
If n_jt <> "" Then n_jt = img_small(n_jt)
ltemp     = vbcrlf & "<table border=0 width='100%' class=tf>"
sql       = "select top " & n_num & " id,username,topic,tim,counter from article where hidden=1 order by counter desc,id desc"
Set rs    = conn.execute(sql)

Do While Not rs.eof
    topic = rs("topic")
    ltemp = ltemp & vbcrlf & "<tr><td height=" & space_mod & ">" & n_jt & "<a href='article_view.asp?id=" & rs("id") & "'" & atb & " title='文章标题：" & code_html(topic,1,0) & "<br>发 布 人：" & rs("username") & "<br>整理时间：" & rs("tim") & "'>" & code_html(topic,1,c_num) & "</a>" & format_end(et,"<font class=red>" & rs("counter") & "</font>") & "</td></tr>"
    rs.movenext
Loop

rs.Close:Set rs = Nothing
ltemp = ltemp & vbcrlf & "</table>"
'response.write kong & format_barc("<font class=end><b>点击排行</b></font>",ltemp,0,0,0,web_var(web_color,2)&"||images/bg2.gif","")
Response.Write format_barc("<font class=end><b>热门文章</b></font>",ltemp,3,0,5)
End Sub

Sub article_left_new(n_jt,n_num,c_num,et,ct,tt)
Dim rs
Dim sql
Dim ltemp
Dim topic
Dim tim
If n_jt <> "" Then n_jt = img_small(n_jt)
ltemp     = vbcrlf & "<table border=0 width='100%' class=tf>"
sql       = "select top " & n_num & " id,username,topic,tim,counter from article where hidden=1 order by id desc"
Set rs    = conn.execute(sql)

Do While Not rs.eof
    topic = rs("topic"):tim = rs("tim")
    ltemp = ltemp & vbcrlf & "<tr><td height=" & space_mod & ">" & n_jt & "<a href='article_view.asp?id=" & rs("id") & "'" & atb & " title='文章标题：" & code_html(topic,1,0) & "<br>发 布 人：" & rs("username") & "<br>整理时间：" & tim & "'>" & code_html(topic,1,c_num) & "</a>" & format_end(et,time_type(tim,tt)) & "</td></tr>"
    rs.movenext
Loop

rs.Close:Set rs = Nothing
ltemp = ltemp & vbcrlf & "</table>"
Response.Write format_barc("<font class=end><b>最近更新</b></font>",ltemp,3,0,7)
End Sub %>