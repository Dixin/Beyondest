<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim val_sort
val_sort = "|news|art|down|gall|web|pro|"

Dim sql2
Dim rs2
Dim del_temp
Dim data_name
Dim id
Dim nummer
Dim sqladd
Dim page
Dim rssum
Dim thepages
Dim viewpage
Dim pageurl
Dim nid
Dim nsort
tit = vbcrlf & "<a href='?'>评论管理</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='?action=delete'>批量删除</a>"
Response.Write header(7,tit)
pageurl = "?":data_name = "review":sqladd = "":nummer = 20
nsort   = Trim(Request.querystring("nsort"))
If InStr(1,val_sort,"|" & nsort & "|") <= 0 Then nsort = ""

If Trim(Request("del_ok")) = "ok" Then
    Response.Write del_selects(Trim(Request.form("del_id")))
End If

Function del_selects(delid)
    Dim del_i
    Dim del_num
    Dim del_dim
    Dim del_sql
    Dim del_rs
    Dim del_username
    Dim picc
    Dim app
    Dim appn
    app             = Trim(Request.form("app"))

    If delid <> "" And Not IsNull(delid) Then
        delid       = Replace(delid," ","")
        del_dim     = Split(delid,",")
        del_num     = UBound(del_dim)

        For del_i = 0 To del_num
            appn    = "删除"
            del_sql = "delete from " & data_name & " where rid=" & del_dim(del_i)
            conn.execute(del_sql)
        Next

        Erase del_dim
        del_selects = vbcrlf & "<script language=javascript>alert(""共" & appn & "了 " & del_num + 1 & " 条记录！"");</script>"
    End If

End Function

Call review_main()

Call close_conn()
Response.Write ender()

Sub review_main()
    Dim rword
    pageurl = pageurl & "nsort=" & nsort & "&" %>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='100%' cellpadding=2>
  <tr valign=top height=350>
    <td width='15%' class=htd><br><a href='?'<% If nsort = "" Then Response.Write " class=red_3" %>>全部评论</a><br>
<a href='?nsort=news'<% If nsort = "news" Then Response.Write " class=red_3" %>>新闻评论</a><br>
<a href='?nsort=art'<% If nsort = "art" Then Response.Write " class=red_3" %>>文栏评论</a><br>
<a href='?nsort=down'<% If nsort = "down" Then Response.Write " class=red_3" %>>下载评论</a><br>
<a href='?nsort=gall'<% If nsort = "gall" Then Response.Write " class=red_3" %>>贴图评论</a><br>
<a href='?nsort=web'<% If nsort = "web" Then Response.Write " class=red_3" %>>网站评论</a><br>
    </td>
    <td width='85%' align=center>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<form name=del_form action='<% Response.Write pageurl %>del_ok=ok' method=post>
<tr><td width='6%'></td><td width='88%'></td><td width='6%'></td></tr>
<%
    If nsort <> "" Then sqladd = " where rsort='" & nsort & "'"
    rssum  = 0
    sql    = "select rid,rusername,rword,rtim from " & data_name & sqladd & " order by rid desc"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql,conn,1,1
    If Not(rs.eof And rs.bof) Then rssum = rs.recordcount
    Call format_pagecute()
    del_temp = nummer
    If rssum = 0 Then del_temp = 0

    If Int(page) = Int(thepages) Then
        del_temp = rssum - nummer*(thepages - 1)
    End If %>
<tr><td colspan=3 align=center height=25>
现有<font class=red><% Response.Write rssum %></font>条信息　<input type=radio name=app value='del' checked> 删除
 <input type=checkbox name=del_all value=1 onClick=selectall('<% Response.Write del_temp %>')> 选中所有　<input type=submit value='操作所选' onclick=""return suredel('<% Response.Write del_temp %>');"">
</td></tr>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<%

    If Int(viewpage) <> 1 Then
        rs.move (viewpage - 1)*nummer
    End If

    For i = 1 To nummer
        If rs.eof Then Exit For
        nid = rs("rid"):rword = rs("rword") %>
<tr<% Response.Write mtr %>>
<td><% Response.Write i + (viewpage - 1)*nummer %>. </td><td>
<a title='<% Response.Write nid %>：<% Response.Write code_html(rword,1,0) %>'><% Response.Write code_html(rword,1,35) %></a>
</td><td align=right>&nbsp;<input type=checkbox name=del_id value='<% Response.Write nid %>' class=bg_1></td></tr>
<%
        rs.movenext
    Next

    rs.Close:Set rs = Nothing %></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %>
</td></tr></table>
</td></tr></table>
<%
End Sub %>