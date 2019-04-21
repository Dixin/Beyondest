<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer
Dim sqladd
Dim page
Dim rssum
Dim thepages
Dim viewpage
Dim pageurl
Dim del_temp
Dim url
Dim types
Dim nname
tit = "<a href='?'>管理上传文件</a> ┋ " & _
"<a href='?types=1'>有效上传</a> ┋ " & _
"<a href='?types=0'>无效上传</a>"
Response.Write header(9,tit)
nummer = 15:rssum = 0:thepages = 0:viewpage = 1
types  = Trim(Request.querystring("types"))
If Not(IsNumeric(types)) Then types =  - 1

Select Case Int(types)
    Case 0
        nname   = "无效上传"
        pageurl = "?types=0&"
    Case 1
        nname   = "有效上传"
        pageurl = "?types=1&"
    Case Else
        types   =  - 1
        nname   = "所有"
        pageurl = "?"
End Select

If Trim(Request("del_ok")) = "ok" Then
    Call del_select(Trim(Request.form("del_id")))
End If

Call upload_main()
Response.Write ender()

Sub del_select(delid)
    'on error resume next
    Dim del_i
    Dim del_num
    Dim del_dim
    Dim del_sql

    If delid <> "" And Not IsNull(delid) Then
        delid       = Replace(delid," ","")
        del_dim     = Split(delid,",")
        del_num     = UBound(del_dim)

        For del_i = 0 To del_num
            del_sql = "select url from upload where id=" & del_dim(del_i)
            Set rs  = conn.execute(del_sql)

            If Not(rs.eof And rs.bof) Then
                Call del_file(rs("url"))
            End If

            rs.Close
            del_sql = "delete from upload where id=" & del_dim(del_i)
            conn.execute(del_sql)
        Next

        Erase del_dim
        Response.Write "<script language=javascript>alert(""共删除了 " & del_num + 1 & " 个文件！"");</script>"
    End If

End Sub

Sub upload_main()
    Dim ntypes
    Dim upload_path
    Dim nnsort
    Dim nsortn
    upload_path = web_var(web_upload,1) %>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=1 width='100%' cellspacing=0 cellpadding=1<% Response.Write table1 %>>
<form name=del_form action='<% Response.Write pageurl %>del_ok=ok' method=post>
<%
    If types >  - 1 Then sqladd = " where types=" & types
    sql    = "select * from upload" & sqladd & " order by id desc"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql,conn,1,1
    If Not(rs.eof And rs.bof) Then rssum = rs.recordcount
    Call format_pagecute()
    del_temp = nummer
    If rssum = 0 Then del_temp = 0
    If Int(page) = Int(thepages) Then del_temp = rssum - nummer*(thepages - 1) %>
<tr bgcolor=<% Response.Write color1 %>><td colspan=8 align=center height=30>
现有 <font class=red><% Response.Write rssum %></font> 个 <font class=red_3><% Response.Write nname %></font> 文件 ┋ 页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>
　<input type=checkbox name=del_all value=1 onClick="javascript:selectall('<% Response.Write del_temp %>');"> 选中所有　<input type=submit value='删除所选' onclick="return suredel('<% Response.Write del_temp %>');"></td></tr>
<tr align=center height=18 bgcolor=<% Response.Write color3 %>>
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
    If Int(viewpage) > 1 Then rs.move (viewpage - 1)*nummer

    For i = 1 To nummer
        If rs.eof Then Exit For
        url = rs("url"):ntypes = rs("types"):nnsort = rs("nsort") %>
<tr align=center<% Response.Write mtr %>>
<td><% Response.Write (viewpage - 1)*nummer + i %>.</td>
<td align=left><a href='<% Response.Write url_true(upload_path,url) %>' target=_blank><% Response.Write url %></a></td>
<td><% Response.Write rs("genre") %></td>
<td align=left><% Response.Write rs("sizes") %></td>
<td><%

        If Int(ntypes) <> 0 Then
            nsortn = format_menu(nnsort)
            If Len(nsortn) < 1 Then nsortn = nnsort
            Response.Write "<font alt='ID：" & rs("iid") & "'>" & nsortn & "</font>"
        Else
            Response.Write "<font class=red_2>无效</font>"
        End If %></td>
<td><% Response.Write format_user_view(rs("username"),1,"") %></td>
<td><% Response.Write time_type(rs("tim"),7) %></td>
<td><input type=checkbox name=del_id value='<% Response.Write rs("id") %>'></td>
</tr>
<%
        rs.movenext
    Next

    rs.Close:Set rs = Nothing %></form>
<tr>
<td colspan=8>分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,6,"#ff0000") %></td>
</tr>
</table>
<%
End Sub %>