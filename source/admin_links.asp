<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #INCLUDE file="INCLUDE/jk_page_cute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

'	fir	sec	txt
Dim id,sort,rssum,nummer,thepages,viewpage,pageurl,page
id   = Trim(Request.querystring("id"))
sort = Trim(Request.querystring("sort"))
tit  = "<a href='?'>友情链接</a>┋" & _
"<a href='?action=main&sort=fir'>首页链接</a>┋" & _
"<a href='?action=main&sort=sec'>内页链接</a>┋" & _
"<a href='?action=main&sort=txt'>文字链接</a>┋" & _
"<a href='?action=add'>新增链接</a>┋" & _
"<a href='?action=list'>重新排序</a>"
Response.Write header(17,tit)

Select Case action
    Case "list"
        Call links_list()
    Case "add"
        Response.Write links_add()
    Case "addchk"
        Response.Write links_addchk()
    Case "order"

        If IsNumeric(id) And ( Trim(Request.querystring("actiones")) = "up" Or Trim(Request.querystring("actiones")) = "down" ) Then
            Response.Write links_order(id)
        Else
            Response.Write links_main()
        End If

    Case "del"

        If IsNumeric(id) Then
            Response.Write links_del(id)
        Else
            Response.Write links_main()
        End If

    Case "hidden"

        If IsNumeric(id) Then
            Response.Write links_hidden(id)
        Else
            Response.Write links_main()
        End If

    Case "edit"

        If IsNumeric(id) Then
            Response.Write links_edit(id)
        Else
            Response.Write links_main()
        End If

    Case "editchk"

        If IsNumeric(id) Then
            Response.Write links_editchk(id)
        Else
            Response.Write links_main()
        End If

    Case Else
        Response.Write links_main()
End Select

Response.Write ender()

Sub links_list()
    Dim rssum,i
    Set rs = Server.CreateObject("adodb.recordset")
    sql    = "select * from links where sort='fir' order by orders,id"
    rs.open sql,conn,1,3

    If rs.eof And rs.bof Then
        rssum = 0
    Else
        rssum = rs.recordcount
    End If

    For i = 1 To rssum
        rs("orders") = i
        rs.update
        rs.movenext
    Next

    rs.Close
    rssum = 0
    sql   = "select * from links where sort='sec' order by orders,id"
    rs.open sql,conn,1,3

    If rs.eof And rs.bof Then
        rssum = 0
    Else
        rssum = rs.recordcount
    End If

    For i = 1 To rssum
        rs("orders") = i
        rs.update
        rs.movenext
    Next

    rs.Close
    rssum = 0
    sql   = "select * from links where sort='txt' order by orders,id"
    rs.open sql,conn,1,3

    If rs.eof And rs.bof Then
        rssum = 0
    Else
        rssum = rs.recordcount
    End If

    For i = 1 To rssum
        rs("orders") = i
        rs.update
        rs.movenext
    Next

    rs.Close:Set rs = Nothing
    Response.Write links_main()
End Sub

Function links_order(id)
    Dim action,sort,tmp_id_1,tmp_id_2,tmp_order_1,tmp_order_2,sqladd,update_ok
    action     = Trim(Request.querystring("actiones"))
    update_ok  = "no":sort = "no"

    If action = "up" Then
        sqladd = " desc"
    Else
        sqladd = ""
    End If

    sql      = "select sort from links where id=" & id
    Set rs   = conn.execute(sql)

    If Not rs.eof Or Not rs.bof Then
        sort = rs("sort")
    End If

    rs.Close:Set rs = Nothing

    If sort <> "no" Then
        sql    = "select * from links where sort='" & sort & "' order by orders" & sqladd
        Set rs = conn.execute(sql)

        Do While Not rs.eof

            If Int(rs("id")) = Int(id) Then
                tmp_id_1    = id
                tmp_order_1 = rs("orders")
                rs.movenext

                If Not rs.eof Then
                    tmp_id_2    = rs("id")
                    tmp_order_2 = rs("orders")
                    update_ok   = "yes"
                    Exit Do
                End If

                Exit Do
            End If

            rs.movenext
        Loop

        rs.Close:Set rs = Nothing
    End If

    If update_ok = "yes" Then
        sql = "update links set orders=" & tmp_order_2 & " where id=" & tmp_id_1
        conn.execute(sql)
        sql = "update links set orders=" & tmp_order_1 & " where id=" & tmp_id_2
        conn.execute(sql)
    End If

    Response.redirect Request.servervariables("http_referer")
End Function

Function links_del(id)
    On Error Resume Next
    conn.execute("delete from links where id=" & id)

    If Err Then
        Err.Clear
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""您的操作有错误（error in del）存在！\n\n点击返回。"");" & _
        vbcrlf & "location='?action=main&sort=" & sort & "'" & _
        vbcrlf & "</script>")
    Else
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""成功删除了一条友情链接！\n\n点击返回。"");" & _
        vbcrlf & "location='?action=main&sort=" & sort & "'" & _
        vbcrlf & "</script>")
    End If

End Function

Function links_hidden(id)
    Dim hid,hh:hh = "no"
    Set rs = conn.execute("select hidden from links where id=" & id)

    If rs.eof And rs.bof Then
        '
    Else
        hid = rs("hidden")
        hh  = "yes"
    End If

    rs.Close:Set rs = Nothing

    If hh = "yes" Then

        If hid = True Then
            hid = 0
        Else
            hid = 1
        End If

        conn.execute("update links set hidden=" & hid & " where id=" & id)
    End If

    Response.redirect Request.servervariables("http_referer")
End Function

Function links_main()
    Dim i,sort,sqladd,sname,iid
    pageurl     = "?"
    sort        = Trim(Request.querystring("sort"))

    If sort = "fir" Or sort = "sec" Or sort = "txt" Then
        sqladd  = " where sort='" & sort & "'"
        pageurl = pageurl & "sort=" & sort & "&"

        Select Case sort
            Case "fir"
                sname = "首页"
            Case "sec"
                sname = "内页"
            Case "txt"
                sname = "文字"
        End Select

    End If

    sql    = "select * from links" & sqladd & " order by orders,id"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql,conn,1,1

    If rs.eof And rs.bof Then
        links_main = "现在还没有！"
    Else
        rssum      = rs.recordcount
        nummer     = 8
        Call format_pagecute()

        links_main = links_main & vbcrlf & "<script language=JavaScript><!--" & _
        vbcrlf & "function Do_del_data(data1)" & _
        vbcrlf & "{" & _
        vbcrlf & "if (confirm(""此操作将删除id为 ""+data1+"" 的友情链接！\n真的要删除吗？\n删除后将无法恢复！""))" & _
        vbcrlf & "  window.location=""" & pageurl & "action=del&id=""+data1" & _
        vbcrlf & "}" & _
        vbcrlf & "//--></script>" & _
        vbcrlf & "<table border=1 width=500 cellspacing=0 cellpadding=1 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>" & _
        vbcrlf & "<tr><td colspan=4 align=center height=30><table border=0 width='100%'cellspacing=0 cellpadding=0>" & _
        vbcrlf & "<tr align=center><td width='40%'>现在有 <font class=red>" & rssum & "</font> 个 <font class=red_4>" & sname & "</font> 链接</td>" & _
        vbcrlf & "<td width='60%'>" & pagecute_fun(viewpage,thepages,pageurl) & "</td></tr></table></td></tr>" & _
        "<tr align=center bgcolor=#ededed><td width='8%'>序号</td><td width='20%'>LOGO</td><td width='35%'>网站名称</td><td width='37%'>操作</td></tr>"

        If Int(viewpage) > 1 Then
            rs.move (viewpage - 1)*nummer
        End If

        pageurl = pageurl & "page=" & viewpage & "&"

        For i = 1 To nummer
            If rs.eof Then Exit For
            iid            = rs("id")
            links_main     = links_main & vbcrlf & "<tr align=center height=40><td>" & i + (viewpage - 1)*nummer & ".</td><td>"

            If rs("sort") = "txt" Then
                links_main = links_main & "txt"
            Else
                links_main = links_main & "<img src='" & rs("pic") & "' width=88 height=31 border=0>"
            End If

            links_main     = links_main & "</td><td><a href='" & rs("url") & "' target=_blank>" & code_html(rs("nname"),1,12) & "</a></td><td>"

            If rs("hidden") = True Then
                links_main = links_main & "<a href='" & pageurl & "action=hidden&id=" & iid & "'>显示</a>┋"
            Else
                links_main = links_main & "<a href='" & pageurl & "action=hidden&id=" & iid & "'><font class=red_2>隐藏</font></a>┋"
            End If

            links_main     = links_main & "<a href='" & pageurl & "action=order&actiones=up&id=" & iid & "'>向上</a>┋<a href='" & pageurl & "action=order&actiones=down&id=" & iid & "'>向下</a>┋<a href='" & pageurl & "action=edit&id=" & iid & "'>修改</a>┋<a href='javascript:Do_del_data(" & iid & ")'>删除</a></td></tr>"
            rs.movenext
        Next

        links_main = links_main & vbcrlf & "</table>"
    End If

    rs.Close:Set rs = Nothing
End Function

Function links_add() %><table border=0 width=450 cellspacing=0 cellpadding=2>
<form action='admin_links.asp?action=addchk' method=post>
  <tr>
    <td colspan=2 align=center height=50><font class=red>新增链接</font></td>
  </tr>
  <tr height=30>
    <td width='20%'>链接类型：</td>
    <td width='80%'><input type=radio name=sort value='fir' checked>首页链接
    <input type=radio name=sort value='sec'>内页链接
    <input type=radio name=sort value='txt'>文字链接</td>
  </tr>
  <tr height=30>
    <td>网站名称：</td>
    <td><input type=text name=nname size=50 maxlength=20></td>
  </tr>
  <tr height=30>
    <td>链接地址：</td>
    <td><input type=text name=url value='http://' size=50 maxlength=100></td>
  </tr>
  <tr height=30>
    <td>链接LOGO：</td>
    <td><input type=text name=pic value='images/links/' size=60 maxlength=100></td>
  </tr>
  <tr height=30 align=center>
    <td colspan=2><input type=submit value='新 增 链 接'></td>
  </tr>
</form></table><%
End Function

Function links_addchk()
    Dim nname,orders
    nname = Trim(Request.form("nname"))
    sort  = Trim(Request.form("sort"))

    If Len( nname) < 1 Or ( sort = "fir" And sort = "sec" And sort = "txt" ) Then
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""网站名称 和 链接类型 是必须要的！\n\n请返回输入。"");" & _
        vbcrlf & "history.back(1)" & _
        vbcrlf & "</script>")
    Else
        Set rs = Server.CreateObject("adodb.recordset")
        sql    = "select top 1 orders from links where sort='" & sort & "' order by orders desc"
        rs.open sql,conn,1,1

        If rs.eof And rs.bof Then
            orders = 0
        Else
            orders = Int(rs("orders"))
        End If

        rs.Close
        orders = Int(orders) + 1

        sql    = "select * from links"
        rs.open sql,conn,1,3
        rs.addnew
        rs("orders") = orders
        rs("sort") = sort
        rs("nname") = nname
        rs("url") = Trim(Request.form("url"))
        rs("pic") = Trim(Request.form("pic"))
        rs("hidden") = True
        rs.update
        rs.Close:Set rs = Nothing
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""成功新增了链接！\n\n点击返回。"");" & _
        vbcrlf & "location='?action=main&sort=" & sort & "'" & _
        vbcrlf & "</script>")
    End If

End Function

Function links_edit(id)
    Dim sss
    sql    = "select * from links where id=" & id
    Set rs = conn.execute(sql)

    If rs.eof And rs.bof Then
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""您的操作有错误（error in edit）存在！\n\n点击返回。"");" & _
        vbcrlf & "location='?action=main&sort=" & sort & "'" & _
        vbcrlf & "</script>")
    Else
        sss = rs("sort") %><table border=0 width=450 cellspacing=0 cellpadding=2>
<form action='admin_links.asp?action=editchk&id=<% Response.Write id %>' method=post>
  <tr>
    <td colspan=2 align=center height=50><font class=red>修改链接</font></td>
  </tr>
  <tr height=30>
    <td width='20%'>链接类型：</td>
    <td width='80%'><input type=radio name=sort value='fir'<% If sss = "fir" Then Response.Write " checked" %>>首页链接
    <input type=radio name=sort value='sec'<% If sss = "sec" Then Response.Write " checked" %>>内页链接
    <input type=radio name=sort value='txt'<% If sss = "txt" Then Response.Write " checked" %>>文字链接</td>
  </tr>
  <tr height=30>
    <td>网站名称：</td>
    <td><input type=text name=nname value='<% Response.Write rs("nname") %>' size=50 maxlength=20></td>
  </tr>
  <tr height=30>
    <td>链接地址：</td>
    <td><input type=text name=url value='<% Response.Write rs("url") %>' size=50 maxlength=100></td>
  </tr>
  <tr height=30>
    <td>链接LOGO：</td>
    <td><input type=text name=pic value='<% Response.Write rs("pic") %>' size=60 maxlength=100></td>
  </tr>
  <tr height=30 align=center>
    <td colspan=2><input type=submit value='修 改 链 接'></td>
  </tr>
</form></table><%
    End If

    rs.Close:Set rs = Nothing
End Function

Function links_editchk(id)
    Dim nname
    nname = Trim(Request.form("nname"))
    sort  = Trim(Request.form("sort"))

    If Len( nname) < 1 Or ( sort = "fir" And sort = "sec" And sort = "txt" ) Then
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""网站名称 和 链接类型 是必须要的！\n\n请返回输入。"");" & _
        vbcrlf & "history.back(1)" & _
        vbcrlf & "</script>")
    Else
        Set rs = Server.CreateObject("adodb.recordset")
        sql    = "select * from links where id=" & id
        rs.open sql,conn,1,3

        If rs.eof And rs.bof Then
            Response.Write("<script language=javascript>" & _
            vbcrlf & "alert(""您的操作有错误（error in editchk）存在！\n\n点击返回。"");" & _
            vbcrlf & "location='?action=main&sort=" & sort & "'" & _
            vbcrlf & "</script>")
        Else
            rs("sort") = sort
            rs("nname") = nname
            rs("url") = Trim(Request.form("url"))
            rs("pic") = Trim(Request.form("pic"))
            rs.update
            rs.Close:Set rs = Nothing
            Response.Write("<script language=javascript>" & _
            vbcrlf & "alert(""成功修改了链接！\n\n点击返回。"");" & _
            vbcrlf & "location='?action=main&sort=" & sort & "'" & _
            vbcrlf & "</script>")
        End If

    End If

End Function %>