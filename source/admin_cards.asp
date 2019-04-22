<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim id
Dim c_name
Dim c_pass
Dim c_emoney
Dim c_hidden
Dim rssum
Dim nummer
Dim thepages
Dim viewpage
Dim pageurl
Dim page
tit = "<a href='?'>会 员 卡</a>"
Response.Write header(1,tit)
id  = Trim(Request.querystring("id"))

If Trim(Request("del_ok")) = "ok" Then
    Response.Write del_select(Trim(Request.form("del_id")))
End If

Function del_select(delid)
    Dim del_i
    Dim del_num
    Dim del_dim
    Dim del_sql
    Dim del_rs
    Dim del_username
    Dim fobj
    Dim picc

    If delid <> "" And Not IsNull(delid) Then
        delid       = Replace(delid," ","")
        del_dim     = Split(delid,",")
        del_num     = UBound(del_dim)

        For del_i = 0 To del_num
            del_sql = "delete from cards where c_id=" & del_dim(del_i)
            conn.execute(del_sql)
        Next

        Erase del_dim
        del_select = vbcrlf & "<script language=javascript>alert(""共删除了 " & del_num + 1 & " 条记录！"");</script>"
    End If

End Function

If (action = "hidden") And IsNumeric(id) Then
    sql    = "select c_hidden from cards where c_id=" & id
    Set rs = conn.execute(sql)

    If Not(rs.eof And rs.bof) Then

        If Int(rs("c_hidden")) = 0 Then
            sql = "update cards set c_hidden=1 where c_id=" & id
        Else
            sql = "update cards set c_hidden=0 where c_id=" & id
        End If

        conn.execute(sql)
    End If

    rs.Close
    action = ""
End If

Select Case action
    Case "del"

        If IsNumeric(id) Then
            Call cards_del()
        Else
            Call cards_main()
        End If

    Case "add"
        Call cards_add()
    Case "edit"

        If IsNumeric(id) Then
            Call cards_edit()
        Else
            Call cards_main()
        End If

    Case Else
        Call cards_main()
End Select

close_conn
Response.Write ender()

Sub cards_edit()
    Dim sql2
    Dim rs2
    Set rs = Server.CreateObject("adodb.recordset")
    sql    = "select * from cards where c_id=" & id
    rs.open sql,conn,1,3

    If rs.eof And rs.bof Then
        rs.Close:Set rs = Nothing
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""您的操作有错误（error in edit）存在！\n\n点击返回。"");" & _
        vbcrlf & "location='?'" & _
        vbcrlf & "</script>")

        Exit Sub
        End If

        If Trim(Request.querystring("chk")) = "ok" Then
            c_name   = code_admin(Request.form("c_name"))
            c_pass   = code_admin(Request.form("c_pass"))
            c_emoney = code_admin(Request.form("c_emoney"))

            If Len(c_name) < 1 Or Len(c_pass) < 1 Or Not(IsNumeric(c_emoney)) Then

                Response.Write "会员卡号、密码和分值不能为空！<br><br>" & go_back:Exit Sub
                End If

                If c_name <> code_admin(Request.form("c_name2")) Then
                    sql2    = "select * from cards where c_name='" & c_name & "'"
                    Set rs2 = conn.execute(sql2)

                    If Not(rs2.eof And rs2.bof) Then
                        rs2.Close:Set rs2 = Nothing

                        Response.Write "会员卡号：" & c_name & " 已存在！请选用其它的代码。<br><br>" & go_back:Exit Sub
                        End If

                        rs2.Close:Set rs2 = Nothing
                    End If

                    rs("c_name") = c_name
                    rs("c_pass") = c_pass
                    rs("c_emoney") = c_emoney

                    If IsNumeric(Trim(Request.form("c_hidden"))) Then

                        If Int(Trim(Request.form("c_hidden"))) = 0 Then
                            rs("c_hidden") = 0
                        Else
                            rs("c_hidden") = 1
                        End If

                    Else
                        rs("c_hidden") = 0
                    End If

                    rs.update
                    rs.Close:Set rs = Nothing
                    Response.Write "<script lanuage=javascrip>alert(""修改会员卡号成功！"");location.href='?page=" & Trim(Request.querystring("page")) & "';</script>"

                    Exit Sub
                    End If %>
<table border=0 align=center>
<form action='?action=edit&chk=ok&page=<% Response.Write Trim(Request.querystring("page")) %>&id=<% Response.Write id %>' method=post>
<tr><td>卡号：&nbsp;<input type=text name=c_name value='<% Response.Write rs("c_name") %>' size=20 maxlength=20></td></tr>
<input type=hidden name=c_name2 value='<% Response.Write rs("c_name") %>'>
<tr><td>密码：&nbsp;<input type=text name=c_pass value='<% Response.Write rs("c_pass") %>' size=20 maxlength=20></td></tr>
<tr><td>分值：&nbsp;<input type=text name=c_emoney value='<% Response.Write rs("c_emoney") %>' size=20 maxlength=20></td></tr>
<tr><td>是否使用：<input type=radio name=c_hidden value='1'<% If Int(rs("c_hidden")) = 1 Then Response.Write " checked" %>>&nbsp;已使用&nbsp;
<input type=radio name=c_hidden value='0'<% If Int(rs("c_hidden")) = 0 Then Response.Write " checked" %>>&nbsp;未使用</td></tr>
<tr><td align=center height=30><input type=submit value='修改会员卡'></td></tr>
</form>
</table>
<%
                End Sub

                Sub cards_add()
                    c_name   = code_admin(Request.form("c_name"))
                    c_pass   = code_admin(Request.form("c_pass"))
                    c_emoney = code_admin(Request.form("c_emoney"))

                    If Len(c_name) < 1 Or Len(c_pass) < 1 Or Not(IsNumeric(c_emoney)) Then

                        Response.Write "会员卡号、密码和分值不能为空！<br><br>" & go_back:Exit Sub
                        End If

                        Set rs = Server.CreateObject("adodb.recordset")
                        sql    = "select * from cards where c_name='" & c_name & "'"
                        rs.open sql,conn,1,3

                        If Not(rs.eof And rs.bof) Then
                            rs.Close:Set rs = Nothing

                            Response.Write "会员卡号：" & c_name & " 已存在！请选用其它的代码。<br><br>" & go_back:Exit Sub
                            End If

                            rs.addnew
                            rs("c_name") = c_name
                            rs("c_pass") = c_pass
                            rs("c_emoney") = c_emoney
                            rs("c_hidden") = 0
                            rs.update
                            rs.Close:Set rs = Nothing
                            Response.Write "<script lanuage=javascrip>alert(""添加会员卡号成功！"");location.href='?';</script>"
                        End Sub

                        Sub cards_main()
                            Dim i
                            Dim hidden
                            Dim sqladd
                            Dim sname
                            Dim iid
                            Dim del_temp
                            hidden  = Trim(Request.querystring("hidden"))
                            pageurl = "?hidden=" & hidden & "&" %>
<script language=javascript src='STYLE/admin_del.js'></script>
<form name=del_form action='<% Response.Write pageurl %>del_ok=ok' method=post>
<%
                            sql     = "select * from cards order by c_id desc"
                            Set rs  = Server.CreateObject("adodb.recordset")
                            rs.open sql,conn,1,1

                            If rs.eof And rs.bof Then
                                rssum = 0
                            Else
                                rssum = rs.recordcount
                            End If

                            nummer    = 15
                            Call format_pagecute()
                            del_temp  = nummer
                            If rssum = 0 Then del_temp = 0

                            If Int(page) = Int(thepages) Then
                                del_temp = rssum - nummer*(thepages - 1)
                            End If %>
<table border=1 width='80%' cellspacing=0 cellpadding=1 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<tr><td colspan=5 align=center height=30>
  <table border=0 width='100%'cellspacing=0 cellpadding=0>
  <tr align=center>
  <td width='40%'>现在有 <font class=red><% Response.Write rssum %></font> 个 <font class=red_3><% Response.Write sname %></font> 会员卡号</td>
  <td width='60%'><input type=checkbox name=del_all value=1 onClick=selectall('<% Response.Write del_temp %>')> 选中所有　<input type=submit value='删除所选' onclick="return suredel('<% Response.Write del_temp %>');"></td>
  </tr>
  </table>
</td></tr>
<tr align=center bgcolor=#ededed>
<td width='8%'>序号</td>
<td width='26%'>会员卡号代码</td>
<td width='26%'>会员卡号类型</td>
<td width='18%'>分值</td>
<td width='24%'>操作</td>
</tr>
<%

                            If Int(viewpage) > 1 Then
                                rs.move (viewpage - 1)*nummer
                            End If

                            For i = 1 To nummer
                                If rs.eof Then Exit For
                                iid = rs("c_id") %>
<tr>
<td align=center><% Response.Write i + (viewpage - 1)*nummer %>.</td>
<td><% Response.Write rs("c_name") %></td>
<td><% Response.Write rs("c_pass") %></td>
<td><% Response.Write rs("c_emoney") %></td>
<td align=center><a href='?action=hidden&page=<% Response.Write viewpage %>&id=<% Response.Write iid %>'>
<%

                                If Int(rs("c_hidden")) = 0 Then
                                    Response.Write "未使用"
                                Else
                                    Response.Write "<font class=red>已使用</font>"
                                End If %></a>&nbsp;
<a href='?action=edit&id=<% Response.Write iid %>'>修改</a>&nbsp;
<input type=checkbox name=del_id value='<% Response.Write iid %>' class=bg_1></td></tr>
<%
                                rs.movenext
                            Next

                            rs.Close:Set rs = Nothing %>
<tr><td colspan=5>页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %></td></tr>
</form>
</table>
<br>
<table border=0 align=center>
<form name=add_frm action='?action=add' method=post>
<tr>
<td>卡号：</td>
<td><input type=text name=c_name size=12 maxlength=20></td>
<td>密码：</td>
<td><input type=text name=c_pass size=12 maxlength=20></td>
<td>分值：</td>
<td><input type=text name=c_emoney size=10 maxlength=20></td>
<td>&nbsp;<input type=submit value='添加会员卡'></td>
</tr>
</form>
</table>
<%
                        End Sub %>