<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim admin_menu
admin_menu = "<a href='admin_forum.asp'>论坛管理</a> ┋ " & _
"<a href='admin_forum.asp?action=mod'>合并论坛</a> ┋ " & _
"<a href='admin_forum.asp?action=order'>重新排序</a>"
Response.Write header(11,admin_menu)

Select Case action
    Case "mod"
        Call forum_mod()
    Case "order"
        Call forum_order()
    Case "forum_add"
        Call forum_add()
    Case "forum_edit"
        Call forum_edit()
    Case "del_forum"
        Call del_forum()
    Case "class_add"
        Call class_add()
    Case "class_edit"
        Call class_edit()
    Case "del_class"
        Call del_class()
    Case Else
        Call forum_main()
End Select

close_conn
Response.Write ender()

Sub forum_order()
    Dim rs,sql,rsf,sqlf,i,j,cid,fid
    i           = 1
    sql         = "select class_id from bbs_class order by class_order,class_id"
    Set rs      = conn.execute(sql)

    Do While Not rs.eof
        j       = 1:cid = rs(0)
        conn.execute("update bbs_class set class_order=" & i & " where class_id=" & cid)
        sqlf    = "select forum_id from bbs_forum where class_id=" & cid & " order by forum_order,forum_id"
        Set rsf = conn.execute(sqlf)

        Do While Not rsf.eof
            fid = rsf(0)
            conn.execute("update bbs_forum set forum_order=" & j & " where forum_id=" & fid)
            rsf.movenext
            j = j + 1
        Loop

        rsf.Close:Set rsf = Nothing
        rs.movenext
        i = i + 1
    Loop

    rs.Close:Set rs = Nothing
    Call forum_main()
End Sub

Sub class_edit()
    Dim classid,rs,strsql,class_name,class_order
    classid = Trim(Request("class_id"))

    If Not(IsNumeric(classid)) Then
        Response.redirect "admin_forum.asp"
        Response.End
    End If

    Set rs = Server.CreateObject("adodb.recordset")
    strsql = "Select * from bbs_class where class_id=" & classid
    rs.open strsql,conn,1,3 %><font class=red>修改论坛分类</font><br><br><br>
<table border=0 width=300><%

    If Trim(Request("edit")) = "ok" Then
        class_name = code_form(Request.form("class_name"))

        If class_name = "" Then
            Response.Write( VbCrLf & "<tr><td height=80 align=center><font class=red_2>论坛分类名称不能为空！</font><br><br>" & go_back & "</td></tr>")
        Else
            rs("class_name") = class_name
            rs.update
            Response.Write( VbCrLf & "<tr><td height=80 align=center>成功的修改了论坛分类：<font class=red>" & class_name & "</font></td></tr>")
        End If

    Else %>
<tr>
<form method=post action='admin_forum.asp?action=class_edit&class_id=<% = classid %>&edit=ok'>
<td width='40%' align=center></td><td width='60%'></td>
</tr>
<tr height=30>
<td align=center>论坛分类名称：</td> 
<td><input type=text name=class_name value='<% = rs("class_name") %>' size=20 maxlength=20></td> 
</tr>
<tr height=30> 
<td colspan=2 align=center height=30><input type=submit value=' 提 交 修 改 '></td>
</form>
</tr><%
    End If

    rs.Close:Set rs = Nothing %></table><%
End Sub

Sub class_add() %><font class=red>添加论坛分类</font><br><br><br>
<table border=0 width=300>
<%

    If Trim(Request.querystring("add")) = "ok" Then
        Dim rs,strsql,class_name,class_order
        class_name = code_form(Request.form("class_name"))

        If class_name = "" Then
            Response.Write( VbCrLf & "<tr><td height=80 align=center><font class=red_2>论坛分类名称不能为空！</font><br><br>" & go_back & "</td></tr>")
        Else
            Set rs = Server.CreateObject("adodb.recordset")
            strsql = "Select top 1 * from bbs_class order by class_order desc"
            rs.open strsql,conn,1,1

            If rs.eof And rs.bof Then
                class_order = 0
            Else
                class_order = rs("class_order")
            End If

            class_order     = class_order + 1
            rs.Close
            strsql          = "Select * from bbs_class"
            rs.open strsql,conn,1,3
            rs.addnew
            rs("class_order") = class_order
            rs("class_name") = class_name
            rs.update
            Response.Write( VbCrLf & "<tr><td height=80 align=center>成功的添加了论坛分类：<font class=red>" & class_name & "</font></td></tr>")
            rs.Close:Set rs = Nothing
        End If

    Else %>
<tr>
<form method=post action='admin_forum.asp?action=class_add&add=ok'>
<td width='40%' align=center></td><td width='60%'></td>
</tr>
<tr height=30>
<td align=center>论坛分类名称：</td> 
<td><input type=text name=class_name size=20 maxlength=20></td> 
</tr>
<tr height=30> 
<td colspan=2 align=center height=30><input type=submit value=' 提 交 添 加 '></td>
</form>
</tr><%
    End If %></table><%
End Sub

Sub forum_edit()
    Dim classid,forumid,rs,strsql,classname,forum_name
    classid = Trim(Request("class_id"))
    forumid = Trim(Request("forum_id"))

    If Not(IsNumeric(classid)) Or Not(IsNumeric(forumid)) Then

        Call forum_main():Exit Sub
        End If

        strsql = "select class_name from bbs_class where class_id=" & classid
        Set rs = conn.execute(strsql)

        If rs.eof And rs.bof Then
            rs.Close:Set rs = Nothing

            Call forum_main():Exit Sub
            End If

            classname = rs("class_name")
            rs.Close:Set rs = Nothing %><font class=red>修改论坛</font>（<font class=blue_1><% = classname %></font>）<br><br><br>
<table border=0 width=400><%
            Set rs    = Server.CreateObject("adodb.recordset")
            strsql    = "Select * from bbs_forum where forum_id=" & forumid
            rs.open strsql,conn,1,3

            If Trim(Request.querystring("edit")) = "ok" Then
                forum_name = code_form(Request.form("forum_name"))

                If forum_name = "" Then
                    Response.Write( VbCrLf & "<tr><td height=80 align=center><font class=red_2>论坛名称不能为空！</font><br><br>" & go_back & "</td></tr>")
                Else
                    rs("class_id")     = classid
                    rs("forum_name")     = forum_name
                    rs("forum_pic")     = Trim(Request.form("forum_pic"))

                    If Request.form("forum_hidden") = "no" Then
                        rs("forum_hidden") = False
                    Else
                        rs("forum_hidden") = True
                    End If

                    rs("forum_type")     = Request.form("forum_type")
                    rs("forum_remark")     = Request.form("forum_remark")
                    rs("forum_power")     = code_form(Request.form("forum_power"))
                    rs.update
                    Response.Write( VbCrLf & "<tr><td height=80 align=center>成功的修改了论坛：<font class=red>" & forum_name & "</font></td></tr>")
                End If

            Else %><form method=post action='admin_forum.asp?action=forum_edit&forum_id=<% = forumid %>&edit=ok'>
<tr><td width='20%' align=center></td><td width='80%'></td></tr>
<tr height=30>
<td align=center>论坛名称：</td> 
<td><input type=text name=forum_name value='<% = rs("forum_name") %>' size=30 maxlength=20></td> 
</tr>
<tr height=30>
<td align=center>所属分类：</td> 
<td><select name=class_id size=1>
<%
                Dim crs,csql,cid,ctype
                csql     = "select * from bbs_class order by class_order"
                Set crs  = conn.execute(csql)

                Do While Not crs.eof
                    cid  = crs("class_id")
                    Response.Write vbcrlf & "<option value='" & cid & "'"

                    If Int(classid) = Int(cid) Then
                        Response.Write " selected class=bg_1"
                    End If

                    Response.Write ">" & crs("class_name") & "</option>"
                    crs.movenext
                Loop

                ctype    = Int(rs("forum_type")) %>
</select></td> 
</tr>
<tr>
<td align=center>论坛说明：</td> 
<td><textarea name=forum_remark rows=5 cols=50><% = rs("forum_remark") %></textarea></td> 
</tr>
<tr>
<td align=center>论坛图片：</td> 
<td><input type=text name=forum_pic value='<% = rs("forum_pic") %>' size=30 maxlength=50></td> 
</tr>
<tr>
<td align=center>论坛类型：</td> 
<td><select name=forum_type size=1>
<%
                Dim tdim,t2
                tdim     = Split(forum_type,"|")

                For i = 0 To UBound(tdim)
                    Response.Write vbcrlf & "<option value='" & i + 1 & "'"
                    If ctype = i + 1 Then Response.Write " selected"
                    Response.Write ">" & Right(tdim(i),Len(tdim(i)) - InStr(tdim(i),":")) & "</option>"
                Next

                Erase tdim %>
</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;是否开放：<input type=checkbox name=forum_hidden value='no'<% If rs("forum_hidden") = False Then Response.Write " checked" %>>&nbsp;（选上为开放）</td> 
</tr>
<tr height=50>
<td align=center>论坛版主：<br><br></td> 
<td><input type=text name=forum_power value='<% = rs("forum_power") %>' size=50 maxlength=50><br>多个请用“|”分开，如：“笼民|apple|5271”</td> 
</tr>
<tr height=30><td colspan=2 align=center height=30><input type=submit value=' 提 交 修 改 '></td></tr>
</form><%
            End If %></table><%
        End Sub

        Sub forum_add()
            Dim rs,strsql,classname,classid,forum_name,forum_order
            classid = Trim(Request("class_id"))

            If Not(IsNumeric(classid)) Then

                Call forum_main():Exit Sub
                End If

                strsql = "select class_name from bbs_class where class_id=" & classid
                Set rs = conn.execute(strsql)

                If rs.eof And rs.bof Then
                    rs.Close:Set rs = Nothing

                    Call forum_main():Exit Sub
                    End If

                    classname      = rs("class_name")
                    rs.Close:Set rs = Nothing %><font class=red>添加论坛</font>（<font class=blue_1><% = classname %></font>）<br><br><br>
<table border=0 width=400>
<%

                    If Trim(Request("add")) = "ok" Then
                        forum_name = code_form(Request.form("forum_name"))

                        If forum_name = "" Then
                            Response.Write( VbCrLf & "<tr><td height=80 align=center><font class=red_2>论坛名称不能为空！</font><br><br>" & go_back & "</td></tr>")
                        Else
                            Set rs = Server.CreateObject("adodb.recordset")
                            strsql = "Select top 1 * from bbs_forum where class_id=" & classid & " order by forum_order desc"
                            rs.open strsql,conn,1,1

                            If rs.eof And rs.bof Then
                                forum_order = 0
                            Else
                                forum_order = rs("forum_order")
                            End If

                            forum_order     = forum_order + 1
                            rs.Close
                            strsql          = "Select * from bbs_forum"
                            rs.open strsql,conn,1,3
                            rs.addnew
                            rs("class_id") = classid
                            rs("forum_order") = forum_order
                            rs("forum_name") = forum_name
                            rs("forum_remark") = Request.form("forum_remark")
                            rs("forum_power") = code_form(Request.form("forum_power"))
                            rs("forum_hidden") = False
                            rs("forum_type") = 1
                            rs("forum_topic_num") = 0
                            rs("forum_data_num") = 0
                            rs("forum_new_info") = "|||"
                            rs.update
                            Response.Write( VbCrLf & "<tr><td height=80 align=center>成功的添加了论坛：<font class=red>" & forum_name & "</font></td></tr>")
                            rs.Close:Set rs = Nothing
                        End If

                    Else %>
<form method=post action='admin_forum.asp?action=forum_add&add=ok&class_id=<% = classid %>'>
<tr><td width='20%' align=center></td><td width='80%'></td></tr>
<tr height=30>
<td align=center>论坛名称：</td> 
<td><input type=text name=forum_name size=30 maxlength=20></td> 
</tr>
<tr>
<td align=center>论坛说明：</td> 
<td><textarea name=forum_remark rows=5 cols=50></textarea></td> 
</tr>
<tr height=50>
<td align=center>论坛版主：<br><br></td> 
<td><input type=text name=forum_power size=50 maxlength=50><br>多个请用“|”分开，如：“笼民|apple|5271”</td> 
</tr>
<tr height=30><td colspan=2 align=center height=30><input type=submit value=' 提 交 添 加 '></td></tr>
</form><%
                    End If

                    Response.Write "</table>"
                End Sub

                Sub forum_mod() %>
<table border=0>
<form action='admin_forum.asp?action=mod' method=post>
<input type=hidden name=modok value='ok'>
<tr><td align=center height=50 colspan=4><font class=red>合并论坛</font></td></tr>
<%

                    If Trim(Request.form("modok")) = "ok" Then
                        Response.Write "<tr><td align=center height=50 colspan=4>"
                        Dim sel1,sel2,rs,sql
                        sel1 = Trim(Request.form("sel_1"))
                        sel2 = Trim(Request.form("sel_2"))

                        If Not(IsNumeric(sel1)) Or Not(IsNumeric(sel2)) Then
                            Response.Write "<font class=red_2>您没有选择要合并的论坛！</font>"
                        Else
                            sql = "update bbs_topic set forum_id=" & Int(sel2) & " where forum_id=" & Int(sel1)
                            conn.execute(sql)
                            sql = "update bbs_data set forum_id=" & Int(sel2) & " where forum_id=" & Int(sel1)
                            conn.execute(sql)
                            Response.Write "<font class=red_3>论坛合并成功！</font>"
                        End If

                        Response.Write "</td></tr>"
                    End If %>
<tr height=50>
<td>从</td>
<td><select name=sel_1><% Call forum_list() %></select></td>
<td>合并到</td>
<td><select name=sel_2><% Call forum_list() %></select></td>
</tr>
<tr><td align=center height=50 colspan=4><input type=submit value='开始合并'></td></tr>
</form>
</table>
<%
                End Sub

                Sub forum_list()
                    Dim strsqlclass,rsclass,strsqlboard,rsboard
                    strsqlclass = "select class_id,class_name from bbs_class order by class_order"
                    Set rsclass = conn.execute(strsqlclass)

                    If Not(rsclass.bof And rsclass.eof) Then

                        Do While Not rsclass.eof
                            Response.Write vbcrlf & "<option class=bg_2>╋ " & rsclass("class_name") & "</option>"
                            strsqlboard = "select forum_id,forum_name from bbs_forum where class_id=" & rsclass("class_id") & " order by forum_order"
                            Set rsboard = conn.execute(strsqlboard)

                            If rsboard.eof And rsboard.bof Then
                                Response.Write vbcrlf & "<option>没有论坛</option>"
                            Else

                                Do While Not rsboard.eof
                                    Response.Write vbcrlf & "<option value='" & rsboard("forum_id") & "'>　├" & rsboard("forum_name") & "</option>"
                                    rsboard.movenext
                                Loop

                            End If

                            rsclass.movenext
                        Loop

                    End If

                    Set rsclass = Nothing:Set rsboard = Nothing
                End Sub

                Sub forum_main() %><table border=1 cellspacing=0 cellpadding=2 width=500 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<%
                    Dim rsclass,strsqlclass,rsboard,strsqlboard,classid,forumid,forumname
                    strsqlclass = "select * from bbs_class order by class_order"
                    Set rsclass = conn.execute(strsqlclass)

                    If rsclass.bof And rsclass.eof Then
                        Response.Write vbcrlf & "<tr><td align=center height=200><font class=red_2>现在好像还没有论坛分类！</font></td></tr>"
                    Else

                        Do While Not rsclass.eof
                            classid     = rsclass("class_id")
                            Response.Write vbcrlf & "<tr height=20 bgcolor=#ffffff align=center><td align=left>" & img_small("fk2") & vbcrlf & "<font class=red_3><b>" & rsclass("class_name") & "</b></font></td><td><a href='admin_forum.asp?action=forum_add&class_id=" & classid & "'>添加论坛</a></td><td><a href='admin_forum.asp?action=class_edit&class_id=" & classid & "'>修改</a></td><td><a href=""javascript:Do_del_class('" & classid & "');"">删除</a></td><td>排序：<a href='admin_class_order.asp?class_id=" & classid & "&action=up'>向上</a> <a href='admin_class_order.asp?class_id=" & classid & "&action=down'>向下</a></td></tr>"
                            strsqlboard = "select forum_id,forum_name,forum_power,forum_hidden from bbs_forum where class_id=" & classid & " order by forum_order"
                            Set rsboard = conn.execute(strsqlboard)

                            If rsboard.eof And rsboard.bof Then
                                Response.Write vbcrlf & "<tr><td colspan=5><font class=gray>　　本分类还没有论坛</font></td></tr>"
                            Else

                                Do While Not rsboard.eof
                                    forumid = rsboard("forum_id"):forumname = rsboard("forum_name")
                                    Response.Write vbcrlf & "<tr align=center><td align=left>　　<font class=blue><b>" & forumname & "</b></font>"
                                    If rsboard("forum_hidden") = True Then Response.Write " <font class=gray>隐藏</font>"
                                    Response.Write "</td><td align=left>（版主：" & rsboard("forum_power") & "）</td><td><a href='admin_forum.asp?action=forum_edit&class_id=" & classid & "&forum_id=" & forumid & "'>编辑</a></td><td><a href=""javascript:Do_del_forum(" & forumid & ");"">删除</a></td><td>排序：<a href='admin_forum_order.asp?forum_id=" & forumid & "&class_id=" & classid & "&action=up'>向上</a> <a href='admin_forum_order.asp?forum_id=" & forumid & "&class_id=" & classid & "&action=down'>向下</a></td></tr>"
                                    rsboard.movenext
                                Loop

                            End If

                            rsclass.movenext
                        Loop

                    End If

                    Set rsclass = Nothing:Set rsboard = Nothing %>
<tr><td align=center height=30 colspan=5><a href='admin_forum.asp?action=class_add'>添加论坛分类</a></td></tr>
</table>
<script language=JavaScript>
<!--
function Do_del_class(data1)
{
if (confirm("此操作将删除id为 "+data1+" 的论坛分类！\n\n真的要删除吗？\n\n删除后将无法恢复！"))
  window.location="admin_forum.asp?action=del_class&class_id="+data1
}

function Do_del_forum(data1)
{
if (confirm("此操作将删除id为 "+data1+" 的论坛！\n\n真的要删除吗？\n\n删除后将无法恢复！"))
  window.location="admin_forum.asp?action=del_forum&forum_id="+data1
}
//-->
</script><%
                End Sub

                Sub del_class()
                    Dim classid,sql,rs,forumid
                    classid = Trim(Request.querystring("class_id"))

                    If Not(IsNumeric(classid)) Then

                        Call forum_main():Exit Sub
                        End If

                        sql         = "delete from bbs_class where class_id=" & classid
                        conn.execute(sql)
                        sql         = "select forum_id from bbs_forum where class_id=" & classid
                        Set rs      = conn.execute(sql)

                        Do While Not rs.eof
                            forumid = rs("forum_id")
                            sql     = "delete from bbs_topic where forum_id=" & forumid
                            conn.execute(sql)
                            sql     = "delete from bbs_data where forum_id=" & forumid
                            conn.execute(sql)
                            rs.movenext
                        Loop

                        sql = "delete from bbs_forum where class_id=" & classid
                        conn.execute(sql)
                        Response.Write "<script language=javascript>alert(""已成功能删除了一个论坛分类！\n\n（包括其所属的论坛的贴子）"");</script>"
                        Call forum_main()
                    End Sub

                    Sub del_forum()
                        Dim classid,forumid,sql
                        forumid = Trim(Request.querystring("forum_id"))

                        If Not(IsNumeric(forumid)) Then

                            Call forum_main():Exit Sub
                            End If

                            sql = "delete from bbs_forum where forum_id=" & forumid
                            conn.execute(sql)
                            sql = "delete from bbs_topic where forum_id=" & forumid
                            conn.execute(sql)
                            sql = "delete from bbs_data where forum_id=" & forumid
                            conn.execute(sql)
                            Response.Write "<script language=javascript>alert(""已成功能删除了一个论坛！\n\n（包括其所属的贴子）"");</script>"
                            Call forum_main()
                        End Sub %>