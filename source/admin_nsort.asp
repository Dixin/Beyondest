<!-- #INCLUDE file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nsort
Dim nsortn
Dim jk_an
tit = vbcrlf & "<a href='?nsort=art'>文栏分类</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='?nsort=down'>下载分类</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='?nsort=news'>新闻分类</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='?nsort=paste'>壁纸分类</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='?nsort=film'>视频分类</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='?nsort=flash'>Flash分类</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='?nsort=web'>网站分类</a>"
Response.Write header(5,tit)
nsort  = Trim(Request.querystring("nsort"))
action = Trim(Request.querystring("action"))

Select Case nsort
    Case "down"
        nsortn = "下载分类"
    Case "news"
        nsortn = "新闻分类"
    Case "web"
        nsortn = "网站分类"
    Case "gall"
        nsortn = "图库分类"
    Case "film"
        nsortn = "视频分类"
    Case "flash"
        nsortn = "Flash分类"
    Case "baner"
        nsortn = "相册分类"
    Case "paste"
        nsortn = "壁纸分类"
    Case Else
        nsort  = "art"
        nsortn = "文栏分类"
End Select

Select Case action
    Case "up","down"
        jk_an = "分类查看"
        Call jk_order()
    Case "del"
        jk_an = "分类查看"
        Call jk_del()
    Case "list"
        jk_an = "分类查看"
        Call jk_list()
    Case "addc"
        jk_an = "添加一级分类"
        Call jk_addc()
    Case "adds"
        jk_an = "添加二级分类"
        Call jk_adds()
    Case "editc"
        jk_an = "修改一级分类"
        Call jk_editc()
    Case "edits"
        jk_an = "修改二级分类"
        Call jk_edits()
    Case Else
        jk_an = "分类查看"
        Call jk_main()
End Select

Response.Write ender()

Sub jk_list()
    Dim i
    Dim j
    Dim cid
    Dim sql2
    Dim rs2:i = 1
    sql         = "select c_id from jk_class where nsort='" & nsort & "' order by c_order,c_id"
    Set rs      = conn.execute(sql)

    Do While Not rs.eof
        cid     = rs(0):j = 1
        conn.execute("update jk_class set c_order=" & i & " where c_id=" & cid)
        sql2    = "select s_id from jk_sort where c_id=" & cid & " order by s_order,s_id"
        Set rs2 = conn.execute(sql2)

        Do While Not rs2.eof
            conn.execute("update jk_sort set s_order=" & j & " where s_id=" & rs2(0))
            rs2.movenext
            j = j + 1
        Loop

        rs2.Close
        rs.movenext
        i = i + 1
    Loop

    rs.Close:Set rs = Nothing:Set rs2 = Nothing
    Call jk_main()
End Sub

Sub jk_del()
    Dim cid
    Dim sid
    cid = Trim(Request.querystring("c_id")):sid = Trim(Request.querystring("s_id"))

    If Not(IsNumeric(cid)) And Not(IsNumeric(sid)) Then Call jk_main():Exit Sub
        If IsNumeric(cid) Then sid = ""

        If sid = "" Then
            sql = "delete from jk_class where c_id=" & cid
            conn.execute(sql)
            sql = "delete from jk_sort where c_id=" & cid
            conn.execute(sql)
        Else
            sql = "delete from jk_sort where s_id=" & sid
            conn.execute(sql)
        End If

        Call jk_main()
    End Sub

    Sub jk_order()
        Dim cid
        Dim sid
        Dim nid
        Dim t1
        Dim t11
        Dim t2
        Dim t22
        Dim sqladd:sqladd = ""
        cid = Trim(Request.querystring("c_id")):sid = Trim(Request.querystring("s_id"))

        If Not(IsNumeric(cid)) And Not(IsNumeric(sid)) Then Call jk_main():Exit Sub
            If IsNumeric(cid) Then sid = ""
            If action = "up" Then sqladd = " desc"

            If sid = "" Then
                t1          = Int(cid)
                sql         = "select c_id,c_order from jk_class where nsort='" & nsort & "' order by c_order" & sqladd & ",c_id" & sqladd
                Set rs      = conn.execute(sql)

                Do While Not rs.eof
                    nid     = Int(rs(0))

                    If Int(cid) = nid Then
                        t22 = rs(1)
                        rs.movenext
                        If rs.eof Then Exit Do
                        t2 = rs(0):t11 = rs(1)
                        conn.execute("update jk_class set c_order=" & t11 & " where c_id=" & t1)
                        conn.execute("update jk_class set c_order=" & t22 & " where c_id=" & t2)
                        Exit Do
                    End If

                    rs.movenext
                Loop

                rs.Close:Set rs = Nothing
            Else
                t1     = Int(sid)
                sql    = "select jk_sort.c_id from jk_class inner join jk_sort on jk_class.c_id=jk_sort.c_id where jk_sort.s_id=" & sid
                Set rs = conn.execute(sql)

                If rs.eof And rs.bof Then
                    rs.Close:Set rs = Nothing

                    Call jk_main():Exit Sub
                    End If

                    cid = Int(rs(0))

                    rs.Close
                    sql         = "select s_id,s_order from jk_sort where c_id=" & cid & " order by s_order" & sqladd & ",s_id" & sqladd
                    Set rs      = conn.execute(sql)

                    Do While Not rs.eof
                        nid     = Int(rs(0))

                        If Int(sid) = nid Then
                            t22 = rs(1)
                            rs.movenext
                            If rs.eof Then Exit Do
                            t2 = rs(0):t11 = rs(1)
                            conn.execute("update jk_sort set s_order=" & t11 & " where s_id=" & t1)
                            conn.execute("update jk_sort set s_order=" & t22 & " where s_id=" & t2)
                            Exit Do
                        End If

                        rs.movenext
                    Loop

                    rs.Close:Set rs = Nothing
                End If

                Call jk_main()
            End Sub

            Sub jk_editc()
                Dim c_name
                Dim cid
                cid        = Trim(Request.querystring("c_id"))

                If Not(IsNumeric(cid)) Then Call jk_main():Exit Sub
                    sql    = "select c_name from jk_class where nsort='" & nsort & "' and c_id=" & cid
                    Set rs = Server.CreateObject("adodb.recordset")
                    rs.open sql,conn,1,3

                    If rs.eof And rs.bof Then
                        rs.Close:Set rs = Nothing

                        Call jk_main():Exit Sub
                        End If

                        Response.Write jk_tit() & "<table border=1 width=350 cellspacing=0 cellpadding=2 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"

                        If Trim(Request.querystring("edit")) = "ok" Then
                            Response.Write vbcrlf & "<tr><td height=100 align=center>"
                            c_name = Replace(Trim(Request.form("c_name")),"'","")

                            If var_null(c_name) = "" Or Len(c_name) > 16 Then
                                Response.Write "<font class=red_2>一级分类名称不能为空（长度不大于16）！</font><br><br>" & go_back
                            Else
                                rs("c_name") = c_name
                                rs.update
                                Response.Write "<font class=red_3>修改一级分类成功！</font><br><br><a href='?nsort=" & nsort & "'>点击返回</a>"
                            End If

                            Response.Write vbcrlf & "</td></tr>"
                        Else %><form action='?nsort=<% Response.Write nsort %>&action=editc&c_id=<% Response.Write cid %>&edit=ok' method=post>
<tr height=50 align=center>
<td>一级分类名称：</td>
<td><input type=text name=c_name value='<% Response.Write rs(0) %>' size=30 maxlength=16></td>
</tr>
<tr><td colspan=2 height=50 align=center><input type=submit value='修改一级分类'></td></tr>
</form><%
                        End If

                        rs.Close:Set rs = Nothing
                        Response.Write "</table>"
                    End Sub

                    Sub jk_edits()
                        Dim s_name
                        Dim pic
                        Dim s_order
                        Dim intro
                        Dim sid
                        Dim cid
                        Dim ccid
                        Dim ncid
                        Dim sqladd
                        sqladd = ""
                        sid    = Trim(Request.querystring("s_id"))
                        If Not(IsNumeric(sid)) Then sid = 0
                        sql    = "select c_id,s_name,pic,intro from jk_sort where s_id=" & sid
                        Set rs = conn.execute(sql)

                        If rs.eof And rs.bof Then
                            rs.Close:Set rs = Nothing

                            Call jk_main():Exit Sub
                            End If

                            cid = rs(0)
                            Response.Write jk_tit() & "<table border=1 width=500 cellspacing=0 cellpadding=2 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"

                            If Trim(Request.querystring("edit")) = "ok" Then
                                Response.Write vbcrlf & "<tr><td height=100 align=center>"
                                ccid   = Trim(Request.form("c_id"))
                                s_name = Replace(Trim(Request.form("s_name")),"'","")
                                pic    = Replace(Trim(Request.form("pic")),"'","")
                                intro  = Replace(Trim(Request.form("intro")),"'","")

                                If Len(s_name) < 1 Or Len(s_name) > 16 Then
                                    Response.Write "<font class=red_2>二级分类名称不能为空（长度不大于16）！</font><br><br>" & go_back
                                Else

                                    If Int(ccid) <> Int(cid) Then
                                        rs.Close
                                        sql         = "select top 1 s_order from jk_sort where c_id=" & ccid & " order by s_order desc"
                                        Set rs      = conn.execute(sql)

                                        If rs.eof And rs.bof Then
                                            s_order = 1
                                        Else
                                            s_order = Int(rs(0)) + 1
                                        End If

                                        sqladd      = ",s_order=" & s_order
                                    End If

                                    sql             = "update jk_sort set intro='" & intro & "',pic='" & pic & "',c_id=" & ccid & ",s_name='" & s_name & "'" & sqladd & " where s_id=" & sid
                                    conn.execute(sql)
                                    Response.Write "<font class=red_3>修改二级分类成功！</font><br><br><a href='?nsort=" & nsort & "'>点击返回</a>"
                                End If

                                Response.Write vbcrlf & "</td></tr>"
                            Else %><form action='?nsort=<% Response.Write nsort %>&action=edits&s_id=<% Response.Write sid %>&edit=ok' method=post>
<tr height=30 align=center>
<td width=100>一级分类类型：</td>
<td><select name=c_id size=1><%
                                pic      = rs(2)
                                intro    = rs(3)
                                s_name   = rs(1):rs.Close
                                sql      = "select c_id,c_name from jk_class where nsort='" & nsort & "' order by c_order,c_id"
                                Set rs   = conn.execute(sql)

                                Do While Not rs.eof
                                    ncid = Int(rs(0))
                                    Response.Write vbcrlf & "<option value='" & ncid & "'"
                                    If cid = ncid Then Response.Write " selected"
                                    Response.Write ">" & rs(1) & "</option>"
                                    rs.movenext
                                Loop %>
</select></td>
</tr>
<tr height=30 align=center>
<td>二级分类名称：</td>
<td><input type=text name=s_name value='<% Response.Write s_name %>' size=30 maxlength=16></td>
</tr>
<tr height=30 align=center>
<td>二级分类图片：</td>
<td><input type=text name=pic value='<% Response.Write pic %>' size=30 maxlength=16></td>
</tr>
<tr height=30 align=center>
<td>二级分类介绍：</td>
<td><textarea rows=6 name=intro cols=70 value=''><% Response.Write intro %></textarea></td>
</tr>
<tr><td colspan=2 height=50 align=center><input type=submit value='修改二级分类'></td></tr>
</form><%
                            End If

                            rs.Close:Set rs = Nothing
                            Response.Write "</table>"
                        End Sub

                        Sub jk_addc()
                            Dim c_name
                            Dim c_order
                            Response.Write jk_tit() & "<table border=1 width=350 cellspacing=0 cellpadding=2 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"

                            If Trim(Request.querystring("add")) = "ok" Then
                                Response.Write vbcrlf & "<tr><td height=100 align=center>"
                                c_name = Replace(Trim(Request.form("c_name")),"'","")

                                If var_null(c_name) = "" Or Len(c_name) > 16 Then
                                    Response.Write "<font class=red_2>一级分类名称不能为空（长度不大于16）！</font><br><br>" & go_back
                                Else
                                    sql         = "select top 1 c_order from jk_class where nsort='" & nsort & "' order by c_order desc"
                                    Set rs      = conn.execute(sql)

                                    If rs.eof And rs.bof Then
                                        c_order = 1
                                    Else
                                        c_order = Int(rs(0)) + 1
                                    End If

                                    rs.Close:Set rs = Nothing
                                    sql = "insert into jk_class(nsort,c_name,c_order) values('" & nsort & "','" & c_name & "'," & c_order & ")"
                                    conn.execute(sql)
                                    Response.Write "<font class=red_3>添加一级分类成功！</font><br><br><a href='?nsort=" & nsort & "'>点击返回</a>"
                                End If

                                Response.Write vbcrlf & "</td></tr>"
                            Else %><form action='?nsort=<% Response.Write nsort %>&action=addc&add=ok' method=post>
<tr height=50 align=center>
<td>一级分类名称：</td>
<td><input type=text name=c_name size=30 maxlength=16></td>
</tr>
<tr><td colspan=2 height=50 align=center><input type=submit value='添加一级分类'></td></tr>
</form><%
                            End If

                            Response.Write "</table>"
                        End Sub

                        Sub jk_adds()
                            Dim s_name
                            Dim s_order
                            Dim cname
                            Dim cid
                            Dim ncid
                            cid = Trim(Request.querystring("c_id"))
                            If Not(IsNumeric(cid)) Then cid = 0
                            cid = Int(cid)
                            Response.Write jk_tit() & "<table border=1 width=350 cellspacing=0 cellpadding=2 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"

                            If Trim(Request.querystring("add")) = "ok" Then
                                Response.Write vbcrlf & "<tr><td height=100 align=center>"
                                s_name = Replace(Trim(Request.form("s_name")),"'","")

                                If Len(s_name) < 1 Or Len(s_name) > 16 Then
                                    Response.Write "<font class=red_2>二级分类名称不能为空（长度不大于16）！</font><br><br>" & go_back
                                Else
                                    cid    = Trim(Request.form("c_id"))
                                    If Not(IsNumeric(cid)) Then cid = 0
                                    sql    = "select c_name from jk_class where nsort='" & nsort & "' and c_id=" & cid
                                    Set rs = conn.execute(sql)

                                    If rs.eof And rs.bof Then
                                        rs.Close:Set rs = Nothing

                                        Call jk_main():Exit Sub
                                        End If

                                        cname = rs(0)
                                        rs.Close

                                        sql         = "select top 1 s_order from jk_sort where c_id=" & cid & " order by s_order desc"
                                        Set rs      = conn.execute(sql)

                                        If rs.eof And rs.bof Then
                                            s_order = 1
                                        Else
                                            s_order = Int(rs(0)) + 1
                                        End If

                                        rs.Close:Set rs = Nothing

                                        sql = "insert into jk_sort(c_id,s_name,s_order) values(" & cid & ",'" & s_name & "'," & s_order & ")"
                                        conn.execute(sql)
                                        Response.Write "<font class=red_3>添加二级分类成功！</font><br><br><a href='?nsort=" & nsort & "'>点击返回</a>"
                                    End If

                                    Response.Write vbcrlf & "</td></tr>"
                                Else %><form action='?nsort=<% Response.Write nsort %>&action=adds&c_id=<% Response.Write cid %>&add=ok' method=post>
<tr height=30 align=center>
<td>一级分类类型：</td>
<td><select name=c_id size=1><%
                                    sql      = "select c_id,c_name from jk_class where nsort='" & nsort & "' order by c_order"
                                    Set rs   = conn.execute(sql)

                                    Do While Not rs.eof
                                        ncid = Int(rs(0))
                                        Response.Write vbcrlf & "<option value='" & ncid & "'"
                                        If cid = ncid Then Response.Write " selected"
                                        Response.Write ">" & rs(1) & "</option>"
                                        rs.movenext
                                    Loop

                                    rs.Close:Set rs = Nothing %>
</select></td>
</tr>
<tr height=30 align=center>
<td>二级分类名称：</td>
<td><input type=text name=s_name size=30 maxlength=16></td>
</tr>
<tr><td colspan=2 height=50 align=center><input type=submit value='添加二级分类'></td></tr>
</form><%
                                End If

                                Response.Write "</table>"
                            End Sub

                            Sub jk_main()
                                Response.Write jk_tit()
                                Dim sql2
                                Dim rs2
                                Dim cid
                                Dim sid
                                Response.Write vbcrlf & "<table border=1 cellspacing=0 cellpadding=2 width=400 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>"
                                sql         = "select c_id,c_name from jk_class where nsort='" & nsort & "' order by c_order,c_id"
                                Set rs      = conn.execute(sql)

                                Do While Not rs.eof
                                    cid     = rs(0)
                                    Response.Write vbcrlf & "<tr bgcolor=#ffffff align=center><td align=left>&nbsp;<font class=red_3><b>" & img_small("jt1") & rs(1) & "</b></font>&nbsp;&nbsp;（<a href='?nsort=" & nsort & "&action=adds&c_id=" & cid & "'>添加二级分类</a>）</td><td><a href='?nsort=" & nsort & "&action=editc&c_id=" & cid & "'>修改</a>&nbsp;&nbsp;<a href=""javascript:Do_del_class('" & cid & "');"">删除</a></td><td>排序：<a href='?nsort=" & nsort & "&action=up&c_id=" & cid & "'>向上</a>&nbsp;&nbsp;<a href='?nsort=" & nsort & "&action=down&c_id=" & cid & "'>向下</a></td></tr>"
                                    sql2    = "select s_id,s_name from jk_sort where c_id=" & cid & " order by s_order,s_id"
                                    Set rs2 = conn.execute(sql2)

                                    Do While Not rs2.eof
                                        sid = rs2(0)
                                        Response.Write vbcrlf & "<tr align=center><td align=left>　　<font class=blue>" & rs2(1) & "</font></td><td><a href='?nsort=" & nsort & "&action=edits&s_id=" & sid & "'>修改</a>&nbsp;&nbsp;<a href=""javascript:Do_del_sort('" & sid & "');"">删除</a></td><td>排序：<a href='?nsort=" & nsort & "&action=up&s_id=" & sid & "'>向上</a>&nbsp;&nbsp;<a href='?nsort=" & nsort & "&action=down&s_id=" & sid & "'>向下</a></td></tr>"
                                        rs2.movenext
                                    Loop

                                    rs2.Close:Set rs2 = Nothing
                                    rs.movenext
                                Loop

                                rs.Close:Set rs = Nothing
                                Response.Write vbcrlf & "<tr><td height=30 align=center colspan=3><a href='?nsort=" & nsort & "&action=addc'>添加一级分类</a>&nbsp;&nbsp;-&nbsp;&nbsp;<a href='?nsort=" & nsort & "&action=list'>重新排序</a></td></tr></table>" %><script language=JavaScript>
<!--
function Do_del_class(data1)
{
if (confirm("此操作将删除id为 "+data1+" 的一级分类！\n\n真的要删除吗？\n\n删除后将无法恢复！"))
  window.location="?nsort=<% Response.Write nsort %>&action=del&c_id="+data1
}

function Do_del_sort(data1)
{
if (confirm("此操作将删除id为 "+data1+" 的二级分类！\n\n真的要删除吗？\n\n删除后将无法恢复！"))
  window.location="?nsort=<% Response.Write nsort %>&action=del&s_id="+data1
}
//-->
</script><%
                            End Sub

                            Function jk_tit()
                                jk_tit = vbcrlf & "<table border=0><tr><td height=30><font class=red>" & nsortn & "</font>&nbsp;&nbsp;-&nbsp;&nbsp;<font class=blue>" & jk_an & "</font></td></tr></table>" & vbcrlf
                            End Function %>