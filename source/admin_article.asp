<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nsort,sql2,rs2,del_temp,data_name,cid,sid,nid,ncid,nsid,id,left_type,now_id,nummer,sqladd,page,rssum,thepages,viewpage,pageurl,topic,csid
tit = vbcrlf & "<a href='?'>文栏管理</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='?action=add'>添加文章</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='admin_nsort.asp?nsort=article'>文栏分类</a>"
Response.Write header(13,tit)
pageurl = "?action=" & action & "&":nsort = "art":data_name = "article":sqladd = "":nummer = 15
Call admin_cid_sid()

If Trim(Request("del_ok")) = "ok" Then
    Response.Write del_select(Trim(Request.form("del_id")))
End If

id         = Trim(Request.querystring("id"))

If (action = "hidden" Or action = "istop") And IsNumeric(id) Then
    sql    = "select " & action & " from " & data_name & " where id=" & id
    Set rs = conn.execute(sql)

    If Not(rs.eof And rs.bof) Then

        If action = "istop" Then

            If Int(rs(action)) = 1 Then
                sql = "update " & data_name & " set " & action & "=0 where id=" & id
            Else
                sql = "update " & data_name & " set " & action & "=1 where id=" & id
            End If

        Else

            If rs(action) = True Then
                sql = "update " & data_name & " set " & action & "=0 where id=" & id
            Else
                sql = "update " & data_name & " set " & action & "=1 where id=" & id
            End If

        End If

        conn.execute(sql)
    End If

    rs.Close:action = ""
End If

Select Case action
    Case "add"
        Call news_add()
    Case "edit"

        If Not(IsNumeric(id)) Then
            Call news_main()
        Else
            Set rs = Server.CreateObject("adodb.recordset")
            sql    = "select * from " & data_name & " where id=" & id
            rs.open sql,conn,1,3
            Call news_edit()
        End If

    Case Else
        Call news_main()
End Select

Call close_conn()
Response.Write ender()

Sub news_edit()

    If Trim(Request.querystring("edit")) = "chk" Then
        topic = code_admin(Request.form("topic"))
        csid  = Trim(Request.form("csid"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>请选择文章类型！</font><br><br>" & go_back
        ElseIf topic = "" Then
            Response.Write "<font class=red_2>文章标题不能为空！</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            rs("c_id")     = cid
            rs("s_id")     = sid
            If Trim(Request.form("username_my")) = "yes" Then rs("username") = login_username
            rs("topic")     = topic
            rs("word")     = Request.form("word")

            If IsNumeric(Trim(Request.form("emoney"))) Then
                rs("emoney") = Trim(Request.form("emoney"))
            Else
                rs("emoney") = 0
            End If

            rs("author")     = code_admin(Request.form("author"))
            rs("power")     = Replace(Replace(Trim(Request.form("power"))," ",""),",",".")
            rs("keyes")     = code_admin(Request.form("keyes"))

            If Trim(Request.form("istop")) = "yes" Then
                rs("istop") = 1
            Else
                rs("istop") = 0
            End If

            If Trim(Request.form("hidden")) = "yes" Then
                rs("hidden") = False
            Else
                rs("hidden") = True
            End If

            If IsNumeric(Trim(Request.form("counter"))) Then rs("counter") = Trim(Request.form("counter"))
            rs.update
            rs.Close:Set rs = Nothing
            Call upload_note(data_name,id)
            Response.Write "<font class=red>已成功修改了一篇文章！</font><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "'>点击返回</a><br><br>"
        End If

    Else
        Dim sql3,rs3 %><table border=0 width='98%' cellspacing=0 cellpadding=1>
<form name='add_frm' action='<% Response.Write pageurl %>c_id=<% Response.Write cid %>&s_id=<% Response.Write sid %>&id=<% Response.Write id %>&edit=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>文章标题：</td><td width='85%'><input type=text size=70 name=topic value='<% = rs("topic") %>' maxlength=40><% = redx %></td></tr>
  <tr><td align=center>文章类型：</td><td><% Call chk_csid(cid,sid):Call chk_emoney(rs("emoney")):Call chk_h_u() %></td></tr>
  <tr><td align=center>浏览权限：</td><td><% Call chk_power(rs("power"),0) %></td></tr>
  <tr><td align=center>文章作者：</td><td><input type=text size=12 name=author value='<% Response.Write rs("author") %>' maxlength=20>&nbsp;&nbsp;关键字：<input type=text name=keyes value='<% Response.Write rs("keyes") %>' size=12 maxlength=20>&nbsp;&nbsp;推荐：<input type=checkbox name=istop value='yes'<% If Int(rs("istop")) = 1 Then Response.Write " checked" %>>&nbsp;&nbsp;人次：<input type=text name=counter value='<% Response.Write rs("counter") %>' size=10 maxlength=10></td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","word","&nbsp;&nbsp;") %></td></tr>
  <tr><td valign=top align=center><br>文章内容：</td><td><textarea name=word rows=15 cols=70><% = rs("word") %></textarea></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=article&upname=a&uptext=word'></iframe></td></tr>
  <tr height=25><td></td><td><input type=submit value=' 修 改 文 章 '></td></tr>
</form>
</table><%
    End If

End Sub

Sub news_add()

    If Trim(Request.querystring("add")) = "chk" Then
        topic = code_admin(Request.form("topic"))
        csid  = Trim(Request.form("csid"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>请选择文章类型！</font><br><br>" & go_back
        ElseIf topic = "" Then
            Response.Write "<font class=red_2>文章标题不能为空！</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            Set rs = Server.CreateObject("adodb.recordset")
            sql    = "select * from " & data_name
            rs.open sql,conn,1,3
            rs.addnew
            rs("c_id")     = cid
            rs("s_id")     = sid
            rs("username")     = login_username
            rs("hidden")     = True
            rs("topic")     = topic
            rs("word")     = Request.form("word")

            If IsNumeric(Trim(Request.form("emoney"))) Then
                rs("emoney") = Trim(Request.form("emoney"))
            Else
                rs("emoney") = 0
            End If

            rs("author")     = code_admin(Request.form("author"))
            rs("power")     = Replace(Replace(Trim(Request.form("power"))," ",""),",",".")
            rs("keyes")     = code_admin(Request.form("keyes"))

            If Trim(Request.form("istop")) = "yes" Then
                rs("istop") = 1
            Else
                rs("istop") = 0
            End If

            rs("tim")     = now_time
            rs("counter")     = 0
            rs.update
            rs.Close:Set rs = Nothing
            Call upload_note(data_name,first_id(data_name))
            Response.Write "<font class=red>已成功添加了一篇文章！</font><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "'>点击返回</a><br><br>"
        End If

    Else %><table border=0 width='98%' cellspacing=0 cellpadding=1>
<form name='add_frm' action='<% Response.Write pageurl %>add=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>文章标题：</td><td width='85%'><input type=text size=70 name=topic maxlength=40><% = redx %></td></tr>
  <tr><td align=center>文章类型：</td><td><% Call chk_csid(cid,sid):Call chk_emoney(0) %></td></tr>
  <tr><td align=center>浏览权限：</td><td><% Call chk_power("",1) %></td></tr>
  <tr><td align=center>文章作者：</td><td><input type=text size=12 name=author maxlength=20>&nbsp;&nbsp;关键字：<input type=text name=keyes size=12 maxlength=20>&nbsp;&nbsp;推荐：<input type=checkbox name=istop value='yes'></td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","word","&nbsp;&nbsp;") %></td></tr>
  <tr><td valign=top align=center><br>文章内容：</td><td><textarea name=word rows=15 cols=70></textarea></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=article&upname=a&uptext=word'></iframe></td></tr>
  <tr><td></td><td height=25><input type=submit value=' 添 加 文 章 '></td></tr>
</form></table><%
    End If

End Sub

Sub news_main() %>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='100%' cellpadding=2>
  <tr valign=top height=350>
    <td width='25%' class=htd><br><% Call left_sort() %></td>
    <td width='75%' align=center>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<form name=del_form action='<% = pageurl %>del_ok=ok' method=post>
<tr><td width='6%'></td><td width='81%'></td><td width='13%'></td></tr>
<%
    Call sql_cid_sid()
    sql    = "select id,c_id,s_id,topic,hidden,istop from " & data_name & sqladd & " order by id desc"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql,conn,1,1

    If rs.eof And rs.bof Then
        rssum = 0
    Else
        rssum = rs.recordcount
    End If

    Call format_pagecute()
    del_temp = nummer
    If rssum = 0 Then del_temp = 0

    If Int(page) = Int(thepages) Then
        del_temp = rssum - nummer*(thepages - 1)
    End If %>
<tr><td colspan=3 align=center height=25>
现有<font class=red><% Response.Write rssum %></font>篇文章　<% Response.Write "<a href='?action=add&c_id=" & cid & "&s_id=" & sid & "'>添加文章</a>" %>
　<input type=checkbox name=del_all value=1 onClick=selectall('<% Response.Write del_temp %>')> 选中所有　<input type=submit value='删除所选' onclick=""return suredel('<% Response.Write del_temp %>');"">
</td></tr>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<%

    If Int(viewpage) <> 1 Then
        rs.move (viewpage - 1)*nummer
    End If

    For i = 1 To nummer
        If rs.eof Then Exit For
        now_id = rs("id"):ncid = rs("c_id"):nsid = rs("s_id")
        Response.Write article_center()
        rs.movenext
    Next

    rs.Close:Set rs = Nothing %></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %>
</td></tr></table>
</td></tr></table>
<%
End Sub

Function article_center()
    article_center = VbCrLf & "<tr" & mtr & ">" & _
    VbCrLf & "<td>" & i + (viewpage - 1)*nummer & ". </td><td>" & _
    VbCrLf & "<a href='?action=edit&c_id=" & ncid & "&s_id=" & nsid & "&id=" & now_id & "'>" & cuted(rs("topic"),30) & "</a>" & _
    "</td><td align=right><a href='?action=hidden&c_id=" & cid & "&s_id=" & sid & "&id=" & now_id & "&page=" & viewpage & "'>"

    If rs("hidden") = True Then
        article_center = article_center & "显"
    Else
        article_center = article_center & "<font class=red_2>隐</font>"
    End If

    article_center = article_center & "</a> <a href='?action=istop&c_id=" & cid & "&s_id=" & sid & "&id=" & now_id & "&page=" & viewpage & "'>"

    If Int(rs("istop")) = 1 Then
        article_center = article_center & "<font class=red>是</font>"
    Else
        article_center = article_center & "否"
    End If

    article_center = article_center & "</a> <input type=checkbox name=del_id value='" & now_id & "' class=bg_1></td></tr>"
End Function %>