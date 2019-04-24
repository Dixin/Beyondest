<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nsort,sql2,rs2,del_temp,data_name,cid,sid,nid,ncid,nsid,id,left_type,now_id,nummer,sqladd,page,rssum,thepages,viewpage,pageurl,csid
tit = vbcrlf & "<a href='?'>行业动态</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='?action=add'>发布新闻</a>&nbsp;┋&nbsp;" & _
vbcrlf & "<a href='admin_nsort.asp?nsort=news'>新闻分类</a>"
Response.Write header(12,tit)
pageurl = "?action=" & action & "&":nsort = "news":data_name = "news":sqladd = "":nummer = 15
Call admin_cid_sid()

If Trim(Request("del_ok")) = "ok" Then
    Response.Write del_select(Trim(Request.form("del_id")))
End If

id         = Trim(Request.querystring("id"))

If (action = "hidden" Or action = "istop") And IsNumeric(id) Then
    sql    = "select " & action & " from " & data_name & " where id=" & id
    Set rs = conn.execute(sql)

    If Not(rs.eof And rs.bof) Then

        If rs(action) = True Then
            sql = "update " & data_name & " set " & action & "=0 where id=" & id
        Else
            sql = "update " & data_name & " set " & action & "=1 where id=" & id
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
    Dim rs3,sql3,topic,comto,istop,word,ispic,pic,keyes

    If Trim(Request.querystring("edit")) = "chk" Then
        topic = code_admin(Request.form("topic"))
        csid  = Trim(Request.form("csid"))
        comto = code_admin(Request.form("comto"))
        keyes = code_admin(Request.form("keyes"))
        istop = Trim(Request.form("istop"))
        word  = Request.form("word")
        ispic = Trim(Request.form("ispic"))
        pic   = Trim(Request.form("pic"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>请选择新闻类型！</font><br><br>" & go_back
        ElseIf Len(topic) < 1 Or Len(word) < 10 Then
            Response.Write "<font class=red_2>新闻标题和内容不能为空！</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            rs("c_id")     = cid
            rs("s_id")     = sid
            If Trim(Request.form("username_my")) = "yes" Then rs("username") = login_username
            rs("topic")     = topic
            rs("comto")     = comto
            rs("keyes")     = keyes
            rs("word")     = word

            If istop = "yes" Then
                rs("istop") = True
            Else
                rs("istop") = False
            End If

            If ispic = "yes" Then
                rs("ispic") = True
            Else
                rs("ispic") = False
            End If

            If Trim(Request.form("hidden")) = "yes" Then
                rs("hidden") = False
            Else
                rs("hidden") = True
            End If

            rs("pic")     = pic
            If IsNumeric(Trim(Request.form("counter"))) Then rs("counter") = Trim(Request.form("counter"))
            rs.update
            rs.Close:Set rs = Nothing
            Call upload_note(data_name,id)
            Response.Write "<font class=red>已成功修改了一篇新闻！</font><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "'>点击返回</a><br><br>"
        End If

    Else %><table border=0 width='98%' cellspacing=0 cellpadding=1>
<form name='add_frm' action='<% Response.Write pageurl %>c_id=<% Response.Write cid %>&s_id=<% Response.Write sid %>&id=<% Response.Write id %>&edit=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>新闻标题：</td><td width='85%'><input type=text size=70 name=topic value='<% Response.Write rs("topic") %>' maxlength=100><% = redx %></td></tr>
  <tr><td align=center>新闻类型：</td><td><% Call chk_csid(cid,sid) %>&nbsp;&nbsp;&nbsp;出处：<input type=text size=20 name=comto value='<% Response.Write rs("comto") %>' maxlength=10>&nbsp;&nbsp;&nbsp;<input type=checkbox name=username_my value='yes'>&nbsp;<font alt='发布人：<% Response.Write rs("username") %>'>修改发布人为我</font></td></tr>
<%
        pic = rs("pic"):ispic = pic
        If InStr(ispic,"/") > 0 Then ispic = Right(ispic,Len(ispic) - InStr(ispic,"/"))
        If InStr(ispic,".") > 0 Then ispic = Left(ispic,InStr(ispic,".") - 1)
        If Len(ispic) < 1 Then ispic = "n" & upload_time(now_time) %>  <tr><td align=center>关 键 字：</td><td><input type=text size=20 name=keyes value='<% Response.Write rs("keyes") %>' maxlength=20>&nbsp;&nbsp;&nbsp;推荐：<input type=checkbox name=istop<% If rs("istop") = True Then Response.Write " checked" %> value='yes'>&nbsp;选为推荐显示&nbsp;&nbsp;&nbsp;隐藏：<input type=checkbox name=hidden<% If rs("hidden") = False Then Response.Write " checked" %> value='yes'>&nbsp;选为隐藏显示</td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","word","&nbsp;&nbsp;") %></td></tr>
  <tr><td align=center valign=top><br>新闻内容：</td><td><textarea name=word rows=15 cols=70><% Response.Write rs("word") %></textarea></td></tr>
  <tr><td align=center>图片新闻：</td><td><input type=checkbox name=ispic<% If rs("ispic") = True Then Response.Write " checked" %> value='yes'>&nbsp;选为图片新闻&nbsp;&nbsp;&nbsp;图片：<input type=test name=pic value='<% Response.Write pic %>' size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=<% Response.Write ispic %>&uptext=pic' target=upload_frame>上传图片</a>&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=n&uptext=word' target=upload_frame>上传至内容</a></td></tr>
  <tr><td align=center>上传图片：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=news&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' 修 改 新 闻 '></td></tr>
</form></table><%
    End If

End Sub

Sub news_add()

    If Trim(Request.querystring("add")) = "chk" Then
        Dim topic,comto,istop,word,ispic,pic,keyes
        topic = code_admin(Request.form("topic"))
        csid  = Trim(Request.form("csid"))
        comto = code_admin(Request.form("comto"))
        keyes = code_admin(Request.form("keyes"))
        istop = Trim(Request.form("istop"))
        word  = Request.form("word")
        ispic = Trim(Request.form("ispic"))
        pic   = Trim(Request.form("pic"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>请选择新闻类型！</font><br><br>" & go_back
        ElseIf Len(topic) < 1 Or Len(word) < 10 Then
            Response.Write "<font class=red_2>新闻标题和内容不能为空！</font><br><br>" & go_back
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
            rs("comto")     = comto
            rs("keyes")     = keyes
            rs("word")     = word

            If istop = "yes" Then
                rs("istop") = True
            Else
                rs("istop") = False
            End If

            If ispic = "yes" Then
                rs("ispic") = True
            Else
                rs("ispic") = False
            End If

            rs("pic")     = pic
            rs("tim")     = now_time
            rs("counter")     = 0
            rs.update
            rs.Close:Set rs = Nothing
            Call upload_note(data_name,first_id(data_name))
            Response.Write "<font class=red>已成功发布了一篇新闻！</font><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "'>点击返回</a><br><br>"
        End If

    Else %><table border=0 width='98%' cellspacing=0 cellpadding=1>
<form name='add_frm' action='<% Response.Write pageurl %>add=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>新闻标题：</td><td width='85%'><input type=text size=70 name=topic maxlength=100><% = redx %></td></tr>
  <tr><td align=center>新闻类型：</td><td><% Call chk_csid(cid,sid) %>&nbsp;&nbsp;&nbsp;&nbsp;出处：<input type=text size=30 name=comto maxlength=10></td></tr>
  <tr><td align=center>关 键 字：</td><td><input type=text size=20 name=keyes maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;推荐：<input type=checkbox name=istop value='yes'>&nbsp;选上为新闻首页显示</td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","word","&nbsp;&nbsp;") %></td></tr>
  <tr><td valign=top align=center><br>新闻内容：</td><td><textarea name=word rows=15 cols=70></textarea></td></tr>
<% ispic = "n" & upload_time(now_time) %>
  <tr><td align=center>图片新闻：</td><td><input type=checkbox name=ispic value='yes'>&nbsp;&nbsp;&nbsp;&nbsp;图片：<input type=test name=pic size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=<% Response.Write ispic %>&uptext=pic' target=upload_frame>上传图片</a>&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=n&uptext=word' target=upload_frame>上传至内容</a></td></tr>
  <tr><td align=center>上传图片：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=news&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' 添 加 新 闻 '></td></tr>
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
    sql    = "select id,c_id,s_id,topic,istop,hidden from " & data_name & sqladd & " order by id desc"
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
现有<font class=red><% Response.Write rssum %></font>篇新闻　<% Response.Write "<a href='?action=add&c_id=" & cid & "&s_id=" & sid & "'>添加新闻</a>" %>
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
        Response.Write news_center()
        rs.movenext
    Next

    rs.Close:Set rs = Nothing %></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %>
</td></tr>
</table>
    </td>
  </tr>
</table>
<%
End Sub

Function news_center()
    news_center = VbCrLf & "<tr" & mtr & ">" & _
    VbCrLf & "<td>" & i + (viewpage - 1)*nummer & ". </td><td>" & _
    VbCrLf & "<a href='?action=edit&c_id=" & ncid & "&s_id=" & nsid & "&id=" & now_id & "'>" & cuted(rs("topic"),30) & "</a>" & _
    "</td><td align=right><a href='?action=hidden&c_id=" & cid & "&s_id=" & sid & "&id=" & now_id & "&page=" & viewpage & "'>"

    If rs("hidden") = True Then
        news_center = news_center & "显"
    Else
        news_center = news_center & "<font class=red_2>隐</font>"
    End If

    news_center = news_center & "</a> <a href='?action=istop&c_id=" & cid & "&s_id=" & sid & "&id=" & now_id & "&page=" & viewpage & "'>"

    If rs("istop") = True Then
        news_center = news_center & "<font class=red>是</font>"
    Else
        news_center = news_center & "否"
    End If

    news_center = news_center & "</a><input type=checkbox name=del_id value='" & now_id & "'></td></tr>"
End Function %>