<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nsort
Dim rs2
Dim sql2
Dim id
Dim j
Dim sqladd
Dim cid
Dim sid
Dim ncid
Dim nsid
Dim nid
Dim now_id
Dim power
Dim pageurl
Dim ispic
Dim data_name
Dim nummer
Dim ddim
Dim genre
Dim os
Dim rssum
Dim thepages
Dim page
Dim viewpage
Dim del_temp
Dim csid
tit = "<a href='?action='>音乐列表</a>&nbsp;┋&nbsp;" & _
"<a href='?action=add'>添加音乐</a>&nbsp;┋&nbsp;" & _
"<a href='admin_nsort.asp?nsort=down'>下载分类</a>&nbsp;┋&nbsp;" & _
"<a href='?action=code'>注册码列表</a>&nbsp;┋&nbsp;" & _
"<a href='?action=code_add'>添加注册码</a>"
Response.Write header(14,tit)
pageurl = "?action=" & action & "&":nsort = "down":sqladd = "":data_name = "down":sqladd = "":nummer = 15
Call admin_cid_sid()

If Trim(Request.querystring("del_ok")) = "ok" Then
    Response.Write del_select(Request.form("del_id"))
End If

id         = Trim(Request.querystring("id"))

If action = "hidden" And IsNumeric(id) Then
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

    rs.Close
    action = ""
End If

Select Case action
    Case "add"
        Call down_add()
    Case "down_edit"
        Call down_edit()
    Case "code"
        Call code_main()
    Case "code_add"
        Call code_add()
    Case "code_edit"
        Call code_edit()
    Case "code_del"
        Call code_del()
    Case Else
        Call down_main()
End Select

close_conn
Response.Write ender()

Sub down_edit()
    Dim sql3
    Dim rs3
    Dim id
    Dim name
    Dim sizes
    Dim url
    Dim url2
    Dim homepage
    Dim remark
    Dim counter
    Dim types
    Dim keyes
    Dim pic
    id = Request.querystring("id")

    If Not(IsNumeric(id)) Then Call down_main():Exit Sub %><table border=0 width=600 cellspacing=0 cellpadding=0>
<tr><td align=center height=300>
<%
        sql    = "select * from " & data_name & " where id=" & id
        Set rs = Server.CreateObject("adodb.recordset")
        rs.open sql,conn,1,3

        If rs.eof And rs.bof Then
            rs.Close:Set rs = Nothing

            Call down_main():Exit Sub
            End If

            If Trim(Request.querystring("types")) = "edit" Then
                csid     = Trim(Request.form("csid"))
                name     = code_admin(Request.form("name"))
                sizes    = code_admin(Request.form("sizes"))
                url      = code_admin(Request.form("url"))
                url2     = code_admin(Request.form("url2"))
                pic      = Request.form("pic")
                If Len(pic) < 3 Then pic = "no_pic.gif"
                homepage = code_admin(Request.form("homepage"))
                keyes    = code_admin(Request.form("keyes"))
                remark   = Request.form("remark")
                counter  = Trim(Request.form("counter"))
                types    = Request.form("types")

                If Len(csid) < 1 Or var_null(name) = "" Or var_null(url) = "" Then
                    Response.Write("<font class=red_2>软件的类型、名称和下载地址不能为空！</font><br><br>" & go_back)
                Else
                    Call chk_cid_sid()
                    rs("c_id") = cid
                    rs("s_id") = sid
                    If Trim(Request.form("username_my")) = "yes" Then rs("username") = login_username

                    If Trim(Request.form("hidden")) = "yes" Then
                        rs("hidden") = False
                    Else
                        rs("hidden") = True
                    End If

                    rs("name")     = name
                    rs("sizes")     = sizes

                    If IsNumeric(Trim(Request.form("emoney"))) Then
                        rs("emoney") = Trim(Request.form("emoney"))
                    Else
                        rs("emoney") = 0
                    End If

                    rs("genre")     = Trim(Request.form("genre"))
                    rs("os")     = Replace(Trim(Request.form("os"))," ","")
                    rs("power")     = Replace(Replace(Trim(Request.form("power"))," ",""),",",".")
                    rs("url")     = url
                    rs("url2")     = url2
                    rs("homepage")     = homepage
                    rs("remark")     = remark
                    rs("keyes")     = keyes
                    rs("pic")     = pic
                    rs("tim")     = now_time
                    rs("types")     = types
                    If IsNumeric(counter) Then rs("counter") = counter
                    rs.update
                    Call upload_note(data_name,id)
                    Response.Write "<font class=red>软件修改成功！</font><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "'>点击返回</a>" & _
                    vbcrlf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=?c_id=" & cid & "&s_id=" & sid & "'>"
                End If

            Else
                cid   = Int(rs("c_id")):sid = Int(rs("s_id")):types = Int(rs("types"))
                ispic = rs("pic"):pic = ispic
                If InStr(ispic,"/") > 0 Then ispic = Right(ispic,Len(ispic) - InStr(ispic,"/"))
                If InStr(ispic,".") > 0 Then ispic = Left(ispic,InStr(ispic,".") - 1)
                If ispic = "no_pic" Then ispic = "n_" & id:pic = "" %><table border=0 width='98%' cellspacing=0 cellpadding=2>
  <tr><td colspan=2 height=50 align=center><font class=red>软件下载属性修改</font></td></tr>
<form name='add_frm' action="?action=down_edit&types=edit&id=<% Response.Write id %>" method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>音乐名称：</td><td width='85%'><input type=text name=name value='<% Response.Write rs("name") %>' size=70 maxlength=40><% Response.Write redx %></td></tr>
  <tr><td align=center>所属类别：</td><td><% Call chk_csid(cid,sid):Call chk_emoney(rs("emoney")):Call chk_h_u() %></td></tr>
  <tr><td align=center>下载权限：</td><td><% Call chk_power(rs("power"),0) %></td></tr>
  <tr><td align=center>文件大小：</td><td><input type=text name=sizes value='<% Response.Write rs("sizes") %>' size=20 maxlength=10>&nbsp;&nbsp;&nbsp;推荐等级：<select name=types size=1>
<option value='0'<% If types = 0 Then Response.Write " selected" %>>没有等级</option>
<option value='1'<% If types = 1 Then Response.Write " selected" %>>一星级</option>
<option value='2'<% If types = 2 Then Response.Write " selected" %>>二星级</option>
<option value='3'<% If types = 3 Then Response.Write " selected" %>>三星级</option>
<option value='4'<% If types = 4 Then Response.Write " selected" %>>四星级</option>
<option value='5'<% If types = 5 Then Response.Write " selected" %>>五星级</option>
</select>&nbsp;&nbsp;&nbsp;音乐类型：<select name=genre size=1><%
                Dim tt:tt = rs("genre"):ddim = Split(web_var(web_down
                Dim 4)
                Dim ":")

                For i = 0 To UBound(ddim)
                    Response.Write vbcrlf & "<option"
                    If tt = ddim(i) Then Response.Write " selected"
                    Response.Write ">" & ddim(i) & "</option>"
                Next

                Erase ddim %></select></td></tr>
  <tr><td align=center>播放软件：</td><td><%
                tt = rs("os"):ddim = Split(web_var(web_down,3),":")

                For i = 0 To UBound(ddim)
                    Response.Write "<input type=checkbox name=os value='" & ddim(i) & "'"
                    If InStr(1,tt,ddim(i)) > 0 Then Response.Write " checked"
                    Response.Write " class=bg_1>" & ddim(i)
                Next

                Erase ddim %></td></tr>
  <tr><td align=center>下载地址1：</td><td><input type=text name=url value='<% Response.Write rs("url") %>' size=70 maxlength=200><% Response.Write redx %></td></tr>
  <tr><td align=center>下载地址2：</td><td><input type=text name=url2 value='<% Response.Write rs("url2") %>' size=70 maxlength=200></td></tr>
  <tr><td align=center>文件来自：</td><td><input type=text name=homepage value='<% Response.Write rs("homepage") %>' size=50 maxlength=50>&nbsp;&nbsp;&nbsp;下载次数：<input type=text name=counter value='<% Response.Write rs("counter") %>' size=4 maxlength=10></td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","remark","&nbsp;&nbsp;") %></td></tr>
  <tr><td align=center valign=top><br>音乐备注：</td><td><% Response.Write("<textarea rows=6 name=remark cols=70>" & rs("remark") & "</textarea>") %></td></tr>
  <tr><td align=center>关 键 字：</td><td><input type=text name=keyes value='<% Response.Write rs("keyes") %>' size=12 maxlength=20>&nbsp;图片：<input type=test name=pic value='<% If ispic <> "no_pic.gif" Then Response.Write pic %>' size=30 maxlength=100>&nbsp;<a href='upload.asp?uppath=down&upname=<% Response.Write ispic %>&uptext=pic' target=upload_frame>上传图片</a>&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=d&uptext=remark' target=upload_frame>上传至内容</a></td></tr>
  <tr><td align=center>上传图片：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=down&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr height=30><td></td><td><input type=submit value=' 提 交 修 改 '></td></tr>
</form></table><%
            End If

            rs.Close:Set rs = Nothing %></td></tr></table><%
        End Sub

        Sub down_add() %><table border="0" width="600" cellspacing="0" cellpadding="0">
<tr><td align=center height=300><%

            If Request.querystring("types") = "add" Then
                Dim name
                Dim sizes
                Dim url
                Dim url2
                Dim homepage
                Dim remark
                Dim types
                Dim keyes
                Dim pic
                csid     = Trim(Request.form("csid"))
                name     = code_admin(Request.form("name"))
                sizes    = code_admin(Request.form("sizes"))
                url      = code_admin(Request.form("url"))
                url2     = code_admin(Request.form("url2"))
                homepage = code_admin(Request.form("homepage"))
                keyes    = code_admin(Request.form("keyes"))
                remark   = Request.form("remark")
                pic      = Request.form("pic")
                If Len(pic) < 3 Then pic = "no_pic.gif"
                types    = Request.form("types")

                If Len(csid) < 1 Or var_null(name) = "" Or var_null(url) = "" Then
                    Response.Write("<font class=red_2>文件的类型、名称和下载地址不能为空！</font><br><br>" & go_back)
                Else
                    Call chk_cid_sid()
                    sql    = "select * from " & data_name
                    Set rs = Server.CreateObject("adodb.recordset")
                    rs.open sql,conn,1,3
                    rs.addnew
                    rs("c_id")     = cid
                    rs("s_id")     = sid
                    rs("username")     = login_username
                    rs("hidden")     = True
                    rs("name")     = name
                    rs("sizes")     = sizes

                    If IsNumeric(Trim(Request.form("emoney"))) Then
                        rs("emoney") = Trim(Request.form("emoney"))
                    Else
                        rs("emoney") = 0
                    End If

                    rs("genre")     = Trim(Request.form("genre"))
                    rs("os")     = Replace(Trim(Request.form("os"))," ","")
                    rs("power")     = Replace(Replace(Trim(Request.form("power"))," ",""),",",".")
                    rs("url")     = url
                    rs("url2")     = url2
                    rs("homepage")     = homepage
                    rs("remark")     = remark
                    rs("keyes")     = keyes
                    rs("pic")     = pic
                    rs("tim")     = now_time
                    rs("counter")     = 0
                    rs("types")     = types
                    rs.update
                    rs.Close:Set rs = Nothing
                    Call upload_note(data_name,first_id(data_name))
                    Response.Write "<font class=red>音乐添加成功！</font>&nbsp;<a href='?c_id=" & cid & "&s_id=" & sid & "'>点击返回</a><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "&action=down_add'>点击继续添加</a>" & _
                    VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=?c_id=" & cid & "&s_id=" & sid & "&action=down_add'>"
                End If

            Else %>
<table border=0 width='98%' cellspacing=0 cellpadding=2>
  <tr><td colspan=2 height=50 align=center><font class=red>添加新软件下载</font></td></tr>
<form name='add_frm' action='?action=add&types=add' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>音乐名称：</td><td width='85%'><input type=text name=name size=70 maxlength=40><% Response.Write redx %></td></tr>
  <tr><td align=center>所属类别：</td><td><% Call chk_csid(cid,sid):Call chk_emoney(0) %></td></tr>
  <tr><td align=center>下载权限：</td><td><% Call chk_power("",1) %></td></tr>
  <tr><td align=center>文件大小：</td><td><input type=text name=sizes value='KB' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;推荐等级：<select name=types size=1>
<option value='0'>没有等级</option>
<option value='1'>一星级</option>
<option value='2'>二星级</option>
<option value='3'>三星级</option>
<option value='4'>四星级</option>
<option value='5'>五星级</option>
</select>&nbsp;&nbsp;&nbsp;音乐类型：<select name=genre size=1><%
                ddim = Split(web_var(web_down,4),":")

                For i = 0 To UBound(ddim)
                    Response.Write vbcrlf & "<option>" & ddim(i) & "</option>"
                Next

                Erase ddim %></select></td></tr>
  <tr><td align=center>播放软件：</td><td><%
                ddim = Split(web_var(web_down,3),":")

                For i = 0 To UBound(ddim)
                    Response.Write "<input type=checkbox name=os value='" & ddim(i) & "' class=bg_1>" & ddim(i)
                Next

                Erase ddim %></td></tr>
  <tr><td align=center>本站下载：</td><td><input type=text name=url size=70 maxlength=200><% Response.Write redx %></td></tr>
  <tr><td align=center>镜像下载：</td><td><input type=text name=url2 value='http://' size=70 maxlength=200></td></tr>
  <tr><td align=center>文件来自：</td><td><input type=text name=homepage value='http://' size=50 maxlength=50></td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","remark","&nbsp;&nbsp;") %></td></tr>
  <tr><td valign=top align=center><br>音乐备注</td><td><textarea rows=6 name=remark cols=70></textarea></td></tr>
<% ispic = "d" & upload_time(now_time) %>
  <tr><td align=center>关 键 字：</td><td><input type=text name=keyes size=12 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;图片：<input type=text name=pic size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=<% Response.Write ispic %>&uptext=pic' target=upload_frame>上传图片</a>&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=d&uptext=remark' target=upload_frame>上传至内容</a></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=down&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr height=30><td></td><td><input type=submit value=' 提 交 添 加 '></td></tr>
</form></table><%
            End If %></td></tr></table><%
        End Sub

        Sub code_del()
            id = Trim(Request.querystring("id"))

            If Not(IsNumeric(id)) Then Call code_main():Exit Sub
                conn.execute("delete from down_code where id=" & id)
                Call code_main()
            End Sub

            Sub code_edit()
                Dim titler
                Dim rs
                Dim strsql

                If id = "" Or IsNull(id) Then Call code_main():Exit Sub %><table border="0" width="600" cellspacing="0" cellpadding="0">
<tr><td align=center height=300><%
                    strsql = "select * from down_code where id=" & id
                    Set rs = Server.CreateObject("adodb.recordset")
                    rs.open strsql,conn,1,3

                    If Request("types") = "edit" Then
                        Dim name
                        Dim username
                        Dim code
                        Dim remark
                        name     = code_form(Trim(Request("name")))
                        username = code_form(Trim(Request("username")))
                        code     = code_form(Trim(Request("code")))
                        remark   = Request("remark")

                        If name = "" Or IsNull(name) Or code = "" Or IsNull(code) Then
                            Response.Write("音乐名称和注 册 码不能为空！<br><br>" & go_back)
                        Else
                            rs("name") = name
                            rs("username") = username
                            rs("code") = code
                            rs("remark") = remark
                            rs.update
                            rs.Close:Set rs = Nothing
                            Response.Write("注册码修改成功！<br><br><a href='admin_down.asp?action=code'>点击返回</a>")
                            Response.Write(VbCrLf & "<meta http-equiv='refresh' content='" & time_go & "; url=admin_down.asp?action=code'>")
                        End If

                    Else %>
<table border="0" width="400" cellspacing="0" cellpadding="2">
  <tr><td colspan=2 height=50 align=center><font class=font_color1>注册码属性修改</font></td></tr>
  <form action="?action=code_edit&types=edit&id=<% = id %>" method=post>
  <tr><td>文件名称</td><td><input type=text name=name value='<% = rs("name") %>' size=50 maxlength=100></td></tr>
  <tr><td>注册名称</td><td><input type=text name=username value='<% = rs("username") %>' size=50 maxlength=100></td></tr>
  <tr><td>注 册 码</td><td><input type=text name=code value='<% = rs("code") %>' size=50 maxlength=100></td></tr>
  <tr><td>备注说明</td><td><% Response.Write("<textarea rows=6 name=remark cols=50>" & rs("remark") & "</textarea>") %></td></tr>
  <tr height=30><td></td><td><input type="submit" value=" 修 改 "></td></tr>
</form></table>
<% End If %></td></tr><tr></table><%

                End Sub

                Sub code_add() %><table border="0" width="600" cellspacing="0" cellpadding="0">
<tr><td align=center height=300><%

                    If Request("types") = "add" Then
                        Dim name
                        Dim username
                        Dim code
                        Dim remark
                        name     = code_form(Trim(Request("name")))
                        username = code_form(Trim(Request("username")))
                        code     = code_form(Trim(Request("code")))
                        remark   = Request("remark")

                        If name = "" Or IsNull(name) Or code = "" Or IsNull(code) Then
                            Response.Write("文件名称和注 册 码不能为空！<br><br>" & go_back)
                        Else
                            Dim rs
                            Dim strsql
                            strsql = "select * from down_code where (id is null)"
                            Set rs = Server.CreateObject("adodb.recordset")
                            rs.open strsql,conn,1,3
                            rs.addnew
                            rs("name") = name
                            rs("username") = username
                            rs("code") = code
                            rs("remark") = remark
                            rs.update
                            rs.Close
                            Set rs = Nothing
                            Response.Write("注册码添加成功！<br><br><a href='admin_down.asp?action=code_add'>点击继续添加</a>")
                            Response.Write(VbCrLf & "<meta http-equiv='refresh' content='" & time_go & "; url=admin_down.asp?action=code_add'>")
                        End If

                    Else %><table border="0" width="400" cellspacing="0" cellpadding="2">
<form action="?action=code_add&types=add" method=post>
  <tr><td colspan=2 height=50 align=center><font class=font_color1>添加新注册码</font></td></tr>
  <tr><td>文件名称</td><td><input type=text name=name size=50 maxlength=100></td></tr>
  <tr><td>注册名称</td><td><input type=text name=username size=50 maxlength=100></td></tr>
  <tr><td>注 册 码</td><td><input type=text name=code size=50 maxlength=100></td></tr>
  <tr><td>备注说明</td><td><textarea rows="6" name=remark cols="50"></textarea></td></tr>
  <tr height=30><td></td><td><input type="submit" value=" 添 加 "></td></tr>
</form></table>
<% End If %></td></tr></table><%

                End Sub

                Sub code_main() %><table border=0 width='95%' cellspacing=0 cellpadding=2><%
                    Dim rs
                    Dim strsql
                    sqladd = ""
                    strsql = "select * from down_code " & sqladd & "order by id desc"
                    Set rs = Server.CreateObject("adodb.recordset")
                    rs.open strsql,conn,1,1

                    If rs.eof And rs.bof Then
                        rssum = 0
                    Else
                        rssum = rs.recordcount
                    End If

                    Call format_pagecute()
                    Response.Write "<tr><td colspan=3 align=center height=30>共有个 <font class=red>" & rssum & "</font> 注册码  分页:" & jk_pagecute(nummer,thepages,viewpage,pageurl,10,"#ff0000") & "</td></tr>" & _
                    "<tr align=center><td width='10%'>序号</td><td width='75%'>类型和名称</td><td width='15%'>操作</td></tr>"

                    If Int(viewpage) > 1 Then
                        rs.move (viewpage - 1)*nummer
                    End If

                    For i = 1 To nummer
                        If rs.eof Then Exit For
                        Response.Write "<tr align=center><td>" & (viewpage - 1)*nummer + i & "</td><td align=left>" & rs("name") & "</td><td><a href='admin_down.asp?action=code_edit&id=" & rs("id") & "'>修改</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href='admin_down.asp?action=code_del&id=" & rs("id") & "'>删除</a></td></tr>"
                        rs.movenext
                    Next

                    rs.Close:Set rs = Nothing %></table><%
                End Sub

                Sub down_main() %><table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr align=center height=300 valign=top>
<td width='20%' class=htd><br><% Call left_sort() %></td>
<td width='80%'>
<table border=1 width='100%' cellspacing=0 cellpadding=1 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<script language=javascript src='STYLE/admin_del.js'></script>
<form name=del_form action='<% Response.Write pageurl %>del_ok=ok' method=post><%
                    Call sql_cid_sid()
                    sql    = "select id,c_id,s_id,name,types,hidden from down" & sqladd & " order by tim desc"
                    Set rs = Server.CreateObject("adodb.recordset")
                    rs.open sql,conn,1,1

                    If rs.eof And rs.bof Then
                        rssum = 0
                    Else
                        rssum = Int(rs.recordcount)
                    End If

                    Call format_pagecute()
                    del_temp = nummer
                    If rssum = 0 Then del_temp = 0

                    If Int(page) = Int(thepages) Then
                        del_temp = rssum - nummer*(thepages - 1)
                    End If %>
<tr><td colspan=3 align=center height=25>
现有<font class=red><% Response.Write rssum %></font>个软件　<% Response.Write "<a href='?action=add&c_id=" & cid & "&s_id=" & sid & "'>添加音乐</a>" %>
　<input type=checkbox name=del_all value=1 onClick=selectall('<% Response.Write del_temp %>')> 选中所有　<input type=submit value='删除所选' onclick=""return suredel('<% Response.Write del_temp %>');"">
</td></tr>
<tr align=center bgcolor=#ffffff><td width='8%'>序号</td><td width='77%'>类型和名称</td><td width='15%'>操作</td></tr>
<%

                    If Int(viewpage) > 1 Then
                        rs.move (viewpage - 1)*nummer
                    End If

                    For i = 1 To nummer
                        If rs.eof Then Exit For
                        now_id = rs("id"):nid = Int(rs("types")):ncid = rs("c_id"):nsid = rs("s_id")
                        Response.Write vbcrlf & "<tr align=center><td>" & (viewpage - 1)*nummer + i & "</td><td align=left><a href='?action=down_edit&id=" & now_id & "'>" & rs("name") & "</a></td><td align=right><a href='?action=hidden&c_id=" & cid & "&s_id=" & sid & "&id=" & now_id & "&page=" & viewpage & "'>"

                        If rs("hidden") = True Then
                            Response.Write "显</a> "
                        Else
                            Response.Write "<font class=red_2>隐</font></a> "
                        End If

                        Response.Write "<font class=red>" & nid & "</font>&nbsp;星&nbsp;<input type=checkbox name=del_id value='" & now_id & "'></td></tr>"
                        rs.movenext
                    Next

                    rs.Close:Set rs = Nothing %></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %>
</td></tr>
</table></td></tr></table><%
                End Sub %>