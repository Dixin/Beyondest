<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nsort,sql2,rs2,del_temp,data_name,cid,sid,ncid,nsid,nid,id,left_type,now_id,nummer,sqladd,page,rssum,thepages,viewpage,pageurl,pic,ispic,types,csid
types   = Trim(Request.querystring("types"))
types   = "film"
pageurl = "?action=" & action & "&types=" & types & "&":nsort = types:data_name = "gallery":sqladd = "":nummer = 30
tit     = vbcrlf & "<a href='?'>��Ƶ����</a>&nbsp;��&nbsp;" & _
vbcrlf & "<a href='?action=add&types=" & nsort & "'>�����Ƶ</a>&nbsp;��&nbsp;" & _
vbcrlf & "<a href='admin_nsort.asp?nsort=film'>��Ƶ����</a>"
Response.Write header(16,tit)
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

Function select_type(st1,st2)
    select_type = vbcrlf & "<option"
    If st1 = st2 Then select_type = select_type & " selected"
    select_type = select_type & ">" & st1 & "</option>"
End Function

Sub news_edit()
    Dim rs3,sql3,name

    If Trim(Request.querystring("edit")) = "chk" Then
        name  = code_admin(Request.form("name"))
        csid  = Trim(Request.form("csid"))
        pic   = code_admin(Request.form("pic"))
        types = Trim(Request.form("types"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>��ѡ���ļ����࣡</font><br><br>" & go_back
        ElseIf Len(name) < 1 Then
            Response.Write "<font class=red_2>�ļ�����˵������Ϊ�գ�</font><br><br>" & go_back
        ElseIf Len(pic) < 3 Then
            Response.Write "<font class=red_2>���ϴ��ļ��������ļ��ĵ�ַ��</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            rs("c_id")     = cid
            rs("s_id")     = sid
            If Trim(Request.form("username_my")) = "yes" Then rs("username") = login_username
            rs("types")     = types
            rs("name")     = name

            If Len(code_admin(Request.form("spic"))) < 3 Then
                rs("spic") = "no_pic.gif"
            Else
                rs("spic") = code_admin(Request.form("spic"))
            End If

            rs("pic")     = pic
            rs("remark")     = Left(Request.form("remark"),250)
            rs("power")     = Replace(Replace(Trim(Request.form("power"))," ",""),",",".")

            If IsNumeric(Trim(Request.form("emoney"))) Then
                rs("emoney") = Trim(Request.form("emoney"))
            Else
                rs("emoney") = 0
            End If

            If Trim(Request.form("istop")) = "yes" Then
                rs("istop") = 1
            Else
                rs("istop") = 0
            End If

            If IsNumeric(Trim(Request.form("counter"))) Then rs("counter") = Trim(Request.form("counter"))

            If Trim(Request.form("hidden")) = "yes" Then
                rs("hidden") = False
            Else
                rs("hidden") = True
            End If

            rs.update
            rs.Close:Set rs = Nothing
            Call upload_note(data_name,id)
            Response.Write "<font class=red>�ѳɹ��޸���һ���ļ���</font><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "&types=" & types & "'>�������</a><br><br>"
        End If

    Else
        types = rs("types") %><table border=0 cellspacing=0 cellpadding=3>
<form action='<% Response.Write pageurl %>c_id=<% Response.Write cid %>&s_id=<% Response.Write sid %>&id=<% Response.Write id %>&edit=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='12%'>�ļ����ƣ�</td><td width='88%'><input type=text size=40 name=name value='<% Response.Write rs("name") %>' maxlength=50><% = redx %></td></tr>
  <tr><td>�ļ����ࣺ</td><td><% Call chk_csid(cid,sid) %>&nbsp;&nbsp;�ļ����ͣ�<select name=types size=1>
<option value='film'<% If types = "film" Then Response.Write " selected" %>>��Ƶ</option>
<option value='logo'<% If types = "logo" Then Response.Write " selected" %>>����</option>
</select><% = redx %>&nbsp;&nbsp;<% Call chk_emoney(rs("emoney")) %></td></tr>
  <tr><td align=center>���Ȩ�ޣ�</td><td><% Call chk_power(rs("power"),0) %></td></tr>
  <tr><td align=center>���������</td><td><input type=text name=counter value='<% Response.Write rs("counter") %>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name=istop value='yes'<% If Int(rs("istop")) = 1 Then Response.Write " checked" %>>&nbsp;�Ƽ�&nbsp;&nbsp;<% Call chk_h_u() %></td></tr>
<%
        pic   = rs("spic")
        If pic = "no_pic.gif" Then pic = ""
        ispic = pic
        If InStr(ispic,"/") > 0 Then ispic = Right(ispic,Len(ispic) - InStr(ispic,"/"))
        If InStr(ispic,".") > 0 Then ispic = Left(ispic,InStr(ispic,".") - 1)
        If Len(ispic) < 1 Then ispic = "n" & upload_time(now_time) %>
  <tr><td>С ͼ Ƭ��</td><td><input type=test name=spic value='<% Response.Write pic %>' size=70 maxlength=100></td></tr>
  <tr><td>�ϴ�ͼƬ��</td><td><iframe frameborder=0 name=upload_frames width='100%' height=30 scrolling=no src='upload.asp?uppath=gallery&upname=<% Response.Write ispic %>&uptext=spic'></iframe></td></tr>
<%
        pic   = rs("pic")
        If pic = "no_pic.gif" Then pic = ""
        ispic = pic
        If InStr(ispic,"/") > 0 Then ispic = Right(ispic,Len(ispic) - InStr(ispic,"/"))
        If InStr(ispic,".") > 0 Then ispic = Left(ispic,InStr(ispic,".") - 1)
        If Len(ispic) < 1 Then ispic = "n" & upload_time(now_time) %>
  <tr><td>�ļ���ַ��</td><td><input type=test name=pic value='<% Response.Write pic %>' size=70 maxlength=100><% Response.Write redx %></td></tr>
  <tr><td>�ϴ��ļ���</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=gallery&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr><td>�ļ�˵����<br><br><=250�ַ�</td><td><textarea name=remark maxlength=250 rows=5 cols=70><% Response.Write rs("remark") %></textarea></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' �� �� �� �� '></td></tr>
</form></table><%
    End If

End Sub

Sub news_add()
    Dim name,csid
    types = Trim(Request.querystring("types"))
    If types <> "flash" And types <> "logo" And types <> "film" Then types = "paste"

    If Trim(Request.querystring("add")) = "chk" Then
        name  = code_admin(Request.form("name"))
        csid  = Trim(Request.form("csid"))
        pic   = code_admin(Request.form("pic"))
        types = Trim(Request.form("types"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>��ѡ���ļ����࣡</font><br><br>" & go_back
        ElseIf Len(name) < 1 Then
            Response.Write "<font class=red_2>�ļ�����˵������Ϊ�գ�</font><br><br>" & go_back
        ElseIf Len(pic) < 3 Then
            Response.Write "<font class=red_2>���ϴ��ļ��������ļ��ĵ�ַ��</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            Set rs = Server.CreateObject("adodb.recordset")
            sql    = "select * from " & data_name
            rs.open sql,conn,1,3
            rs.addnew
            rs("c_id")     = cid
            rs("s_id")     = sid
            rs("username")     = login_username
            rs("types")     = types
            rs("name")     = name

            If Len(code_admin(Request.form("spic"))) < 3 Then
                rs("spic") = "no_pic.gif"
            Else
                rs("spic") = code_admin(Request.form("spic"))
            End If

            rs("pic")     = pic
            rs("remark")     = Left(Request.form("remark"),250)
            rs("power")     = Replace(Replace(Trim(Request.form("power"))," ",""),",",".")

            If IsNumeric(Trim(Request.form("emoney"))) Then
                rs("emoney") = Trim(Request.form("emoney"))
            Else
                rs("emoney") = 0
            End If

            If Trim(Request.form("istop")) = "yes" Then
                rs("istop") = 1
            Else
                rs("istop") = 0
            End If

            rs("counter")     = 0
            rs("tim")     = now_time
            rs("hidden")     = True
            rs.update
            rs.Close:Set rs = Nothing
            Call upload_note(data_name,first_id(data_name))
            Response.Write "<font class=red>�ѳɹ������һ���ļ���</font><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "&types=" & types & "'>�������</a><br><br>"
        End If

    Else %><table border=0 cellspacing=0 cellpadding=3>
<form action='<% Response.Write pageurl %>add=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='12%' align=center>�ļ����ƣ�</td><td width='88%'><input type=text size=70 name=name maxlength=50><% = redx %></td></tr>
  <tr><td align=center>�ļ����ࣺ</td><td><% Call chk_csid(cid,sid) %>&nbsp;&nbsp;�ļ����ͣ�<select name=types size=1>
<option value='film'<% If types = "film" Then Response.Write " selected" %>>��Ƶ</option>
<option value='logo'<% If types = "logo" Then Response.Write " selected" %>>����</option>
</select><% Response.Write redx %>&nbsp;&nbsp;<% Call chk_emoney(0) %></td></tr>
  <tr><td align=center>���Ȩ�ޣ�</td><td><% Call chk_power("",1) %></td></tr>
<% ispic = "gs" & upload_time(now_time) %>
  <tr><td align=center>С ͼ Ƭ��</td><td><input type=test name=spic size=70 maxlength=100></td></tr>
  <tr><td align=center>�ϴ�ͼƬ��</td><td><iframe frameborder=0 name=upload_frames width='100%' height=28 scrolling=no src='upload.asp?uppath=gallery&upname=<% Response.Write ispic %>&uptext=spic'></iframe></td></tr>
<% ispic = "g" & upload_time(now_time) %>
  <tr><td align=center>�ļ���ַ��</td><td><input type=test name=pic size=70 maxlength=100><% Response.Write redx %></td></tr>
  <tr><td align=center>�ϴ��ļ���</td><td><iframe frameborder=0 name=upload_frame width='100%' height=28 scrolling=no src='upload.asp?uppath=gallery&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr><td align=center>�ļ�˵����<br><br><=250�ַ�</td><td><textarea name=remark rows=5 cols=70></textarea></td></tr>
  <tr><td colspan=2 align=center height=30><input type=submit value=' �� �� �� �� '></td></tr>
</form></table><%
    End If

End Sub

Sub news_main() %>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='100%' cellpadding=2>
  <tr valign=top height=350>
    <td width='25%' class=htd><br>�ļ����ͣ�<br>

<a href='?types=film'<% If types = "film" Then Response.Write " class=red_3" %>>��Ƶ</a><br>
<a href='?types=logo'<% If types = "logo" Then Response.Write " class=red_3" %>>����</a><br>
<br>�ļ����ࣺ<br><% Call left_sort2() %></td>
    <td width='75%' align=center>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<form name=del_form action='<% = pageurl %>del_ok=ok' method=post>
<tr><td width='6%'></td><td width='80%'></td><td width='14%'></td></tr>
<%
    Call sql_cid_sid()

    If Len(sqladd) < 1 Then
        sqladd = " where types='" & types & "'"
    Else
        sqladd = sqladd & " and types='" & types & "'"
    End If

    sql        = "select id,c_id,s_id,name,pic,hidden,istop from " & data_name & sqladd & " order by id desc"
    Set rs     = Server.CreateObject("adodb.recordset")
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
����<font class=red><% Response.Write rssum %></font>���ļ���<% Response.Write "<a href='?action=add&c_id=" & cid & "&s_id=" & sid & "'>����ļ�</a>" %>
��<input type=checkbox name=del_all value=1 onClick=selectall('<% Response.Write del_temp %>')> ѡ�����С�<input type=submit value='ɾ����ѡ' onclick=""return suredel('<% Response.Write del_temp %>');"">
</td></tr>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<%

    If Int(viewpage) > 1 Then
        rs.move (viewpage - 1)*nummer
    End If

    For i = 1 To nummer
        If rs.eof Then Exit For
        now_id = rs("id"):ncid = rs("c_id"):nsid = rs("s_id")
        Response.Write gallery_center()
        rs.movenext
    Next

    rs.Close:Set rs = Nothing %></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>ҳ�Σ�<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>
��ҳ��<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %>
</td></tr></table>
</td></tr></table>
<%
End Sub

Function gallery_center()
    gallery_center = vbcrlf & "<tr" & mtr & ">" & _
    vbcrlf & "<td><a href='" & url_true(web_var(web_upload,1),rs("pic")) & "' target=_blank title='������ļ�'>" & i + (viewpage - 1)*nummer & ".</a> </td><td>" & _
    vbcrlf & "<a href='?action=edit&c_id=" & rs(1) & "&s_id=" & rs(2) & "&id=" & now_id & "'>" & rs("name") & "</a></td><td align=center><a href='?action=hidden&c_id=" & cid & "&s_id=" & sid & "&id=" & now_id & "&types=" & types & "&page=" & viewpage & "'>"

    If rs("hidden") = True Then
        gallery_center = gallery_center & "��"
    Else
        gallery_center = gallery_center & "<font class=red_2>��</font>"
    End If

    gallery_center = gallery_center & "</a> <a href='?action=istop&c_id=" & cid & "&s_id=" & sid & "&id=" & now_id & "&types=" & types & "&page=" & viewpage & "'>"

    If Int(rs("istop")) = 1 Then
        gallery_center = gallery_center & "<font class=red>��</font>"
    Else
        gallery_center = gallery_center & "��"
    End If

    gallery_center = gallery_center & "</a> <input type=checkbox name=del_id value='" & now_id & "'></td></tr>"
End Function %>