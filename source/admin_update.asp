<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #INCLUDE file="include/jk_page_cute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim id
Dim nsort
Dim rssum
Dim nummer
Dim thepages
Dim viewpage
Dim pageurl
Dim page
nsort = Trim(Request("nsort"))

Select Case nsort
    Case "forum"
        nsort = nsort
    Case Else
        nsort = "news"
End Select

sql = "select * from bbs_cast"			' where sort='"&nsort&"'"

tit = "<a href='admin_update.asp?'>��վ����</a> �� " & _
"<a href='admin_data.asp'>���ݸ���</a> �� " & _
"<a href='admin_update.asp?nsort=news'>���¹���</a> �� " & _
"<a href='admin_update.asp?nsort=forum'>��̳����</a> �� " & _
"<a href='admin_update.asp?action=add'>��Ӹ���</a>"

Response.Write header(7,tit)
id = Trim(Request.querystring("id"))

Select Case action
    Case "add"
        Response.Write news_add()
    Case "addchk"
        Response.Write news_addchk()
    Case "del"

        If IsNumeric(id) Then
            Response.Write news_del(id)
        Else
            Response.Write news_main()
        End If

    Case "edit"

        If IsNumeric(id) Then
            Response.Write news_edit(id)
        Else
            Response.Write news_main()
        End If

    Case "editchk"

        If IsNumeric(id) Then
            Response.Write news_editchk(id)
        Else
            Response.Write news_main()
        End If

    Case Else
        Response.Write news_main()
End Select

Response.Write ender()

Function news_del(id)
    On Error Resume Next
    conn.execute("delete from bbs_cast where sort='" & nsort & "' and id=" & id)
    Call upload_del("update",id)

    If Err Then
        Err.Clear
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""���Ĳ����д���error in del�����ڣ�\n\n������ء�"");" & _
        vbcrlf & "location='?nsort=" & nsort & "'" & _
        vbcrlf & "</script>")
    Else
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""�ɹ�ɾ����һ�����£�\n\n������ء�"");" & _
        vbcrlf & "location='?nsort=" & nsort & "'" & _
        vbcrlf & "</script>")
    End If

End Function

Function news_main()
    pageurl = "?nsort=" & nsort & "&action=main&"
    Set rs  = Server.CreateObject("adodb.recordset")
    sql     = sql & " where sort='" & nsort & "' order by id desc"
    rs.open sql,conn,1,1

    If Not(rs.eof And rs.bof) Then
        rssum  = rs.recordcount
        nummer = 15
        Call format_pagecute

        news_main = news_main & vbcrlf & "<script language=JavaScript><!--" & _
        vbcrlf & "function Do_del_data(data1)" & _
        vbcrlf & "{" & _
        vbcrlf & "if (confirm(""�˲�����ɾ��idΪ ""+data1+"" ��չ����Ϣ��\n���Ҫɾ����\nɾ�����޷��ָ���""))" & _
        vbcrlf & "  window.location=""?nsort=" & nsort & "&action=del&id=""+data1" & _
        vbcrlf & "}" & _
        vbcrlf & "//--></script>" & _
        vbcrlf & "<table border=1 width=500 cellspacing=0 cellpadding=1 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>" & _
        vbcrlf & "<tr><td colspan=3 align=center height=30>������ <font class=red>" & rssum & "</font> ������</td></tr>" & _
        "<tr align=center><td width='8%'>���</td><td width='75%'>����</td><td width='17%'>����</td></tr>"

        If Int(viewpage) > 1 Then
            rs.move (viewpage - 1)*nummer
        End If

        For i = 1 To nummer
            If rs.eof Then Exit For
            news_main = news_main & vbcrlf & "<tr align=center><td>" & i + (viewpage - 1)*nummer & ".</td><td align=left>" & code_html(rs("topic"),1,28) & "</td><td><a href='?nsort=" & nsort & "&action=edit&id=" & rs("id") & "'>�޸�</a> �� <a href='javascript:Do_del_data(" & rs("id") & ")'>ɾ��</a></td></tr>"
            rs.movenext
        Next

        news_main = news_main & vbcrlf & "</table>" & kong & pagecute_fun(viewpage,thepages,pageurl)
    End If

    rs.Close:Set rs = Nothing
End Function

Function news_add() %><table border=0 width='98%' cellspacing=0 cellpadding=2>
<form name='add_frm' action='?action=addchk' method=post>
<input type=hidden name=upid value=''>
  <tr><td colspan=2 align=center height=50><font class=red>��ӹ������</font></td></tr>
  <tr><td width='15%' align=center>���±��⣺</td><td width='85%'><input type=text name=topic size=65 maxlength=50></td></tr>
  <tr><td align=center height=30>�������ͣ�</td><td><input type=radio name=nsort value='news' checked>&nbsp;��վ����&nbsp;&nbsp;<input type=radio name=nsort value='forum'>&nbsp;��̳����</td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","word","&nbsp;&nbsp;") %></td></tr>
  <tr><td align=center valign=top><br>�����ڿգ�</td><td><textarea name=word rows=15 cols=65></textarea></td></tr>
  <tr><td align=center>�ϴ��ļ���</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=f&uptext=word'></iframe></td></tr>
  <tr height=30 align=center><td colspan=2><input type=submit value='�� �� �� ��'>������<input type=reset value='������д'></td></tr>
</form></table><%
End Function

Function news_addchk()
    Dim topic
    topic = Trim(Request.form("topic"))

    If Len(topic) < 1 Then
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""���±��� �Ǳ���Ҫ�ģ�\n\n�뷵�����롣"");" & _
        vbcrlf & "history.back(1)" & _
        vbcrlf & "</script>")
    Else
        Set rs = Server.CreateObject("adodb.recordset")
        rs.open sql,conn,1,3
        rs.addnew
        rs("sort") = nsort
        rs("username") = login_username
        rs("topic") = topic
        rs("word") = Request.form("word")
        rs("tim") = Now
        rs.update
        rs.Close:Set rs = Nothing
        Call upload_note("update",first_id("bbs_cast"))
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""�ɹ������˸��£�\n\n������ء�"");" & _
        vbcrlf & "location='?nsort=" & nsort & "'" & _
        vbcrlf & "</script>")
    End If

End Function

Function news_edit(id)
    sql    = sql & " where id=" & id
    Set rs = conn.execute(sql)

    If rs.eof And rs.bof Then
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""���Ĳ����д���error in edit�����ڣ�\n\n������ء�"");" & _
        vbcrlf & "location='?nsort=" & nsort & "'" & _
        vbcrlf & "</script>")
    Else
        Dim msort:msort = rs("sort") %><table border=0 width='98%' cellspacing=0 cellpadding=2>
<form name='add_frm' action='?action=editchk&id=<% Response.Write id %>' method=post>
<input type=hidden name=upid value=''>
  <tr><td colspan=2 align=center height=50><font class=red>�޸ĸ���</font></td></tr>
  <tr><td width='15%' align=center>���±��⣺</td><td width='85%'><input type=text name=topic value='<% Response.Write rs("topic") %>' size=65 maxlength=50></td></tr>
  <tr><td height=30 align=center>�������ͣ�</td><td><input type=radio name=nsort value='news'<% If msort = "news" Then Response.Write "checked" %>>&nbsp;��վ����&nbsp;&nbsp;<input type=radio name=nsort value='forum'<% If msort = "forum" Then Response.Write "checked" %>>&nbsp;��̳����</td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","word","&nbsp;&nbsp;") %></td></tr>
  <tr><td align=center>�����ڿգ�</td><td><textarea name=word rows=15 cols=65><% Response.Write rs("word") %></textarea></td></tr>
  <tr><td align=center>�ϴ��ļ���</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=forum&upname=f&uptext=word'></iframe></td></tr>
  <tr height=30 align=center><td colspan=2><input type=submit value='�� �� �� ��'>������<input type=reset value='������д'></td></tr>
</form></table><%
    End If

    rs.Close:Set rs = Nothing
End Function

Function news_editchk(id)
    Dim topic:topic = Trim(Request.form("topic"))
    Call upload_note("update",id)

    If Len(topic) < 1 Then
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""���±��� �Ǳ���Ҫ�ģ�\n\n�뷵�����롣"");" & _
        vbcrlf & "history.back(1)" & _
        vbcrlf & "</script>")
    Else
        Set rs = Server.CreateObject("adodb.recordset")
        sql    = sql & " where id=" & id
        rs.open sql,conn,1,3

        If rs.eof And rs.bof Then
            Response.Write("<script language=javascript>" & _
            vbcrlf & "alert(""���Ĳ����д���error in editchk�����ڣ�\n\n������ء�"");" & _
            vbcrlf & "location='?nsort=" & nsort & "'" & _
            vbcrlf & "</script>")
        Else
            rs("sort") = nsort
            rs("topic") = topic
            rs("word") = Request.form("word")
            rs.update
            rs.Close:Set rs = Nothing
            Response.Write("<script language=javascript>" & _
            vbcrlf & "alert(""�ɹ��޸��˸��£�\n\n������ء�"");" & _
            vbcrlf & "location='?nsort=" & nsort & "'" & _
            vbcrlf & "</script>")
        End If

    End If

End Function %>