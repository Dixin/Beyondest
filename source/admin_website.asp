<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<!-- #INCLUDE file="include/functions.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim website_menu
Dim nsort
Dim sql2
Dim rs2
Dim del_temp
Dim data_name
Dim cid
Dim sid
Dim ncid
Dim nsid
Dim nid
Dim id
Dim left_type
Dim now_id
Dim nummer
Dim sqladd
Dim page
Dim rssum
Dim thepages
Dim viewpage
Dim pageurl
Dim pic
Dim ispic
Dim csid
website_menu = vbcrlf & "<a href='?'>��վ�Ƽ�</a>&nbsp;��&nbsp;" & _
vbcrlf & "<a href='?action=add'>�����վ</a>&nbsp;��&nbsp;" & _
vbcrlf & "<a href='admin_nsort.asp?nsort=web'>��վ����</a>"
Response.Write header(15,website_menu)
pageurl = "?action=" & action & "&":nsort = "web":data_name = "website":sqladd = "":nummer = 15
Call admin_cid_sid()

If Trim(Request("del_ok")) = "ok" Then
    Response.Write del_select(Trim(Request.form("del_id")))
End If

Function del_select(delid)
    Dim del_i
    Dim del_num
    Dim del_dim
    Dim del_sql
    Dim fobj
    Dim picc

    If delid <> "" And Not IsNull(delid) Then
        delid   = Replace(delid," ","")
        del_dim = Split(delid,",")
        del_num = UBound(del_dim)

        For del_i = 0 To del_num
            Call upload_del(data_name,del_dim(del_i))
            del_sql = "delete from " & data_name & " where id=" & del_dim(del_i)
            conn.execute(del_sql)
        Next

        Erase del_dim
        del_select = vbcrlf & "<script language=javascript>alert(""��ɾ���� " & del_num + 1 & " ����¼��"");</script>"
    End If

End Function

id         = Trim(Request.querystring("id"))

If (action = "hidden" Or action = "isgood") And IsNumeric(id) Then
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

close_conn()
Response.Write ender()

Function select_type(st1,st2)
    select_type = vbcrlf & "<option"
    If st1 = st2 Then select_type = select_type & " selected"
    select_type = select_type & ">" & st1 & "</option>"
End Function

Sub news_edit()
    Dim rs3
    Dim sql3
    Dim name
    Dim url
    Dim isgood
    Dim country
    Dim lang
    Dim remark

    If Trim(Request.querystring("edit")) = "chk" Then
        name    = code_admin(Request.form("name"))
        csid    = Trim(Request.form("csid"))
        url     = code_admin(Request.form("url"))
        isgood  = Trim(Request.form("isgood"))
        remark  = Request.form("remark")
        country = Trim(Request.form("country"))
        lang    = Trim(Request.form("lang"))
        pic     = Trim(Request.form("pic"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>��ѡ����վ���ͣ�</font><br><br>" & go_back
        ElseIf Len(name) < 1 Or Len(url) < 1 Then
            Response.Write "<font class=red_2>��վ���ƺ͵�ַ����Ϊ�գ�</font><br><br>" & go_back
        ElseIf Len(remark) > 250 Then
            Response.Write "<font class=red_2>��վ˵�����ܳ���250���ַ���</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            rs("c_id")     = cid
            rs("s_id")     = sid
            If Trim(Request.form("username_my")) = "yes" Then rs("username") = login_username
            rs("name")     = name
            rs("url")     = url
            rs("country")     = country
            rs("lang")     = lang
            rs("remark")     = remark

            If isgood = "yes" Then
                rs("isgood") = True
            Else
                rs("isgood") = False
            End If

            If Trim(Request.form("hidden")) = "yes" Then
                rs("hidden") = False
            Else
                rs("hidden") = True
            End If

            If Len(pic) < 3 Then
                rs("pic") = "no_pic.gif"
            Else
                rs("pic") = pic
            End If

            rs("tim")     = now_time
            rs.update
            rs.Close:Set rs = Nothing
            Call upload_note(data_name,id)
            Response.Write "<font class=red>�ѳɹ��޸���һ����վ��</font><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "'>�������</a><br><br>"
        End If

    Else %><table border=0 cellspacing=0 cellpadding=3>
<form action='<% Response.Write pageurl %>c_id=<% Response.Write cid %>&s_id=<% Response.Write sid %>&id=<% Response.Write id %>&edit=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='12%'>��վ���ƣ�</td><td width='88%'><input type=text size=70 name=name value='<% Response.Write rs("name") %>' maxlength=50><% = redx %></td></tr>
  <tr><td>��վ���ͣ�</td><td><% Call chk_csid(cid,sid):Call chk_h_u() %></td></tr>
  <tr><td>��վ��ַ��</td><td><input type=text size=70 name=url value='<% Response.Write rs("url") %>' maxlength=100><% = redx %></td></tr>
  <tr><td>���ҵ�����</td><td><select name=country size=1>
<%
        pic       = rs("pic")
        If pic = "no_pic.gif" Then pic = ""
        ispic     = pic

        If InStr(ispic,"/") > 0 Then
            ispic = Right(ispic,Len(ispic) - InStr(ispic,"/"))
        End If

        If InStr(ispic,".") > 0 Then
            ispic = Left(ispic,InStr(ispic,".") - 1)
        End If

        If Len(ispic) < 1 Then ispic = "n" & upload_time(now_time)
        tit = rs("country")
        Response.Write select_type("�й�",tit)
        Response.Write select_type("���",tit)
        Response.Write select_type("̨��",tit)
        Response.Write select_type("����",tit)
        Response.Write select_type("Ӣ��",tit)
        Response.Write select_type("�ձ�",tit)
        Response.Write select_type("����",tit)
        Response.Write select_type("���ô�",tit)
        Response.Write select_type(">�Ĵ�����",tit)
        Response.Write select_type("������",tit)
        Response.Write select_type("����˹",tit)
        Response.Write select_type("�����",tit)
        Response.Write select_type("����",tit)
        Response.Write select_type("������",tit)
        Response.Write select_type("�¹�",tit)
        Response.Write select_type("��������",tit) %>
</select>&nbsp;&nbsp;&nbsp;&nbsp;վ�����ԣ�<select name=lang size=1>
<%
        tit = rs("lang")
        Response.Write select_type("��������",tit)
        Response.Write select_type("��������",tit)
        Response.Write select_type("English",tit)
        Response.Write select_type("��������",tit) %>
</select>&nbsp;&nbsp;&nbsp;�Ƽ���<input type=checkbox name=isgood<% If rs("isgood") = True Then Response.Write " checked" %> value='yes'></td></tr>
  <tr><td>ͼƬ��ַ��</td><td><input type=test name=pic value='<% Response.Write pic %>' size=70 maxlength=100></td></tr>
  <tr><td>�ϴ�ͼƬ��</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=website&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr><td valign=top class=htd><br>��վ���ݣ�<br><=250B</td><td><textarea name=remark rows=5 cols=70><% Response.Write rs("remark") %></textarea></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' �� �� �� վ '></td></tr>
</form></table><%
    End If

End Sub

Sub news_add()

    If Trim(Request.querystring("add")) = "chk" Then
        Dim name
        Dim url
        Dim isgood
        Dim country
        Dim lang
        Dim remark
        name    = code_admin(Request.form("name"))
        csid    = Trim(Request.form("csid"))
        url     = code_admin(Request.form("url"))
        isgood  = Trim(Request.form("isgood"))
        remark  = Request.form("remark")
        country = Trim(Request.form("country"))
        lang    = Trim(Request.form("lang"))
        pic     = Trim(Request.form("picg"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>��ѡ����վ���ͣ�</font><br><br>" & go_back
        ElseIf Len(name) < 1 Or Len(url) < 1 Then
            Response.Write "<font class=red_2>��վ���ƺ͵�ַ����Ϊ�գ�</font><br><br>" & go_back
        ElseIf Len(remark) > 250 Then
            Response.Write "<font class=red_2>��վ˵�����ܳ���250���ַ���</font><br><br>" & go_back
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
            rs("name")     = name
            rs("url")     = url
            rs("country")     = country
            rs("lang")     = lang
            rs("remark")     = remark

            If isgood = "yes" Then
                rs("isgood") = True
            Else
                rs("isgood") = False
            End If

            rs("username")     = login_username

            If Len(pic) < 3 Then
                rs("pic") = "no_pic.gif"
            Else
                rs("pic") = pic
            End If

            rs("tim")     = now_time
            rs("counter")     = 0
            rs.update
            rs.Close:Set rs = Nothing
            Call upload_note(data_name,first_id(data_name))
            Response.Write "<font class=red>�ѳɹ������һ����վ��</font><br><br><a href='?c_id=" & cid & "&s_id=" & sid & "'>�������</a><br><br>"
        End If

    Else %><table border=0 cellspacing=0 cellpadding=3>
<form action='<% Response.Write pageurl %>add=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='12%'>��վ���ƣ�</td><td width='88%'><input type=text size=70 name=name maxlength=50><% = redx %></td></tr>
  <tr><td>��վ���ͣ�</td><td><% Call chk_csid(cid,sid) %></td></tr>
  <tr><td>��վ��ַ��</td><td><input type=text size=70 name=url value='http://' maxlength=100><% = redx %></td></tr>
  <tr><td>���ҵ�����</td><td><select name=country size=1>
<option>�й�</option>
<option>���</option>
<option>̨��</option>
<option>����</option>
<option>Ӣ��</option>
<option>�ձ�</option>
<option>����</option>
<option>���ô�</option>
<option>�Ĵ�����</option>
<option>������</option>
<option>����˹</option>
<option>�����</option>
<option>����</option>
<option>������</option>
<option>�¹�</option>
<option>��������</option>
</select>&nbsp;&nbsp;&nbsp;&nbsp;վ�����ԣ�<select name=lang size=1>
<option>��������</option>
<option>��������</option>
<option>English</option>
<option>��������</option>
</select>&nbsp;&nbsp;&nbsp;�Ƽ���<input type=checkbox name=isgood value='yes'></td></tr>
<% ispic = "w" & upload_time(now_time) %>
  <tr><td>ͼƬ��ַ��</td><td><input type=test name=pic size=70 maxlength=100></td></tr>
  <tr><td>�ϴ�ͼƬ��</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=website&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr><td valign=top class=htd><br>��վ���ݣ�<br><=250B</td><td><textarea name=remark rows=5 cols=70></textarea></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' �� �� �� վ '></td></tr>
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
    sql    = "select id,c_id,s_id,name,url,isgood,hidden from " & data_name & sqladd & " order by id desc"
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
����<font class=red><% Response.Write rssum %></font>����վ��<% Response.Write "<a href='?action=add&c_id=" & cid & "&s_id=" & sid & "'>�����վ</a>" %>
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
        Response.Write website_center()
        rs.movenext
    Next

    rs.Close:Set rs = Nothing %></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>ҳ�Σ�<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>
��ҳ��<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %>
</td></tr></table>
    </td>
  </tr>
</table>
<%
End Sub

Function website_center()
    website_center = VbCrLf & "<tr" & mtr & ">" & _
    VbCrLf & "<td><a href='" & rs("url") & "' target=_blank title='�������վ'>" & i + (viewpage - 1)*nummer & ".</a> </td><td>" & _
    VbCrLf & "<a href='?action=edit&c_id=" & ncid & "&s_id=" & nsid & "&id=" & now_id & "'>" & rs("name") & "</a></td><td align=right><a href='?action=hidden&c_id=" & cid & "&s_id=" & sid & "&id=" & now_id & "&page=" & viewpage & "'>"

    If rs("hidden") = True Then
        website_center = website_center & "��"
    Else
        website_center = website_center & "<font class=red_2>��</font>"
    End If

    website_center = website_center & "</a> <a href='?action=isgood&c_id=" & cid & "&s_id=" & sid & "&id=" & now_id & "&page=" & viewpage & "'>"

    If rs("isgood") = True Then
        website_center = website_center & "<font class=red>��</font>"
    Else
        website_center = website_center & "��"
    End If

    website_center = website_center & "</a><input type=checkbox name=del_id value='" & now_id & "'></td></tr>"
End Function %>