<!-- #include file="config.asp" -->
<!-- #include file="config_nsort.asp" -->
<!-- #include file="skin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim atb
Dim nid
Dim sqladd
Dim name
atb       = " target=_blank":sk_bar = 12:sk_class = "end"
index_url = "down":n_sort = "down"
tit_fir   = format_menu(index_url)

Sub down_class_sort(t1,t2)
    Response.Write class_sortp(n_sort,index_url,t1,t2)
End Sub

Sub down_intro(introid,introsn)
    Dim tempix
    Dim sqlx
    Dim theintrox
    Dim thepicx
    Dim rsx
    tempix    = "<table border=0 width='100%' cellspacing=0 cellpadding=12><tr><td width='40%' align=center valign=top>"
    sqlx      = "select intro,pic from jk_sort where s_id=" & introid
    Set rsx   = conn.execute(sqlx)
    theintrox = rsx(0)
    thepicx   = rsx(1)
    tempix    = tempix & "<img src=images/down/" & thepicx & ".jpg></td><td>" & kong & "<font class=big><b>" & introsn & "</b></font>" & kong & "&nbsp;&nbsp;&nbsp;&nbsp;" & code_jk(theintrox) & "</td></tr></table>"
    rsx.Close:Set rsx = Nothing
    Response.Write tempix

End Sub

Sub down_class_sortt(t1,t2)
    Response.Write format_barc("<font class=" & sk_class & "><b>ר���б�</b></font>",class_sort(n_sort,index_url,t1,t2),3,0,6)
End Sub

Sub down_new_hot(n_jt,nnhead,nmore,nsql,nt,n_num,n_m,c_num,et,tt)
    Dim rs
    Dim sql
    Dim di
    Dim temp1
    Dim tim
    Dim counter
    Dim nhead:nhead = nnhead
    If n_jt <> "" Then n_jt = img_small(n_jt)
    temp1 = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
    sql   = "select top " & n_num + n_m & " id,name,username,tim,counter from down where hidden=1" & nsql

    Select Case nt
        Case "hot"
            sql = sql & " order by counter desc,id desc"
            If nhead = "" Then nhead = "��������"
        Case "good"
            sql = sql & " and types=5 order by id desc"
            If nhead = "" Then nhead = "�����Ƽ�"
        Case Else
            sql = sql & " order by id desc"
            If nhead = "" Then nhead = "���ڸ���"
    End Select

    Set rs = conn.execute(sql)

    For di = 1 To n_m
        If rs.eof Or rs.bof Then Exit For
        rs.movenext
    Next

    'if n_m>0 then rs.move(n_m)

    Do While Not rs.eof
        name  = rs("name"):tim = rs("tim"):counter = rs("counter")
        temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & " class=bw>" & n_jt & "<a href='down_view.asp?id=" & rs("id") & "'" & atb & " title='�������ƣ�" & code_html(name,1,0) & "<br>�� �� �ˣ�" & rs("username") & "<br>�����˴Σ�" & counter & "<br>����ʱ�䣺" & time_type(tim,88) & "'>" & code_html(name,1,c_num) & "</a>"
        If tt > 0 Then temp1 = temp1 & format_end(et,time_type(tim,tt) & ",<font class=blue>" & counter & "</font>")
        temp1 = temp1 & "</td></tr>"
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing
    temp1 = temp1 & vbcrlf & "</table>"
    Response.Write kong & format_barc("<font class=" & sk_class & "><b>" & nhead & "</b></font>",temp1,2,0,8)
End Sub

Sub down_new_hotr(n_jt,nnhead,nmore,nsql,nt,n_num,n_m,c_num,et,tt)
    Dim rs
    Dim sql
    Dim di
    Dim temp1
    Dim tim
    Dim counter
    Dim nhead:nhead = nnhead
    If n_jt <> "" Then n_jt = img_small(n_jt)
    If n_jt = "" Then n_jt = img_small("jt0")
    temp1 = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
    sql   = "select top " & n_num + n_m & " id,name,username,tim,counter,order from down where hidden=1" & nsql

    Select Case nt
        Case "hot"
            sql = sql & " order by counter desc,id desc"
            If nhead = "" Then nhead = "��������"
        Case "good"
            sql = sql & " and types=5 order by id desc"
            If nhead = "" Then nhead = "�����Ƽ�"
        Case Else
            sql = sql & " order by [order],id"
            If nhead = "" Then nhead = "���ڸ���"
    End Select

    Set rs = conn.execute(sql)

    For di = 1 To n_m
        If rs.eof Or rs.bof Then Exit For
        rs.movenext
    Next

    'if n_m>0 then rs.move(n_m)

    Do While Not rs.eof
        name  = rs("name"):tim = rs("tim"):counter = rs("counter")
        temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & " class=bw>" & n_jt & "<a href='down_view.asp?id=" & rs("id") & "'" & atb & " title='�������ƣ�" & code_html(name,1,0) & "<br>�� �� �ˣ�" & rs("username") & "<br>�����˴Σ�" & counter & "<br>����ʱ�䣺" & time_type(tim,88) & "'>" & code_html(name,1,c_num) & "</a>"
        If tt > 0 Then temp1 = temp1 & format_end(et,time_type(tim,tt) & ",<font class=blue>" & counter & "</font>")
        temp1 = temp1 & "</td></tr>"
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing
    temp1 = temp1 & vbcrlf & "</table>"
    Response.Write format_barc("<font class=" & sk_class & "><b>" & nhead & "</b></font>",temp1,3,0,8)
End Sub

Sub down_new_hotrn(n_jt,nnhead,nmore,nsql,nt,n_num,n_m,c_num,et,tt)
    Dim rs
    Dim sql
    Dim di
    Dim temp1
    Dim tim
    Dim counter
    Dim nhead:nhead = nnhead
    If n_jt <> "" Then n_jt = img_small(n_jt)
    temp1 = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
    sql   = "select top " & n_num + n_m & " id,name,username,tim,counter from down where hidden=1" & nsql

    Select Case nt
        Case "hot"
            sql = sql & " order by counter desc,id desc"
            If nhead = "" Then nhead = "��������"
        Case "good"
            sql = sql & " and types=5 order by id desc"
            If nhead = "" Then nhead = "�����Ƽ�"
        Case Else
            sql = sql & " order by id desc"
            If nhead = "" Then nhead = "���ڸ���"
    End Select

    Set rs = conn.execute(sql)

    For di = 1 To n_m
        If rs.eof Or rs.bof Then Exit For
        rs.movenext
    Next

    'if n_m>0 then rs.move(n_m)

    Do While Not rs.eof
        name  = rs("name"):tim = rs("tim"):counter = rs("counter")
        temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & " class=bw>" & n_jt & "<a href='down_view.asp?id=" & rs("id") & "'" & atb & " title='�������ƣ�" & code_html(name,1,0) & "<br>�� �� �ˣ�" & rs("username") & "<br>�����˴Σ�" & counter & "<br>����ʱ�䣺" & time_type(tim,88) & "'>" & code_html(name,1,c_num) & "</a>"
        If tt > 0 Then temp1 = temp1 & format_end(et,time_type(tim,tt) & ",<font class=blue>" & counter & "</font>")
        temp1 = temp1 & "</td></tr>"
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing
    temp1 = temp1 & vbcrlf & "</table>"
    Response.Write format_barc("<font class=" & sk_class & "><b>" & nhead & "</b></font>",temp1,1,1,8)
End Sub

Sub down_pic(nnhead,dsql,nt,n_num,c_num)
    Dim rs
    Dim sql
    Dim temp1
    Dim nhead:nhead = nnhead
    temp1 = "<table border=0 width='100%' cellspacing=0 cellpadding=2><tr align=center valign=top>"
    sql   = "select top " & n_num & " id,name,tim,pic from down where hidden=1" & dsql

    Select Case nt
        Case "hot"
            sql = sql & " order by counter desc,id desc"
            If nhead = "" Then nhead = "�ȵ�����"
        Case "good"
            sql = sql & " and types=5 order by id desc"
            If nhead = "" Then nhead = "��Ʒ�Ƽ�"
        Case Else
            sql = sql & " order by id desc"
            If nhead = "" Then nhead = "��������"
    End Select

    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        name  = rs("name"):nid = rs("id")
        temp1 = temp1 & vbcrlf & "<td width='" & Int(100\n_num) & "%'><table border=0 cellspacing=0 cellpadding=2 width='100%' class=tf><tr><td align=center><a href='down_view.asp?id=" & nid & "'" & atb & "><img src='images/down/" & rs("pic") & "' border=0 ></a></td></tr>" & _
        vbcrlf & "<tr><td align=center class=bw><a href='down_view.asp?id=" & nid & "'" & atb & " class=red_3><b>" & code_html(name,1,0) & "</b></a></td></tr></table></td>"
        rs.movenext
    Loop

    If temp1 = "<table border=0 width='100%' cellspacing=0 cellpadding=2><tr align=center valign=top>" Then temp1 = temp1 & "<td>��</td>"
    rs.Close:Set rs = Nothing
    temp1 = temp1 & "</tr></table>"
    Response.Write format_barc("<font class=" & sk_class & "><b>" & nhead & "</b></font>",temp1,3,0,5)
End Sub

Sub down_remark(njt)
    Dim temp1
    temp1 = vbcrlf & "<table border=0 width='98%' align=center>" & _
    vbcrlf & "<tr><td>" & img_small(njt) & "��վ�Ƽ�ʹ�� <a href='file/soft/flashget.rar'>���ʿ쳵</a> �������֣�һ������������ء�</td></tr>" & _
    vbcrlf & "<tr><td>" & img_small(njt) & "��������ֱ�վ���κ�������������⣬��<a href='gbook.asp?action=write'" & atb & ">����֪ͨ��</a>��лл��</td></tr>" & _
    vbcrlf & "<tr><td>" & img_small(njt) & "��վ������ļ����� <a href='" & web_var(web_down,5) & "/soft/winrar.exe'>WinRAR</a> ѹ�������ڴ��������°汾��</td></tr>" & _
    vbcrlf & "<tr><td class=red>" & img_small(njt) & "��������ӱ�վ�ļ�����ע�����ԣ�<a href='" & web_var(web_config,2) & "'" & atb & ">" & web_var(web_config,1) & "</a>��лл����֧�֣�</td></tr>" & _
    vbcrlf & "<tr><td>" & img_small(njt) & "��վ�ṩ���������ؽ���������������Ȩ���뼰ʱ <a href='gbook.asp?action=write'" & atb & ">֪ͨ��</a> ��<font color='#ff0000'>ϣ�����֧�����档</font></td></tr>" & _
    vbcrlf & "<tr><td>" & img_small(njt) & "��ӭ��ҵ���վ <a href='forum.asp'>��̳</a> ����ͽ������ļ��⡣��л���ķ��ʣ�</td></tr>" & _
    vbcrlf & "</table>"
    Response.Write format_barc("<font class=" & sk_class & "><b>��������˵��</b></font>",temp1,4,1,"")
End Sub

Sub down_tool()
    Dim temp1
    temp1 = vbcrlf & "<table border=0 cellspacing=0 cellpadding=2><tr><td height=5></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_winrar.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/winrar.exe'>WinRAR</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_qq.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/qq.rar'>QQ2004(ȥ�����IP)</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_winamp.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/winamp.rar'>Winamp</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_realone.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/realoneplayer.rar'>RealOnePlayer</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_wmp.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/wmp2k.rar'>Windows Midia Player(2k&98)</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_wmp.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/wmpxp.rar'>Windows Midia Player(xp)</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_flashget.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/flashget.rar'>Flashget</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_cuteftp.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/flashfxp.rar'>FlashFXP</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_wopti.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/wom.rar'>Windows�Ż���ʦ</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_norton.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/norton.rar'>Norton Antivirus 2004</a></td></tr>" & _
    vbcrlf & "<tr><td><img src='images/down/tool_norton.gif' border=0 align=absmiddle>&nbsp;<a href='" & web_var(web_down,5) & "/soft/nortonsp.rar'>Norton���²�����</a></td></tr>" & _
    vbcrlf & "</table>"
    Response.Write format_barc("<font class=" & sk_class & "><b>���ù���</b></font>",temp1,1,1,1)
End Sub

Sub down_atat()
    Dim temp1
    Dim num1
    Dim num2
    Dim num3
    Dim sq
    Dim rs
    sql    = "select count(id) from down where hidden=1 and tim>=#" & FormatDateTime(FormatDateTime(now_time,2)) & "#"
    Set rs = conn.execute(sql)
    num1   = rs(0)
    rs.Close
    sql    = "select num_down from configs where id=1"
    'sql="select count(id) from down where hidden=1"
    Set rs = conn.execute(sql)
    num2   = rs(0)
    rs.Close
    sql    = "select sum(counter) from down where hidden=1"
    Set rs = conn.execute(sql)
    num3   = rs(0)
    rs.Close:Set rs = Nothing
    temp1  = vbcrlf & "<table border=0 cellspacing=0 cellpadding=3><tr><td height=5></td></tr>" & _
    vbcrlf & "<tr><td>���ո��£�<font class=red>" & num1 & "</font>������</td></tr>" & _
    vbcrlf & "<tr><td>����������<font class=red>" & num2 & "</font>������</td></tr>" & _
    vbcrlf & "<tr><td>�����أ�<font class=red>" & num3 & "</font>�˴�</td></tr>" & _
    vbcrlf & "<tr><td>[ <a href='down_list.asp'>�� ������ַ���</a> ]</td></tr>" & _
    vbcrlf & "<tr><td>[ <a href='gbook.asp?action=write'>�� �������ӱ���</a> ]</td></tr>" & _
    vbcrlf & "<tr><td>" & put_type("down") & "</td></tr>" & _
    vbcrlf & "</table>"
    Response.Write format_barc("<font class=" & sk_class & "><b>��Ŀͳ��</b></font>",temp1,2,0,5)
End Sub

Sub down_main()
    Dim rs2
    Dim sql2

    If cid = 0 Then
        sql2    = "select c_id,c_name from jk_class where nsort='" & n_sort & "' order by c_order"
        Set rs2 = conn.execute(sql2)

        Do While Not rs2.eof
            nid = rs2("c_id"):sqladd = " and c_id=" & nid %>
<tr align=center valign=top>
<td width='60%'><% Call down_new_hotr("jt0","<a href='down_list.asp?c_id=" & nid & "'><font class=" & sk_class & ">" & rs2("c_name") & "</font></a>","<a href='down_list.asp?c_id=" & nid & "&action=more'><font class=" & sk_class & ">����...</font></a>",sqladd,"new",15,0,20,1,8) %></td>
<td width=1 bgcolor='<% = web_var(web_color,3) %>'></td>
<td bgcolor='<% = web_var(web_color,1) %>'><%
            Call down_new_hotr("","��������","",sqladd,"hot",5,0,11,1,0)
            Call down_pic("վ���Ƽ�",sqladd,"good",1,10) %></td>
</tr>
<%
            rs2.movenext
        Loop

        rs2.Close:Set rs2 = Nothing
    Else

        If sid = 0 Then
            sql2    = "select s_id,s_name from jk_sort where c_id=" & cid & " order by s_order"
            Set rs2 = conn.execute(sql2)
            Response.Write "<tr height=1><td colspan=3 align=center>" & format_img("rdown.jpg") & "</td></tr>"

            Do While Not rs2.eof
                nid = rs2("s_id"):sqladd = " and c_id=" & cid & " and s_id=" & nid %>
<tr height=1><td colspan=3 bgcolor="<% Response.Write web_var(web_color,3) %>"></td></tr>
<tr align=center><td colspan=3>
<% Call down_intro(nid,rs2("s_name")) %>
</td></tr>
<tr align=center valign=top>
<td width=400><% Call down_new_hotr("jt0","<a href='down_list.asp?c_id=" & cid & "&s_id=" & nid & "'><font class=" & sk_class & ">" & rs2("s_name") & "</font></a>","<a href='down_list.asp?c_id=" & cid & "&s_id=" & nid & "&action=more'><font class=" & sk_class & ">����...</font></a>",sqladd,"new",40,0,20,1,8) %></td>
<td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
<td><%
                Call down_new_hotrn("jt0","��������","",sqladd,"hot",40,0,11,1,0)
                'call down_pic("վ���Ƽ�",sqladd,"good",1,10) %></td>
</tr>
<%
                rs2.movenext
            Loop

            rs2.Close:Set rs2 = Nothing
        Else
            sql2    = "select jk_class.c_name,jk_sort.s_name from jk_sort inner join jk_class on jk_sort.c_id=jk_class.c_id where jk_sort.c_id=" & cid & " and jk_sort.s_id=" & sid
            Set rs2 = conn.execute(sql2)

            If rs2.eof And rs2.bof Then
                rs2.Close:Set rs2 = Nothing
                cid = 0:sid = 0

                Call down_main():Exit Sub
                End If

                sqladd = " and c_id=" & cid & " and s_id=" & sid %>
<tr align=center>
<td colspan=3><% Call down_intro(sid,rs2("s_name")) %></td>
</tr>
<tr align=center>
<td colspan=3><% Call down_pic("վ���Ƽ�",sqladd,"good",5,20) %></td>
</tr>
<tr align=center valign=top>
<td width=400><% Call down_new_hotr("jt0","<a href='down_list.asp?c_id=" & cid & "'><font class=" & sk_class & ">" & rs2("c_name") & "</font></a> �� <a href='down_list.asp?c_id=" & cid & "&s_id=" & sid & "'><font class=" & sk_class & ">" & rs2("s_name") & "</font></a>","<a href='down_list.asp?c_id=" & cid & "&s_id=" & sid & "&action=more'><font class=" & sk_class & ">����...</font></a>",sqladd,"new",40,0,20,1,8) %></td>
<td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
<td><%
                Call down_new_hotrn("jt0","��������","",sqladd,"hot",40,0,11,1,0) %></td>
</tr>
<%
                rs2.Close:Set rs2 = Nothing
            End If

        End If

    End Sub

    Sub down_more(c_num,tt)
        Dim temp1
        Dim tim
        Dim cnum
        Dim sql2
        Dim mhead
        Dim name
        Dim c1
        Dim c2
        Dim sql
        Dim rs
        Dim cname
        Dim sname
        c1       = web_var(web_color,6):c2 = web_var(web_color,1)
        pageurl  = "?action=more&"
        keyword  = code_form(Request.querystring("keyword"))
        sea_type = Trim(Request.querystring("sea_type"))
        If sea_type <> "username" Then sea_type = "name"
        Call cid_sid_sql(2,sea_type)

        temp1 = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=4><tr><td colspan=5 height=5></td></tr>" & _
        vbcrlf & "<tr align=left height=20 valign=bottom>" & _
        vbcrlf & "<td width='6%'>���</td>" & _
        vbcrlf & "<td width='44%'>��������</td>" & _
        vbcrlf & "<td width='28%'>��������</td>" & _
        vbcrlf & "<td width='12%'>�Ƽ��ȼ�</td>" & _
        vbcrlf & "<td width='10%'>���ش���</td>" & _
        vbcrlf & "</tr>" & _
        vbcrlf & "<tr><td colspan=5 height=1 background='images/bg_dian.gif'></td></tr>"
        sql          = "select id,name,username,tim,counter,types from down where hidden=1 " & sqladd

        If cid > 0 Then
            sql      = sql & " and c_id=" & cid

            If sid > 0 Then
                sql  = sql & " and s_id=" & sid
                sql2 = "select jk_class.c_name,jk_sort.s_name from jk_sort inner join jk_class on jk_sort.c_id=jk_class.c_id where jk_sort.c_id=" & cid & " and jk_sort.s_id=" & sid
            Else
                sql2 = "select c_name from jk_class where c_id=" & cid
            End If

        End If

        sql = sql & " order by id desc"

        If cid > 0 Then
            Set rs    = conn.execute(sql2)

            If rs.eof And rs.bof Then rs.Close:Set rs = Nothing:Call down_main():Exit Sub
                cname = code_html(rs("c_name"),1,0)
                If sid > 0 Then sname = code_html(rs("s_name"),1,0)
                rs.Close
            Else
                cname = "�������"
            End If

            mhead     = "<a href='down_list.asp?c_id=" & cid & "'><b><font class=" & sk_class & ">" & cname & "</font></b></a>"
            If cid > 0 And sid > 0 Then mhead = mhead & "&nbsp;<font class=" & sk_class & ">��</font>&nbsp;<a href='down_list.asp?c_id=" & cid & "&s_id=" & sid & "'><b><font class=" & sk_class & ">" & sname & "</font></b></a>"

            Set rs = Server.CreateObject("adodb.recordset")
            rs.open sql,conn,1,1

            If rs.eof And rs.bof Then
                rssum = 0
            Else
                rssum = rs.recordcount
            End If

            Call format_pagecute()

            If Int(viewpage) > 1 Then
                rs.move (viewpage - 1)*nummer
            End If

            For i = 1 To nummer
                If rs.eof Then Exit For
                name  = rs("name"):tim = rs("tim")
                temp1 = temp1 & vbcrlf & "<tr onmouseover=""javascript:this.bgColor='" & c1 & "';"" onmouseout=""javascript:this.bgColor='';""><td>" & i + (viewpage - 1)*nummer & ".</td>" & _
                vbcrlf & "<td><a href='down_view.asp?id=" & rs("id") & "'" & atb & " title='�������ƣ�" & code_html(name,1,0) & "<br>�� �� �ˣ�" & rs("username") & "<br>����ʱ�䣺" & tim & "'>" & code_html(name,1,c_num) & "</a></td>" & _
                vbcrlf & "<td>" & time_type(tim,tt) & "</td>" & _
                vbcrlf & "<td><img src='images/down/star" & rs("types") & ".gif' border=0></td>" & _
                vbcrlf & "<td align=center class=blue>" & rs("counter") & "</td></tr>" & _
                vbcrlf & "<tr><td colspan=5 height=1 background='images/bg_dian.gif'></td></tr>"
                rs.movenext
            Next

            rs.Close:Set rs = Nothing
            temp1 = temp1 & vbcrlf & "<tr><td colspan=5 height=25 valign=bottom>" & _
            vbcrlf & "����&nbsp;<font class=red>" & rssum & "</font>&nbsp;���ļ�&nbsp;" & _
            vbcrlf & "ҳ�Σ�<font class=red>" & viewpage & "</font>/<font class=red>" & thepages & "</font>&nbsp;" & _
            vbcrlf & "��ҳ��" & jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000") & _
            vbcrlf & "</td></tr></table>"
            Response.Write "<tr><td colspan=3 align=center>" & format_barc(mhead,temp1,3,0,11) & "</td></tr>"
        End Sub

        Sub down_sea()
            Dim temp1
            Dim nid
            Dim nid2
            Dim rs
            Dim sql
            Dim rs2
            Dim sql2
            temp1 = vbcrlf & "<table border=0 cellspacing=0 cellpadding=0 align=center>" & _
            vbcrlf & "<script language=javascript><!--" & _
            vbcrlf & "function down_sea()" & _
            vbcrlf & "{" & _
            vbcrlf & "  if (down_sea_frm.keyword.value==""������ؼ���"")" & _
            vbcrlf & "  {" & _
            vbcrlf & "    alert(""������������ǰ������Ҫ��ѯ�� �ؼ��� ��"");" & _
            vbcrlf & "    down_sea_frm.keyword.focus();" & _
            vbcrlf & "    return false;" & _
            vbcrlf & "  }" & _
            vbcrlf & "}" & _
            vbcrlf & "--></script>" & _
            vbcrlf & "<form name=down_sea_frm action='down_list.asp' method=get onsubmit=""return down_sea()"">" & _
            vbcrlf & "<input type=hidden name=action value='more'><tr><td height=3></td></tr>" & _
            vbcrlf & "<tr><td>" & _
            vbcrlf & "  <table border=0><tr><td colspan=2><input type=text name=keyword value='������ؼ���' onfocus=""if (value =='������ؼ���'){value =''}"" onblur=""if (value ==''){value='������ؼ���'}"" size=20 maxlength=20></td></tr>" & _

            vbcrlf & "  </table>" & _
            vbcrlf & "</td></tr><tr><td>" & _
            vbcrlf & "  <table border=0><tr>" & _
            vbcrlf & "  <td><select name=c_id sizs=1><option value=''>ȫ�����</option>"
            sql           = "select c_id,c_name from jk_class where nsort='" & n_sort & "' order by c_order,c_id"
            Set rs        = conn.execute(sql)

            Do While Not rs.eof
                nid       = Int(rs(0))
                temp1     = temp1 & vbcrlf & "<option value='" & nid & "' class=bg_2"
                If cid = nid Then temp1 = temp1 & " selected"
                temp1     = temp1 & ">" & rs(1) & "</option>"
                sql2      = "select s_id,s_name from jk_sort where c_id=" & nid & " order by s_order,s_id"
                Set rs2   = conn.execute(sql2)

                Do While Not rs2.eof
                    nid2  = rs2(0)
                    temp1 = temp1 & vbcrlf & "<option value='" & nid & "&s_id=" & nid2 & "'"
                    If sid = nid2 Then temp1 = temp1 & " selected"
                    temp1 = temp1 & ">��" & rs2(1) & "</option>"
                    rs2.movenext
                Loop

                rs2.Close:Set rs2 = Nothing
                rs.movenext
            Loop

            rs.Close:Set rs = Nothing
            temp1 = temp1 & vbcrlf & "</select></td>" & _
            vbcrlf & "  <td></td></tr>" & _
            vbcrlf & "  <tr height=25><td><select name=sea_type size=1><option value='name'>��������</option><option value='username'>������</option></select></td>" & _
            vbcrlf & "  <td align=left><input type=image src='images/small/search_go.gif' border=0 height=25 width=40></td>" & _
            vbcrlf & "  </tr></table>" & _
            vbcrlf & "</td></tr>" & _
            vbcrlf & "</form><tr><td height=1></td></tr></table>"
            Response.Write kong & format_barc("<font class=" & sk_class & "><b>��������</b></font>",temp1,2,0,4)
        End Sub %>