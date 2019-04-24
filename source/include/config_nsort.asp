<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim n_sort,cid,sid
cid = 0:sid = 0

Function class_sort(sns,surl,idc,ids)
    Dim nid,rs,sql,rs2,sql2,nnid,j,k:k = 12
    class_sort         = vbcrlf & "<table border=0><tr><td height=3></td></tr>"
    sql                = "select c_name,c_id from jk_class where nsort='" & sns & "' order by c_order"
    Set rs             = conn.execute(sql)

    Do While Not rs.eof
        nid            = rs("c_id")
        class_sort     = class_sort & vbcrlf & "<tr valign=top><td><table border=0><tr><td height=18>|&nbsp;<a href='" & surl & "_list.asp?c_id=" & nid & "'"
        If idc = nid Then class_sort = class_sort & " class=red"
        class_sort     = class_sort & ">" & rs("c_name") & "</a>&nbsp;|</td></tr></table></td><td>" & vbcrlf & "<table border=0>"
        sql2           = "select s_name,s_id from jk_sort where c_id=" & nid & " order by s_order"
        Set rs2        = conn.execute(sql2)

        Do While Not rs2.eof
            class_sort = class_sort & "<tr>"

            For j = 1 To k
                If rs2.eof Then Exit For
                nnid       = rs2("s_id")
                class_sort = class_sort & vbcrlf & "<td>&nbsp;<a href='" & surl & "_list.asp?c_id=" & nid & "&s_id=" & rs2("s_id") & "'"
                If ids = nnid Then class_sort = class_sort & " class=red_3"
                class_sort = class_sort & ">" & rs2("s_name") & "</a>&nbsp;</td>"
                rs2.movenext
            Next

            class_sort = class_sort & "</tr>"
        Loop

        class_sort     = class_sort & "</table></td></tr>"
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing
    class_sort = class_sort & vbcrlf & "<tr><td height=1></td></tr></table>"
End Function

Function class_sortp(sns,surl,idc,ids)
    Dim nid,rs,sql,rs2,sql2,stt,con,spic,nnid,j,k:k = 7
    class_sortp  = ""
    sql          = "select c_name,c_id from jk_class where nsort='" & sns & "' order by c_order"
    Set rs       = conn.execute(sql)

    Do While Not rs.eof
        nid      = rs("c_id")
        stt      = "&nbsp;<a href='" & surl & "_list.asp?c_id=" & nid & "'"
        If idc = nid Then stt = stt & " class=red"
        stt      = stt & ">" & rs("c_name") & "</a>"
        sql2     = "select s_name,s_id,pic from jk_sort where c_id=" & nid & " order by s_order"
        Set rs2  = conn.execute(sql2)
        con      = "<table wdth=100% border=0 cellspacing=0 cellpadding=0><tr align=left>"

        Do While Not rs2.eof
            nnid = rs2("s_id")
            spic = rs2("pic")
            con  = con & "<td width=113 valign=left>" & kong & "<table border=0 cellspacing=0 cellpadding=0><tr><td align=left>&nbsp;<a href='" & surl & "_list.asp?c_id=" & nid & "&s_id=" & rs2("s_id") & "'"
            If ids = nnid Then con = con & " class=red_3"
            con  = con & "><img src=images/down/" & spic & "x.jpg width=80 height=80 border=0></a></td><tr><td height=20 align=left>&nbsp;<a href='" & surl & "_list.asp?c_id=" & nid & "&s_id=" & rs2("s_id") & "'"
            If ids = nnid Then con = con & " class=red_3"
            con  = con & ">" & rs2("s_name") & "</a>&nbsp;</td></tr></table></td><td width=1 bgcolor=" & web_var(web_color,3) & "></td>"
            rs2.movenext
        Loop

        con         = con & "<td></td></tr></table>"
        class_sortp = class_sortp & format_barc(stt,con,5,0,71)
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing
End Function

Function nsort_left(n_sort,cc_id,ss_id,link_url,left_type)
    Dim rs1,sql1,rs2,sql2,ccid,ssid:cc_id = Int(cc_id):ss_id = Int(ss_id)
    nsort_left             = vbcrlf & "<table border=0><tr><td height=1></td></tr>"
    sql1                   = "select c_id,c_name from jk_class where nsort='" & n_sort & "' order by c_order,c_id"
    Set rs1                = conn.execute(sql1)

    Do While Not rs1.eof
        ccid               = Int(rs1(0))

        If cc_id = ccid Or left_type = 0 Then
            nsort_left     = nsort_left & vbcrlf & "<tr><td>"

            If cc_id = ccid Then
                nsort_left = nsort_left & img_small("jt1")
            Else
                nsort_left = nsort_left & img_small("jt12")
            End If

            nsort_left     = nsort_left & "<a href='" & link_url & "c_id=" & ccid & "'>" & rs1(1) & "</a></td></tr>"
            sql2           = "select s_id,s_name from jk_sort where c_id=" & ccid & " order by s_order,s_id"
            Set rs2        = conn.execute(sql2)

            Do While Not rs2.eof
                ssid       = Int(rs2(0))
                nsort_left = nsort_left & vbcrlf & "<tr><td>&nbsp;&nbsp;" & img_small("jt0") & "<a href='" & link_url & "c_id=" & ccid & "&s_id=" & rs2(0) & "'"
                If ssid = ss_id Then nsort_left = nsort_left & " class=red_3"
                nsort_left = nsort_left & ">" & rs2(1) & "</a></td></tr>"
                rs2.movenext
            Loop

            rs2.Close:Set rs2 = Nothing
        Else
            nsort_left = nsort_left & vbcrlf & "<tr><td>" & img_small("jt12") & "<a href='" & link_url & "c_id=" & ccid & "'>" & rs1(1) & "</a></td></tr>"
        End If

        rs1.movenext
    Loop

    rs1.Close:Set rs1 = Nothing
    nsort_left = nsort_left & vbcrlf & "<tr><td height=1></td></tr></table>"
End Function

Sub cid_sid_sql(csst,csstt)
    Dim temp:temp = csstt

    If cid > 0 Then
        sqladd      = sqladd & " and c_id=" & cid
        pageurl     = pageurl & "c_id=" & cid & "&"

        If sid > 0 Then
            sqladd  = sqladd & " and s_id=" & sid
            pageurl = pageurl & "s_id=" & sid & "&"
        End If

    End If

    If csst = 1 Or csst = 2 Then

        If Len(keyword) > 0 Then
            sqladd  = sqladd & " and " & temp & " like '%" & keyword & "%'"
            pageurl = pageurl & "keyword=" & Server.urlencode(keyword) & "&"
            If csst = 2 Then pageurl = pageurl & "sea_type=" & sea_type & "&"
        End If

    End If

End Sub

Sub cid_sid()
    cid = Trim(Request.querystring("c_id"))
    sid = Trim(Request.querystring("s_id"))

    If Not(IsNumeric(cid)) Then

        If Len(cid) > 0 Then
            cid = Replace(cid,"&s_id=",",")
            sid = Mid(cid,InStr(cid,",") + 1,Len(cid))
            cid = Mid(cid,1,InStr(cid,",") - 1)
        Else
            cid = 0
        End If

    End If

    If Not(IsNumeric(cid)) Then cid = 0
    If Not(IsNumeric(sid)) Then sid = 0
    cid = Int(cid):sid = Int(sid)
End Sub

Function put_type(pts)
    put_type = ""

    Select Case pts
        Case "article"
            put_type = "我要发表文章"
        Case "news"
            put_type = "我要发布新闻"
        Case "down"
            put_type = "我要添加音乐"
        Case "website"
            put_type = "我要推荐网站"
        Case "gallery"
            put_type = "我要上传贴图"
    End Select

    If put_type <> "" Then put_type = "[ <a href='user_put.asp?action=" & pts & "'>→ " & put_type & "</a> ]"
End Function %>