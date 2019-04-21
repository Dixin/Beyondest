<!-- #include file="include/onlogin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim classid
classid = Trim(Request("class_id"))
action  = Trim(Request("action"))

If Not(IsNumeric(classid)) Or (action <> "up" And action <> "down") Then
    Response.redirect "admin_forum.asp"
    Response.End
End If %>
<!-- #include file="include/conn.asp" -->
<%
Dim tmp_id_1
Dim tmp_id_2
Dim tmp_order_1
Dim tmp_order_2
Dim sqladd
Dim update_ok
update_ok  = "no"

If action = "up" Then
    sqladd = " desc"
Else
    sqladd = ""
End If

sql    = "select * from bbs_class where class_id=" & classid
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing
    close_conn
    Response.redirect "admin_forum.asp"
    Response.End
End If

rs.Close:Set rs = Nothing

sql    = "select * from bbs_class order by class_order" & sqladd
Set rs = conn.execute(sql)

Do While Not rs.eof

    If Int(rs("class_id")) = Int(classid) Then
        tmp_id_1    = classid
        tmp_order_1 = rs("class_order")
        rs.movenext

        If Not rs.eof Then
            tmp_id_2    = rs("class_id")
            tmp_order_2 = rs("class_order")
            update_ok   = "yes"
            Exit Do
        End If

        Exit Do
    End If

    rs.movenext
Loop

rs.Close:Set rs = Nothing

If update_ok = "yes" Then
    sql = "update bbs_class set class_order=" & tmp_order_2 & " where class_id=" & tmp_id_1
    conn.execute(sql)
    sql = "update bbs_class set class_order=" & tmp_order_1 & " where class_id=" & tmp_id_2
    conn.execute(sql)
End If

close_conn

Response.redirect "admin_forum.asp" %>