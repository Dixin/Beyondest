<!-- #include file="include/onlogin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim classid,forumid
classid = Trim(Request("class_id"))
forumid = Trim(Request("forum_id"))
action  = Trim(Request("action"))

If Not(IsNumeric(classid)) Or Not(IsNumeric(forumid)) Or (action <> "up" And action <> "down") Then
    Response.redirect "admin_forum.asp"
    Response.End
End If %>
<!-- #include file="include/conn.asp" -->
<%
Dim tmp_id_1,tmp_id_2,tmp_order_1,tmp_order_2,sqladd,update_ok,rssum
update_ok  = "no"

If action = "up" Then
    sqladd = " desc"
Else
    sqladd = ""
End If

sql    = "select forum_order from bbs_forum where forum_id=" & forumid & " and class_id=" & classid
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing
    close_conn
    Response.redirect "admin_forum.asp"
    Response.End
End If

rs.Close:Set rs = Nothing

sql    = "select forum_id,forum_order from bbs_forum where class_id=" & classid & " order by forum_order" & sqladd & ",forum_id desc"
Set rs = conn.execute(sql)

Do While Not rs.eof

    If Int(rs("forum_id")) = Int(forumid) Then
        tmp_id_1    = forumid
        tmp_order_1 = rs("forum_order")
        rs.movenext

        If Not rs.eof Then
            tmp_id_2    = rs("forum_id")
            tmp_order_2 = rs("forum_order")
            update_ok   = "yes"
            Exit Do
        End If

        Exit Do
    End If

    rs.movenext
Loop

rs.Close:Set rs = Nothing

If update_ok = "yes" Then
    sql = "update bbs_forum set forum_order=" & tmp_order_2 & " where forum_id=" & tmp_id_1
    Response.Write sql
    conn.execute(sql)
    sql = "update bbs_forum set forum_order=" & tmp_order_1 & " where forum_id=" & tmp_id_2
    conn.execute(sql)
End If

close_conn

Response.redirect "admin_forum.asp" %>