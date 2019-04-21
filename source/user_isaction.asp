<!-- #include file="include/config.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim cancel
Dim old_url
Dim username
cancel   = Trim(Request.querystring("cancel"))
old_url  = Request.servervariables("http_referer")
If Len(old_url) < 3 Then old_url = "user_main.asp"
username = Trim(Request.querystring("username"))

If symbol_name(username) <> "yes" Or (action <> "locked" And action <> "shield") Then
    Response.redirect old_url
    Response.End
End If %>
<!-- #include file="include/skin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
Call web_head(2,2,0,0,0)
If format_user_power(login_username,login_mode,"") <> "yes" Then Call close_conn():Call cookies_type("power")

sql    = "select power,popedom from user_data where username='" & username & "'"
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    Response.Write username
    Response.End
    rs.Close:Set rs = Nothing
    Call close_conn()
    Response.redirect old_url
    Response.End
End If

Dim user_popedom
Dim u_power
Dim aname
Dim fname
Dim popedom_true
u_power      = rs("power")
user_popedom = rs("popedom")
rs.Close:Set rs = Nothing

If Int(format_power(u_power,2)) = 1 Then
    Call close_conn()
    Call cookies_type("power")
    Response.End
End If

popedom_true = "yes"
If cancel = "yes" Then fname = "解除"

Select Case action
    Case "shield"
        aname = "屏蔽"
        Call useres_popedom(42)
    Case "locked"
        aname = "锁定"
        Call useres_popedom(41)
End Select

Call useres_msg()

Call close_conn()
'response.redirect old_url
'response.end
Sub useres_popedom(pn)
    Dim temp1
    Dim temp2
    Dim temp3

    If Len(user_popedom) <> 50 Or pn > Len(user_popedom) Then popedom_true = "no":Exit Sub
        temp1 = Left(user_popedom,pn - 1)
        temp2 = popedom_format(user_popedom,pn)
        temp3 = Right(user_popedom,Len(user_popedom) - pn)

        If cancel = "yes" Then
            temp2 = "0"
        Else
            temp2 = "1"
        End If

        sql = "update user_data set popedom='" & temp1 & temp2 & temp3 & "' where username='" & username & "'"
        conn.execute(sql)
    End Sub

    Sub useres_msg()

        If popedom_true = "yes" Then
            Response.Write "<script language=javascript>alert(""已成对用户（" & username & "）进行了如下操作：\n\n" & fname & " " & aname & "\n\n点击返回！"");location.href='" & old_url & "';</script>"
        Else
            Response.Write "<script language=javascript>alert(""在对用户（" & username & "）进行操作时出现了严重错误！\n\n请与站长联系！\n\n点击返回！"");location.href='" & old_url & "';</script>"
        End If

    End Sub %>