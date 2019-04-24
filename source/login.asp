<!-- #include file="INCLUDE/config_login.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

If action = "logout" Then

    If login_username <> "" Then
        conn.execute("delete from user_login where l_username='" & login_username & "'")
    End If

    Response.cookies(web_cookies)("login_username") = ""
    Response.cookies(web_cookies)("login_password") = ""
    Response.cookies(web_cookies)("iscookies") = ""

    If Trim(Request.servervariables("http_referer")) <> "" Then
        Call close_conn()
        Response.redirect Trim(Request.servervariables("http_referer"))
        Response.End
    End If

End If

If login_username <> "" And login_password <> "" Then
    Call close_conn()
    Call format_redirect("user_main.asp")
    Response.End
End If

Select Case action
    Case "register"
        tit = "用户注册"
    Case "nopass"
        tit = "忘记密码"
    Case Else
        tit = "用户登陆"
End Select

Call web_head(0,0,3,0,0)
'-----------------------------------center---------------------------------

Select Case action
    Case "register"
        Call register_main()
    Case "login_chk"
        Call login_chk()
    Case "nopass"
        Call nopass()
    Case Else
        Call login_main()
End Select

'---------------------------------center end-------------------------------
Call web_end(0) %>