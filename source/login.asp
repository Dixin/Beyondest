<!-- #include file="INCLUDE/config_login.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

if action="logout" then
  if login_username<>"" then
    conn.execute("delete from user_login where l_username='"&login_username&"'")
  end if
  response.cookies(web_cookies)("login_username")=""
  response.cookies(web_cookies)("login_password")=""
  response.cookies(web_cookies)("iscookies")=""
  if trim(request.servervariables("http_referer"))<>"" then
    call close_conn()
    response.redirect trim(request.servervariables("http_referer"))
    response.end
  end if
end if

if login_username<>"" and login_password<>"" then
  call close_conn()
  call format_redirect("user_main.asp")
  response.end
end if

select case action
case "register"
  tit="用户注册"
case "nopass"
  tit="忘记密码"
case else
  tit="用户登陆"
end select

call web_head(0,0,3,0,0)
'-----------------------------------center---------------------------------

select case action
case "register"
  call register_main()
case "login_chk"
  call login_chk()
case "nopass"
  call nopass()
case else
  call login_main()
end select

'---------------------------------center end-------------------------------
call web_end(0)
%>