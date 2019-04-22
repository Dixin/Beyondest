<!-- #include file="include/config_user.asp" -->
<!-- #include file="include/jk_md5.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim err_head
tit      = "修改资料"
err_head = img_small("jt0")

Call web_head(2,0,0,0,0)

If Int(popedom_format(login_popedom,41)) Then Call close_conn():Call cookies_type("locked")
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong & table1 & vbcrlf & "<tr" & table2 & " height=25><td height=20 class=end background=images/" & web_var(web_config,5) & "/bar_3_bg.gif>&nbsp;" & img_small(us) & "&nbsp;&nbsp;<b>修改我的个人资料</b></td></tr><tr" & table3 & "><td height=150 align=center>"

sql    = "select * from user_data where username='" & login_username & "'"
Set rs = Server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing
    Call close_conn()
    Call format_redirect("login.asp")
    Response.End
End If

Select Case Trim(Request.form("user_edit"))
    Case "yes"
        Response.Write edit_chk()
    Case Else
        Response.Write edit_main()
End Select

rs.Close

Response.Write vbcrlf & "<tr" & table2 & " height=25><td height=20 class=end background=images/" & web_var(web_config,5) & "/bar_3_bg.gif><a name='pass'></a>&nbsp;" & img_small(us) & "&nbsp;&nbsp;<b>修改我的登陆密码</b></td></tr><tr" & table3 & "><td height=150 align=center>"

Select Case Trim(Request("user_pass"))
    Case "yes"
        Response.Write pass_chk()
    Case Else
        Response.Write pass_main()
End Select

Response.Write vbcrlf & "</td></tr></table><br>"
'---------------------------------center end-------------------------------
Call web_end(0)

Function edit_main()
    Dim seboy
    Dim segirl
    Dim rsface
    Dim rfs
    Dim fff:fff = 0
    edit_main = edit_main & vbcrlf & "<table border=0 width='98%'>" & _
    vbcrlf & "<form name=user_edit_frm action='?' method=post><input type=hidden name=user_edit value='yes'>" & _
    vbcrlf & "<tr><td width='100%' colspan=3 align=center height=30><font class=red><b>注意：</b></font>以下星号（" & redx & "）标出的项目必需填写.</td></tr>" & _
    vbcrlf & "<tr><td width='16%'>您的头衔：</td><td width='84%' colspan=2><input type=text name=nname value='" & code_form(rs("nname")) & "' size=28 maxlength=20></td></tr>"

    If rs("sex") = False Then
        segirl = " checked":seboy = ""
    Else
        seboy  = " checked":segirl = ""
    End If

    edit_main  = edit_main & vbcrlf & "<script language=javascript>function showimage(){ document.images.face_img.src=""images/face/""+document.user_edit_frm.face.options[document.user_edit_frm.face.selectedIndex].value+"".gif""; }</script>" & _
    vbcrlf & "<tr><td width='16%'>你的姓别：</td><td width='45%'> <input type=radio value=true name=sex" & seboy & " class=bg_1>&nbsp;Boy　&nbsp;<input type=radio name=sex value=false" & segirl & " class=bg_1>&nbsp;Girl</td>" & _
    vbcrlf & "<td width='39%' align=center><a href='user_face.asp' target=_blank>→查看所有头像</a>&nbsp;&nbsp;" & _
    vbcrlf & "<select size=1 name=face style='width: 50;' onChange=""showimage()"">"
    rsface        = rs("face")

    For i = 0 To web_var(web_num,11)
        rfs       = ""
        If Int(rsface) = i Then rfs = " selected":fff = 1
        edit_main = edit_main & vbcrlf & "<option value='" & i & "'" & rfs & ">" & i & "</option>"
    Next

    If fff = 0 Then edit_main = edit_main & vbcrlf & "<option value='" & rsface & "' selected>" & rsface & "</option>"
    edit_main = edit_main & vbcrlf & "</select></td></tr>" & _
    vbcrlf & "<tr><td height=30>你的生日：</td><td><select name=b_year>"
    Dim bires
    Dim birse:bires = rs("birthday")
    If Not(IsDate(bires)) Then bires = #1982/6/16#

    For i = 1900 To Year(Now)
        birse     = ""
        If Int(Year(bires)) = Int(i) Then birse = " selected"
        edit_main = edit_main & vbcrlf & "<option value='" & i & "'" & birse & ">" & i & "</option>"
    Next

    edit_main     = edit_main & vbcrlf & "</select>年 <select name=b_month>"

    For i = 1 To 12
        birse     = ""
        If Int(Month(bires)) = Int(i) Then birse = " selected"
        edit_main = edit_main & vbcrlf & "<option value='" & i & "'" & birse & ">" & i & "</option>"
    Next

    edit_main     = edit_main & vbcrlf & "</select>月 <select name=b_day>"

    For i = 1 To 31
        birse     = ""
        If Int(Day(bires)) = Int(i) Then birse = " selected"
        edit_main = edit_main & vbcrlf & "<option value='" & i & "'" & birse & ">" & i & "</option>"
    Next

    edit_main     = edit_main & vbcrlf & "</select>日</td><td rowspan=5 align=center><img border=0 src='images/face/" & rsface & ".gif' name=face_img></td></tr>" & _
    vbcrlf & "<tr><td width='16%'>电子邮件：</td><td width='45%'><input type=text name=email value='" & rs("email") & "' size=28 maxlength=50>" & redx & "</td></tr>" & _
    vbcrlf & "<tr><td>你的QQ：</td><td><input type=text name=qq value='" & rs("qq") & "' size=28 maxlength=15></td></tr>" & _
    vbcrlf & "<tr><td>你的主页：</td><td><input type=text name=url value='" & code_form(rs("url")) & "' size=28 maxlength=100></td></tr>" & _
    vbcrlf & "<tr><td>来自哪里：</td><td><input type=text name=whe value='" & code_form(rs("whe")) & "' size=28 maxlength=20></td></tr>" & _
    vbcrlf & "<tr><td valign=top><br>个人介绍：</td><td colspan=2 valign=top>" & _
    vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=0><tr><td width='69%'>" & _
    vbcrlf & "<textarea rows=7 name=remark cols=42>" & rs("remark") & "</textarea></td>" & _
    vbcrlf & "<td width='31%' valign=top><br>" & redx & "注意：<br><br><br>" & web_var(web_error,3) & _
    vbcrlf & "</td></tr></table>" & _
    vbcrlf & "</td></tr>" & _
    vbcrlf & "<tr><td></td><td colspan=2 height=50>" & _
    vbcrlf & "<input type=submit value=' 更 新 资 料 '>　　　<input type=reset value=' 重 新 修 改 '>" & _
    vbcrlf & "</td></form></tr></table><br>"
End Function

Function edit_chk()
    Dim nname
    Dim sex
    Dim birthday
    Dim face
    Dim email
    Dim qq
    Dim url
    Dim whe
    Dim remark
    Dim founderr
    nname        = code_form(Trim(Request.form("nname")))
    sex          = Trim(Request.form("sex"))
    birthday     = Trim(Request.form("b_year")) & "-" & Trim(Request.form("b_month")) & "-" & Trim(Request.form("b_day"))
    face         = Trim(Request.form("face"))
    email        = code_form(Trim(Request.form("email")))
    qq           = Trim(Request.form("qq"))
    url          = code_form(Trim(Request.form("url")))
    whe          = code_form(Trim(Request.form("whe")))
    remark       = code_form(Request.form("remark"))

    founderr     = ""

    If Not(IsDate(birthday)) Then
        founderr = founderr & err_head & "您选择的 <font class=red_3>生日</font> 不是一个有效的日期格式！<br>"
    End If

    If email_ok(email) <> "yes" Or Len(email) > 50 Then
        founderr = founderr & err_head & "您输入的 <font class=red_3>E-mail</font> 为空或不符合邮件规则！<br>"
    End If

    If qq <> "" And Not IsNull(qq) Then

        If Not(IsNumeric(qq)) Or Len(qq) > 15 Then
            founderr = founderr & err_head & "您输入的 <font class=red_3>QQ</font> 不是数字或长度超过15位！<br>"
        End If

    End If

    If Len(remark) > 250 Then
        founderr = founderr & err_head & "您输入的 <font class=red_3>个人介绍</font> 太多了！不能超过250个字符。<br>"
    End If

    If founderr = "" Then
        rs("nname")     = nname
        rs("sex")     = sex
        rs("birthday")     = birthday
        rs("face")     = face
        rs("email")     = email

        If qq <> "" And Not IsNull(qq) Then
            rs("qq") = qq
        End If

        rs("url")     = url
        rs("whe")     = whe
        rs("remark")     = remark
        rs.update

        edit_chk = "<font class=red>您已成功修改了您的基本资料！</font>" & VbCrLf & "<br><br><a href='user_main.asp'>返回" & tit_fir & "</a>" & vbcrLf & "<br><br>（系统将在 " & web_var(web_num,5) & " 秒钟后自动返回）" & _
        VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=user_main.asp'>"
        Exit Function
    Else
        edit_chk = found_error(founderr,300):Exit Function
    End If

End Function

Function pass_main() %>
<table border=0 width=300 cellspacing=0 cellpadding=2>
<form action='#pass' method=post>
<input type=hidden name=user_pass value='yes'>
<tr height=10><td colspan=2></td></tr>
<tr align=center>
<td width='30%'>登陆密码：</td>
<td width='70%'><input type=password name=password size=25 maxlength=20></td>
</tr>
<tr align=center>
<td>重复密码：</td>
<td><input type=password name=password2 size=25 maxlength=20></td>
</tr>
<tr align=center>
<td>密码钥匙：</td>
<td><input type=text name=passwd size=25 maxlength=20></td>
</tr>
<tr height=30><td colspan=2 align=center><input type=submit value=' 提 交 更 改 '></td></tr>
</form>
</table>
<%
End Function

Function pass_chk()
    Dim password
    Dim password2
    Dim passwd
    Dim founderr
    Dim rs
    Dim sql
    password     = Trim(Request.form("password"))
    password2    = Trim(Request.form("password2"))
    passwd       = Trim(Request.form("passwd"))

    founderr     = ""

    If symbol_ok(password) <> "yes" Then
        founderr = founderr & err_head & "您输入的 <font class=red_3>登陆密码</font> 为空或不符合相关规则！<br>"
    Else

        If password <> password2 Then
            founderr = founderr & err_head & "您输入的 <font class=red_3>登陆密码</font> 和 <font class=founderr>确认密码</font> 不一致！<br>"
        End If

    End If

    If symbol_name(passwd) <> "yes" Then
        founderr = founderr & err_head & "您输入的 <font class=red_3>密码钥匙</font> 为空或不符合相关规则！<br>"
    End If

    If founderr = "" Then
        Set rs = Server.CreateObject("adodb.recordset")
        sql    = "select password,passwd from user_data where username='" & login_username & "' and password='" & login_password & "'"
        rs.open sql,conn,1,3

        If rs.eof And rs.bof Then
            pass_chk = "<font class=red_2>在修改过程中出现在了登陆信息有误的意外！</font><br><br>请查阅 <a href='help.asp?action=register' class=red_3>会员注册注意事项</a> 查看有关事项<br><br>请 <a href='login.asp?action=logout'>重新说锹?/a> 并再次进行修改<br><br>（系统将在 " & web_var(web_num,5) & " 秒钟后自动重登陆）" & _
            VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=logout.asp'>"
            rs.Close:Set rs = Nothing:Exit Function
        Else
            password                                 = jk_md5(password,"short")
            rs("password")                           = password
            rs("passwd")                           = jk_md5(passwd,"short")
            rs.update
            Response.cookies("beyondest_online")("login_password") = password

            If Request.cookies("beyondest_online")("iscookies") = "yes" Then
                Response.cookies("beyondest_online").expires = Date + 365
            End If

            pass_chk                                 = "<font class=red>您已成功修改了您的 登陆密码 和 密码钥匙！</font>" & VbCrLf & "<br><br><a href='user_main.asp'>返回用户中心</a>" & VbCrLf & "<br><br>（系统将在 " & web_var(web_num,5) & " 秒钟后自动返回）" & _
            VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=user_main.asp'>"
            rs.Close:Set rs = Nothing:Exit Function
        End If

        rs.Close:Set rs = Nothing
    Else
        founderr = founderr & err_head & "请查阅有关 <a href='help.asp?action=register' class=red_3>会员注册注意事项</a> 并重新填写。"
        pass_chk = found_error(founderr,280):Exit Function
    End If

End Function %>