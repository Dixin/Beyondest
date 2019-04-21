<!-- #include file="include/config_other.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim wt:wt = Int(Mid(web_setup
Dim 3
Dim 1))
tit = "与我在线"

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Response.Write left_action("jt13",4)
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong %>
<table border=0 width='98%'>
<tr><td align=center height=30><% Call online_main() %></td></tr>
<tr><td align=center height=30><% Response.Write user_power_type(0) %></td></tr>
<tr><td align=center class=htd><%

Select Case wt
    Case 1
        Response.Write "<font class=red>本站现在已开启 <font class=blue>" & web_var(web_stamp,wt + 1) & "</font> 模式！所有登陆和浏览本站的人被并记录在线列表。</font>"
    Case 2
        Response.Write "<font class=red>本站现在已开启 <font class=blue>" & web_var(web_stamp,wt + 1) & "</font> 模式！本站的注册用户可以登陆，并记录用户在线列表。</font>"
    Case Else
        Response.Write "<font class=red>本站现在已开启 <font class=blue>" & web_var(web_stamp,wt + 1) & "</font> 模式！本站的注册用户可以登陆，不记录在线列表。</font>"
End Select

Response.Write "<br>有关 <a href='help.asp?action=web'>网站模式</a> 的详细说明，请进入 <a href='help.asp?action=web'>网站帮助</a> 查看相关信息。" %></td></tr>
<tr><td align=center height=5></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
Call web_end(0)

Sub online_main()
    Dim rssum
    Dim l_username

    If var_null(login_username) = "" Then
        Response.Write "<font class=blue>"

        If wt = 1 Then
            Response.Write Request.cookies("beyondest_online")("guest_name")
        Else
            Response.Write "游客"
        End If

        Response.Write "</font>，您好！" & web_var(web_error,2)
    Else
        Response.Write "<font class=blue>" & login_username & "</font>，您好！欢迎您注册并登陆本站！您现在可以点击浏览其它栏目的详细内容。"
    End If

    If wt = 0 Then
        Response.Write "<tr><td></td></tr><tr><td height=200>"

        Exit Sub
        End If

        Set rs = Server.CreateObject("adodb.recordset")

        If wt = 1 Or wt = 2 Then
            sql = "select user_login.*,user_data.power from user_data inner join user_login on user_login.l_username=user_data.username where user_login.l_type=0 order by user_login.l_id"
            rs.open sql,conn,1,1

            If rs.eof And rs.bof Then
                rssum = 0
            Else
                rssum = rs.recordcount
            End If %>
</td></tr>
<tr><td height=5></td></tr>
<tr><td><% Response.Write img_small("jt1") %><font class=red_3><b>在线会员</b></font>&nbsp;（<font class=red><% Response.Write rssum %></font>&nbsp;人）</td></tr>
<tr><td align=center height=150 valign=top>
  <table border=0 width='100%'>
  <tr><td width='20%'></td><td width='20%'></td><td width='20%'></td><td width='20%'></td><td width='20%'></td></tr>
<%

            Do While Not rs.eof
                Response.Write "<tr>"

                For i = 1 To 5
                    If rs.eof Then Exit For
                    l_username = rs("l_username")
                    Response.Write "<td>" & img_small("icon_" & rs("power")) & "<a href='user_view.asp?username=" & Server.urlencode(l_username) & "' title='目前位置：" & rs("l_where") & "<br>来访时间：" & rs("l_tim_login") & "<br>活动时间：" & rs("l_tim_end") & "<br>真实IP：" & ip_types(rs("l_ip"),l_username,0) & "<br>" & view_sys(rs("l_sys")) & "' target=_blank>" & l_username & "</a></td>"
                    rs.movenext
                Next

                Response.Write "</tr>"
            Loop

            rs.Close
            Response.Write "</table>"
        End If

        If wt = 1 Then
            sql = "select * from user_login where l_type=1 order by l_id"
            rs.open sql,conn,1,1

            If rs.eof And rs.bof Then
                rssum = 0
            Else
                rssum = rs.recordcount
            End If %>
</td></tr>
<tr><td><% Response.Write img_small("jt12") %><font class=red_3><b>在线游客</b></font>&nbsp;（<font class=red><% Response.Write rssum %></font>&nbsp;人）</td></tr>
<tr><td align=center height=150 valign=top>
  <table border=0 width='98%'>
  <tr><td width='25%'></td><td width='25%'></td><td width='25%'></td><td width='25%'></td></tr>
<%

            Do While Not rs.eof
                Response.Write "<tr>"

                For i = 1 To 4
                    If rs.eof Then Exit For
                    l_username = rs("l_username")
                    Response.Write "<td>" & img_small("icon_other") & "<font title='目前位置：" & rs("l_where") & "<br>来访时间：" & rs("l_tim_login") & "<br>活动时间：" & rs("l_tim_end") & "<br>" & ip_types(rs("l_ip"),l_username,0) & "<br>" & view_sys(rs("l_sys")) & "' target=_blank>" & l_username & "</font></td>"
                    rs.movenext
                Next

                Response.Write "</tr>"
            Loop

            rs.Close
            Response.Write "</table>"
        End If

        Set rs = Nothing
    End Sub %>