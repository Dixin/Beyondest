<!-- #include file="include/onlogin.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/jk_md5.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim udim,unum,id,rssum,thepages,viewpage,page,nummer,pageurl,j,sqladd,admin_user,user_tit,now_username,now_id,now_power,now_hidden,del_temp,checkbox_val,power,hidden,keyword
tit     = "<a href='?'>用户管理</a>　┋"
udim    = Split(user_power,"|"):unum = UBound(udim) + 1

For i = 0 To unum - 1
    tit = tit & "<a href='?power=" & Left(udim(i),InStr(udim(i),":") - 1) & "'>" & Right(udim(i),Len(udim(i)) - InStr(udim(i),":")) & "</a>┋"
Next

Erase udim
tit = tit & "　<a href='?hidden=true'>正常用户</a>┋" & _
"<a href='?hidden=false'>未审核用户</a>"
Response.Write header(1,tit) %>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='98%' cellspacing=0 cellpadding=2 align=center>
<tr><td align=center valign=top height=350>
<%
pageurl      = "?":sqladd = "":user_tit = format_power2(unum,2)
admin_user   = web_var(web_config,3)
id           = Trim(Request.querystring("id"))
power        = format_power(Trim(Request.querystring("power")),0)

If power <> "" And Not IsNull(power) Then
    sqladd   = "where power='" & power & "' "
    pageurl  = pageurl & "power=" & power & "&"
    user_tit = format_power(power,1)
End If

If IsNumeric(id) Then

    Select Case action
        Case "hidden"
            Call user_hidden()
        Case "locked"
            Call useres_popedom(41)
        Case "shield"
            Call useres_popedom(42)
    End Select

End If

Sub user_hidden()
    Dim rs,sql,hid:hid = ""
    sql    = "select username,hidden from user_data where id=" & id
    Set rs = conn.execute(sql)

    If Not (rs.eof And rs.bof) Then

        If rs(0) = web_var(web_config,3) Then Exit Sub

            If rs("hidden") = True Then
                hid = " hidden=0"
            Else
                hid = " hidden=1"
            End If

        End If

        rs.Close:Set rs = Nothing
        If hid <> "" Then conn.execute("update user_data set" & hid & " where username<>'" & admin_user & "' and id=" & id)
    End Sub

    Sub useres_popedom(pn)
        Dim sql,rs,temp1,temp2,temp3,user_popedom
        sql              = "select popedom from user_data where id=" & id
        Set rs           = conn.execute(sql)

        If rs.eof And rs.bof Then rs.Close:Set rs = Nothing:Exit Sub
            user_popedom = rs("popedom")
            rs.Close:Set rs = Nothing

            If Len(user_popedom) <> 50 Then
                user_popedom = "00000000000000000000000000000000000000000000000000"
            Else

                If pn > Len(user_popedom) Then Exit Sub
                    temp1     = Left(user_popedom,pn - 1)
                    temp2     = popedom_format(user_popedom,pn)
                    temp3     = Right(user_popedom,Len(user_popedom) - pn)

                    If Int(temp2) = 0 Then
                        temp2 = 1
                    Else
                        temp2 = 0
                    End If

                End If

                sql = "update user_data set popedom='" & temp1 & temp2 & temp3 & "' where id=" & id
                conn.execute(sql)
            End Sub

            If Trim(Request("del_ok")) = "ok" Then Response.Write del_select()

            Function del_select()
                Dim delid,del_i,del_num,del_dim,del_sql
                delid       = Request("del_id")

                If delid <> "" And Not IsNull(delid) Then
                    delid   = Replace(delid," ","")
                    del_dim = Split(delid,",")
                    del_num = UBound(del_dim)

                    For del_i = 0 To del_num
                        Call delete_userdata(del_dim(del_i))
                        del_sql = "delete from user_data where username='" & del_dim(del_i) & "'"
                        conn.execute(del_sql)
                    Next

                    Erase del_dim
                    del_select = vbcrlf & "<script language=javascript>alert(""共删除了 " & del_num + 1 & " 条记录！"");</script>"
                Else
                    del_select = vbcrlf & "<script language=javascript>alert(""没有删除记录！"");</script>"
                End If

            End Function

            Sub delete_userdata(username)

                If Len(username) < 1 Then Response.Write 1:Exit Sub
                    Dim sql,rs,nn,tnum,dnum:tnum = 0:dnum = 0
                    sql      = "select id,forum_id,re_counter from bbs_topic where username='" & username & "' order by id"
                    Set rs   = conn.execute(sql)

                    Do While Not rs.eof
                        nn   = rs("re_counter") + 1
                        tnum = tnum + 1:dnum = dnum + nn
                        sql  = "delete from bbs_data where reply_id=" & rs("id")
                        conn.execute(sql)
                        sql  = "update bbs_forum set forum_topic_num=forum_topic_num-1,forum_data_num=forum_data_num-" & nn & " where forum_id=" & rs("forum_id")
                        conn.execute(sql)
                        rs.movenext
                    Loop

                    rs.Close
                    sql = "delete from bbs_topic where username='" & username & "'"
                    conn.execute(sql)

                    sql      = "select forum_id,reply_id from bbs_data where username='" & username & "' order by id"
                    Set rs   = conn.execute(sql)

                    Do While Not rs.eof
                        dnum = dnum + 1
                        sql  = "update bbs_topic set re_counter=re_counter-1 where id=" & rs("reply_id")
                        conn.execute(sql)
                        sql  = "update bbs_forum set forum_data_num=forum_data_num-1 where forum_id=" & rs("forum_id")
                        conn.execute(sql)
                        rs.movenext
                    Loop

                    rs.Close:Set rs = Nothing
                    sql = "delete from bbs_data where username='" & username & "'"
                    conn.execute(sql)
                    sql = "update configs set num_topic=num_topic-" & tnum & ",num_data=num_data-" & dnum & " where id=1"
                    conn.execute(sql)
                End Sub

                Select Case action
                    Case "edit"

                        If IsNumeric(id) Then
                            Call user_edit()
                        Else
                            Call user_main()
                        End If

                    Case Else
                        Call user_main()
                End Select

                Call close_conn()
                Response.Write ender()

                Sub user_edit()
                    Dim hidden,h1,h2,password,password2,passwd,passwd2,bbs_counter,counter,integral,emoney,u_popedom
                    Set rs = Server.CreateObject("adodb.recordset")
                    sql    = "select * from user_data where id=" & id
                    rs.open sql,conn,1,3

                    If rs.eof And rs.bof Then rs.Close:Set rs = Nothing:Call user_main():Exit Sub

                        If rs("username") = web_var(web_config,3) Then rs.Close:Set rs = Nothing:Call user_main():Exit Sub
                            u_popedom = rs("popedom")

                            If Trim(Request.querystring("edit")) = "ok" Then
                                Dim temp1,temp2,temp3

                                If Len(u_popedom) <> 50 Then
                                    u_popedom = "00000000000000000000000000000000000000000000000000"
                                Else
                                    temp1     = Left(u_popedom,40)
                                    temp2     = Trim(Request.form("locked")) & Trim(Request.form("shield"))
                                    temp3     = Right(u_popedom,8)
                                    u_popedom = temp1 & temp2 & temp3
                                End If

                                password      = Trim(Request.form("password"))
                                password2     = Trim(Request.form("password2"))
                                passwd        = Trim(Request.form("passwd"))
                                passwd2       = Trim(Request.form("passwd2"))
                                power         = Trim(Request.form("power"))
                                hidden        = Trim(Request.form("hidden"))
                                If password <> password2 Then rs("password") = jk_md5(password,"short")
                                If passwd <> passwd2 Then rs("passwd") = jk_md5(passwd,"short")
                                bbs_counter = Trim(Request.form("bbs_counter"))
                                counter     = Trim(Request.form("counter"))
                                integral    = Trim(Request.form("integral"))
                                emoney      = Trim(Request.form("emoney"))
                                '-2147483648 +2147483647

                                If IsNumeric(bbs_counter) Then
                                    bbs_counter = Int(bbs_counter)
                                    If bbs_counter <> Int(Request.form("bbs_counter2")) And bbs_counter > 0 And bbs_counter <= 2147483647 Then rs("bbs_counter") = bbs_counter
                                End If

                                If IsNumeric(counter) Then
                                    counter = Int(counter)
                                    If counter <> Int(Request.form("counter2")) And counter > 0 And counter <= 2147483647 Then rs("counter") = counter
                                End If

                                If IsNumeric(integral) Then
                                    integral = Int(integral)
                                    If integral <> Int(Request.form("integral2")) And integral > 0 And integral <= 2147483647 Then rs("integral") = integral
                                End If

                                If IsNumeric(emoney) Then
                                    emoney = Int(emoney)
                                    If emoney <> Int(Request.form("emoney2")) And emoney > 0 And emoney <= 2147483647 Then rs("emoney") = emoney
                                End If

                                rs("power") = power
                                rs("hidden") = hidden
                                rs("popedom") = u_popedom
                                rs.update
                                Response.Write "<br><br><br><br><br><br><font class=red>用户信息修改成功！</font><br><br><a href='?power=" & power & "'>点击返回</a>"
                            Else
                                power = rs("power"):hidden = rs("hidden") %>
<table border=0 width=300>
  <form action='?action=edit&edit=ok&power=<% Response.Write power %>&id=<% Response.Write id %>' method=post>
  <tr><td colspan=2 align=center height=50><font class=red>用户管理修改</font></td></tr>
  <tr><td width='30%'>用户名称：</td><td width='70%'><input type=text value='<% Response.Write rs("username") %>' readonly size=25></td></tr>
  <tr><td>用户密码：</td><td><input type=text name=password value='<% Response.Write rs("password") %>' size=25 maxlength=20><input type=hidden name=password2 value='<% Response.Write rs("password") %>'></td></tr>
  <tr><td>密码钥匙：</td><td><input type=text name=passwd value='<% Response.Write rs("passwd") %>' size=25 maxlength=20><input type=hidden name=passwd2 value='<% Response.Write rs("passwd") %>'></td></tr>
  <tr><td>论坛发贴：</td><td><input type=text name=bbs_counter value='<% Response.Write rs("bbs_counter") %>' size=15 maxlength=10></td></tr><input type=hidden name=bbs_counter2 value='<% Response.Write rs("bbs_counter") %>'>
  <tr><td>文栏发贴：</td><td><input type=text name=counter value='<% Response.Write rs("counter") %>' size=15 maxlength=10></td></tr><input type=hidden name=counter2 value='<% Response.Write rs("counter") %>'>
  <tr><td>用户积分：</td><td><input type=text name=integral value='<% Response.Write rs("integral") %>' size=15 maxlength=10></td></tr><input type=hidden name=integral2 value='<% Response.Write rs("integral") %>'>
  <tr><td>用户金钱：</td><td><input type=text name=emoney value='<% Response.Write rs("emoney") %>' size=15 maxlength=10></td></tr><input type=hidden name=emoney2 value='<% Response.Write rs("emoney") %>'>
  <tr><td>用户类型：</td><td><select name=power size=1><%

                                For i = 1 To unum
                                    Response.Write vbcrlf & "<option value='" & format_power2(i,1) & "'"
                                    If power = format_power2(i,1) Then Response.Write " selected"
                                    Response.Write ">" & format_power2(i,2) & "</option>"
                                Next %></select>（<% Response.Write power %>）</td></tr>
  <tr><td>注册审核：</td><td><%

                                If hidden = True Then
                                    h1 = " checked":h2 = ""
                                Else
                                    h1 = "":h2 = " checked"
                                End If %><input type=radio name=hidden value=true<% Response.Write h1 %>>正常<input type=radio name=hidden value=false<% Response.Write h2 %>>未审核</td></tr>
  <tr><td>是否锁定：</td><td><%

                                If Int(popedom_format(u_popedom,41)) = 0 Then
                                    h1 = " checked":h2 = ""
                                Else
                                    h1 = "":h2 = " checked"
                                End If %><input type=radio name=locked value='0'<% Response.Write h1 %>>正常<input type=radio name=locked value='1'<% Response.Write h2 %>>锁定</td></tr>
  <tr><td>论坛屏蔽：</td><td><%

                                If Int(popedom_format(u_popedom,42)) = 0 Then
                                    h1 = " checked":h2 = ""
                                Else
                                    h1 = "":h2 = " checked"
                                End If %><input type=radio name=shield value='0'<% Response.Write h1 %>>正常<input type=radio name=shield value='1'<% Response.Write h2 %>>屏蔽</td></tr>
  <tr><td colspan=2 align=center height=30><input type=submit value=' 提 交 修 改 '></td></tr>
  </form>
</table>
<%
                            End If

                        End Sub

                        Sub user_main()
                            Dim u_popedom
                            hidden       = Trim(Request.querystring("hidden"))

                            If hidden = "true" Then
                                sqladd   = "where hidden=1 "
                                pageurl  = pageurl & "hidden=true&"
                                user_tit = "正常用户"
                            ElseIf hidden = "false" Then
                                sqladd   = "where hidden=0 "
                                pageurl  = pageurl & "hidden=false&"
                                user_tit = "未审核用户"
                            End If

                            keyword      = Trim(Request.querystring("keyword"))

                            If keyword <> "" And Not IsNull(keyword) Then

                                If sqladd <> "" Then
                                    sqladd = sqladd & "and username like '%" & keyword & "%' "
                                Else
                                    sqladd = sqladd & "where username like '%" & keyword & "%' "
                                End If

                                pageurl    = pageurl & "keyword=" & Server.urlencode(keyword) & "&"
                            End If

                            Set rs = Server.CreateObject("adodb.recordset")
                            sql    = "select id,username,tim,power,hidden,popedom from user_data " & sqladd & " order by id desc"
                            rs.open sql,conn,1,1

                            If rs.eof And rs.bof Then
                                Response.Write "<br><br><br><br><br><br><br>还没有" & user_tit

                                rs.Close:Set rs = Nothing:Exit Sub
                                End If

                                rssum    = rs.recordcount
                                nummer   = 15
                                Call format_pagecute()
                                del_temp = nummer %>
<table border=1 width='98%' cellspacing=0 cellpadding=2 align=center bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<form name=del_form action='<% = pageurl %>del_ok=ok' method=post>
  <tr align=center height=30>
  <td colspan=2>现在有 <font class=red><% = rssum %></font> 位<% = user_tit %></td>
  <td colspan=6><% = jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000") %></td>
  </tr>
  <tr align=center bgcolor=#ededed height=20>
    <td width='6%'>序号</td>
    <td width='30%'>会员名称</td>
    <td width='22%'>注册时间</td>
    <td width='12%'>类型</td>
    <td width='8%'>审核</td>
    <td width='7%'>锁定</td>
    <td width='7%'>屏蔽</td>
    <td width='8%'>操作</td>
  </tr>
<%
                                If Int(viewpage) > 1 Then rs.move (viewpage - 1)*nummer

                                For i = 1 To nummer
                                    If rs.eof Then Exit For
                                    checkbox_val = ""
                                    now_id       = rs("id"):now_username = rs("username")
                                    now_power    = rs("power"):now_hidden = rs("hidden"):u_popedom = rs("popedom") %>
  <tr align=center<% Response.Write mtr %>>
    <td align=left><a href='user_view.asp?username=<% Response.Write Server.urlencode(now_username) %>' target=_blank><font color=#000000><% Response.Write i + (viewpage - 1)*nummer %>.</font></a></td>
    <td align=left><a href='?power=<% Response.Write power %>&action=edit&id=<% Response.Write now_id %>'><font class=blue_1><% Response.Write now_username %></font></a></td>
    <td align=left><% Response.Write time_type(rs("tim"),8) %></td>
    <td><%

                                    For j = 1 To unum

                                        If now_power = format_power2(j,1) Then

                                            Select Case j
                                                Case 1
                                                    Response.Write "<font class=red>" & format_power2(j,2) & "</font>"
                                                    checkbox_val = "no":del_temp = del_temp - 1:Exit For
                                                Case 2
                                                    Response.Write "<font class=red_2>" & format_power2(j,2) & "</font>"
                                                    checkbox_val = "no":del_temp = del_temp - 1:Exit For
                                                Case 3
                                                    Response.Write "<font class=red_3>" & format_power2(j,2) & "</font>":Exit For
                                                Case 4
                                                    Response.Write "<font class=blue>" & format_power2(j,2) & "</font>":Exit For
                                                Case Else
                                                    Response.Write format_power2(j,2):Exit For
                                            End Select

                                        End If

                                    Next %></td>
    <td><a href='?power=<% Response.Write power %>&hidden=<% Response.Write hidden %>&action=hidden&id=<% Response.Write now_id %>'><%

                                    If now_hidden = True Then
                                        Response.Write "正常"
                                    Else
                                        Response.Write "<font class=red_2>未审核</font>"
                                    End If %></a></td>
    <td><a href='?power=<% Response.Write power %>&hidden=<% Response.Write hidden %>&action=locked&id=<% Response.Write now_id %>'><%

                                    If Int(popedom_format(u_popedom,41)) = 0 Then
                                        Response.Write "正常"
                                    Else
                                        Response.Write "<font class=red_2>锁定</font>"
                                    End If %></a></td>
    <td><a href='?power=<% Response.Write power %>&hidden=<% Response.Write hidden %>&action=shield&id=<% Response.Write now_id %>'><%

                                    If Int(popedom_format(u_popedom,42)) = 0 Then
                                        Response.Write "正常"
                                    Else
                                        Response.Write "<font class=red_2>屏蔽</font>"
                                    End If %></a></td>
    <td><%

                                    If checkbox_val <> "no" Then
                                        Response.Write "<input type=checkbox name=del_id value='" & now_username & "'>"
                                    Else
                                        Response.Write "&nbsp;"
                                    End If %></td>
  </tr>
<%
                                    rs.movenext
                                Next %>
  <tr align=center height=30>
  <td colspan=2><input type=submit value='删除所选' onclick="return suredel('<% Response.Write del_temp %>');"> &nbsp;<input type=checkbox name=del_all value=1 onClick=selectall('<% Response.Write del_temp %>')>&nbsp;选择所有</td>
</form>
  <td colspan=6>
    <table border=0>
    <form name=sea_frm action='<% Response.Write pageurl %>'>
    <tr>
    <td>关键字：</td>
    <td><input type=text name=keyword value='<% Response.Write keyword %>' size=20 maxlength=20>&nbsp;</td>
    <td>&nbsp;<input type=submit value=' 搜 索 '>&nbsp;</td>
    </tr>
    </form>
    </table>
  </td>
  </tr>
</table>
<%
                                rs.Close:Set rs = Nothing %>
</td></tr></table>
<%
                            End Sub %>