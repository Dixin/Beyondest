<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim admin_menu,udim,unum
admin_menu     = "<a href='admin_user_list.asp'>用户管理</a>　　┋"
udim           = Split(user_power,"|"):unum = UBound(udim) + 1

For i = 0 To unum - 1
    admin_menu = admin_menu & "<a href='admin_user_list.asp?power=" & Left(udim(i),InStr(udim(i),":") - 1) & "'>" & Right(udim(i),Len(udim(i)) - InStr(udim(i),":")) & "</a>┋"
Next

admin_menu     = admin_menu & "　　<a href='admin_user_list.asp?hidden=true'>正常用户</a>┋" & _
"<a href='admin_user_list.asp?hidden=false'>锁定用户</a>"

Response.Write header(1,admin_menu) %>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='98%' cellspacing=0 cellpadding=2 align=center>
<tr><td align=center valign=top height=350>
<%
Dim id,rssum,thepages,viewpage,page,nummer,pageurl,sqladd,user_tit,now_username,now_id,now_power,now_hidden,del_temp,checkbox_val,power,hidden,keyword
pageurl      = "admin_user_list.asp?"
sqladd       = ""
user_tit     = format_power2(unum,2)
power        = format_power(Trim(Request.querystring("power")),0)

If power <> "" And Not IsNull(power) Then
    sqladd   = "where power='" & power & "' "
    pageurl  = pageurl & "power=" & power & "&"
    user_tit = format_power(power,1)
End If

id           = Trim(Request.querystring("id"))
hidden       = Trim(Request.querystring("hidden"))

If hidden = "true" Then
    sqladd   = "where hidden=1 "
    pageurl  = pageurl & "hidden=true&"
    user_tit = "正常用户"
ElseIf hidden = "false" Then
    sqladd   = "where hidden=0 "
    pageurl  = pageurl & "hidden=false&"
    user_tit = "锁定用户"
End If

keyword = Trim(Request.querystring("keyword"))

If keyword <> "" And Not IsNull(keyword) Then

    If sqladd <> "" Then
        sqladd = sqladd & "and username like '%" & keyword & "%' "
    Else
        sqladd = sqladd & "where username like '%" & keyword & "%' "
    End If

    pageurl    = pageurl & "keyword=" & Server.urlencode(keyword) & "&"
End If

If action = "hidden" And IsNumeric(id) Then Call user_hidden()

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

        If hid <> "" Then conn.execute("update user_data set" & hid & " where id=" & id)
    End Sub

    If Trim(Request("del_ok")) = "ok" Then Response.Write del_select()

    Function del_select()
        Dim delid,del_i,del_num,del_dim,del_sql
        delid           = Request("del_id")

        If delid <> "" And Not IsNull(delid) Then
            delid       = Replace(delid," ","")
            del_dim     = Split(delid,",")
            del_num     = UBound(del_dim)

            For del_i = 0 To del_num
                del_sql = "delete from user_data where id=" & del_dim(del_i)
                conn.execute(del_sql)
            Next

            Erase del_dim
            del_select = vbcrlf & "<script language=javascript>alert(""共删除了 " & del_num + 1 & " 条记录！"");</script>"
        Else
            del_select = vbcrlf & "<script language=javascript>alert(""没有删除记录！"");</script>"
        End If

    End Function

    Set rs = Server.CreateObject("adodb.recordset")
    sql    = "select id,username,tim,power,hidden from user_data " & sqladd & " order by id desc"
    rs.open sql,conn,1,1

    If rs.eof And rs.bof Then
        Response.Write "<br><br><br><br><br><br><br>还没有" & user_tit
    Else
        rssum    = rs.recordcount
        nummer   = 15
        Call format_pagecute()
        del_temp = nummer %>
<table border=1 width='98%' cellspacing=0 cellpadding=0 align=center bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<form name=del_form action='<% = pageurl %>del_ok=ok' method=post>
  <tr align=center height=30>
  <td colspan=2>现在有 <font class=red><% = rssum %></font> 位<% = user_tit %></td>
  <td colspan=4><% = jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000") %></td>
  </tr>
  <tr align=center bgcolor=#ededed height=25>
    <td width='6%'>序号</td>
    <td width='30%'>会员名称</td>
    <td width='28%'>注册时间</td>
    <td width='14%'>类型</td>
    <td width='12%'>状态</td>
    <td width='10%'>操作</td>
  </tr>
<%

        If Int(viewpage) > 1 Then
            rs.move (viewpage - 1)*nummer
        End If

        For i = 1 To nummer
            If rs.eof Then Exit For
            checkbox_val = ""
            now_id       = rs("id")
            now_username = rs("username")
            now_power    = rs("power")
            now_hidden   = rs("hidden") %>
  <tr align=center>
    <td align=left><a href='user_view.asp?username=<% Response.Write Server.urlencode(now_username) %>' target=_blank><font color=#000000><% = i + (viewpage - 1)*nummer %>.</font></a></td>
    <td><a href='admin_user_edit.asp?id=<% = now_id %>'><font class=blue_1><% = now_username %></font></a></td>
    <td><% = rs("tim") %></td>
    <td><%

            Select Case now_power
                Case format_power2(1,1)
                    Response.Write "<font class=red>" & format_power2(1,2) & "</font>"
                    checkbox_val = "no"
                    del_temp     = del_temp - 1
                Case format_power2(2,1)
                    Response.Write "<font class=red_2>" & format_power2(2,2) & "</font>"
                    checkbox_val = "no"
                    del_temp     = del_temp - 1
                Case format_power2(3,1)
                    Response.Write "<font class=red_3>" & format_power2(3,2) & "</font>"
                Case format_power2(4,1)
                    Response.Write "<font class=red_4>" & format_power2(4,2) & "</font>"
                Case Else
                    Response.Write format_power2(5,2)
            End Select %></td>
    <td><a href='admin_user_list.asp?power=<% Response.Write power %>&hidden=<% Response.Write hidden %>&action=hidden&id=<% Response.Write now_id %>'><%

            Select Case now_hidden
                Case True
                    Response.Write "正常"
                Case Else
                    Response.Write "<font class=red_2>锁定</font>"
            End Select %></a></td>
    <td><%

            If checkbox_val <> "no" Then
                Response.Write "<input type=checkbox name=del_id value='" & now_id & "'>"
            Else
                Response.Write "&nbsp;"
            End If %></td>
  </tr>
<%
            rs.movenext
        Next %>
  <tr align=center height=30>
  <td colspan=2><input type=submit value='删除所选' onclick="return suredel('<% = del_temp %>');"> &nbsp;<input type=checkbox name=del_all value=1 onClick=selectall('<% = del_temp %>')>&nbsp;选择所有</td>
</form>
  <td colspan=4>
<table border=0>
<form name=sea_frm action='<% = pageurl %>'>
<tr>
<td>关键字：</td>
<td><input type=text name=keyword value='<% = keyword %>' size=20 maxlength=20>&nbsp;</td>
<td>&nbsp;<input type=submit value=' 搜 索 '>&nbsp;</td>
</tr>
</form>
</table>
  </td>
  </tr>
</table>
<%
    End If

    rs.Close:Set rs = Nothing %>
</td></tr></table>
<%
    Call close_conn()
    Response.Write ender() %>