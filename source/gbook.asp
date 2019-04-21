<!-- #include file="include/config_other.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer
Dim rssum
Dim sumrs
Dim page
Dim viewpage
Dim thepages
Dim pageurl
Dim id
Dim nname
Dim sex
Dim qq
Dim email
Dim url
Dim whe
Dim topic
Dim ip
Dim re_word
Dim hidden
pageurl     = "?":nummer = web_var(web_num,4)

index_url   = "gbook"

If action = "write" Then
    tit     = "签写留言"
    tit_fir = format_menu(index_url)
Else
    tit     = format_menu(index_url)
    tit_fir = ""
End If

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Call gbook_left()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center--------------------------------- %>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
    <td align=center><%
Response.Write format_img("rgbook.jpg") & gang
Response.Write ukong

If action = "write" Then
    Call gbook_write()
Else
    Call gbook_main()
End If

Response.Write kong %><td></tr></table><%
'---------------------------------center end-------------------------------
Call web_end(0)

Sub gbook_left()
    Dim temp1
    temp1 = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=2 align=center>" & _
    vbcrlf & "<tr><td height=5></td></tr>" & _
    vbcrlf & "<tr><td height=30 align=center><a href='gbook.asp?action=write'>签写我的留言</a></td></tr>" & _

    vbcrlf & "<tr><td align=left>任何人都可以在这里留言<br>只有注册并登陆后才可以看到其它留言者的系统信息</td></tr>" & _
    vbcrlf & "<tr><td align=left>系统支持：" & Replace(web_var(web_error,3),"<br>","，") & "</td></tr>" & _
    vbcrlf & "</table>"
    Response.Write "<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>" & kong & format_barc("<img src='images/" & web_var(web_config,5) & "/left_gbook.gif' border=0>",temp1,2,0,5) & "</td></tr></table>"
    Response.Write left_action("jt13",4)
End Sub

Sub gbook_main()
    Set rs = Server.CreateObject("adodb.recordset")
    sql    = "select * from gb_data order by id desc"
    rs.open sql,conn,1,1

    If rs.eof And rs.bof Then
        rs.Close:Set rs = Nothing
        Call close_conn()
        Response.redirect "gbook.asp?action=write"
        Response.End
    End If

    rssum = rs.recordcount
    Call format_pagecute()

    Response.Write table1 %>
<tr<% Response.Write table4 %>><td height=25 align=center>
共有<font class=red><% Response.Write rssum %></font>条留言，每页<font class=red><% Response.Write nummer %></font>条，页次<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>&nbsp;
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %>
</td></tr>
</table>
<%
    Response.Write kong

    If Int(viewpage) > 1 Then
        rs.move (viewpage - 1)*nummer
    End If

    For i = 1 To nummer
        If rs.eof Then Exit For
        Response.Write gbook_view()
        rs.Movenext
    Next

    rs.Close:Set rs = Nothing
End Sub

Function gbook_view()
    id      = rs("id")
    nname   = rs("nname")
    sex     = rs("sex")

    If sex = "girl" Then
        sex = "女孩"
    Else
        sex = "男孩"
    End If

    qq      = rs("qq")
    email   = rs("email")
    url     = rs("url")
    whe     = rs("whe")
    topic   = rs("topic")
    ip      = rs("ip")
    re_word = code_jk(rs("re_word"))
    hidden  = rs("hidden")
    Response.Write table1 %>
<tr<% Response.Write table2 %>><td valign=bottom background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<b><font class=end>No.<% Response.Write rssum - (i + (viewpage - 1)*nummer) + 1 %>&nbsp;主题：</font><font class=end title='<% Response.Write code_html(topic,1,0) %>'><% Response.Write code_html(topic,1,36) %></font></b></td></tr>
<tr<% Response.Write table3 %>><td>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
  <tr>
    <td align=center>
      <table border=0 width='100%' cellspacing=2 cellpadding=0 height='100%'>
        <tr align=center valign=top>
          <td width=120>
            <table border=0 width='100%'>
              <tr><td align=center><% = nname %></td></tr>
              <tr><td align=center><img src='images/face/<% = rs("face") %>.gif' title='<% = nname & " 是 " & sex %>' border=0></td></tr>
              <tr><td align=center><% If whe <> "nuller" Then Response.Write "来自：" & code_html(whe,1,0) %></td></tr>
            </table>
          </td>
          <td width=1 bgcolor='<% = web_var(web_color,3) %>'></td>
          <td width=458 height='100%'>
            <table border=0 width='100%' height='100%'>
            <tr><td colspan=2 valign=top>
              <table border=0 width='100%' class=tf cellspacing=4 cellpadding=2><tr><td class=bw>
<%

    If hidden = True And login_mode <> format_power2(1,1) Then
        Response.Write "<br><br><center><font class=red_3>^-^ 此条留言只有站长才可以看哦 ^-^</font></center><br><br>"
    Else
        Response.Write "<img src='images/icon/" & rs("icon") & ".gif' border=0>&nbsp;"

        If hidden = True Then
            Response.Write "<font class=red_3>[隐藏]</font>&nbsp;"
        End If

        Response.Write code_jk(rs("word"))
    End If %></td></tr></table>

</td></tr>
<tr><td height=5></td></tr>
<tr><td height=1 colspan=2 bgcolor=<% Response.Write web_var(web_color,2) %>></td></tr>
<tr height=20>
<td width='60%'>&nbsp;<img src='IMAGES/SMALL/TIM.GIF' align=absmiddle title='签写时间' border=0>　<% Response.Write rs("tim") %></td>
<td width='40%' align=right><%

    If qq <> "nuller" Then
        Response.Write "<a href='http://search.tencent.com/cgi-bin/friend/user_show_info?ln=" & qq & "' target=_blank><img src='images/small/qq.gif' title='" & nname & " 的QQ是：" & qq & "' border=0></a>&nbsp;"
    End If

    If var_null(url) <> "" And url <> "nuller" And url <> "http://" Then
        Response.Write "<a href='" & url & "' target=_blank><img src='images/small/url.gif' title='访问 " & nname & " 的主页' border=0></a>&nbsp;"
    End If

    If email <> "nuller" Then
        Response.Write "<a href='mailto:" & email & "'><img src='images/small/email.gif' title='给 " & nname & " 发电子邮件' border=0></a>&nbsp;"
    End If

    If login_username <> "" And login_password <> "" And login_mode <> "" Then
        Response.Write ip_types(ip,nname,1) & "&nbsp;" & _
        "<img src='images/small/sys.gif' align=absMiddle title='" & view_sys(rs("sys")) & "' border=0>"

        If login_mode = "admin" Then
            Response.Write "&nbsp;<a href='gbook_action.asp?action=reply&id=" & rs("id") & "&page=" & viewpage & "'><img src='images/small/reply.gif' alt='回复这条留言' border=0></a>&nbsp;" & _
            "<a href='gbook_action.asp?action=delete&id=" & rs("id") & "&page=" & viewpage & "'><img src='images/small/del.gif' alt='删除这条留言' border=0></a>"
        End If

    End If %></td>
</tr>
<%

    If Len(re_word) > 0 Then
        Response.Write vbcrlf & "<tr><td colspan=2>" & table1 & "<tr" & table4 & "><td class=bw bgcolor=" & web_var(web_color,6) & ">" & _
        vbcrlf & "<font class=red>站长回复：</font>&nbsp;&nbsp;&nbsp;&nbsp;（时间：" & rs("re_tim") & "）<br>" & _
        vbcrlf & "<img src='images/icon/" & rs("re_icon") & ".gif' border=0>&nbsp;" & re_word & _
        vbcrlf & "</td></tr></table></td></tr>"
    End If %>
</table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
 </table>
</td></tr>
</table>
<%
    Response.Write kong
End Function

Sub gbook_write()
    Response.Write vbcrlf & "<table border=0 width='100%'><tr><td align=center height=300>"

    If post_chk() = "no" Then
        Response.Write web_var(web_error,1) & "<br><br><br>" & "<a href='gbook.asp?action=write'>点击此处返回留言处</a>"
    ElseIf Request.form("gb_write") = "ok" Then
        Response.Write write_chk()
    Else
        Response.Write write_type()
    End If

    Response.Write vbcrlf & "</td></tr></table>"
End Sub

Function write_type()
    write_type = vbcrlf & "<script language=javascript src='style/form_gbook.js'></script><script language=javascript src='style/form_ok.js'></script>" & _
    vbcrlf & "<table border=0 cellpadding=0 cellspacing=0 width=500>" & _
    vbcrlf & "<tr><form name=write_frm method=POST action='gbook.asp?action=write' onsubmit=""frm_submitonce(this);"">" & _
    vbcrlf & "<input type=hidden name=gb_write value='ok'>" & _
    vbcrlf & "<td align=center height=50 colspan=4><font class=red><b>注意：</b></font>以下星号（" & redx & "）标出的项目必需填写</td>" & _
    vbcrlf & "</tr><tr><td align=center width=80 height=25>你的名字：</td>" & _
    vbcrlf & "<td width=224><input type=text name=wrname value='" & login_username & "' size=27 maxlength=9>" & redx & "</td>" & _
    vbcrlf & "<td width=100 align=center>&nbsp;<a title='查看所有头像' href='user_face.asp' target=_blank>→查看所有头像</a></td>" & _
    vbcrlf & "<td width=71 align=center>" & _
    vbcrlf & "<select size=1 name=wrface style=""width： 50; border： 1px solid #C0C0C0"" onChange=""showimage()"">" & _
    vbcrlf & "<option value=0 selected>0</option>"

    For i = 1 To web_var(web_num,11)
        write_type = write_type & "<option value='" & i & "'>" & i & "</option>"
    Next

    write_type     = write_type & vbcrlf & "</select></td></tr>" & _
    vbcrlf & "<tr><td align=center width=80 height=25>姓别：</td>" & _
    vbcrlf & "<td width=224>&nbsp;Boy <input type=radio value='boy' name=wrsex checked class=bg_1>&nbsp;&nbsp; Girl <input type=radio name=wrsex value='girl' class=bg_1></td>" & _
    vbcrlf & "<td width=196 rowspan=5 align=center colspan=2><img border=0 src='images/face/0.gif' name=wrimg></td></tr>" & _
    vbcrlf & "<tr><td align=center height=25>QQ：</td>" & _
    vbcrlf & "<td width=224><input type=text name=wrqq size=28 maxlength=15></td></tr>" & _
    vbcrlf & "<tr><td align=center height=25>电子邮件：</td>" & _
    vbcrlf & "<td width=224><input type=text name=wremail size=28 maxlength=50></td></tr>" & _
    vbcrlf & "<tr><td align=center height=25>你的主页： </td>" & _
    vbcrlf & "<td width=224><input type=text name=wrurl size=28 value='http://' maxlength=50></td></tr>" & _
    vbcrlf & "<tr><td align=center height=25>来自：</td>" & _
    vbcrlf & "<td width=224><input type=text name=wrwhe size=28 maxlength=20></td></tr>" & _
    vbcrlf & "<tr><td align=center height=25>留言主题：</td>" & _
    vbcrlf & "<td width=420 colspan=3><input type=text name=wrtopic size=38 maxlength=50>" & redx & "</td></tr>" & _
    vbcrlf & "<tr><td align=center height=25>表情图标： </td>" & _
    vbcrlf & "<td align=left width=420 colspan=3>" & icon_type(7,1) & "</td></tr>" & _
    vbcrlf & "<tr><td align=center width=80 valign=top><br>你的留言：<br><br></td>" & _
    vbcrlf & "<td width=420 colspan=3>" & _
    vbcrlf & "<table border=0 cellpadding=0 cellspacing=0 width='100%'>" & _
    vbcrlf & "<tr><td width='69%'><textarea rows=7 name=wrword cols=60 maxlength=1000 title='按 Ctrl+Enter 可直接发送' onkeydown=""javascript:frm_quicksubmit();""></textarea></td></tr></table>" & _
    vbcrlf & "</td></tr><tr>" & _
    vbcrlf & "<td align=center width=80 height=25>是否隐藏：</td>" & _
    vbcrlf & "<td width=420 colspan=3><input type=radio name=wrhidden value='no' checked class=bg_1>正常<input type=radio name=wrhidden value='yes' class=bg_1>隐藏" & redx & "选择隐藏后，此留言只有站长才可以看到。</td></tr>" & _
    vbcrlf & "<tr height=50><td></td><td colspan=3>" & _
    vbcrlf & "<input type=submit name=wsubmit value=' 可 以 发 送 啦 ' onclick=""return check(write_frm)"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
    vbcrlf & "<input type=reset value=' 重 新 再 来 ' onclick=""return reset()"">" & _
    vbcrlf & "</td></form></tr></table><br>"
End Function

Function write_chk()
    Call time_load(1,0,1)
    Dim founderr
    Dim drname
    Dim drsex
    Dim drqq
    Dim dremail
    Dim drurl
    Dim drwhe
    Dim drtopic
    Dim dricon
    Dim drface
    Dim drword
    Dim drremark
    Dim drhidden
    drname       = Trim(Request.form("wrname"))
    drname       = code_form(drname)'---------------
    drsex        = Trim(Request.form("wrsex"))
    drqq         = Trim(Request.form("wrqq"))
    If drqq = "" Or IsNull(drqq) Then drqq = "nuller"
    dremail      = Trim(Request.form("wremail"))
    If dremail = "" Or IsNull(dremail) Then dremail = "nuller"
    drurl        = Trim(Request.form("wrurl"))
    If drurl = "http://" Or IsNull(drurl) Then drurl = "nuller"
    drwhe        = Trim(Request.form("wrwhe"))
    drwhe        = code_form(drwhe)'---------------
    If drwhe = "" Or IsNull(drwhe) Then drwhe = "nuller"
    dricon       = Trim(Request.form("icon"))
    drface       = Trim(Request.form("wrface"))
    drtopic      = Trim(Request.form("wrtopic"))
    drtopic      = code_form(drtopic)'---------------
    drword       = Request.form("wrword")
    drremark     = Request.form("wrremark")
    drhidden     = Trim(Request.form("wrhidden"))

    founderr     = ""

    If symbol_name(drname) = "no" Then
        founderr = founderr & "<br><li>请输入您的 <font class=founderr>用户名</font>（长度不能大于20）！"
    End If

    If drqq <> "nuller" Then

        If Not(IsNumeric(drqq)) Then
            founderr = founderr & "<br><li>您的 <font class=founderr>QQ</font>> 只能为数字！"
        End If

    End If

    If dremail <> "nuller" Then

        If email_ok(dremail) = False Then
            founderr = founderr & "<br><li>您输入的 <font class=founderr>Email</font> 格式有错误！"
        End If

    End If

    If drtopic = "" Or IsNull(drtopic) Then
        founderr = founderr & "<br><li><font class=founderr>主题</font> 是必需要的！请输入。"
    End If

    If drword = "" Or IsNull(drword) Then
        founderr = founderr & "<br><li><font class=founderr>留言内容</font> 是必需要的！请输入。"
    End If

    If founderr = "" Then
        Dim rs
        Dim strsql
        Set rs = Server.CreateObject("adodb.recordset")
        strsql = "select * from gb_data where (id is null)"
        rs.open strsql,conn,1,3
        rs.addnew
        rs("nname") = drname
        rs("sex") = drsex
        rs("whe") = drwhe
        rs("qq") = drqq
        rs("email") = dremail
        rs("url") = drurl
        rs("ip") = ip_sys(1,1)
        rs("sys") = ip_sys(3,0)
        rs("icon") = dricon
        rs("face") = drface
        rs("topic") = drtopic
        rs("word") = drword
        rs("tim") = now_time
        rs("re_icon") = "0"

        If drhidden = "yes" Then
            rs("hidden") = True
        Else
            rs("hidden") = False
        End If

        rs.update
        rs.Close:Set rs = Nothing
        Call time_load(0,0,1)
        Response.Write VbCrLf & "<br>！！！<font class=red>谢谢你的留言</font>！！！" & VbCrLf & "<br><br><a href='gbook.asp'>返回留言</a>" & VbCrLf & "<br><br>（系统将在 " & web_var(web_num,5) & " 秒钟后自动返回）" & _
        VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=gbook.asp'>"
    Else
        Response.Write found_error(founderr,"250")
    End If

End Function %>