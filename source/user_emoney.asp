<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim integral,unit_num,emoney_1,chk,errs
tit = "虚拟货币"

Call web_head(2,0,0,0,0)
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong
Call emoney_top()

Call emoney_main()

Response.Write ukong
'---------------------------------center end-------------------------------
Call web_end(0)

Sub emoney_main()
    Dim emoneys,emoney_2,e_num,e_all,c_name,c_pass,c_emoney,c_id,userp
    unit_num = Int(web_var(web_num,14)):errs = "":emoney_2 = 0:c_id = 0
    Set rs   = conn.execute("select integral from user_data where hidden=1 and username='" & login_username & "'")
    integral = rs("integral")
    rs.Close:Set rs = Nothing
    emoney_1 = integral\unit_num:userp = format_power(login_mode,2)
    If Not(IsNumeric(userp)) Then userp = 0
    userp    = Int(userp)
    chk      = Trim(Request.querystring("chk"))
    If action <> "virement" And action <> "card" Then action = "converion"

    If (action = "converion" Or action = "virement") And chk = "yes" Then
        e_num   = Trim(Request.form("e_num")):e_all = Trim(Request.form("e_all"))
        emoneys = emoney_1
        If action = "virement" Then emoneys = login_emoney

        If e_all = "yes" Then
            emoney_2 = emoneys
        Else

            If Not(IsNumeric(e_num)) Then
                errs = "no"
            Else

                If InStr(1,e_num,".") > 0 Then
                    errs = "no"
                Else

                    If Int(e_num) < 1 Or Int(e_num) > Int(emoneys) Then
                        errs     = "no"
                    Else
                        emoney_2 = e_num
                    End If

                End If

            End If

        End If

        If action = "converion" And Int(emoney_2) > 0 Then
            conn.execute("update user_data set integral=integral-" & emoney_2*unit_num & ",emoney=emoney+" & emoney_2 & " where username='" & login_username & "'")
            integral = integral - emoney_2*unit_num:login_emoney = login_emoney + emoney_2:emoney_1 = emoney_1 - emoney_2
            Response.Write "<script language=javascript>alert(""您已成功换算了 " & emoney_2 & " " & m_unit & "！\n\n您的积分消耗了：" & emoney_2*unit_num & " 分\n\n目前的积分换算率为：每 " & unit_num & " 分可换算 1 " & m_unit & """);</script>"
        End If

        If action = "virement" And Int(emoney_2) > 0 Then
            Dim username2:username2 = Trim(Request.form("username2"))

            If symbol_name(username2) <> "yes" Then
                errs   = "no"
            Else
                Set rs = conn.execute("select username from user_data where username='" & username2 & "'")
                If rs.eof And rs.bof Then errs = "no"
                rs.Close:Set rs = Nothing
            End If

            If errs = "" Then
                conn.execute("update user_data set emoney=emoney-" & emoney_2 & " where username='" & login_username & "'")
                conn.execute("update user_data set emoney=emoney+" & emoney_2 & " where username='" & username2 & "'")
                login_emoney = login_emoney - emoney_2
                Response.Write "<script language=javascript>alert(""您已成功的给 " & username2 & " 转帐了 " & emoney_2 & " " & m_unit & "！\n\n您的拥有的" & tit & "也减少了：" & emoney_2 & " " & m_unit & """);</script>"
                sql          = "insert into user_mail(send_u,accept_u,topic,word,tim,types,isread) " & _
                "values('" & login_username & "','" & username2 & "','[系统]货币转帐信息提示','" & login_username & " 已成功的给 您 转帐了 " & emoney_2 & " " & m_unit & "！','" & now_time & "',1,0)"
                conn.execute(sql)
            End If

        End If

    End If

    If action = "card" And chk = "yes" Then
        c_name = code_form(Trim(Request.form("c_name")))
        c_pass = code_form(Trim(Request.form("c_pass")))
        If Len(c_name) < 1 Or Len(c_pass) < 1 Then errs = "no"

        If errs = "" Then
            sql      = "select c_id,c_emoney from cards where c_name='" & c_name & "' and c_pass='" & c_pass & "' and c_hidden=0"
            Set rs   = conn.execute(sql)

            If rs.eof And rs.bof Then
                errs = "no"
            Else
                c_id = rs("c_id"):c_emoney = rs("c_emoney")
            End If

            rs.Close:Set rs = Nothing
        End If

        If errs = "" Then
            Dim ok_msg:ok_msg = ""
            conn.execute("update cards set c_hidden=1 where c_id=" & c_id)
            sql          = "update user_data set emoney=emoney+" & c_emoney
            If Int(userp) > 3 Then sql = sql & ",power='" & format_power2(3,1) & "'":ok_msg = "\n\n您也同时升级为 VIP 会员！"
            sql          = sql & " where username='" & login_username & "'"
            conn.execute(sql)
            login_emoney = login_emoney + c_emoney
            Response.Write "<script language=javascript>alert(""您已成功的用会员卡（卡号：" & c_name & "）给您充值了 " & c_emoney & " " & m_unit & "！" & ok_msg & """);</script>"
        End If

    End If

    Select Case action
        Case "virement"
            Call emoney_virement()
            Call emoney_card()
            Call emoney_converion()
        Case "card"
            Call emoney_card()
            Call emoney_converion()
            Call emoney_virement()
        Case Else
            Call emoney_converion()
            Call emoney_virement()
            Call emoney_card()
    End Select

    Response.Write ukong & table1 %>
<tr<% Response.Write table2 %>><td>&nbsp;<% Response.Write img_small("fk00") %>&nbsp;<font class=end><b>相关说明</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25><font class=red>注意：</font></td><td>您输入的换算的<% Response.Write m_unit %>数值不能超过您目前可以换算的最大值（<font class=red><% Response.Write emoney_1 & "</font>&nbsp;" & m_unit %>）</td></tr>
  <tr><td height=25></td><td>您输入的要转帐的<% Response.Write m_unit %>数值不能超过您目前拥有的最大值（<font class=red><% Response.Write login_emoney & "</font>&nbsp;" & m_unit %>）</td></tr>
  <tr><td height=25></td><td>您在这里进行的<font class=blue>积分换算</font>和<font class=blue>货币转帐</font>为<font class=red>不可逆操作</font>！请在操作前注意一下。</td></tr>
  </table>
</td></tr>
</table><%
    Response.Write ukong
End Sub

Sub emoney_converion()
    Response.Write ukong & table1 %>
<tr<% Response.Write table2 %>><td>&nbsp;<% Response.Write img_small("fk00") %>&nbsp;<font class=end><b>积分换算</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>您目前拥有的<% Response.Write tit %>为：<font class=red><% Response.Write login_emoney & "</font>&nbsp;" & m_unit %></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>目前的积分换算率为：每&nbsp;<font class=red_3><b><% Response.Write unit_num %></b></font>&nbsp;分可换算&nbsp;<font class=red><b>1</b></font>&nbsp;<% Response.Write m_unit %></td></tr>
  <tr><td height=25>您目前的社区积分为：<font class=red_3><% Response.Write integral %></font>&nbsp;分</td></tr>
  <tr><td height=25>您目前可以换算：<font class=red><% Response.Write emoney_1 & "</font>&nbsp;" & m_unit %></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
<% If action = "converion" And chk = "yes" And errs <> "" Then %>
  <tr><td height=50><font class=red_2>换算失败：</font>请输入一个不大于 <font class=red><% Response.Write emoney_1 %></font> 的正整数！
&nbsp;&nbsp;&nbsp;&nbsp;<% Response.Write go_back %></td></tr>
<% Else %>
  <form name=emoney_frm_1 action='?action=converion&chk=yes' method=post>
  <tr><td height=50>请输入您要换算的<% Response.Write m_unit %>数值：&nbsp;
<input type=text name=e_num size=12 maxlength=10 value=''>&nbsp;&nbsp;&nbsp;
<input type=checkbox name=e_all value='yes'>&nbsp;全部换算&nbsp;&nbsp;&nbsp;
<input type=submit value='进行换算'></td></tr>
  </form>
<% End If %>
  </table>
</td></tr>
</table><%

End Sub

Sub emoney_virement()
Response.Write ukong & table1 %>
<tr<% Response.Write table2 %>><td>&nbsp;<% Response.Write img_small("fk00") %>&nbsp;<font class=end><b>货币转帐</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>您目前拥有的<% Response.Write tit %>为：<font class=red><% Response.Write login_emoney & "</font>&nbsp;" & m_unit %></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
<% If action = "virement" And chk = "yes" And errs <> "" Then %>
  <tr><td height=50><font class=red_2>转帐失败：</font></td><td>请输入一个不大于 <font class=red><% Response.Write emoney_1 %></font> 的正整数&nbsp;或&nbsp;您要转入的注册用户不存在！&nbsp;&nbsp;<% Response.Write go_back %></td></tr>
<% Else %>
  <form name=emoney_frm_2 action='?action=virement&chk=yes' method=post>
  <tr><td height=10></td></tr>
  <tr><td height=30>请输入您要转帐的注册用户：&nbsp;
<input type=text name=username2 size=15 maxlength=20 value=''>&nbsp;&nbsp;&nbsp;
<% Response.Write friend_select() %>
</td></tr>
  <tr><td height=30>请输入您要转帐的<% Response.Write m_unit %>数值：&nbsp;
<input type=text name=e_num size=12 maxlength=10 value=''>&nbsp;&nbsp;&nbsp;
<input type=checkbox name=eall value='yes'>&nbsp;全部转帐&nbsp;&nbsp;&nbsp;
<input type=submit value='进行转帐'></td></tr>
  <tr><td height=10></td></tr>
  </form>
<% End If %>
  </table>
</td></tr>
</table><%

End Sub

Sub emoney_card()
Response.Write ukong & table1 %>
<tr<% Response.Write table2 %>><td>&nbsp;<% Response.Write img_small("fk00") %>&nbsp;<font class=end><b>会员卡充值</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>您目前拥有的<% Response.Write tit %>为：<font class=red><% Response.Write login_emoney & "</font>&nbsp;" & m_unit %></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
<% If action = "card" And chk = "yes" And errs <> "" Then %>
  <tr><td height=50><font class=red_2>会员卡充值失败：</font></td><td>您输入的会员 <font class=red>卡号</font> 或 <font class=red>密码</font> 有错误！&nbsp;&nbsp;<% Response.Write go_back %></td></tr>
<% Else %>
  <form name=emoney_frm_3 action='?action=card&chk=yes' method=post>
  <tr><td height=50>
    <table border=0>
    <tr>
    <td>卡号：&nbsp;<input type=text name=c_name size=15 maxlength=20></td>
    <td>&nbsp;&nbsp;密码：&nbsp;<input type=password name=c_pass size=15 maxlength=20></td>
    <td>&nbsp;&nbsp;<input type=submit value='会员卡充值'></td>
    </tr>
    </table>
  </td><tr>
  </form>
<% End If %>
  </table>
</td></tr>
</table><%

End Sub

Sub emoney_top() %>
<table border=0>
<tr align=center>
<td height=50><a href='?action=converion'><img src='IMAGES/SMALL/emoney_converion.gif' border=0></a></td>
<td width=50></td>
<td><a href='?action=virement'><img src='IMAGES/SMALL/emoney_virement.gif' border=0></a></td>
<td width=50></td>
<td><a href='?action=card'><img src='IMAGES/SMALL/emoney_card.gif' border=0></a></td>
</tr>
</table>
<%
End Sub

Function friend_select()
Dim sql,rs,ttt
friend_select = vbcrlf & "<script language=javascript>" & _
vbcrlf & "function Do_accept(addaccept) {" & _
vbcrlf & "  if (addaccept!=0) { document.emoney_frm_2.username2.value=addaccept; }" & _
vbcrlf & "  return;" & _
vbcrlf & "}</script>" & _
vbcrlf & "<select name=friend_select size=1 onchange=Do_accept(this.options[this.selectedIndex].value)>" & _
vbcrlf & "<option value='0'>选择我的好友</option>"
sql           = "select username2 from user_friend where username1='" & login_username & "' order by id"
Set rs        = conn.execute(sql)

Do While Not rs.eof
ttt           = rs(0)
friend_select = friend_select & vbcrlf & "<option value='" & ttt & "'>" & ttt & "</option>"
rs.movenext
Loop

rs.Close
friend_select = friend_select & vbcrlf & "</select>"
End Function %>