<!-- #include file="include/config_forum.asp" -->
<% If Not(IsNumeric(forumid)) Or Not(IsNumeric(viewid)) Then Call cookies_type("view_id") %>
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Call forum_first()
Call forum_word()

Dim ii
Dim rssum
Dim nummer
Dim thepages
Dim viewpage
Dim page
Dim pageurl
Dim view_temp
Dim iid
Dim qid
Dim table_bg
Dim id
Dim money_name
Dim thetopictt
Dim u_username
Dim u_nname
Dim u_sex
Dim u_whe
Dim u_qq
Dim u_remark
Dim u_bbs_counter
Dim u_integral
Dim u_emoney
Dim u_power
Dim u_popedom
Dim fir_id
Dim fir_topic
Dim fir_top
Dim counter
Dim re_counter
Dim fir_istop
Dim fir_isgood
Dim fir_islock
Dim del_type
tit     = forumname & "��������ӣ�"
pageurl = "?forum_id=" & forumid & "&view_id=" & viewid & "&"
nummer  = web_var(web_num,3):view_temp = "":money_name = web_var(web_config,8)

Set rs  = Server.CreateObject("adodb.recordset")
sql     = "select bbs_data.*,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.islock,bbs_topic.istop,bbs_topic.isgood,user_data.username as u_username," & _
"user_data.nname as u_nname,user_data.sex as u_sex,user_data.whe as u_whe,user_data.qq as u_qq,user_data.email as u_email," & _
"user_data.url as u_url,user_data.face as u_face,user_data.tim as u_tim,user_data.remark as u_remark,user_data.bbs_counter as u_bbs_counter,user_data.emoney as u_emoney,user_data.integral as u_integral,user_data.power as u_power,user_data.popedom as u_popedom " & _
"from user_data inner join ( bbs_topic inner join bbs_data on bbs_data.reply_id=bbs_topic.id )" & _
" on bbs_data.username=user_data.username where bbs_data.forum_id=" & forumid & " and bbs_data.reply_id=" & viewid & " order by bbs_data.id"
rs.open sql,conn,1,1

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing:close_conn
    Call cookies_type("view_id")
End If

Call web_head(0,0,2,0,0)
'-----------------------------------center--------------------------------- %>
<script language=JavaScript>
<!--
function forum_do_del(data1,data2)
{
  if (confirm("�˲�����ɾ��idΪ "+data2+" �Ļ�����\n\n���Ҫɾ����\nɾ�����޷��ָ���"))
    window.location = "forum_isaction.asp?isaction=del&forum_id="+data1+"&del_id="+data2
}
function forum_do_delete(data1,data2)
{
  if (confirm("�˲�����ɾ��idΪ "+data2+" �����ӣ�\n\n���Ҫɾ����\nɾ�����޷��ָ���"))
    window.location = "forum_isaction.asp?isaction=delete&forum_id="+data1+"&del_id="+data2
}
//-->
</script>

<%
thetopictt = forum_table1 & "<tr height=26><td background=images/" & web_var(web_config,5) & "/bar_1_bg.gif colspan=2>"
view_temp  = "</td></tr>"

rssum = rs.recordcount
Call format_pagecute()

If Int(viewpage) > 1 Then
    fir_id     = rs("reply_id")
    fir_topic  = rs("topic")
    fir_islock = rs("islock"):fir_istop = rs("istop"):fir_isgood = rs("isgood")
    fir_istop  = Int(fir_istop):fir_isgood = Int(fir_isgood):fir_islock = Int(fir_islock)
    fir_top    = fir_topic
    fir_top    = code_html(fir_top,1,0)
    counter    = rs("counter")
    re_counter = rs("re_counter")
    rs.move (viewpage - 1)*nummer
End If

For ii = 1 To nummer
    If rs.eof Then Exit For
    iid            = rs("id")
    qid            = iid
    id             = rs("reply_id")
    u_username     = rs("u_username")
    u_nname        = code_html(rs("u_nname"),1,0)
    u_sex          = rs("u_sex")
    u_whe          = code_html(rs("u_whe"),1,0)
    u_qq           = rs("u_qq")
    u_remark       = code_jk2(rs("u_remark"))
    u_bbs_counter  = rs("u_bbs_counter")
    u_integral     = rs("u_integral")
    u_emoney       = rs("u_emoney")
    u_power        = rs("u_power")
    u_popedom      = rs("u_popedom")
    del_type       = "forum_do_del"

    If Int(ii) = 1 And Int(viewpage) = 1 Then
        fir_id     = id
        fir_topic  = rs("topic")
        fir_islock = rs("islock"):fir_istop = rs("istop"):fir_isgood = rs("isgood")
        fir_istop  = Int(fir_istop):fir_isgood = Int(fir_isgood):fir_islock = Int(fir_islock)
        fir_top    = fir_topic
        fir_top    = code_html(fir_top,1,0)
        counter    = rs("counter")
        re_counter = rs("re_counter")
        iid        = viewid
        del_type   = "forum_do_delete"
    End If

    view_temp = view_temp & view_type()
    rs.movenext
Next

rs.Close:Set rs = Nothing
view_temp = view_temp & "</td></tr></table>"

fir_istop = Int(fir_istop)
If fir_istop <> 0 And fir_istop <> 1 And fir_istop <> 2 Then fir_istop = 0

Response.Write forum_top("������� ���ظ���<font class=red>" & re_counter & "</font>&nbsp;�����<font class=red>" & counter + 1 & "</font>��") %>

<table border=0 width='98%' cellspacing=0 cellpadding=0><tr><td align=left width='15%'><a href='forum_write.asp?forum_id=<% = forumid %>'><img src='images/<% = web_var(web_config,5) %>/new_topic.gif' align=absMiddle border=0 title='�� <% = forumname %> �﷢���ҵ�����'></a></td><td align=right width='85%'></td></tr></table>



<% Response.Write kong & thetopictt %> 

<table boder=0 width='100%' cellspacing=0 cellpadding=0>
<tr><td width='80%'>&nbsp;���⣺<b><font class=end title='<% Response.Write fir_top %>'><% Response.Write code_html(fir_topic,1,30) %></font></b></td>
<td align=center width='20%'><table border=0 cellspacing=0 cellpadding=0><tr align=center>
  <td width=50><a href='javascript:;' onclick="javascript:document.location.reload()"><% Response.Write img_small("page_refresh") %></a></td>
  <td width=55><a href="javascript:window.external.AddFavorite('<% Response.Write web_var(web_config,2) & pageurl %>','<% Response.Write web_var(web_config,1) & " - " & forumname & "�����ӣ�" & fir_topic & "��" %>')"><% Response.Write img_small("page_fav") %></td>
  </tr></table>
  
</td></tr></table>

<% Response.Write view_temp & kong & format_table(1,2) %>



<tr height=30<% Response.Write forum_table3 %>>
<td width='75%'>&nbsp;��ҳ��<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,6,"#ff0000") %></td>
<td width='25%' align=center><% Response.Write forum_go() %></td>
</tr>
<tr height=30 align=center<% Response.Write format_table(3,1) %>><td>���������ͣ�<font class=blue><%

If fir_istop <> 0 Or fir_isgood <> 0 Or fir_islock <> 0 Then

    If fir_istop = 1 Then
        Response.Write "[ �̶� ]&nbsp;"
    ElseIf fir_istop = 2 Then
        Response.Write "[ �̶ܹ� ]&nbsp;"
    End If

    If fir_isgood <> 0 Then Response.Write "[ ���� ]&nbsp;"
    If fir_islock <> 0 Then Response.Write "[ ���� ]&nbsp;"
Else
    Response.Write "[ ���� ]&nbsp;"
End If

Response.Write "</font>"

If format_user_power(login_username,login_mode,forumpower) = "yes" Then %>&nbsp;��ز�����
<a href='forum_isaction.asp?isaction=is&forum_id=<% Response.Write forumid %>&view_id=<% Response.Write id %>&action=istop<%

    Select Case fir_istop
        Case 1
            Response.Write "&cancel=yes' class=red_3>ȡ���̶�</a>&nbsp;��" & _
            "<a href='forum_isaction.asp?isaction=is&forum_id=" & forumid & "&view_id=" & id & "&action=istops'>�̶ܹ�</a>"
        Case 2
            Response.Write "'>�̶�</a>&nbsp;��" & _
            "<a href='forum_isaction.asp?isaction=is&forum_id=" & forumid & "&view_id=" & id & "&action=istops&cancel=yes' class=red_3>ȡ���̶ܹ�</a>"
        Case Else
            Response.Write "'>�̶�</a>&nbsp;��" & _
            "<a href='forum_isaction.asp?isaction=is&forum_id=" & forumid & "&view_id=" & id & "&action=istops'>�̶ܹ�</a>"
    End Select %>&nbsp;��
<a href='forum_isaction.asp?isaction=is&forum_id=<% Response.Write forumid %>&view_id=<% Response.Write id %>&action=isgood<%

    If fir_isgood = 0 Then
        Response.Write "'>"
    Else
        Response.Write "&cancel=yes' class=red_3>ȡ��"
    End If %>����</a>&nbsp;��
<a href='forum_isaction.asp?isaction=is&forum_id=<% Response.Write forumid %>&view_id=<% Response.Write id %>&action=islock<%

    If fir_islock = 0 Then
        Response.Write "'>"
    Else
        Response.Write "&cancel=yes' class=red_3>ȡ��"
    End If %>����</a>&nbsp;��
<a href='forum_isaction.asp?isaction=delete&forum_id=<% Response.Write forumid %>&del_id=<% Response.Write id %>'>ɾ��</a>
<% End If %>
</td>
<td><% Response.Write forum_move(forumid,viewid) %></td></tr>
</table>
<script language=javascript src='style/forum_ok.js'></script>
<% Response.Write kong & forum_table1 %>
<form name=write_frm action='forum_reply.asp?forum_id=<% = forumid %>&view_id=<% = qid %>' method=post onsubmit="frm_submitonce(this);">
<input type=hidden name=reply value='ok'>
<tr<% Response.Write forum_table2 %>><td height=25 valign=bottom colspan=2 background=images/<% = web_var(web_config,5) %>/bar_1_bg.gif>
<%

If fir_islock <> 1 Then

    If login_mode = "" Then
        Response.Write "<div align=center>" & web_var(web_error,2) & "</div>"
    Else %>
&nbsp;�� ���ٻظ���<b><font class=red_3><% Response.Write fir_top %></b></font>
</td></tr>
<tr height=30<% Response.Write format_table(3,1) %>>
<td width='20%' align=center bgcolor='<% = web_var(web_color,6) %>'>�û���Ϣ��</td>
<td width='80%'>&nbsp;&nbsp;�û�����<input type=username name=username value='<% Response.Write login_username %>' size=18 maxlength=20>&nbsp;&nbsp;
���룺<input type=password name=password value='<% Response.Write login_password %>' size=18 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='user_main.asp?username=<% Response.Write Server.htmlencode(login_username) %>'>�û�����</a> ]</font>&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='login.asp?action=logout'>�˳���½</a> ]</font></td>
</tr>
<tr height=30<% Response.Write forum_table3 %>>
<td align=center bgcolor='<% = web_var(web_color,6) %>'>������ţ�</td>
<td>&nbsp;&nbsp;<% Response.Write icon_type(9,3) %></td>
</tr>
<tr align=center<% Response.Write format_table(3,1) %>>
<td bgcolor='<% = web_var(web_color,6) %>'><table border=0><tr><td class=htd>�������ݣ�<% Response.Write redx %><br><% Response.Write word_remark %></td></tr></table></td>
<td><table border=0><tr><td><textarea name=jk_word rows=8 cols=95 title='�� Ctrl+Enter ��ֱ�ӷ���' onkeydown="javascript:frm_quicksubmit();"></textarea></td></tr></table></td>
</tr>
<script language=javascript src='style/em_type.js'></script>
<tr height=30<% Response.Write forum_table3 %>>
<td align=center bgcolor='<% = web_var(web_color,6) %>'>E M ��ͼ��</td>
<td>&nbsp;<script language=javascript>jk_em_type('s');</script></td>
</tr>
<tr<% Response.Write format_table(3,1) %>><td colspan=2 align=center height=60>&nbsp;&nbsp;<script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<% Response.Write forum_table3 %>>
<td bgcolor='<% = web_var(web_color,6) %>'>���ٻظ���</td>
<td><input type=submit name=wsubmit value='���ٷ����ҵĻ���'>��&nbsp;<input type=button value='Ԥ���ҵĻظ�'>��&nbsp;<input type=reset value='�����д'>��&nbsp;���� Ctrl + Enter �ɿ��ٻظ���
<%
    End If

Else
    Response.Write "<div align=center><font class=red_2>��������ѱ������������ٶ�����лظ�</font></div>"
End If %>
</td></tr></form></table>
<br>
<%
sql = "update bbs_topic set counter=counter+1 where forum_id=" & forumid & " and id=" & viewid
conn.execute(sql)
'---------------------------------center end-------------------------------
Call web_end(0) %>