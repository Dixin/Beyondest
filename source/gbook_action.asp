<!--#include file="include/onlogin.asp"-->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim html_temp,id,viewpage,affirm,reicon,reword,retim
id        = Trim(Request.querystring("id"))
viewpage  = Trim(Request.querystring("page"))
affirm    = Trim(Request.form("affirm"))
tit       = "<a href='?action=reply&id=" & id & "&page=" & viewpage & "'>�ظ�����</a>&nbsp;��&nbsp;<a href='?action=delete&id=" & id & "&page=" & viewpage & "'>ɾ������</a>"

html_temp = header( - 1,tit)

If action = "delete" Then
    html_temp = html_temp & "<script language=JavaScript><!--" & _
    vbcrlf & "function confirm_delete()" & _
    vbcrlf & "{" & _
    vbcrlf & "if (confirm(""��ȷ��Ҫɾ������������"")){ return true; }" & _
    vbcrlf & "return false;" & _
    vbcrlf & "}" & _
    vbcrlf & "//--></script>" & _
    vbcrlf & "<table border=0 width=100% height=100% cellpadding=0 cellspacing=0>" & _
    vbcrlf & "<tr><td width=100% align=center>" & _
    vbcrlf & "  <table border=1 cellpadding=0 cellspacing=0 bordercolor=" & web_var(web_color,2) & " width=400>" & _
    vbcrlf & "  <tr><td width=100% align=center>"
Else
    html_temp = html_temp & "<script language=javascript><!--" & _
    vbcrlf & "function check(reply_form)" & _
    vbcrlf & "{" & _
    vbcrlf & "if( reply_form.reply_word.innertext == """" )" & _
    vbcrlf & "  { alert(""�ظ������Ǳ���Ҫ�ġ�"");return false; }" & _
    vbcrlf & "if (reply_form.reply_word.value.length > 10000)" & _
    vbcrlf & "  { alert(""�Բ��𣬻ظ����ݲ��ܳ��� 10000 ���ֽڣ�"");return false; }" & _
    vbcrlf & "}" & _
    vbcrlf & "function reset(reply_form)" & _
    vbcrlf & "{" & _
    vbcrlf & "  if (confirm(""�������Ҫ���ȫ�������ݣ���ȷ��Ҫ�����?"")){ return true; }" & _
    vbcrlf & "  return false;" & _
    vbcrlf & "}" & _
    vbcrlf & "--></script>" & _
    vbcrlf & "<table border=0 width=100% height=100% cellpadding=0 cellspacing=0>" & _
    vbcrlf & "<tr><td width=100% align=center>" & _
    vbcrlf & "<table border=1 cellpadding=0 cellspacing=0 bordercolor=" & web_var(web_color,2) & " width=506>" & _
    vbcrlf & "<tr><td width=100% align=center>"
End If

If affirm = "ok" Then

    If action = "delete" Then
        conn.Execute("Delete from gb_data where ID = " & id)
        html_temp = html_temp & VbCrLf & "<br>�ѳɹ���ɾ��IDΪ<font class=red>" & id & " </font>�����ԣ�<br><br>" & VbCrLf & "<a href=gbook.asp?page=" & viewpage & ">�������Բ�</a><br><br>" & VbCrLf & "��ϵͳ���� " & web_var(web_num,5) & " ���Ӻ��Զ����أ�<br><br>" & _
        VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=gbook.asp?page=" & viewpage & "'>"
    Else
        reicon = Request.form("reply_icon")
        reword = Request.form("reply_word")
        retim  = Now
        Set rs = Server.CreateObject("adodb.recordset")
        sql    = "Select re_icon,re_word,re_tim from gb_data where ID=" & id
        rs.open sql,conn,1,3

        If rs.eof And rs.bof Then
            html_temp      = html_temp & vbcrlf & "<p>���������Ҳ������Ϊ" & id & "�����ԣ���˲��ܽ��лظ�������</p>" & closer
        Else
            rs("re_icon") = reicon
            rs("re_word") = reword
            rs("re_tim") = retim
            rs.update
            rs.Close:Set rs = Nothing
            html_temp = html_temp & vbcrlf & "<br>�ѳɹ��Ļظ�IDΪ<font class=red> " & id & " </font>�����ԣ�" & VbCrLf & "<br><br><a href=gbook.asp?page=" & viewpage & ">�������Բ�</a><br><br>" & VbCrLf & "��ϵͳ���� " & web_var(web_num,5) & " ���Ӻ��Զ����أ�<br><br>" & _
            VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=gbook.asp?page=" & viewpage & "'>"
        End If

    End If

Else

    If action = "delete" Then
        html_temp = html_temp & VbCrLf & "<table border=0 width=100% cellpadding=0 cellspacing=0>" & _
        vbcrlf & "<tr><td align=center bgcolor=" & web_var(web_color,2) & " height=20>" & _
        vbcrlf & "<font class=end><b>ɾ �� �� ��</b></font>" & _
        vbcrlf & "</td></tr>" & _
        vbcrlf & "<tr><form name=reply_form method=POST action='?action=" & action & "&id=" & id & "&page=" & viewpage & "'>" & _
        vbcrlf & "<input type=hidden name=affirm value='ok'>" & _
        vbcrlf & "<td height=50 align=center>" & _
        vbcrlf & "�˲�����ɾ��idΪ<font class=red> " & id & " </font>�����ԣ��Ƿ�ȷ������" & _
        vbcrlf & "</td></tr>" & _
        vbcrlf & "<tr><td align=center>" & _
        vbcrlf & "<input type=submit value=' ȷ �� ' onclick=""return confirm_delete()"">&nbsp;&nbsp;" & _
        vbcrlf & "<input type=button value=' ȡ �� ' onclick='javascript:history.back(1)'> " & _
        vbcrlf & "</td></tr>"
    Else
        html_temp = html_temp & vbcrlf & "<table border=0 cellpadding=0 cellspacing=0 width=550>" & _
        vbcrlf & "<tr><form name=reply_form method=POST action='?action=" & action & "&id=" & id & "&page=" & viewpage & "'>" & _
        vbcrlf & "<input type=hidden name=affirm value='ok'>" & _
        vbcrlf & "<td align=center colspan=2 bgcolor=" & web_var(web_color,2) & " height=20>" & _
        vbcrlf & "<font class=end><b>�� �� �� ��</b></font></td></tr>" & _
        vbcrlf & "<tr><td height=15 colspan=2>  </td></tr>" & _
        vbcrlf & "<tr><td align=center height=30 colspan=2>Щ�������ظ�IDΪ<font class=red>  " & Request("id") & " </font>������</td></tr>" & _
        vbcrlf & "<tr><td align=center width=80 height=10>��</td><td align=left width=440>  </td></tr>" & _
        vbcrlf & "<tr><td align=center height=25>����ͼ��: </td><td align=left>" & _
        vbcrlf & "<img border=0 src='images/icon/0.gif'>" & _
        vbcrlf & "<input type=radio value=0 name=reply_icon checked class=bg_1>" & _
        vbcrlf & "<img border=0 src='images/icon/1.gif'>" & _
        vbcrlf & "<input type=radio value=1 name=reply_icon class=bg_1>" & _
        vbcrlf & "<img border=0 src='images/icon/2.gif'>" & _
        vbcrlf & "<input type=radio value=2 name=reply_icon class=bg_1>" & _
        vbcrlf & "<img border=0 src='images/icon/3.gif'>" & _
        vbcrlf & "<input type=radio value=3 name=reply_icon class=bg_1>" & _
        vbcrlf & "<img border=0 src='images/icon/4.gif'>" & _
        vbcrlf & "<input type=radio value=4 name=reply_icon class=bg_1>" & _
        vbcrlf & "<img border=0 src='images/icon/5.gif'>" & _
        vbcrlf & "<input type=radio value=5 name=reply_icon class=bg_1>" & _
        vbcrlf & "<img border=0 src='images/icon/6.gif'>" & _
        vbcrlf & "<input type=radio value=6 name=reply_icon class=bg_1>" & _
        vbcrlf & "<img border=0 src='images/icon/7.gif'>" & _
        vbcrlf & "<input type=radio value=7 name=reply_icon class=bg_1>" & _
        vbcrlf & "<img border=0 src='images/icon/8.gif'>" & _
        vbcrlf & "<input type=radio value=8 name=reply_icon class=bg_1>" & _
        vbcrlf & "<img border=0 src='images/icon/9.gif'>" & _
        vbcrlf & "<input type=radio value=9 name=reply_icon class=bg_1>" & _
        vbcrlf & "</td></tr>" & _
        vbcrlf & "<tr><td align=center valign=top><br>�ظ�����:<br><br>" & web_var(web_error,3) & "<br><br>����<=10KB</td>" & _
        vbcrlf & "<td align=left>" & _
        vbcrlf & "<table border=0 cellpadding=0 cellspacing=0 width='100%'>" & _
        vbcrlf & "<tr><td>" & _
        vbcrlf & "<textarea rows=8 name=reply_word cols=70></textarea></td></tr></table>" & _
        vbcrlf & "</td></tr>" & _
        vbcrlf & "<tr><td align=center colspan=2></td></tr>" & _
        vbcrlf & "<tr><td align=center colspan=2 height=50>" & _
        vbcrlf & "<input type=submit value=' �� �� �� �� ' onclick=""return check(this.form)"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & _
        vbcrlf & "<input type=reset value=' �� д ' onclick=""return reset()"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type=button value=' ȡ �� ' onclick='javascript:history.back(1)'>" & _
        vbcrlf & "</td></form></tr>" & _
        vbcrlf & "</table>"
    End If

End If

Call close_conn()
Response.Write html_temp & ender() %>