<!-- #include file="INCLUDE/config_other.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim old_url,web_skin:web_skin = web_var(web_config,5)
index_url   = "error"
tit         = "������Ϣ��ʾ"
tit_fir     = ""

action      = Trim(Request.cookies("beyondest_online")("error_action"))
old_url     = Trim(Request.cookies("beyondest_online")("old_url"))

If var_null(old_url) = "" Then
    old_url = "main.asp"
End If

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Response.Write left_action("jt13",4)
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------

Select Case action
    Case "loading"
        tit = "<font class=red>�Բ��𣬱�վ����ά��������С���<br><br>����ʱ����ע����½��վ��<br><br>���Ե�Ƭ�̡���<br><br>���������Ĳ��㣬������£�����</font>"
    Case "username"
        tit = "<font class=red>�����鿴��ϸ���û���Ϣ���û����������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "login"
        tit = "<font class=red>������û��ע��͵�½��վ���½��Ϣ����</font><br><br><font class=red_3>Ϊ֧�ֱ�վ�ķ�չ�����ӱ�վ��Ա���Ͷ��ɹ���<br>��վ�Ĵ󲿷���Դ����̳�����������ء����ŵȹ��ܷ���<br>��Ҫע�Ტ��ȷ��½���ܽ��С�"
    Case "power"
        tit = "<font class=red>����Ȩ��̫�ͣ�ϵͳ�����������иղŵĲ�����<br>��������Ҫ�鿴������������Լ���̳��������輶��ϸߡ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "locked"
        tit = "<font class=red>����Ŀǰ�ѱ���վ����Ա������ֻ�ܽ��е�½������Ȳ�����<br>ԭ���������֮ǰ�����˲��ѺõĲ�������Ҫ���������������վ����Ա��ϵ��</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "post"
        tit = post_error & "<br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "effect_id"
        tit = "<font class=red>�����鿴����ЧID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "islock"
        tit = "<font class=red>�����ظ��������ѱ�������</font><br><br>���������ٶԸ������лظ�������"
    Case "mail_id"
        tit = "<font class=red>�����鿴���ظ���ת����ɾ���Ķ���ID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "edit_id"
        tit = "<font class=red>�����༭������ID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "del_id"
        tit = "<font class=red>����ɾ��������ID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "forum_id"
        tit = "<font class=red>�����鿴�򷢱����ӵ���̳ID�������йع���򲻴��ڣ�<br>���ܸ����Ѿ���ɾ�������̳�Ѿ�����ʱ�رգ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "time_load"
        tit = "<font class=red>��վ�ѿ�����ˢ�»��ƣ������� " & web_var(web_num,16) & " �������ظ�����</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case "view_id"
        tit = "<font class=red>�����鿴�򷢱��������������ID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
    Case Else
        tit = "<font class=red>����δ֪����</font>�������Ա��ϵ��<br>����δ <a href=login.asp?action=register>ע��</a> ���� <a href=login.asp>��½</a>�����߲��߱�ʹ�õ�ǰ���ܵ�Ȩ�ޡ�<br><br><a href='gbook.asp?action=write'>�� �������� ��</a>"
End Select

If action <> "loading" Then
    tit = tit & "<br><br><br><a href='" & old_url & "'>����˴��ɷ��س���ҳ��ǰһҳ</a>"
End If

'response.cookies("beyondest_online")("error_action")="" %>
<!-----------------------------------center---------------------------------->
<table border=0 width=590 cellspacing=0 cellpadding=0 class=fr>
<tr><td align=right><img src='images/<% Response.Write web_skin %>/center_error.gif' border=0></td></tr>
<tr><td align=center height=380>
<table border=0 cellpadding=0 cellspacing=0 width=534>
  <tr>
   <td colspan=3><img src=images/<% Response.Write web_skin %>/error_r1_c1.gif width=534 height=42 border=0></td>
   <td><img src=images/error/spacer.gif width=1 height=42 border=0></td>
  </tr>
  <tr>
   <td rowspan=2><img src=images/<% Response.Write web_skin %>/error_r2_c1.gif width=43 height=239 border=0></td>
   <td width=479 height=228 align=center bgcolor=#f7f7f7 class=htd><% Response.Write tit %></td>
   <td rowspan=2><img src=images/<% Response.Write web_skin %>/error_r2_c3.gif width=12 height=239 border=0></td>
   <td><img src=images/<% Response.Write web_skin %>/spacer.gif width=1 height=228 border=0></td>
  </tr>
  <tr>
   <td><img src=images/<% Response.Write web_skin %>/error_r3_c2.gif width=479 height=11 border=0></td>
   <td><img src=images/<% Response.Write web_skin %>/spacer.gif width=1 height=11 border=0></td>
  </tr>
</table><br>
</td></tr></table>
<%
'---------------------------------center end-------------------------------
Call web_end(1) %>