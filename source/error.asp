<!-- #include file="INCLUDE/config_other.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim old_url,web_skin:web_skin=web_var(web_config,5)
index_url="error"
tit="������Ϣ��ʾ"
tit_fir=""

action=trim(request.cookies("beyondest_online")("error_action"))
old_url=trim(request.cookies("beyondest_online")("old_url"))
if var_null(old_url)="" then
  old_url="main.asp"
end if

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
select case action
case "loading"
  tit="<font class=red>�Բ��𣬱�վ����ά��������С���<br><br>����ʱ����ע����½��վ��<br><br>���Ե�Ƭ�̡���<br><br>���������Ĳ��㣬������£�����</font>"
case "username"
  tit="<font class=red>�����鿴��ϸ���û���Ϣ���û����������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "login"
  tit="<font class=red>������û��ע��͵�½��վ���½��Ϣ����</font><br><br><font class=red_3>Ϊ֧�ֱ�վ�ķ�չ�����ӱ�վ��Ա���Ͷ��ɹ���<br>��վ�Ĵ󲿷���Դ����̳�����������ء����ŵȹ��ܷ���<br>��Ҫע�Ტ��ȷ��½���ܽ��С�"
case "power"
  tit="<font class=red>����Ȩ��̫�ͣ�ϵͳ�����������иղŵĲ�����<br>��������Ҫ�鿴������������Լ���̳��������輶��ϸߡ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "locked"
  tit="<font class=red>����Ŀǰ�ѱ���վ����Ա������ֻ�ܽ��е�½������Ȳ�����<br>ԭ���������֮ǰ�����˲��ѺõĲ�������Ҫ���������������վ����Ա��ϵ��</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "post"
  tit=post_error&"<br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "effect_id"
  tit="<font class=red>�����鿴����ЧID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "islock"
  tit="<font class=red>�����ظ��������ѱ�������</font><br><br>���������ٶԸ������лظ�������"
case "mail_id"
  tit="<font class=red>�����鿴���ظ���ת����ɾ���Ķ���ID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "edit_id"
  tit="<font class=red>�����༭������ID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "del_id"
  tit="<font class=red>����ɾ��������ID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "forum_id"
  tit="<font class=red>�����鿴�򷢱����ӵ���̳ID�������йع���򲻴��ڣ�<br>���ܸ����Ѿ���ɾ�������̳�Ѿ�����ʱ�رգ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "time_load"
  tit="<font class=red>��վ�ѿ�����ˢ�»��ƣ������� "&web_var(web_num,16)&" �������ظ�����</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case "view_id"
  tit="<font class=red>�����鿴�򷢱��������������ID�������йع���򲻴��ڣ�</font><br><br>�����Ҹ���վ�ĳ����ύ�Ƿ�������"
case else
  tit="<font class=red>����δ֪����</font>�������Ա��ϵ��<br>����δ <a href=login.asp?action=register>ע��</a> ���� <a href=login.asp>��½</a>�����߲��߱�ʹ�õ�ǰ���ܵ�Ȩ�ޡ�<br><br><a href='gbook.asp?action=write'>�� �������� ��</a>"
end select

if action<>"loading" then
  tit=tit&"<br><br><br><a href='"&old_url&"'>����˴��ɷ��س���ҳ��ǰһҳ</a>"
end if

'response.cookies("beyondest_online")("error_action")=""
%>
<!-----------------------------------center---------------------------------->
<table border=0 width=590 cellspacing=0 cellpadding=0 class=fr>
<tr><td align=right><img src='images/<%response.write web_skin%>/center_error.gif' border=0></td></tr>
<tr><td align=center height=380>
<table border=0 cellpadding=0 cellspacing=0 width=534>
  <tr>
   <td colspan=3><img src=images/<%response.write web_skin%>/error_r1_c1.gif width=534 height=42 border=0></td>
   <td><img src=images/error/spacer.gif width=1 height=42 border=0></td>
  </tr>
  <tr>
   <td rowspan=2><img src=images/<%response.write web_skin%>/error_r2_c1.gif width=43 height=239 border=0></td>
   <td width=479 height=228 align=center bgcolor=#f7f7f7 class=htd><%response.write tit%></td>
   <td rowspan=2><img src=images/<%response.write web_skin%>/error_r2_c3.gif width=12 height=239 border=0></td>
   <td><img src=images/<%response.write web_skin%>/spacer.gif width=1 height=228 border=0></td>
  </tr>
  <tr>
   <td><img src=images/<%response.write web_skin%>/error_r3_c2.gif width=479 height=11 border=0></td>
   <td><img src=images/<%response.write web_skin%>/spacer.gif width=1 height=11 border=0></td>
  </tr>
</table><br>
</td></tr></table>
<%
'---------------------------------center end-------------------------------
call web_end(1)
%>