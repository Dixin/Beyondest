<!-- #INCLUDE file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/fso_file.asp" -->
<!-- #INCLUDE file="INCLUDE/common_other.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim tit_menu
tit_menu="<a href='?'>�����޸�</a>&nbsp;��&nbsp;"&_
	 "<a href='javascript:Do_b_data();'>��������</a>&nbsp;��&nbsp;"&_
	 "<a href='javascript:Do_h_data();'>��ԭ����</a>"
response.write header(3,tit_menu)
%>
<script language=JavaScript><!--
function Do_b_data()
{
if (confirm("�˲����� ���� ���е���վ���ã�\n\n���Ҫ������\n���ݺ��޷��ָ���"))
  window.location="?action=b"
}
function Do_h_data()
{
if (confirm("�˲����� ��ԭ ���е���վ���ã�\n\n���Ҫ������\n��ԭ���޷��ָ���"))
  window.location="?action=h"
}
//--></script>
<%
if action="b" or action="h" then call config_bh(action)

select case trim(request.querystring("edit"))
case "chk"
  call config_chk()
case else
  call config_main()
end select

response.write ender()

sub config_main()
  dim t1,t2,j,tdim,udim,tt:tt=0
%>
<table border=1 width='92%' cellspacing=0 cellpadding=2<%response.write table1%>>
<tr><td colspan=3 align=center>
  <table border=0 cellspacing=0 cellpadding=3>
  <tr>
  <td><%response.write img_small("jt1")%><a href='?action=config'>������Ϣ����</a></td>
  <td><%response.write img_small("jt1")%><a href='?action=config2'>������������</a></td>
  <td><%response.write img_small("jt1")%><a href='?action=num'>ҳ��ʾ������</a></td>
  <td><%response.write img_small("jt12")%><a href='?action=info'>������ʾ����</a></td>
  <td><%response.write img_small("jt1")%><a href='?action=down_up'>�����ϴ�����</a></td>
  </tr>
  <tr>
  <td><%response.write img_small("jt1")%><a href='?action=menu'>��Ŀ�˵�����</a></td>
  <td><%response.write img_small("jt1")%><a href='?action=color'>��վ��ɫ����</a></td>
  <td><%response.write img_small("jt12")%><a href='?action=user'>�� �� �����</a></td>
  <td><%response.write img_small("jt12")%><a href='?action=grade'>�û��ȼ�����</a></td>
  <td><%response.write img_small("jt12")%><a href='?action=forum'>��̳�������</a></td>
  </tr>
  </table>
</td></tr>
<tr align=center bgcolor=<%response.write color2%>>
<td width='16%'>��������</td>
<td width='50%'>����</td>
<td width='34%'>���˵��</td>
</tr>
<form action='?action=<%response.write action%>&edit=chk' method=post>
<input type=hidden name=web_news_art value='<%response.write web_news_art%>'>
<input type=hidden name=web_shop value='<%response.write web_shop%>'>
<%
if action="config" then
  tt=1
%>
<tr>
<td>��վ���ƣ�</td>
<td><input type=text name=web_config_1 value='<%response.write web_var(web_config,1)%>' size=38 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>��վ��ַ��</td>
<td><input type=text name=web_config_2 value='<%response.write web_var(web_config,2)%>' size=38 maxlength=50></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>��������Ա��</td>
<td><input type=text name=web_config_3 value='<%response.write web_var(web_config,3)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>����Ŀ¼��</td>
<td><input type=text name=web_config_4 value='<%response.write web_var(web_config,4)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>��վSKIN��</td>
<td><input type=text name=web_config_5 value='<%response.write web_var(web_config,5)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>Cookies���ƣ�</td>
<td><input type=text name=web_cookies value='<%response.write web_cookies%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>���ݿ����ƣ�</td>
<td><input type=text name=web_config_6 value='<%response.write web_var(web_config,6)%>' size=38 maxlength=50></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>��վ������</td>
<td><input type=text name=web_config_7 value='<%response.write web_var(web_config,7)%>' size=20 maxlength=20></td>
<td background='images/<%response.write web_var(web_config,7)%>.gif'>&nbsp;</td>
</tr>
<tr>
<td>������ң�</td>
<td><input type=text name=web_config_8 value='<%response.write web_var(web_config,8)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;��������������</td>
</tr>
<% else %>
<input type=hidden name=web_config_1 value='<%response.write web_var(web_config,1)%>'>
<input type=hidden name=web_config_2 value='<%response.write web_var(web_config,2)%>'>
<input type=hidden name=web_config_3 value='<%response.write web_var(web_config,3)%>'>
<input type=hidden name=web_config_4 value='<%response.write web_var(web_config,4)%>'>
<input type=hidden name=web_config_5 value='<%response.write web_var(web_config,5)%>'>
<input type=hidden name=web_cookies value='<%response.write web_cookies%>'>
<input type=hidden name=web_config_6 value='<%response.write web_var(web_config,6)%>'>
<input type=hidden name=web_config_7 value='<%response.write web_var(web_config,7)%>'>
<input type=hidden name=web_config_8 value='<%response.write web_var(web_config,8)%>'>
<%
end if

if action="config2" then
  tt=1
%>
<tr>
<td>��վ״̬��</td>
<td><input type=radio name=web_login value='1'<% if int(web_login)=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_login value='0'<% if int(web_login)<>1 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;�Ƿ񿪷���վ</td>
</tr>
<% t1=web_var_num(web_setup,1,1) %>
<tr>
<td>��½�����</td>
<td colspan=2><input type=radio name=web_setup_1 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;��&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_1 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;<font class=gray>�Ƿ�Ҫ��½�ſ�������»����������</font></td>
</tr>
<% t1=web_var_num(web_setup,2,1) %>
<tr>
<td>ע����ˣ�</td>
<td><input type=radio name=web_setup_2 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;��&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_2 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;��</td>
<td class=gray>&nbsp;�Ƿ����ע��������</td>
</tr>
<%
t1=web_var_num(web_setup,3,1)
if t1<>1 and t1<>2 then t1=0
%>
<tr>
<td>��վģʽ��</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_1 value='<%response.write web_var(web_stamp,1)%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;��վ��ע���û����Ե�½������¼�����б�</td>
<tr>
<td>&nbsp;</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_2 value='<%response.write web_var(web_stamp,2)%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;���е�½�������վ���˱�����¼�����б�</td>
<tr>
<td>&nbsp;</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='2'<% if t1=2 then response.write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_3 value='<%response.write web_var(web_stamp,3)%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;��վ��ע���û����Ե�½������¼�û������б�</td>
</tr>
<% t1=web_var_num(web_setup,4,1) %>
<tr>
<td>��Ϣ���ˣ�</td>
<td><input type=radio name=web_setup_4 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;��&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_4 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;��</td>
<td class=gray>&nbsp;�Ƿ����̳�Ƚ����ַ�����</td>
</tr>
<%
t1=web_var_num(web_setup,5,1)
if t1<>0 and t1<>1 then t1=2
%>
<tr>
<td>��ʾ IP��</td>
<td><input type=radio name=web_setup_5 value='0'<% if t1=0 then response.write " checked" %> class=bg_1> ��ȫ����
<input type=radio name=web_setup_5 value='1'<% if t1=1 then response.write " checked" %> class=bg_1> ��ʾ����
<input type=radio name=web_setup_5 value='2'<% if t1=2 then response.write " checked" %> class=bg_1> ��ȫ����</td>
<td class=gray>&nbsp;�Թ���Ա������ȫ����</td>
</tr>
<% t1=web_var_num(web_setup,6,1) %>
<tr>
<td>������ʾ��</td>
<td><input type=radio name=web_setup_6 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;��������&nbsp;<input type=radio name=web_setup_6 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;�����˵���</td>
<td class=gray>&nbsp;��̳������ʾģʽ</td>
</tr>
<% t1=web_var_num(web_setup,7,1) %>
<tr>
<td>������ʽ��</td>
<td><input type=radio name=web_setup_7 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;��¼���&nbsp;<input type=radio name=web_setup_7 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;��¼һ��</td>
<td class=gray>&nbsp;��վ�����ķ�ʽ</td>
</tr>
<% t1=web_var_num(web_var(web_config,9),1,1) %><tr>
<td>�������ţ�</td>
<td><input type=radio name=web_config_9_1 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_1 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr><% t1=web_var_num(web_var(web_config,9),2,1) %>
<tr>
<td>�������£�</td>
<td><input type=radio name=web_config_9_2 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_2 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr><% t1=web_var_num(web_var(web_config,9),3,1) %>
<tr>
<td>������֣�</td>
<td><input type=radio name=web_config_9_3 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_3 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr><% t1=web_var_num(web_var(web_config,9),4,1) %>
<tr>
<td>�ϴ���ͼ��</td>
<td><input type=radio name=web_config_9_4 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_4 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr><% t1=web_var_num(web_var(web_config,9),5,1) %>
<tr>
<td>�Ƽ���վ��</td>
<td><input type=radio name=web_config_9_5 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_5 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr>
<% else %>
<input type=hidden name=web_login value='<%response.write web_login%>'>
<input type=hidden name=web_stamp_1 value='<%response.write web_var(web_stamp,1)%>'>
<input type=hidden name=web_stamp_2 value='<%response.write web_var(web_stamp,2)%>'>
<input type=hidden name=web_stamp_3 value='<%response.write web_var(web_stamp,3)%>'>
<input type=hidden name=web_setup_1 value='<%response.write web_var_num(web_setup,1,1)%>'>
<input type=hidden name=web_setup_2 value='<%response.write web_var_num(web_setup,2,1)%>'>
<input type=hidden name=web_setup_3 value='<%response.write web_var_num(web_setup,3,1)%>'>
<input type=hidden name=web_setup_4 value='<%response.write web_var_num(web_setup,4,1)%>'>
<input type=hidden name=web_setup_5 value='<%response.write web_var_num(web_setup,5,1)%>'>
<input type=hidden name=web_setup_6 value='<%response.write web_var_num(web_setup,6,1)%>'>
<input type=hidden name=web_setup_7 value='<%response.write web_var_num(web_setup,7,1)%>'>
<input type=hidden name=web_config_9_1 value='<%response.write web_var_num(web_var(web_config,9),1,1)%>'>
<input type=hidden name=web_config_9_2 value='<%response.write web_var_num(web_var(web_config,9),2,1)%>'>
<input type=hidden name=web_config_9_3 value='<%response.write web_var_num(web_var(web_config,9),3,1)%>'>
<input type=hidden name=web_config_9_4 value='<%response.write web_var_num(web_var(web_config,9),4,1)%>'>
<input type=hidden name=web_config_9_5 value='<%response.write web_var_num(web_var(web_config,9),5,1)%>'>
<%
end if

if action="num" then
  tt=1
%>
<tr>
<td>�û������ȣ�</td>
<td><input type=text name=web_num_1 value='<%response.write web_var(web_num,1)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>ÿҳ��������</td>
<td><input type=text name=web_num_2 value='<%response.write web_var(web_num,2)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;��̳����Աÿҳ��ʾ��</td>
</tr>
<tr>
<td>ÿҳ��ʾ����</td>
<td><input type=text name=web_num_3 value='<%response.write web_var(web_num,3)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;�鿴�������ݵ�</td>
</tr>
<tr>
<td>ÿҳ��������</td>
<td><input type=text name=web_num_4 value='<%response.write web_var(web_num,4)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;ÿҳ������Ŀ��</td>
</tr>
<tr>
<td>�Զ����أ�</td>
<td><input type=text name=web_num_5 value='<%response.write web_var(web_num,5)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;�Զ�����ʱ�䣬��λΪ��</td>
</tr>
<tr>
<td>���ⳤ�ȣ�</td>
<td><input type=text name=web_num_6 value='<%response.write web_var(web_num,6)%>' size=20 maxlength=20>&nbsp;KB</td>
<td class=gray>&nbsp;��̳��������ȳ���</td>
</tr>
<tr>
<td>ͼƬ��ȣ�</td>
<td><input type=text name=web_num_7 value='<% response.write web_var(web_num,7) %>' size=20 maxlength=20>&nbsp;����</td>
<td class=gray>&nbsp;���ء�ͼ���ͼƬ��ʾ���</td>
</tr>
<tr>
<td>ͼƬ�߶ȣ�</td>
<td><input type=text name=web_num_8 value='<% response.write web_var(web_num,8) %>' size=20 maxlength=20>&nbsp;����</td>
<td class=gray>&nbsp;���ء�ͼ���ͼƬ��ʾ�߶�</td>
</tr>
<tr>
<td>����ȣ�</td>
<td><input type=text name=web_num_9 value='<% response.write web_var(web_num,9) %>' size=20 maxlength=20>&nbsp;����</td>
<td class=gray>&nbsp;��ͼ��FLASH�������</td>
</tr>
<tr>
<td>���߶ȣ�</td>
<td><input type=text name=web_num_10 value='<% response.write web_var(web_num,10) %>' size=20 maxlength=20>&nbsp;����</td>
<td class=gray>&nbsp;��ͼ��FLASH�����߶�</td>
</tr>
<tr>
<td>�û�ͷ������</td>
<td><input type=text name=web_num_11 value='<% response.write web_var(web_num,11) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;�û�ͷ�������</td>
</tr>
<tr>
<td>���ϵ����</td>
<td><input type=text name=web_num_12 value='<% response.write web_var(web_num,12) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;����еļ��</td>
</tr>
<tr>
<td>��½��ʱ��</td>
<td><input type=text name=web_num_13 value='<% response.write web_var(web_num,13) %>' size=20 maxlength=20>&nbsp;����</td>
<td class=gray>&nbsp;�û���½��ʱ��ʱ��</td>
</tr>
<tr>
<td>���ֻ��㣺</td>
<td><input type=text name=web_num_14 value='<% response.write web_var(web_num,14) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;���ֻ������</td>
</tr>
<tr>
<td>�����ӷ֣�</td>
<td><input type=text name=web_num_15 value='<% response.write web_var(web_num,15) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;ǰ̨������Ϣ�ӷ�ֵ</td>
</tr>
<tr>
<td>��ˢʱ�䣺</td>
<td><input type=text name=web_num_16 value='<% response.write web_var(web_num,16) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;��λΪ����</td>
</tr>
<% else %>
<input type=hidden name=web_num_1 value='<%response.write web_var(web_num,1)%>'>
<input type=hidden name=web_num_2 value='<%response.write web_var(web_num,2)%>'>
<input type=hidden name=web_num_3 value='<%response.write web_var(web_num,3)%>'>
<input type=hidden name=web_num_4 value='<%response.write web_var(web_num,4)%>'>
<input type=hidden name=web_num_5 value='<%response.write web_var(web_num,5)%>'>
<input type=hidden name=web_num_6 value='<%response.write web_var(web_num,6)%>'>
<input type=hidden name=web_num_7 value='<%response.write web_var(web_num,7)%>'>
<input type=hidden name=web_num_8 value='<%response.write web_var(web_num,8)%>'>
<input type=hidden name=web_num_9 value='<%response.write web_var(web_num,9)%>'>
<input type=hidden name=web_num_10 value='<%response.write web_var(web_num,10)%>'>
<input type=hidden name=web_num_11 value='<%response.write web_var(web_num,11)%>'>
<input type=hidden name=web_num_12 value='<%response.write web_var(web_num,12)%>'>
<input type=hidden name=web_num_13 value='<%response.write web_var(web_num,13)%>'>
<input type=hidden name=web_num_14 value='<%response.write web_var(web_num,14)%>'>
<input type=hidden name=web_num_15 value='<%response.write web_var(web_num,15)%>'>
<input type=hidden name=web_num_16 value='<%response.write web_var(web_num,16)%>'>
<%
end if

if action="menu" then
  tt=1
  tdim=split(web_menu,"|")
  for i=0 to ubound(tdim)
%>
<tr>
<td>��վ�˵� <%response.write i+1%>��</td>
<td colspan=2><input type=text name=web_menu_<%response.write i+1%> value='<%response.write tdim(i)%>' size=40 maxlength=20></td>
</tr>
<%
  next
  erase tdim
%>
<input type=hidden name=web_menu_num value='<%response.write i%>'>
<tr>
<td>�����²˵���</td>
<td colspan=2><input type=text name=web_menu_new value='' size=40 maxlength=20>&nbsp;����abc:����ѧϰ</td>
</tr>
<tr>
<td colspan=3 class=htd><font class=red>��վ�˵��޸�˵����</font>�������ӡ��޸ġ�ɾ����վ�˵������ı�����ֻ�����ԡ�abc:����ѧϰ������ʽ����Ŀ��:�˵��������ڣ����򽫳�������<br>
<font class=red_3>��Ӳ˵�</font>���ڡ������²˵����а��������Ҫ����Ĳ˵���ֻ��һ��һ�������ӣ�<br>
<font class=red_3>�޸Ĳ˵�</font>���޸����еģ���Ҫ����ֻҪ�������ݻ������ɣ�����Ŀ���Ͳ˵���Ҫͬʱ������<br>
<font class=red_3>ɾ���˵�</font>��һ��ɾ��һ������������Ϊ��Ҫɾ������Ŀ�˵������������գ��ٽ���Ŀ���ϴ�����С����ת�ơ�</td>
</tr>
<% else %>
<input type=hidden name=web_menu value='<%response.write web_menu%>'>
<%
end if

if action="color" then
  tt=1
%>
<tr>
<td>��վ������</td>
<td><input type=text name=web_color_1 value='<%response.write web_var(web_color,1)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,1)%>>&nbsp;</td>
</tr>
<tr>
<td>���ɫһ��</td>
<td><input type=text name=web_color_2 value='<%response.write web_var(web_color,2)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,2)%>>&nbsp;</td>
</tr>
<tr>
<td>���ɫ����</td>
<td><input type=text name=web_color_3 value='<%response.write web_var(web_color,3)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,3)%>>&nbsp;</td>
</tr>
<tr>
<td>���ɫ����</td>
<td><input type=text name=web_color_4 value='<%response.write web_var(web_color,4)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,4)%>>&nbsp;</td>
</tr>
<tr>
<td>���ɫ�ģ�</td>
<td><input type=text name=web_color_5 value='<%response.write web_var(web_color,5)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,5)%>>&nbsp;</td>
</tr>
<tr>
<td>��߱�����</td>
<td><input type=text name=web_color_6 value='<%response.write web_var(web_color,6)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,6)%>>&nbsp;</td>
</tr>
<tr>
<td>������ɫ��</td>
<td><input type=text name=web_color_7 value='<%response.write web_var(web_color,7)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,7)%>>&nbsp;</td>
</tr>
<tr>
<td>ͻ������һ��</td>
<td><input type=text name=web_color_8 value='<%response.write web_var(web_color,8)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,8)%>>&nbsp;</td>
</tr>
<tr>
<td>��ɫ���壺</td>
<td><input type=text name=web_color_9 value='<%response.write web_var(web_color,9)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,9)%>>&nbsp;</td>
</tr>
<tr>
<td>��ɫ����һ��</td>
<td><input type=text name=web_color_10 value='<%response.write web_var(web_color,10)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,10)%>>&nbsp;</td>
</tr>
<tr>
<td>��ɫ�������</td>
<td><input type=text name=web_color_11 value='<%response.write web_var(web_color,11)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,11)%>>&nbsp;</td>
</tr>
<tr>
<td>��ɫ��������</td>
<td><input type=text name=web_color_12 value='<%response.write web_var(web_color,12)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,12)%>>&nbsp;</td>
</tr>
<tr>
<td>��վ���壺</td>
<td colspan=2><input type=text name=web_font_family value='<%response.write web_font_family%>' size=60 maxlength=100></td>
</tr>
<tr>
<td>�����С��</td>
<td colspan=2><input type=text name=web_font_size value='<%response.write web_font_size%>' size=20 maxlength=10></td>
</tr>
<% else %>
<input type=hidden name=web_color_1 value='<%response.write web_var(web_color,1)%>'>
<input type=hidden name=web_color_2 value='<%response.write web_var(web_color,2)%>'>
<input type=hidden name=web_color_3 value='<%response.write web_var(web_color,3)%>'>
<input type=hidden name=web_color_4 value='<%response.write web_var(web_color,4)%>'>
<input type=hidden name=web_color_5 value='<%response.write web_var(web_color,5)%>'>
<input type=hidden name=web_color_6 value='<%response.write web_var(web_color,6)%>'>
<input type=hidden name=web_color_7 value='<%response.write web_var(web_color,7)%>'>
<input type=hidden name=web_color_8 value='<%response.write web_var(web_color,8)%>'>
<input type=hidden name=web_color_9 value='<%response.write web_var(web_color,9)%>'>
<input type=hidden name=web_color_10 value='<%response.write web_var(web_color,10)%>'>
<input type=hidden name=web_color_11 value='<%response.write web_var(web_color,11)%>'>
<input type=hidden name=web_color_12 value='<%response.write web_var(web_color,12)%>'>
<input type=hidden name=web_font_family value='<%response.write web_font_family%>'>
<input type=hidden name=web_font_size value='<%response.write web_font_size%>'>
<%
end if

if action="down_up" then
  tt=1
%>
<tr>
<td>���ͼƬ��</td>
<td><input type=text name=web_down_1 value='<%response.write web_var(web_down,1)%>' size=20 maxlength=10>&nbsp;����</td>
<td class=gray>&nbsp;���ͼƬ�Ŀ��</td>
</tr>
<tr>
<td>���ͼƬ�ߣ�</td>
<td><input type=text name=web_down_2 value='<%response.write web_var(web_down,2)%>' size=20 maxlength=10>&nbsp;����</td>
<td class=gray>&nbsp;���ͼƬ�ĸ߶�</td>
</tr>
<tr>
<td>���Ŀ¼��</td>
<td><input type=text name=web_down_5 value='<%response.write web_var(web_down,5)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>���л�����</td>
<td colspan=2><input type=text name=web_down_3 value='<%response.write web_var(web_down,3)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>��Ȩ���ͣ�</td>
<td colspan=2><input type=text name=web_down_4 value='<%response.write web_var(web_down,4)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>�ϴ�·����</td>
<td><input type=text name=web_upload_1 value='<% response.write web_var(web_upload,1) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;Ϊ���д�������Ҫ�޸Ĵ���</td>
</tr>
<tr>
<td>�ļ����ͣ�</td>
<td><input type=text name=web_upload_2 value='<% response.write web_var(web_upload,2) %>' size=35 maxlength=50></td>
<td class=gray>&nbsp;��������á�,���ֿ�</td>
</tr>
<tr>
<td>�ļ���С��</td>
<td><input type=text name=web_upload_3 value='<% response.write web_var(web_upload,3) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;��λΪKB�����С��500</td>
</tr>
<td>����Ŀ¼��</td>
<td><input type=text name=web_upload_3 value='<% response.write web_var(web_lct,1) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<td>���Ŀ¼��</td>
<td><input type=text name=web_upload_3 value='<% response.write web_var(web_lct,2) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<td>��ƵĿ¼��</td>
<td><input type=text name=web_upload_3 value='<% response.write web_var(web_lct,3) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<td>FlashĿ¼��</td>
<td><input type=text name=web_upload_3 value='<% response.write web_var(web_lct,4) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<td>��ֽĿ¼��</td>
<td><input type=text name=web_upload_3 value='<% response.write web_var(web_lct,5) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<% else %>
<input type=hidden name=web_down_1 value='<%response.write web_var(web_down,1)%>'>
<input type=hidden name=web_down_2 value='<%response.write web_var(web_down,2)%>'>
<input type=hidden name=web_down_5 value='<%response.write web_var(web_down,5)%>'>
<input type=hidden name=web_down_3 value='<%response.write web_var(web_down,3)%>'>
<input type=hidden name=web_down_4 value='<%response.write web_var(web_down,4)%>'>
<input type=hidden name=web_upload_1 value='<% response.write web_var(web_upload,1) %>'>
<input type=hidden name=web_upload_2 value='<% response.write web_var(web_upload,2) %>'>
<input type=hidden name=web_upload_3 value='<% response.write web_var(web_upload,3) %>'>
<%
end if

if action="info" then
  tt=1
%>
<tr><td colspan=3 class=red_3>&nbsp;�����ַ�����</td></tr>
<tr>
<td>�Ƿ��ַ���</td>
<td colspan=2><input type=text name=web_safety_1 value='<%response.write replace(web_var(web_safety,1),"'","")%>' size=30 maxlength=100>&nbsp;&nbsp;������(')��˫����(")�ѱ�ϵͳ����</td>
</tr>
<tr>
<td>��������</td>
<td colspan=2><input type=text name=web_safety_2 value='<%response.write web_var(web_safety,2)%>' size=66 maxlength=200></td>
</tr>
<tr>
<td>ע����ã�</td>
<td colspan=2><input type=text name=web_safety_3 value='<%response.write web_var(web_safety,3)%>' size=66 maxlength=200></td>
</tr>
<tr>
<td>�������ַ���</td>
<td colspan=2><input type=text name=web_safety_4 value='<%response.write web_var(web_safety,4)%>' size=66 maxlength=200></td>
</tr>
<tr><td colspan=3 class=red_3>&nbsp;��Ϣ��ʾ����</td></tr>
<tr>
<td>�ⲿ�ύ��</td>
<td colspan=2><input type=text name=web_error_1 value='<%response.write web_var(web_error,1)%>' size=66 maxlength=200></td>
</tr>
<tr>
<td>δע���½��</td>
<td colspan=2><input type=text name=web_error_2 value='<%response.write web_var(web_error,2)%>' size=66 maxlength=200></td>
</tr>
<tr>
<td>֧����Ϣ��</td>
<td colspan=2><input type=text name=web_error_3 value='<%response.write web_var(web_error,3)%>' size=66 maxlength=200></td>
</tr>
<tr>
<td>��վ�ײ���</td>
<td colspan=2><input type=text name=web_error_4 value='<%response.write web_var(web_error,4)%>' size=66 maxlength=200></td>
</tr>
<% else %>
<input type=hidden name=web_safety_1 value='<%response.write replace(web_var(web_safety,1),"'","")%>'>
<input type=hidden name=web_safety_2 value='<%response.write web_var(web_safety,2)%>'>
<input type=hidden name=web_safety_3 value='<%response.write web_var(web_safety,3)%>'>
<input type=hidden name=web_safety_4 value='<%response.write web_var(web_safety,4)%>'>
<input type=hidden name=web_error_1 value='<%response.write web_var(web_error,1)%>'>
<input type=hidden name=web_error_2 value='<%response.write web_var(web_error,2)%>'>
<input type=hidden name=web_error_3 value='<%response.write web_var(web_error,3)%>'>
<input type=hidden name=web_error_4 value='<%response.write web_var(web_error,4)%>'>
<%
end if

if action="user" then
  tt=1
  tdim=split(user_power,"|")
  for i=0 to ubound(tdim)
%>
<tr>
<td>�û��� <%response.write i+1%>��</td>
<td><input type=text name=user_power_<%response.write i+1%> value='<%response.write tdim(i)%>' size=30 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<%
  next
  erase tdim
%>
<input type=hidden name=user_power_num value='<%response.write i%>'>
<tr>
<td>���û��飺</td>
<td><input type=text name=user_power_new value='' size=30 maxlength=20></td>
<td class=gray>&nbsp;����huser:�߼��û�</td>
</tr>
<tr>
<td colspan=3 class=htd><font class=red>�û����޸�˵����</font>�������ӡ��޸ġ�ɾ���û��飬���ı�����ֻ�����ԡ�huser:�߼��û�������ʽ�����:�û����������ڣ����򽫳�������
<font class=red>�ڳ����У��û�����ԽС��Ȩ��Խ��ǰ�����û��������˽���޸ģ����޸����û��������ԣ�����վ��ʽ���к��������޸��û�����Է������û����᲻�ܱ�������ȷʵ��</font><br>
<font class=red_3>����û���</font>���ڡ��������û��顱�а��������Ҫ������û��飬ֻ�ܵ��������ӣ�<br>
<font class=red_3>�޸��û���</font>���޸����еģ���Ҫ����ֻҪ�������ݻ������ɣ����û������û�����Ҫͬʱ������<br>
<font class=red_3>ɾ���û���</font>��һ��ɾ��һ������������Ϊ��Ҫɾ�����û�������������գ��ٽ���Ŀ���ϴ�����С����ת�ơ�</td>
</tr>
<% else %>
<input type=hidden name=user_power value='<%response.write user_power%>'>
<%
end if

if action="grade" then
  tt=1
  tdim=split(user_grade,"|")
  for i=0 to ubound(tdim)
%>
<tr>
<td>�û��ȼ� <%response.write i%>��</td>
<td><input type=text name=user_grade_<%response.write i+1%> value='<%response.write tdim(i)%>' size=30 maxlength=20></td>
<td class=gray>&nbsp;<img src='images/star/star_<%response.write i%>.gif' border=0></td>
</tr>
<%
  next
  erase tdim
%>
<input type=hidden name=user_grade_num value='<%response.write i%>'>
<tr>
<td>���û��ȼ���</td>
<td><input type=text name=user_grade_new value='' size=30 maxlength=20></td>
<td class=gray>&nbsp;����10000:����</td>
</tr>
<tr>
<td colspan=3 class=htd><font class=red>�û��ȼ��޸�˵����</font>�������ӡ��޸ġ�ɾ���û��ȼ������ı�����ֻ�����ԡ�10000:����������ʽ���������:�ȼ����ƣ����ڣ����򽫳�������<br>
<font class=red_3>����û��ȼ�</font>���ڡ��������û��顱�а��������Ҫ������û��ȼ���ֻ�ܵ��������ӣ�<br>
<font class=red_3>�޸��û��ȼ�</font>���޸����еģ���Ҫ����ֻҪ�������ݻ������ɣ���������ֺ͵ȼ�����Ҫͬʱ������<br>
<font class=red_3>ɾ���û��ȼ�</font>��һ��ɾ��һ������������Ϊ��Ҫɾ�����û��ȼ������������գ��ٽ���Ŀ���ϴ�����С����ת�ơ�</td>
</tr>
<% else %>
<input type=hidden name=user_grade value='<%response.write user_grade%>'>
<%
end if

if action="forum" then
  tt=1
  udim=split(user_power,"|")
  tdim=split(forum_type,"|")
  for i=0 to ubound(tdim)
    t2=left(tdim(i),instr(tdim(i),":")-1)
%>
<tr>
<td>��̳���� <%response.write i+1%>��</td>
<td><input type=text name=forum_type_<%response.write i+1%>_2 value='<%response.write right(tdim(i),len(tdim(i))-instr(tdim(i),":"))%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr><td class=gray>��̳Ȩ��</td><td colspan=2><%
for j=0 to ubound(udim)
  response.write vbcrlf&"<input type=checkbox name=forum_type_"&i+1&"_1 value='"&j+1&"' class=bg_1"
  if instr(1,"."&t2&".","."&j+1&".")>0 then response.write " checked"
  response.write ">"&right(udim(j),len(udim(j))-instr(udim(j),":"))
next
%><input type=checkbox name=forum_type_<%response.write i+1%>_1 value='0' class=bg_1<%if instr(1,"."&t2&".",".0.")>0 then response.write " checked"%>>�ο�</td></tr><%
  next
  erase tdim
%>
<input type=hidden name=forum_type_num value='<%response.write i%>'>
<tr>
<td>����̳���ࣺ</td>
<td><input type=text name=forum_type_new_2 value='' size=30 maxlength=20></td>
<td class=gray>&nbsp;����������̳</td>
</tr>
<tr><td class=gray>��̳Ȩ��</td><td colspan=2><%
for j=0 to ubound(udim)
  response.write vbcrlf&"<input type=checkbox name=forum_type_new_1 value='"&j+1&"' class=bg_1>"&right(udim(j),len(udim(j))-instr(udim(j),":"))
next
%><input type=checkbox name=forum_type_new_1 value='0' class=bg_1>�ο�</td></tr>
<tr>
<td colspan=3 class=htd><font class=red>��̳�����޸�˵����</font>�������ӡ��޸ġ�ɾ����̳���࣬���ı�����ֻ�����ԡ�����ѧϰ������ʽ����̳�������ƣ����ڣ����򽫳�������
<font class=red>���û�����ɾ�����޸���Ȩ�޺������·�����̳Ȩ�ޣ�����̳���༰Ȩ�����޸ģ��������̳���������·�����̳���</font><br>
<font class=red_3>�����̳����</font>���ڡ�����̳���ࡱ�а��������Ҫ�������̳���ֻ࣬�ܵ��������ӣ�<br>
<font class=red_3>�޸���̳����</font>���޸����еģ��������̳��Ȩ��û��Ӱ�죻<br>
<font class=red_3>ɾ����̳����</font>��һ��ɾ��һ������������Ϊ��Ҫɾ������̳������������ɾ�����ɡ�</td>
</tr>
<%
  erase udim
else
%>
<input type=hidden name=forum_type value='<%response.write forum_type%>'>
<%
end if

if tt=0 then
%>
<tr><td colspan=3 align=center height=150 class=htd><font class=red>����ѡ�������޸ĵ����ͣ��������޸������ļ�ǰ��<a href='javascript:Do_b_data();'>��������</a>��<br>������޸�����ʱ�����˴�������������<a href='javascript:Do_h_data();'>��ԭ����</a>��<br>�����Ϊ����һʱ��С�ģ��������ļ��޸Ĵ����뾡����FTP���ϲ��ÿ��õ������ļ����Ǵ�����ļ���include/common.asp��</font></td></tr>
<% else %>
<tr><td colspan=3 align=center height=30><input type=submit value='�� �� �� ��'>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=reset value=' �� �� '>
</td></tr>
<% end if %>
</form>
</table>
<%
end sub

sub config_chk()
  dim t1,t2,tn
  web_font_family=code_config(request.form("web_font_family"),0):web_font_size=code_config(request.form("web_font_size"),0)
  web_config=code_config(request.form("web_config_1"),3) &"|"& code_config(request.form("web_config_2"),3) &"|"& _
	     code_config(request.form("web_config_3"),3) &"|"& code_config(request.form("web_config_4"),3) &"|"& _
	     code_config(request.form("web_config_5"),3) &"|"& code_config(request.form("web_config_6"),3) &"|"& _
	     code_config(request.form("web_config_7"),3) &"|"& code_config(request.form("web_config_8"),3) &"|"& _
	     trim(request.form("web_config_9_1")) & trim(request.form("web_config_9_2")) & _
	     trim(request.form("web_config_9_3")) & trim(request.form("web_config_9_4")) & _
	     trim(request.form("web_config_9_5"))
  web_cookies=code_config(request.form("web_cookies"),3):web_login=trim(request.form("web_login"))
  web_setup=trim(request.form("web_setup_1"))&trim(request.form("web_setup_2"))&trim(request.form("web_setup_3"))&_
	    trim(request.form("web_setup_4"))&trim(request.form("web_setup_5"))&trim(request.form("web_setup_6"))&_
	    trim(request.form("web_setup_7"))
  web_num=code_num(trim(request.form("web_num_1")),20) &"|"& code_num(trim(request.form("web_num_2")),20) &"|"& _
	  code_num(trim(request.form("web_num_3")),10) &"|"& code_num(trim(request.form("web_num_4")),5) &"|"& _
	  code_num(trim(request.form("web_num_5")),3) &"|"& code_num(trim(request.form("web_num_6")),12) &"|"& _
	  code_num(trim(request.form("web_num_7")),150) &"|"& code_num(trim(request.form("web_num_8")),112) &"|"& _
	  code_num(trim(request.form("web_num_9")),430) &"|"& code_num(trim(request.form("web_num_10")),350) &"|"& _
	  code_num(trim(request.form("web_num_11")),24) &"|"& code_num(trim(request.form("web_num_12")),16) &"|"& _
	  code_num(trim(request.form("web_num_13")),20) &"|"& code_num(trim(request.form("web_num_14")),20) &"|"& _
	  code_num(trim(request.form("web_num_15")),2) &"|"& code_num(trim(request.form("web_num_16")),15)
  if action="menu" then
    web_menu="":tn=int(trim(request.form("web_menu_num")))
    for i=1 to tn
      t1=code_config(request.form("web_menu_"&i),4)
      if len(t1)>1 then if instr(1,t1,":")>0 then web_menu=web_menu&t1&"|"
    next
    t1=code_config(request.form("web_menu_new"),4)
    if len(t1)>1 then if instr(1,t1,":")>0 then web_menu=web_menu&t1
    if right(web_menu,1)="|" then web_menu=left(web_menu,len(web_menu)-1)
  else
    web_menu=trim(request.form("web_menu"))
  end if
  web_color=code_config(request.form("web_color_1"),2)&"|"&code_config(request.form("web_color_2"),2)&"|"&code_config(request.form("web_color_3"),2)&"|"&_
	    code_config(request.form("web_color_4"),2)&"|"&code_config(request.form("web_color_5"),2)&"|"&code_config(request.form("web_color_6"),2)&"|"&_
	    code_config(request.form("web_color_7"),2)&"|"&code_config(request.form("web_color_8"),2)&"|"&code_config(request.form("web_color_9"),2)&"|"&_
	    code_config(request.form("web_color_10"),2)&"|"&code_config(request.form("web_color_11"),2)&"|"&code_config(request.form("web_color_12"),2)
  web_upload=code_config(request.form("web_upload_1"),2)&"|"&code_config(request.form("web_upload_2"),2)&"|"&code_config(request.form("web_upload_3"),2)
  web_safety=replace(trim(request.form("web_safety_1")),"""","""&chr(34)&""")&"'"&"|"&trim(request.form("web_safety_2"))&"|"&trim(request.form("web_safety_3"))&"|"&trim(request.form("web_safety_4"))
  web_error=code_config(request.form("web_error_1"),3)&"|"&code_config(request.form("web_error_2"),3)&"|"&code_config(request.form("web_error_3"),3)&"|"&code_config(request.form("web_error_4"),3)
  web_news_art=trim(request.form("web_news_art"))
  web_down=code_num(trim(request.form("web_down_1")),95)&"|"&code_num(trim(request.form("web_down_2")),75)&"|"&_
	   code_config(request.form("web_down_3"),4)&"|"&code_config(request.form("web_down_4"),4)&"|"&code_config(request.form("web_down_5"),4)
  web_shop=trim(request.form("web_shop"))
  web_stamp=code_config(request.form("web_stamp_1"),3)&"|"&code_config(request.form("web_stamp_2"),3)&"|"&code_config(request.form("web_stamp_3"),3)
  if action="user" then
    user_power="":tn=int(trim(request.form("user_power_num")))
    for i=1 to tn
      t1=code_config(request.form("user_power_"&i),4)
      if len(t1)>1 then if instr(1,t1,":")>0 then user_power=user_power&t1&"|"
    next
    t1=code_config(request.form("user_power_new"),4)
    if len(t1)>1 then if instr(1,t1,":")>0 then user_power=user_power&t1
    if right(user_power,1)="|" then user_power=left(user_power,len(user_power)-1)
  else
    user_power=trim(request.form("user_power"))
  end if
  if action="grade" then
    user_grade="":tn=int(trim(request.form("user_grade_num")))
    for i=1 to tn
      t1=code_config(request.form("user_grade_"&i),4)
      if len(t1)>1 then if instr(1,t1,":")>0 then user_grade=user_grade&t1&"|"
    next
    t1=code_config(request.form("user_grade_new"),4)
    if len(t1)>1 then if instr(1,t1,":")>0 then user_grade=user_grade&t1
    if right(user_grade,1)="|" then user_grade=left(user_grade,len(user_grade)-1)
  else
    user_grade=trim(request.form("user_grade"))
  end if
  if action="forum" then
    forum_type="":tn=int(trim(request.form("forum_type_num")))
    for i=1 to tn
      t1=replace(trim(request.form("forum_type_"&i&"_1"))," ","")
      t2=code_config(request.form("forum_type_"&i&"_2"),2)
      if len(t1)>0 and len(t2)>0 then t1=replace(t1,",","."):forum_type=forum_type&t1&":"&t2&"|"
    next
    t1=replace(trim(request.form("forum_type_new_1"))," ","")
    t2=code_config(request.form("forum_type_new_2"),2)
    if len(t1)>0 and len(t2)>0 then t1=replace(t1,",","."):forum_type=forum_type&t1&":"&t2
    if right(forum_type,1)="|" then forum_type=left(forum_type,len(forum_type)-1)
  else
    forum_type=trim(request.form("forum_type"))
  end if
  
  call config_file()
  if action="color" then call config_css():call config_mouse_on_title()
  response.write "<script language=javascript>alert(""�����޸ĳɹ���"");</script>"
  call config_main()
end sub

sub config_bh(bht)
  dim vv,filetype,file_name1,file_name2,filetemp,fileos,filepath:filetype=""
  if bht="h" then
    file_name1="include/back_common.asp"
    file_name2="include/common.asp"
    vv="��ԭ"
  else
    file_name1="include/common.asp"
    file_name2="include/back_common.asp"
    vv="����"
  end if
  
  set fileos=CreateObject("Scripting.FileSystemObject")
  filepath=server.mappath(file_name1)
  set filetemp=fileos.OpenTextFile(filepath,1,true)
  filetype=filetemp.ReadAll
  filetemp.close
  filepath=server.mappath(file_name2)
  set filetemp=fileos.createtextfile(filepath,true)
  filetemp.writeline( filetype )
  filetemp.close
  set filetemp=nothing
  set fileos=nothing
  response.write "<script language=javascript>alert("""&vv&" ��վ���óɹ���"");</script>"
end sub

function code_num(strers,cnum)
  dim strer:strer=trim(strers)
  if not(isnumeric(strer)) then strer=cnum
  if int(strer)<1 then strer=cnum
  code_num=strer
end function

function code_config(strers,ct)
  dim strer:strer=trim(strers)
  if isnull(strer) or strer="" then code_config="":exit function
  select case ct
  case 1
    strer=replace(strer,"'","&#39;")
    strer=replace(strer,CHR(34),"&quot;")
  case 2
    strer=replace(strer,CHR(34),"")
    strer=replace(strer,"'","")
    strer=replace(strer,":","")
    strer=replace(strer,"|","")
  case 3
    strer=replace(strer,"'","&#39;")
    strer=replace(strer,CHR(34),"&quot;")
    strer=replace(strer,"|","")
  case 4
    strer=replace(strer,CHR(34),"")
    strer=replace(strer,"'","")
    strer=replace(strer,"|","")
  case else
    strer=replace(strer,"'","")
    strer=replace(strer,CHR(34),"")
  end select
  code_config=strer
end function
%>