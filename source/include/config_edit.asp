<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

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
<td>��վ״̬��</td>
<td><input type=radio name=web_login value='1'<% if int(web_login)=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_login value='0'<% if int(web_login)<>1 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;�Ƿ񿪷���վ</td>
</tr>
<% t1=int(mid(web_setup,1,1)) %>
<tr>
<td>��½�����</td>
<td colspan=2><input type=radio name=web_setup_1 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;��&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_1 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;<font class=gray>�Ƿ�Ҫ��½�ſ�������»����������</font></td>
</tr>
<%
t1=int(mid(web_setup,3,1))
if t1<>1 and t1<>2 then t1=0
%>
<tr>
<td>��վģʽ��</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='1'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_1 value='<%response.write web_var(web_stamp,1)%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;��վ��ע���û����Ե�½������¼�����б�</td>
<tr>
<td>&nbsp;</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_2 value='<%response.write web_var(web_stamp,2)%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;���е�½�������վ���˱�����¼�����б�</td>
<tr>
<td>&nbsp;</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='2'<% if t1=2 then response.write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_3 value='<%response.write web_var(web_stamp,3)%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;��վ��ע���û����Ե�½������¼�û������б�</td>
</tr>
<% t1=int(mid(web_setup,4,1)) %>
<tr>
<td>��Ϣ���ˣ�</td>
<td><input type=radio name=web_setup_4 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;��&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_4 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;��</td>
<td class=gray>&nbsp;�Ƿ����̳�Ƚ����ַ�����</td>
</tr>
<%
t1=int(mid(web_setup,5,1))
if t1<>0 and t1<>1 then t1=2
%>
<tr>
<td>��ʾ IP��</td>
<td><input type=radio name=web_setup_5 value='0'<% if t1=0 then response.write " checked" %> class=bg_1> ��ȫ����
<input type=radio name=web_setup_5 value='1'<% if t1=1 then response.write " checked" %> class=bg_1> ��ʾ����
<input type=radio name=web_setup_5 value='2'<% if t1=2 then response.write " checked" %> class=bg_1> ��ȫ����</td>
<td class=gray>&nbsp;�Թ���Ա������ȫ����</td>
</tr>
<% t1=int(mid(web_setup,6,1)) %>
<tr>
<td>������ʾ��</td>
<td><input type=radio name=web_setup_6 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;��������&nbsp;<input type=radio name=web_setup_6 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;�����˵���</td>
<td class=gray>&nbsp;��̳������ʾģʽ</td>
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
<input type=hidden name=web_login value='<%response.write web_login%>'>
<input type=hidden name=web_setup_1 value='<%response.write mid(web_setup,1,1)%>'>
<input type=hidden name=web_setup_3 value='<%response.write mid(web_setup,3,1)%>'>
<input type=hidden name=web_stamp_1 value='<%response.write web_var(web_stamp,1)%>'>
<input type=hidden name=web_stamp_2 value='<%response.write web_var(web_stamp,2)%>'>
<input type=hidden name=web_stamp_3 value='<%response.write web_var(web_stamp,3)%>'>
<input type=hidden name=web_setup_4 value='<%response.write mid(web_setup,4,1)%>'>
<input type=hidden name=web_setup_5 value='<%response.write mid(web_setup,5,1)%>'>
<input type=hidden name=web_setup_6 value='<%response.write mid(web_setup,6,1)%>'>
<%
end if

if action="put" then
  tt=1
  t1=int(mid(web_var(web_config,9),1,1))
%><tr>
<td>�������ţ�</td>
<td><input type=radio name=web_config_9_1 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_1 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr><% t1=int(mid(web_var(web_config,9),2,1)) %>
<tr>
<td>�������£�</td>
<td><input type=radio name=web_config_9_2 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_2 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr><% t1=int(mid(web_var(web_config,9),3,1)) %>
<tr>
<td>������֣�</td>
<td><input type=radio name=web_config_9_3 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_3 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr><% t1=int(mid(web_var(web_config,9),4,1)) %>
<tr>
<td>�ϴ���ͼ��</td>
<td><input type=radio name=web_config_9_4 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_4 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr><% t1=int(mid(web_var(web_config,9),5,1)) %>
<tr>
<td>�Ƽ���վ��</td>
<td><input type=radio name=web_config_9_5 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;����&nbsp;<input type=radio name=web_config_9_5 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;�ر�</td>
<td class=gray>&nbsp;</td>
</tr>
<% else %>
<input type=hidden name=web_config_9_1 value='<%response.write mid(web_var(web_config,9),1,1)%>'>
<input type=hidden name=web_config_9_2 value='<%response.write mid(web_var(web_config,9),2,1)%>'>
<input type=hidden name=web_config_9_3 value='<%response.write mid(web_var(web_config,9),3,1)%>'>
<input type=hidden name=web_config_9_4 value='<%response.write mid(web_var(web_config,9),4,1)%>'>
<input type=hidden name=web_config_9_5 value='<%response.write mid(web_var(web_config,9),5,1)%>'>
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
<%
end if

if action="upload" then
  tt=1
%>
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
<% else %>
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
<td colspan=2><input type=text name=web_safety_2 value='<%response.write web_var(web_safety,2)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>ע����ã�</td>
<td colspan=2><input type=text name=web_safety_3 value='<%response.write web_var(web_safety,3)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>�������ַ���</td>
<td colspan=2><input type=text name=web_safety_4 value='<%response.write web_var(web_safety,4)%>' size=66 maxlength=100></td>
</tr>
<tr><td colspan=3 class=red_3>&nbsp;��Ϣ��ʾ����</td></tr>
<tr>
<td>�ⲿ�ύ��</td>
<td colspan=2><input type=text name=web_error_1 value='<%response.write web_var(web_error,1)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>δע���½��</td>
<td colspan=2><input type=text name=web_error_2 value='<%response.write web_var(web_error,2)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>֧����Ϣ��</td>
<td colspan=2><input type=text name=web_error_3 value='<%response.write web_var(web_error,3)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>��վ�ײ���</td>
<td colspan=2><input type=text name=web_error_4 value='<%response.write web_var(web_error,4)%>' size=66 maxlength=100></td>
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

if action="down" then
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
<td>��Ȩ���ͣ�</td>
<td colspan=2><input type=text name=web_down_3 value='<%response.write web_var(web_down,3)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>���л�����</td>
<td colspan=2><input type=text name=web_down_4 value='<%response.write web_var(web_down,4)%>' size=66 maxlength=100></td>
</tr>
<% else %>
<input type=hidden name=web_down_1 value='<%response.write web_var(web_down,1)%>'>
<input type=hidden name=web_down_2 value='<%response.write web_var(web_down,2)%>'>
<input type=hidden name=web_down_5 value='<%response.write web_var(web_down,5)%>'>
<input type=hidden name=web_down_3 value='<%response.write web_var(web_down,3)%>'>
<input type=hidden name=web_down_4 value='<%response.write web_var(web_down,4)%>'>
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
%>