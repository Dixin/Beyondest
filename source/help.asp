<!-- #include file="INCLUDE/config_other.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

index_url="user_main"
tit="��������"
tit_fir=format_menu(index_url)

call web_head(0,1,0,0,0)
'------------------------------------left----------------------------------
call format_login()
call help_left("jt12")
response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong

select case action
case "about"
  call help_about()
case "register"
  call help_register()
case "put"
  call help_put()
case "mail"
  call help_mail()
case "forum"
  call help_forum()
case "ubb"
  call help_ubb()
case else
  call help_about()
  'call help_main()
end select

response.write kong
'---------------------------------center end-------------------------------
call web_end(1)

sub help_left(ljt)
  if ljt<>"" then ljt=img_small(ljt)
  tit=vbcrlf&"<table border=0 width='96%' cellpadding=0 cellspacing=6 align=center>" & _
  vbcrlf&"<tr><td width='50%'></td><td width='50%'></td></tr>" & _
  vbcrlf&"<tr><td>"&ljt&"<a href='?action=about'>��������</a></td><td>"&ljt&"<a href='?action=register'>ע��˵��</a></td></tr>" & _
  vbcrlf&"<tr><td>"&ljt&"<a href='?action=put'>������Ϣ</a></td><td>"&ljt&"<a href='?action=mail'>վ�ڶ���</a></td></tr>" & _
  vbcrlf&"<tr><td>"&ljt&"<a href='?action=forum'>��̳����</a></td><td>"&ljt&"<a href='?action=ubb'>UBB�﷨</a></td></tr>" & _
  vbcrlf&"</table>"
  call left_type(tit,"help",1)
end sub

sub help_main()
  response.write table1
%>
<tr<%response.write table2%>><td>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>��������</b></font></td></tr>
<tr<%response.write table3%>><td class=htd></td></tr>
<tr<%response.write table3%>><td align=center>

</td></tr>
</table>
<%
end sub

sub help_register()
  response.write table1
%>
<tr<%response.write table2%>><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>ע��˵��</b></font></td></tr>
<tr<%response.write table3%>><td align=center>
  <br>
  <table border=0 width='94%'>
  <tr><td class=htd>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ӭ�����뱾վ��<a href='<%=web_var(web_config,2)%>'><%=web_var(web_config,1)%></a>���μӽ��������ۣ�<a href='<%=web_var(web_config,2)%>'><%=web_var(web_config,1)%></a>Ϊ��ȫ��Ӯ���ԡ���ҵ�Ե���վ��<font color="#FF0000">���ǵ�Ŀ�����ƹ�Beyond�����֣�����Beyond�ľ����о���صļ������������⡣���ǳ�ŵ�Ը��õ�Ϊ�����������ṩ���ַ���ͷ���Ϊ��ּ��������ǵ����з�������ѵģ����Ǿ������κ��������û���ȡ�κη��ã��������û������κ���Ʒ��</font><br><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ϊά�����Ϲ������������Ծ�������������<br><br>
��һ���������ñ�վΣ�����Ұ�ȫ��й¶�������ܣ������ַ����ҡ���ᡢ����͹���ĺϷ�Ȩ�棬�������ñ�վ���������ƺʹ���������Ϣ�� <br>
������һ��ɿ�����ܡ��ƻ��ܷ��ͷ���ʵʩ�ģ�<br>
����������ɿ���߸�������Ȩ���Ʒ���������ƶȵģ�<br>
����������ɿ�����ѹ��ҡ��ƻ�����ͳһ�ģ�<br>
�������ģ�ɿ�������ޡ��������ӣ��ƻ������Ž�ģ�<br>
�������壩�������������ʵ��ɢ��ҥ�ԣ������������ģ�<br>
��������������⽨���š����ࡢɫ�顢�Ĳ�����������ɱ���ֲ�����������ģ�<br>
�������ߣ���Ȼ�������˻���������ʵ�̰����˵ģ����߽����������⹥���ģ�<br>
�������ˣ��𺦱�վ�����ģ�<br>
�������ţ�����ʹ�����Ի���ģ�<br>
������ʮ��������ҵ���ʵ���Ϊ�ġ�<br>
�������������أ����Լ������ۺ���Ϊ����<br>
�������������ǵ��Ͷ��ɹ���<br>
������һ��ת�ر�վ������ע��������<br>
�������������𽫱�վ�ṩ������������ҵ��;��<br>

  </td></tr>
  <form name=form_reg action='login.asp?action=register' method=post>
  <input type=hidden name=reg_action value='reg_main'>
  <tr><td align=center height=30><input type=submit value='�����Ķ���ͬ����������'></td></tr>
  </form>
  </table>
  <br>
</td></tr>
</table>
<%
end sub

sub help_put()
  response.write table1
%>
<tr<%response.write table2%> height=25><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>������Ϣ</b></font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
  <tr valign=top>
  <td width='30%'>
    <table border=0>
    <tr><td class=htd>�����������ڱ�վ����һЩ�����»������ͼ����Ϣ��ͨ������Ա���֮��Ϳ�������վ����ʾ�����ʹ�ҷ����ˣ����ǻ�ӭ����л��Ϊ��վ�ṩ�������ϡ�<br>�������ڷ�����Ϣǰ��<a href='login.asp?action=register'>ע��</a>����<a href='login.asp'>��½</a>��վ�����������������Ľ�����Ϣ������</td></tr>
    </table>
  </td>
  <td width='3%'></td>
  <td width='67%'>
    <table border=0>
    <tr><td height=1 width='5%'></td><td width='95%'></td></tr>
    <tr><td colspan=2 height=20><%response.write img_small("jt1")%><a href='user_put.asp?action=news'>�����ҵ�����</a></td></tr>
    <tr><td></td><td class=htd>��������Beyond�����ţ����������У����⡢���ݡ��������ؼ��֡�ͼƬ�����ϴ����ȣ������Ա��ˡ�</td></tr>
    <tr><td height=5></td></tr>
    <tr><td colspan=2 height=20><%response.write img_small("jt1")%><a href='user_put.asp?action=article'>�����ҵ�����</a></td></tr>
    <tr><td></td><td class=htd>��������Beyond�����£����������У����⡢���ͺ����ݵȣ������Ա��ˡ�</td></tr>
    <tr><td height=5></td></tr>
    <tr><td colspan=2 height=20><%response.write img_small("jt1")%><a href='user_put.asp?action=down'>�����ҵ�����</a></td></tr>
    <tr><td></td><td class=htd>�ʹ�ҷ���Beyond�ľ��ʣ����������У����ơ����͡���С���Ƽ��ȼ������������˵�����ؼ��֡�ͼƬ�����ϴ����ȣ������Ա��ˡ�</td></tr>
    <tr><td height=5></td></tr>
    <tr><td colspan=2 height=20><%response.write img_small("jt1")%><a href='user_put.asp?action=gallery'>�ϴ��ҵ�ͼƬ</a></td></tr>
    <tr><td></td><td class=htd>�ϴ�Beyond��ͼƬ��FLASH�����������У����ơ����͡�˵����ͼƬ�����ϴ����ȣ������Ա��ˡ�</td></tr>
    <tr><td height=5></td></tr>
    <tr><td colspan=2 height=20><%response.write img_small("jt1")%><a href='user_put.asp?action=website'>��Ҫ�Ƽ���վ</a></td></tr>
    <tr><td></td><td class=htd>�Ƽ���ص���վ����ҳ�����������У����ơ����͡���ַ�����ҵ�����վ�����ԡ�˵����ͼƬ�����ϴ����ȣ������Ա��ˡ�</td></tr>
    </table>
  </td></tr>
  </table>
</td></tr>
</table>
<%
end sub

sub help_mail()
  response.write table1
%>
<tr<%response.write table2%>><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>վ�ڶ���</b></font></td></tr>
<tr<%response.write table3%>><td class=htd align=center height=30>վ�ڶ��ſ�ʹ�����磬��ȫ���շ�˽����Ϣ�����ᱻ���˼�����鿴����</td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='80%'>
  <tr><td width='5%'></td><td width='95%'></td></tr>
  <tr><td colspan=2 class=blue><%response.write img_small("jt0")%>����վ�ڶ���</td></tr>
  <tr><td></td><td class=htd>���������<a href='user_main.asp'>���û����ġ�</a>�����
  վ�ڶ��š����������е����������Ϣ����ť�������ռ��˵����ֺ���Ϣ���⡣������֧������ͼ�ͣ��������ͼ�ʹ��뽫���Զ�ת��Ϊ��ӦͼƬ��ע�⣬�ڰ����ͼ�ǰ��ȷ������д�����е���Ŀ��</td></tr>
  <tr><td colspan=2 class=blue><%response.write img_small("jt0")%>�ռ���</td></tr>
  <tr><td></td><td class=htd>�����ռ����д�����з������˽����Ϣ���������Ķ�����ɾ�����ǡ�</td></tr>
  <tr><td colspan=2 class=blue><%response.write img_small("jt0")%>������</td></tr>
  <tr><td></td><td class=htd>���������������������͹���ȫ����Ϣ��¼����ʹ�������˭���͹�ʲô��Ϣ�������Ķ��⣬��ɾ�����ǣ�</td></tr>
  <tr><td colspan=2 class=blue><%response.write img_small("jt0")%>�ر�����</td></tr>
  <tr><td></td><td class=htd>�벻Ҫ�ô���ʹ�������Ļ���ʹ�˲�������Ϣ���������ˣ�Ҳ�������Լ���</td></tr>
  </table>
</td></tr>
</table>
<%
end sub

sub help_forum()
  response.write table1
%>
<tr<%response.write table2%>><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>��̳����</b></font></td></tr>
<tr<%response.write table4%>><td class=htd align=center bgcolor=<%=web_var(web_color,6)%>><font class=red_3>������ע�ᡢ������������ɾ�����ӵȲ������û���ֵ��Ӱ������˵����ʾ��</font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='95%'>
  <tr><td colspan=2 class=btd>&nbsp;<font class=blue>��һ������</font></td></tr>
  <tr><td width='5%'></td><td width='95%'>ע���ʼ���ӣ�<font class=red>0</font>&nbsp;&nbsp;�����������ӣ�<font class=red>1</font>&nbsp;�����������ӣ�<font class=red>1</font>&nbsp;&nbsp;ɾ���������ӣ�<font class=red>1</font></td></tr>
  </table>
</td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='95%'>
  <tr><td colspan=2 class=btd>&nbsp;<font class=blue>����������</font></td></tr>
  <tr><td width='5%'></td><td width='95%'>ע���ʼ���֣�<font class=red>0</font>&nbsp;&nbsp;�������ӻ��֣�<font class=red>2</font>&nbsp;&nbsp;�������ӻ��֣�<font class=red>1</font>&nbsp;&nbsp;ɾ�����ٻ��֣�����&nbsp;<font class=red>3</font>&nbsp;&nbsp;����&nbsp;<font class=red>2</font></td></tr>
  </table>
</td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='95%'>
  <tr><td colspan=2 class=btd>&nbsp;<font class=blue>��������Ǯ</font></td></tr>
  <tr><td width='5%'></td><td width='95%'>ע���ʼ��Ǯ��<font class=red>0</font>&nbsp;&nbsp;&nbsp;&nbsp;<font class=gray>�����������</font></td></tr>
  </table>
</td></tr>
<tr<%response.write table4%>><td class=htd align=center bgcolor=<%=web_var(web_color,6)%>><font class=red_3>�������û����֣��ȼ���ͼ��ѡ������˵����ʾ��</font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0>
  <tr align=center>
  <td height=30 width=100><%response.write img_small("icon_admin")%>����Ա</td>
  <td width=100><img src='IMAGES/STAR/star_admin.gif'></td>
  <td width=100><%response.write img_small("icon_super")%>��̳����</td>
  <td width=100><img src='IMAGES/STAR/star_super.gif'></td>
  </tr>
  </table>
  <table border=0>
  <tr align=center>
  <td height=30 width=50 align=left>����</td>
  <td><%response.write img_small("icon_user")%>��ͨ/<%response.write img_small("icon_puser")%>��Ա�û�</td>
  <td><%response.write img_small("icon_vip")%>VIP�û�</td>
  <td width=80>�ȼ�����</td>
  <td>�������</td>
  </tr>
<%
dim sdim,sn,su:su=0
sdim=split(user_grade,"|")
for sn=0 to ubound(sdim)
%>
  <tr>
  <td><%response.write sn%>��</td>
  <td><img src='images/star/star_<%response.write sn%>.gif'></td>
  <td><img src='images/star/star_p<%response.write sn%>.gif'></td>
  <td align=center><%response.write right(sdim(sn),len(sdim(sn))-instr(sdim(sn),":"))%></td>
  <td><%
if sn=int(ubound(sdim)) then
  response.write left(sdim(sn),instr(sdim(sn),":")-1)&"������"
else
  response.write left(sdim(sn),instr(sdim(sn),":")-1)&"-"&(left(sdim(sn+1),instr(sdim(sn+1),":")-1)-1)
end if
%></td>
  </tr>
<%
next
erase sdim
%>
  </table>
</td></tr>
<tr<%response.write table4%>><td class=htd align=center bgcolor=<%=web_var(web_color,6)%>><font class=red_3>�������û������͸���ǩ�������벻����ѡ������˵����ʾ��</font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='90%'>
  <tr><td class=htd><%response.write web_var(web_error,3)&"<br>С��"&web_var(web_num,6)&"KB"%></td></tr>
  </table>
</td></tr>
</table>
<% response.write kong&table1 %>
<tr<%response.write table2%>><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>��̳ͼ��</b></font></td></tr>
<tr<%response.write table3%>><td align=center height=30><% response.write ip_sys(0,0) %></td></tr>
<tr<%response.write table3%>><td align=center height=30><%response.write user_power_type(0)%></td></tr>
<tr<%response.write table3%>><td align=center height=30>
<%response.write img_small("isok")%>&nbsp;���ŵ�����&nbsp;&nbsp;
<%response.write img_small("ishot")%>&nbsp;�ظ�����10��&nbsp;&nbsp;
<%response.write img_small("islock")%>&nbsp;����������&nbsp;&nbsp;
<%response.write img_small("istop")%>&nbsp;�̶����˵�����&nbsp;&nbsp;
<%response.write img_small("isgood")%>&nbsp;��������
</td></tr>
</table>
<%
end sub

sub help_ubb()
  response.write table1
%>
<tr<%response.write table2%>><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>UBB�﷨</b></font></td></tr>
<tr<%response.write table3%>><td class=htd>��������Ϊ��վʹ�õ�UBB�﷨�ľ���ʹ��˵������Ϊ��Ҫ��������һЩ�Ľ���UBB��ǩ���ǲ�����ʹ��HTML�﷨������£�ͨ������ת��������������֧���������õġ���Σ���Ե�HTMLЧ����ʾ��</td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
  <tr><td class=htd>
<li><font color=red>[B]</font><B>����</B><font color=red>[/B]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ����Ч����</li>
<li><font color=red>[I]</font><I>����</I><font color=red>[/I]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪб��Ч����</li>
<li><font color=red>[U]</font><U>����</U><font color=red>[/U]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ�»���Ч����</li>
<li><font color=red>[ALIGN=center]</font>����<font color=red>[/ALIGN]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ���centerλ��center��ʾ���У�left��ʾ����right��ʾ���ҡ�</li>
<li><font color=red>[COLOR=��ɫ����]</font>����<font color=red>[/COLOR]</font>������������ɫ���룬�ڱ�ǩ���м�������ֿ���ʵ��������ɫ�ı䡣</li>
<li><font color=red>[SIZE=����]</font>����<font color=red>[/SIZE]</font>���������������С���ڱ�ǩ���м�������ֿ���ʵ�����ִ�С�ı䡣</li>
<li><font color=red>[FACE=����]</font>����<font color=red>[/FACE]</font>����������Ҫ�����壬�ڱ�ǩ���м�������ֿ���ʵ����������ת����</li>
<li><font color=red>[FLY]</font>���������<font color=red>[/FLY]</font>���ڱ�ǩ���м�������ֿ���ʵ�����ַ���Ч������������ơ�</li>
<li><font color=red>[MOVE]</font>�ƶ�������<font color=red>[/MOVE]</font>���ڱ�ǩ���м�������ֿ���ʵ�������ƶ�Ч����Ϊ����Ʈ����</li>
<li><font color=red>[GLOW=255,red,2]</font>����<font color=red>[/GLOW]</font>���ڱ�ǩ���м�������ֿ���ʵ�����ַ�����Ч��glow����������Ϊ��ȡ���ɫ�ͱ߽��С��</li>
<li><font color=red>[SHADOW=255,red,2]</font>����<font color=red>[/SHADOW]</font>���ڱ�ǩ���м�������ֿ���ʵ��������Ӱ��Ч��shadow����������Ϊ��ȡ���ɫ�ͱ߽��С��</li>
<li><font color=red>[URL]</font><A href="<%response.write web_var(web_config,2)%>"><%response.write web_var(web_config,2)%></A><font color=red>[/URL]</font></li>
<li><font color=red>[URL=<%response.write web_var(web_config,2)%>]</font><A href="<%response.write web_var(web_config,2)%>"><%response.write web_var(web_config,1)%></A><font color=red>[/URL]</font>�������ַ������Լ��볬�����ӣ��������Ӿ����ַ�����������ӡ�</li>
<li><font color=red>[EMAIL]</font><A href="mailto:dixinyan@live.com">dixinyan@live.com</A><font color=red>[/EMAIL]</font></li>
<li><font color=red>[EMAIL=dixinyan@live.com]</font><A href="mailto:dixinyan@live.com">����</A><font color=red>[/EMAIL]</font>�������ַ������Լ����ʼ����ӣ��������Ӿ����ַ�����������ӡ�</li>
<li><font color=red>[IMG]images/logo.gif[/IMG]</font> ���ڱ�ǩ���м����ͼƬ��ַ����ʵ�ֲ�ͼЧ����
<li><font color=red>[DOWNLOAD]http://beyondest.com/music/test.rar[/DOWNLOAD]</font>���ڱ�ǩ���м�����ṩ���ص��ļ���ַ����ʵ���ļ�����Ч����

<li><font color=red>[FLASH=���,�߶�]</font>Flash���ӵ�ַ<font color=red>[/FLASH]</font>���ڱ�ǩ���м����FlashͼƬ��ַ����ʵ�ֲ���Flash��</li>
<li><font color=red>[CODE]</font>����<font color=red>[/CODE]</font>���ڱ�ǩ��д�����ֿ�ʵ��html�б��Ч����</li>
<li><font color=red>[OTE]</font>����<font color=red>[/QUOTE]</font>���ڱ�ǩ���м�������ֿ���ʵ��HTMl����������Ч����</li>
<li><font color=red>[RM=���,�߶�]</font>http://<font color=red>[/RM]</font>��Ϊ����realplayer��ʽ��rm�ļ����м������Ϊ��Ⱥͳ��ȡ�</li>
<li><font color=red>[MP=���,�߶�]</font>http://<font color=red>[/MP]</font>��Ϊ����Ϊmidia player��ʽ���ļ����м������Ϊ��Ⱥͳ��ȡ�</li>
<li><font color=red>[DIR=���,�߶�]</font>http://<font color=red>[/DIR]</font>��Ϊ����shockwave��ʽ�ļ����м������Ϊ��Ⱥͳ��ȡ�</li>
<li><font color=red>[QT=500,350]</font>http://<font color=red>[/QT]</font>��Ϊ����ΪQuick time��ʽ���ļ����м������Ϊ��Ⱥͳ��ȡ�</li>
  </td></tr>
  </table>
</td></tr>
<tr<%response.write table2%>><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>EM ��ͼ</b></font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0>
  <tr><td width=80></td><td></td></tr>
  <tr align=center>
  <td>С EM ��ͼ<br>��1-8��</td>
  <td>
    <table border=0>
    <tr align=center><%
for i=1 to 8
  response.write vbcrlf&"    <td width=50><img src='images/icon/em"&i&".gif' border=0></td>"
next
%></tr>
    <tr align=center><%
for i=1 to 8
  response.write vbcrlf&"    <td>[em"&i&"]</td>"
next
%></tr>
    </table>
  </td>
  </tr>
  <tr><td colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr align=center>
  <td>С EM ��ͼ<br>��9-16��</td>
  <td>
    <table border=0>
    <tr align=center><%
for i=9 to 16
  response.write vbcrlf&"    <td width=50><img src='images/icon/em"&i&".gif' border=0></td>"
next
%></tr>
    <tr align=center><%
for i=9 to 16
  response.write vbcrlf&"    <td>[em"&i&"]</td>"
next
%></tr>
    </table>
  </td>
  </tr>
  <tr><td colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr align=center>
  <td>�� EM ��ͼ<br>��1-7��</td>
  <td>
    <table border=0>
    <tr><%
for i=1 to 7
  response.write vbcrlf&"    <td width=60><img src='images/icon/emb"&i&".gif' border=0></td>"
next
%></tr>
    <tr><%
for i=1 to 7
  response.write vbcrlf&"    <td>[emb"&i&"]</td>"
next
%></tr>
    </table>
  </td>
  </tr>
  <tr><td colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr>
  <td align=center>�� EM ��ͼ<br>��8-13��</td>
  <td>
    <table border=0>
    <tr><%
for i=8 to 13
  response.write vbcrlf&"    <td width=60><img src='images/icon/emb"&i&".gif' border=0></td>"
next
%></tr>
    <tr><%
for i=8 to 13
  response.write vbcrlf&"    <td>[emb"&i&"]</td>"
next
%></tr>
    </table>
  </td>
  </tr>
  </table>
</td></tr>
</table>
<%
end sub

sub help_about()
  response.write table1
%>
<tr<%response.write table2%>><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>��������</b></font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
  <tr>
  <td width='100%' class=htd>������վ��<a href='<%=web_var(web_config,2)%>'><%=web_var(web_config,1)%></a>����ʽ������1999��12�£������ҹ��Ƚ����һ��Beyond��վ����ʱ����İ汾Ϊ����̬��html�������ص�������Ӿ�����ϣ��������νϴ��ģ��������Ľ������ǵ�Խ��Խע�ؼ�������ߡ����ھ���������ѧҵ��ԭ�򣬴�2000���°�����վ�ķ�չһֱ�����ں�С�ķ�Χ�ڣ������շǳ���ǿ��ά�ֵ���2003�ꡣ���Ǿ���4��������ռ��ͼ�����ߣ����ǵĲ�иŬ���ͻ�������ʹ<a href='<%=web_var(web_config,2)%>'><%=web_var(web_config,1)%></a>���߹�ģ����վ�ֲ�����asp+access������������Լ���Ϊ�����ۺ�����վ��չ��<br>
����<a href='<%=web_var(web_config,2)%>'><%=web_var(web_config,1)%></a>��������ʼ�ռ�֡����ɡ��ľ�������ȫ��Ӯ���ԡ�����ҵ�Ե���վ�������Ը��õ�Ϊ�����������ṩ���ַ���ͷ���Ϊ��ּ��������ǵ�������վ���ݺͷ���Ͷ�����ѵģ����Ǿ������κ��������û���ȡ�κη��ã��������û������κ���Ʒ������������һ�У�Ŀ�Ľ������ƹ�Beyond�����֣�����Beyond�ľ����о���صļ������������⡣
</td>
  </tr>
  </table>
</td></tr>
</table>

<br>

<%response.write table1
%>
<tr<%response.write table2%>><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk0") %>&nbsp;<font class=end><b>������</b></font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
  <tr>
        <td width='100%' class=htd><br><center><img src=images/yandixin.jpg></center>
          <br>�����е�������վ������Ϊ���������һ����Ʒ�����е�������վ������Ϊ����һЩ��Ҫ˵��������������ʿ������ֱ��ҵ���������<br>
<br>����ÿÿ����Щʧ�ߵ�ҹ������������һЩ�ϸ裬������Ǹ�����һ����ȥ���˺��£�������Щ�����ݱ���ӳⷽ�ٵ�ʱ�ڡ�����һ���˻�����ҹ��ֻ���һ��ڣ����֣����֣���������������������͸ʱ�գ���͸�ҵ����࣡<br>
<br>


�������õ����֣�����õ��������




</td>
  </tr>
  </table>
</td></tr>
</table>
<%
end sub
%>