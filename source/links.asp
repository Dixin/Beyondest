<!-- #include file="INCLUDE/config_other.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nummer:nummer=5
tit="��������"

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
call links_left()
'response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong&table1
response.write vbcrlf&"<tr"&table2&"><td class=end background=images/"&web_var(web_config,5)&"/bar_3_bg.gif>&nbsp;"&img_small(us)&"&nbsp;&nbsp;<b>����վ��</b></td></tr><tr"&table3&"><td align=center>"
%>
  <table border=0 width='100%' cellspacing=0 cellpadding=0>
  <tr><td align=center><%call links_main("fir",nummer)%></td></tr>
  </table>
<%
response.write vbcrlf&"</td></tr>"
response.write vbcrlf&"<tr"&table2&"><td class=end background=images/"&web_var(web_config,5)&"/bar_3_bg.gif>&nbsp;"&img_small(us)&"&nbsp;&nbsp;<b>��������</b></td></tr><tr"&table3&"><td align=center>"
%>
  <table border=0 width='100%' cellspacing=0 cellpadding=0>
  <tr><td align=center><%call links_main("sec",nummer)%></td></tr>
  </table>
<%
response.write vbcrlf&"</td></tr>"
response.write vbcrlf&"<tr"&table2&"><td class=end background=images/"&web_var(web_config,5)&"/bar_3_bg.gif>&nbsp;"&img_small(us)&"&nbsp;&nbsp;<b>��������</b></td></tr><tr"&table3&"><td align=center>"
%>
  <table border=0 width='100%' cellspacing=0 cellpadding=0>
  <tr><td align=center><%call links_main("txt",nummer)%></td></tr>
  </table>
<%
response.write vbcrlf&"</td></tr>"
response.write vbcrlf&"<tr"&table2&"><td class=end background=images/"&web_var(web_config,5)&"/bar_3_bg.gif>&nbsp;"&img_small(us)&"&nbsp;&nbsp;<b>����˵��</b></td></tr><tr"&table3&"><td align=center>"
%>
<table border=0 width=450>
<tr><td>
<table border=0><tr><td class=htd>1��վ�����������ʵ����������������ʵ��ˡ�����ӡ�<br>
2�������޸������Ĭ��ҳ��ע������ҳ�������ӡ�<br>
3�������κη������ݻ�ɫ�����ݵĲ������ӡ�</td></tr></table>
</td></tr>
</table>
<%
response.write vbcrlf&"</td></tr></table><br>"
'---------------------------------center end-------------------------------
call web_end(0)

sub links_left()
  tit=vbcrlf&"<table border=0 width='100%' cellpadding=0 cellspacing=0>" & _
      vbcrlf&"<tr><td align=center height=50><a href='" & web_var(web_config,2) & "' target=_blank><img border=0 src='images/"&web_var(web_config,5)&"/logo.gif' width=120 heigh=40 alt='" & web_var(web_config,1) & "'></a></td></tr>" & _
      vbcrlf&"<tr><td align=center valign=top><textarea name=flink_main rows=5 cols=21  onfocus=this.select() onmouseover=this.focus()><a href='" & web_var(web_config,2) & "' target=_blank><img border=0 src='" & web_var(web_config,2) & "images/"&web_var(web_config,5)&"/logo.gif' width=88 heigh=31 alt='" & web_var(web_config,1) & "'></a></textarea></td></tr>" & _
      vbcrlf&"</table>"
  call left_type(tit,"links",1)
end sub
%>