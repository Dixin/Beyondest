<!--#include file="include/config.asp"-->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
%>
<html>
<head>
<title><%response.write web_var(web_config,1)%> - �����б�</title>
<meta name="Description"  content="Beyondest">
<meta name="keywords" content="��ȫ��Beyond����,��õ�Beyond��վ,asp,Beyondest,����">
<meta name="author" content="Beyondest">
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel=stylesheet href="include/beyondest.css" type=text/css>
</head>
<body leftmargin=0 topmargin=0>
<body topmargin=0 leftmargin=0 bgcolor=<%
dim ttt,vid
vid=trim(request.querystring("vid"))
response.write web_var(web_color,1)
ttt=web_var(web_config,7)
if ttt<>"" then
  response.write " background='images/"&ttt&".gif'"
end if
%>>
<center><table border=0 width='100%' height='100%' cellspacing=0 cellpadding=0>
<tr><td width='100%' height='100%' align=center>
  <table border=0 width='100%' height='100%' cellspacing=0 cellpadding=0>
  <tr><td width=20 height=16><img src='IMAGES/VOTE/vote_r1_c1.gif' width=20 height=16 border=0></td><td background='IMAGES/VOTE/vote_r1_c2.gif'></td><td width=16><img src='IMAGES/VOTE/vote_r1_c6.gif' width=16 height=16 border=0></td></tr>
  <tr><td background='IMAGES/VOTE/vote_r2_c1.gif' valign=bottom><img src='IMAGES/VOTE/vote_r4_c1.gif' width=20 border=0 height=8></td><td bgcolor=#ffffff align=center>
<%
if not(isnumeric(vid)) then
  call vote_error()
else
  select case action
  case "save"
    call vote_save()
  case else
    call vote_view()
  end select
end if

call close_conn()
%>
  </td><td background='IMAGES/VOTE/vote_r2_c6.gif'></td></tr>
  <tr><td height=16 background='IMAGES/VOTE/vote_r5_c4.gif' colspan=2><a href="javascript:window.close();"><img src='IMAGES/VOTE/vote_r5_c1.gif' width=84 height=18 border=0></a></td><td width=16><img src='IMAGES/VOTE/vote_r5_c6.gif' width=16 height=18 border=0></td></tr>
  </table>
</td></tr></table></center>
</body>
</html>
<%
sub vote_save()
  dim go_tim,vvid:go_tim=web_var(web_num,5)
  if trim(request.cookies("beyondest_online")("vote_vid"))="v"&vid then
%><font class=red_2>���Ѿ�Ͷ��һƱ���������ظ���Ͷ����</font><br><br>
<a href='votetype.asp?type=view&vid=<%response.write vid%>'>�鿴ͶƱ���</a><br><br>
<font class=gray>��ϵͳ���� <font class=red><%response.write go_tim%></font> ���Ӻ��Զ����룩</font><br><br>
<meta http-equiv='refresh' content='<%response.write go_tim%>; url=votetype.asp?type=view&vid=<%response.write vid%>'>
<%  exit sub
  end if
  sql="select id from vote where vid="&vid
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call vote_error():exit sub
  end if
  rs.close:set rs=nothing
  dim vote_id,ddim,j:j=0
  vote_id=trim(request.form("vote_id"))
  vote_id=replace(vote_id," ","")
  if len(vote_id)<1 then call vote_error():exit sub
  ddim=split(vote_id,",")
  for i=0 to ubound(ddim)
    if isnumeric(ddim(i)) then
      sql="update vote set counter=counter+1 where vid="&vid&" and vtype<>0 and id="&ddim(i)
      conn.execute(sql)
      j=j+1
    end if
  next
  erase ddim
  if j=0 then call vote_error():exit sub
  response.cookies("beyondest_online")("vote_vid")="v"&vid
  call cookies_yes()
%>������<font class=red>лл���֧�������</font>������<br><br>
<a href='votetype.asp?type=view&vid=<%response.write vid%>'>�鿴ͶƱ���</a><br><br>
<font class=gray>��ϵͳ���� <font class=red><%response.write go_tim%></font> ���Ӻ��Զ����룩</font><br><br>
<meta http-equiv='refresh' content='<%response.write go_tim%>; url=votetype.asp?type=view&vid=<%response.write vid%>'>
<%
end sub

sub vote_view()
  dim rssum,dimc,dimn,num,t
  rssum=0:num=0
  set rs=server.createobject("adodb.recordset")
  sql="select id,vname,counter from vote where vid="&vid&" order by id"
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call vote_error():exit sub
  end if
  rssum=int(rs.recordcount)
  rssum=rssum-1
  redim dimc(rssum),dimn(rssum)
  for i=0 to rssum
    if rs.eof then exit for
    dimc(i)=rs("counter")
    dimn(i)=rs("vname")
    num=dimc(i)+num
    rs.movenext
  next
  rs.close:set rs=nothing
%>
<table border=0 width='96%' cellpadding=0 cellspacing=2 align=center>
<tr><td align=center colspan=3 class=red_3 height=20><b><%response.write code_html(dimn(0),1,0)%></b></td></tr>
<tr><td align=center colspan=3 class=gray>Ŀǰ���� <font class=blue><%response.write num%></font> �˲�����ͶƱ</td></tr>
<tr>
<td></td>
<td></td>
<td width='15%'></td>
</tr>
<%
  for i=1 to rssum
    if int(dimc(i))=0 then
      t="0%"
    else
      t=FormatPercent(dimc(i)/num,1)
    end if
%>
<tr>
<td height=18><%response.write i%>��<%response.write code_html(dimn(i),1,0)%> <font class=gray>(<font class=blue><%response.write dimc(i)%></font>)</font></td>
<td align=right><img src='IMAGES/VOTE/BAR.GIF' width=<%response.write dimc(i)%> height=10></td>
<td align=right><%response.write t%></td>
</tr>
<% next %>
</table>
<%
  erase dimc:erase dimn
end sub

sub vote_error()
%>
<font class=red>������û��ѡ����ص���Ʊѡ�</font><br><br><font class=red_2>����зǷ����ύ��ͶƱ���ݣ�</font>
<br><br>
<%
  response.write closer
end sub
%>
