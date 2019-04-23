<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nummer,page,rssum,thepages,viewpage,pageurl
thepages=0:viewpage=1:pageurl="?"
tit="浏览头像"

call web_head(0,0,3,0,0)
'------------------------------------left----------------------------------
response.write ukong

call user_face()

response.write kong
'---------------------------------center end-------------------------------
call web_end(0)

sub user_face()
  dim fc,j,fnum,cnum,rnum,tt,nnum
  cnum=5:rnum=3:fc=0
  fnum=int(web_var(web_num,11))+1
  'fc=fnum\cnum
  'if fnum mod cnum >0 then fc=fc+1
  
  nummer=cnum*rnum
  rssum=fnum
  call format_pagecute()
  if int(viewpage)>1 then
    fc=(viewpage-1)*nummer
  end if
%>
<table border=0 cellpadding=0 cellspacing=4 align=center>
<tr align=center>
<td>本站共有 <font class=red><%response.write rssum%></font> 个头像</td>
<td width=10></td>
<td>页次：<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font></td>
<td width=10></td>
<td>分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%></td>
</tr>
</table>
<table border=0 width='98%' cellpadding=0 cellspacing=8 align=center>
<%
  for j=1 to rnum
    response.write vbcrlf&"<tr align=center>"
    for i=1 to cnum
      nnum=cnum*(j-1)+i-1+fc
      if nnum>=rssum then exit for
      tt="<table border=0><tr><td align=center><img src='images/face/"&nnum&".gif' border=0></td></tr><tr><td align=center><b>"&nnum&"</b></td></tr></table>"
      response.write vbcrlf&"<td>"&format_k(tt,1,5,120,120)&"</td>"
    next
    response.write vbcrlf&"</tr>"
  next
%>
</table>
<%
end sub
%>