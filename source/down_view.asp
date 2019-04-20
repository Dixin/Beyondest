<!-- #include file="INCLUDE/config_down.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V4.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim id:id=trim(request.querystring("id"))
if not(isnumeric(id)) then
  call format_redirect("down.asp")
  reponse.end
end if
%>
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="INCLUDE/config_review.asp" -->
<!-- #include file="include/conn.asp" -->
<%
dim cname,sname,temp1,keyes,power,userp,emoney,url1,url2,sql2,rs2
set rs=server.createobject("ADODB.recordset")
sql="select * from down where hidden=1 and id="&id
rs.open sql,conn,1,1
if rs.eof and rs.bof then
  rs.close:set rs=nothing
  call close_conn()
  call format_redirect("down.asp")
  response.end
end if
cid=rs("c_id")
sid=rs("s_id")
keyes=rs("keyes")
power=rs("power")
emoney=rs("emoney")

cname="音乐浏览":sname=""
if cid>0 then
  if sid>0 then
    sql2="select jk_class.c_name,jk_sort.s_name from jk_sort inner join jk_class on jk_sort.c_id=jk_class.c_id where jk_sort.c_id="&cid&" and jk_sort.s_id="&sid
  else
    sql2="select c_name from jk_class where c_id="&cid
  end if
  set rs2=conn.execute(sql2)
  if not (rs2.eof and rs2.bof) then
    cname=rs2("c_name"):tit=cname
    if sid>0 then sname=rs2("s_name"):tit=sname&"（"&cname&"）"
  end if
  rs2.close:set rs2=nothing
end if

if action="download" then
  call web_head(1,0,0,0,0)
else
  call web_head(0,0,0,0,0)
end if
'--------------------------------download---------------------------------
userp=int(format_power(login_mode,2))


if action="download" then
  call emoney_notes(power,emoney,n_sort,id,"js",1,1,"?id="&id)
  if trim(request.querystring("url"))="download2" then
    index_url=rs("url2")
  else
    index_url=rs("url")
  end if
  rs.close:set rs=nothing
  sql="update down set counter=counter+1 where id="&id
  conn.execute(sql)
  call close_conn()
  
  response.redirect ""&url_true(web_var(web_down,5),index_url)&""
  response.end
end if
'------------------------------------left----------------------------------
%>
<table border=0 width='96%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center><%call format_login()%></td></tr>
<tr><td align=center><%call down_sea()%></td></tr>
<tr><td align=center><%call down_new_hot("jt0","","","","good",10,0,13,1,0)%></td></tr>
<tr><td align=center><%call down_new_hot("jt0","","","","hot",10,0,13,1,0)%></td></tr>
<tr><td align=center><%call down_new_hot("jt0","","","","new",10,0,13,1,0)%></td></tr>
</table>
<%
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td align=center><% call down_intro(sid,sname) %></td></tr>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center>
  <table border=1 cellspacing=0 cellpadding=4 width='98%' bordercolorlight=<%response.write web_var(web_color,3)%> bordercolordark=<%response.write web_var(web_color,5)%>>
  <tr bgcolor=<%response.write web_var(web_color,6)%> bordercolordark=<%response.write web_var(web_color,5)%>>
  <td align=center colspan=3 height=30><font size=3 class=blue><b><%response.write rs("name")%></b></font></td></tr>
  <tr><td align=center width='15%' bgcolor=<%response.write web_var(web_color,5)%>>专辑类型：</td><td width='40%'><%response.write rs("genre")%>&nbsp;</td>
  <td align=center width='45%' rowspan=8>
   <%
        response.write "<img src='images/down/"&rs("pic")&"' border=0>"
     
%>
</td></tr>
  <tr><td align=center bgcolor=<%response.write web_var(web_color,5)%>>播放软件：</td><td><%
if rs("os")="Realone" then response.write "<a href="&web_var(web_down,5)&"/soft/realoneplayer.rar><img src=images/down/tool_realone.gif alt='Real One Player'  border='0'></a>"
if rs("os")="WinMediaPlayer" then response.write "<a href="&web_var(web_down,5)&"/soft/wmp2k.rar><img src=images/down/TOOL_WMP.gif alt='Windows Media Player for 98 & Me & 2k'  border='0'></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="&web_var(web_down,5)&"/soft/wmpxp.rar><img src=images/down/TOOL_WMP.gif alt='Windows Media Player for XP'  border='0'></a>"
if rs("os")="Winamp" then response.write "<a href="&web_var(web_down,5)&"/soft/Winamp.rar><img src=images/down/tool_winamp.gif alt='Winamp'  border='0'></a>"
%>&nbsp;</td></tr>
  <tr><td align=center bgcolor=<%response.write web_var(web_color,5)%>>文件大小：</td><td><%response.write rs("sizes")%></td></tr>
  <tr><td align=center bgcolor=<%response.write web_var(web_color,5)%>>推荐等级：</td><td><img src='images/down/star<%response.write rs("types")%>.gif' border=0></td></tr>
  <tr><td align=center bgcolor=<%response.write web_var(web_color,5)%>>下载次数：</td><td><font class=red><%response.write rs("counter")%></font></td></tr>
  <tr><td align=center bgcolor=<%response.write web_var(web_color,5)%>>发&nbsp;布&nbsp;人：</td><td><%response.write format_user_view(rs("username"),1,1)%></td></tr>
  <tr><td align=center bgcolor=<%response.write web_var(web_color,5)%>>上传日期：</td><td><%response.write time_type(rs("tim"),88)%></td></tr>
  <tr><td align=center bgcolor=<%response.write web_var(web_color,5)%>>文件来自：</td><td><%
temp1=rs("homepage")
if temp1="" or isnull(temp1) or temp1="http://" then
  response.write "<a href='"&web_var(web_config,2)&"' target=_blank>"&web_var(web_config,2)&"</a>"
else
  response.write "<a href='"&temp1&"' target=_blank>"&temp1&"</a>"
end if
%></td></tr>
  <tr><td align=center bgcolor=<%response.write web_var(web_color,5)%>>下载权限：</td><td colspan=2>&nbsp;注册用户</td></tr>
  <tr><td align=center bgcolor=<%response.write web_var(web_color,5)%>>下载地址：</td><td colspan=2>&nbsp;&nbsp;&nbsp;<a href='?action=download&id=<%response.write id%>'<%response.write atb%>><img src='IMAGES/DOWN/DOWNLOAD.GIF' border=0></a>&nbsp;
<% if len(rs("url2"))>8 then %>
&nbsp;&nbsp;&nbsp;<a href='?action=download&url=download2&id=<%response.write id%>'<%response.write atb%>><img src='IMAGES/DOWN/download2.gif' border=0></a>
<% end if %></td></tr>
  <tr height=50 valign=top><td align=center bgcolor=<%response.write web_var(web_color,5)%>>作品备注：</td><td colspan=2><table borer=0 width='100%' class=tf><tr><td><%
temp1=rs("remark")
if len(temp1)<3 then
  temp1="<font class=gray>好像没有关于该音乐的介绍哦！</font>"
else
  temp1=code_jk(temp1)
end if
response.write temp1
rs.close
%></td></tr></table></td></tr>
  <tr valign=top><td align=center bgcolor=<%response.write web_var(web_color,5)%>>相关音乐：</td><td colspan=2><table border=0><%dim tempsn,tempcn,sqls,sqlt,rss,rst
sql="select id,name,tim,counter,c_id,s_id from down where hidden=1 and keyes like '%"&keyes&"%' and id<>"&id&" order by counter desc"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  response.write vbcrlf&"<tr><td class=gray>没有与之相关的作品</td></tr>"
else
  do while not rs.eof
    temp1=rs("name")
    sqls="select s_name from jk_sort where s_id="&rs("s_id")
    set rss=conn.execute(sqls)
    tempsn=rss("s_name")
    rss.close:set rss=nothing
    sqlt="select c_name from jk_class where c_id="&rs("c_id")
    set rst=conn.execute(sqlt)
    tempcn=rst("c_name")
    rst.close:set rst=nothing
    response.write vbcrlf&"<tr><td><img src=images/small/jt0.gif>&nbsp;"&tempsn&"（"&tempcn&"）：<a href='down_view.asp?id="&rs("id")&"' title='"&code_html(temp1,1,0)&"'>"&code_html(temp1,1,30)&"</a></td></tr>"
    rs.movenext
  loop
end if
rs.close:set rs=nothing
%></table></td></tr>
  </table>
</td></tr>

<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td height=10></td></tr>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><% call review_type(n_sort,id,"down_view.asp?id="&id,1) %></td></tr>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td height=5></td></tr>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%call down_class_sortt(cid,sid)%></td></tr>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%call down_remark("jt0")%></td></tr>

</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)
%>