<!-- #include file="INCLUDE/config_forum.asp" -->
<% if not(isnumeric(forumid)) or not(isnumeric(viewid)) then call cookies_type("view_id") %>
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

call forum_first()
call forum_word()

dim ii,rssum,nummer,thepages,viewpage,page,pageurl,view_temp,iid,qid,table_bg,id,money_name,thetopictt
dim u_username,u_nname,u_sex,u_whe,u_qq,u_remark,u_bbs_counter,u_integral,u_emoney,u_power,u_popedom
dim fir_id,fir_topic,fir_top,counter,re_counter,fir_istop,fir_isgood,fir_islock,del_type
tit=forumname&"（浏览贴子）"
pageurl="?forum_id="&forumid&"&view_id="&viewid&"&"
nummer=web_var(web_num,3):view_temp="":money_name=web_var(web_config,8)

set rs=server.createobject("adodb.recordset")
sql="select bbs_data.*,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.islock,bbs_topic.istop,bbs_topic.isgood,user_data.username as u_username," & _
    "user_data.nname as u_nname,user_data.sex as u_sex,user_data.whe as u_whe,user_data.qq as u_qq,user_data.email as u_email," & _
    "user_data.url as u_url,user_data.face as u_face,user_data.tim as u_tim,user_data.remark as u_remark,user_data.bbs_counter as u_bbs_counter,user_data.emoney as u_emoney,user_data.integral as u_integral,user_data.power as u_power,user_data.popedom as u_popedom " & _
    "from user_data inner join ( bbs_topic inner join bbs_data on bbs_data.reply_id=bbs_topic.id )" & _
    " on bbs_data.username=user_data.username where bbs_data.forum_id="&forumid&" and bbs_data.reply_id="&viewid&" order by bbs_data.id"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
  rs.close:set rs=nothing:close_conn
  call cookies_type("view_id")
end if

call web_head(0,0,2,0,0)
'-----------------------------------center---------------------------------
%>
<script language=JavaScript>
<!--
function forum_do_del(data1,data2)
{
  if (confirm("此操作将删除id为 "+data2+" 的回贴！\n\n真的要删除吗？\n删除后将无法恢复！"))
    window.location = "forum_isaction.asp?isaction=del&forum_id="+data1+"&del_id="+data2
}
function forum_do_delete(data1,data2)
{
  if (confirm("此操作将删除id为 "+data2+" 的贴子！\n\n真的要删除吗？\n删除后将无法恢复！"))
    window.location = "forum_isaction.asp?isaction=delete&forum_id="+data1+"&del_id="+data2
}
//-->
</script>

<%
thetopictt=forum_table1&"<tr height=26><td background=images/"&web_var(web_config,5)&"/bar_1_bg.gif colspan=2>"
view_temp="</td></tr>"




rssum=rs.recordcount
call format_pagecute()

if int(viewpage)>1 then
  fir_id=rs("reply_id")
  fir_topic=rs("topic")
  fir_islock=rs("islock"):fir_istop=rs("istop"):fir_isgood=rs("isgood")
  fir_istop=int(fir_istop):fir_isgood=int(fir_isgood):fir_islock=int(fir_islock)
  fir_top=fir_topic
  fir_top=code_html(fir_top,1,0)
  counter=rs("counter")
  re_counter=rs("re_counter")
  rs.move (viewpage-1)*nummer
end if
for ii=1 to nummer
  if rs.eof then exit for
  iid=rs("id")
  qid=iid
  id=rs("reply_id")
  u_username=rs("u_username")
  u_nname=code_html(rs("u_nname"),1,0)
  u_sex=rs("u_sex")
  u_whe=code_html(rs("u_whe"),1,0)
  u_qq=rs("u_qq")
  u_remark=code_jk2(rs("u_remark"))
  u_bbs_counter=rs("u_bbs_counter")
  u_integral=rs("u_integral")
  u_emoney=rs("u_emoney")
  u_power=rs("u_power")
  u_popedom=rs("u_popedom")
  del_type="forum_do_del"
  if int(ii)=1 and int(viewpage)=1 then
    fir_id=id
    fir_topic=rs("topic")
    fir_islock=rs("islock"):fir_istop=rs("istop"):fir_isgood=rs("isgood")
    fir_istop=int(fir_istop):fir_isgood=int(fir_isgood):fir_islock=int(fir_islock)
    fir_top=fir_topic
    fir_top=code_html(fir_top,1,0)
    counter=rs("counter")
    re_counter=rs("re_counter")
    iid=viewid
    del_type="forum_do_delete"
  end if
  
  view_temp=view_temp&view_type()
  rs.movenext
next
rs.close:set rs=nothing
view_temp=view_temp&"</td></tr></table>"


fir_istop=int(fir_istop)
if fir_istop<>0 and fir_istop<>1 and fir_istop<>2 then fir_istop=0

response.write forum_top("浏览贴子 （回复：<font class=red>"&re_counter&"</font>&nbsp;浏览：<font class=red>"&counter+1&"</font>）")%>

<table border=0 width='98%' cellspacing=0 cellpadding=0><tr><td align=left width='15%'><a href='forum_write.asp?forum_id=<%=forumid%>'><img src='images/<%=web_var(web_config,5)%>/new_topic.gif' align=absMiddle border=0 title='在 <%=forumname%> 里发表我的新贴'></a></td><td align=right width='85%'></td></tr></table>



<%response.write kong&thetopictt%> 

<table boder=0 width='100%' cellspacing=0 cellpadding=0>
<tr><td width='80%'>&nbsp;主题：<b><font class=end title='<%response.write fir_top%>'><%response.write code_html(fir_topic,1,30)%></font></b></td>
<td align=center width='20%'><table border=0 cellspacing=0 cellpadding=0><tr align=center>
  <td width=50><a href='javascript:;' onclick="javascript:document.location.reload()"><%response.write img_small("page_refresh")%></a></td>
  <td width=55><a href="javascript:window.external.AddFavorite('<%response.write web_var(web_config,2)&pageurl%>','<%response.write web_var(web_config,1)&" - "&forumname&"（贴子："&fir_topic&"）"%>')"><%response.write img_small("page_fav")%></td>
  </tr></table>
  
</td></tr></table>

<%response.write view_temp&kong&format_table(1,2)
%>



<tr height=30<%response.write forum_table3%>>
<td width='75%'>&nbsp;分页：<% response.write jk_pagecute(nummer,thepages,viewpage,pageurl,6,"#ff0000") %></td>
<td width='25%' align=center><% response.write forum_go() %></td>
</tr>
<tr height=30 align=center<%response.write format_table(3,1)%>><td>主题贴类型：<font class=blue><%
if fir_istop<>0 or fir_isgood<>0 or fir_islock<>0 then
  if fir_istop=1 then
    response.write "[ 固顶 ]&nbsp;"
  elseif fir_istop=2 then
    response.write "[ 总固顶 ]&nbsp;"
  end if
  if fir_isgood<>0 then response.write "[ 精华 ]&nbsp;"
  if fir_islock<>0 then response.write "[ 锁定 ]&nbsp;"
else
  response.write "[ 正常 ]&nbsp;"
end if
response.write "</font>"

if format_user_power(login_username,login_mode,forumpower)="yes" then
%>&nbsp;相关操作：
<a href='forum_isaction.asp?isaction=is&forum_id=<% response.write forumid %>&view_id=<% response.write id %>&action=istop<%
select case fir_istop
case 1
  response.write "&cancel=yes' class=red_3>取消固顶</a>&nbsp;┋" & _
                 "<a href='forum_isaction.asp?isaction=is&forum_id="&forumid&"&view_id="&id&"&action=istops'>总固顶</a>"
case 2
  response.write "'>固顶</a>&nbsp;┋" & _
                 "<a href='forum_isaction.asp?isaction=is&forum_id="&forumid&"&view_id="&id&"&action=istops&cancel=yes' class=red_3>取消总固顶</a>"
case else
  response.write "'>固顶</a>&nbsp;┋" & _
                 "<a href='forum_isaction.asp?isaction=is&forum_id="&forumid&"&view_id="&id&"&action=istops'>总固顶</a>"
end select
%>&nbsp;┋
<a href='forum_isaction.asp?isaction=is&forum_id=<% response.write forumid %>&view_id=<% response.write id %>&action=isgood<%
  if fir_isgood=0 then
    response.write "'>"
  else
    response.write "&cancel=yes' class=red_3>取消"
  end if
%>精华</a>&nbsp;┋
<a href='forum_isaction.asp?isaction=is&forum_id=<% response.write forumid %>&view_id=<% response.write id %>&action=islock<%
  if fir_islock=0 then
    response.write "'>"
  else
    response.write "&cancel=yes' class=red_3>取消"
  end if
%>锁定</a>&nbsp;┋
<a href='forum_isaction.asp?isaction=delete&forum_id=<% response.write forumid %>&del_id=<% response.write id %>'>删除</a>
<% end if %>
</td>
<td><% response.write forum_move(forumid,viewid) %></td></tr>
</table>
<script language=javascript src='style/forum_ok.js'></script>
<% response.write kong & forum_table1 %>
<form name=write_frm action='forum_reply.asp?forum_id=<%=forumid%>&view_id=<%=qid%>' method=post onsubmit="frm_submitonce(this);">
<input type=hidden name=reply value='ok'>
<tr<%response.write forum_table2%>><td height=25 valign=bottom colspan=2 background=images/<%=web_var(web_config,5)%>/bar_1_bg.gif>
<%
if fir_islock<>1 then
  if login_mode="" then
    response.write "<div align=center>"&web_var(web_error,2)&"</div>"
  else
%>
&nbsp;→ 快速回复：<b><font class=red_3><% response.write fir_top %></b></font>
</td></tr>
<tr height=30<%response.write format_table(3,1)%>>
<td width='20%' align=center bgcolor='<%=web_var(web_color,6)%>'>用户信息：</td>
<td width='80%'>&nbsp;&nbsp;用户名：<input type=username name=username value='<%response.write login_username%>' size=18 maxlength=20>&nbsp;&nbsp;
密码：<input type=password name=password value='<%response.write login_password%>' size=18 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='user_main.asp?username=<%response.write server.htmlencode(login_username)%>'>用户中心</a> ]</font>&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='login.asp?action=logout'>退出登陆</a> ]</font></td>
</tr>
<tr height=30<%response.write forum_table3%>>
<td align=center bgcolor='<%=web_var(web_color,6)%>'>表情符号：</td>
<td>&nbsp;&nbsp;<%response.write icon_type(9,3)%></td>
</tr>
<tr align=center<%response.write format_table(3,1)%>>
<td bgcolor='<%=web_var(web_color,6)%>'><table border=0><tr><td class=htd>贴子内容：<%response.write redx%><br><%response.write word_remark%></td></tr></table></td>
<td><table border=0><tr><td><textarea name=jk_word rows=8 cols=95 title='按 Ctrl+Enter 可直接发送' onkeydown="javascript:frm_quicksubmit();"></textarea></td></tr></table></td>
</tr>
<script language=javascript src='style/em_type.js'></script>
<tr height=30<%response.write forum_table3%>>
<td align=center bgcolor='<%=web_var(web_color,6)%>'>E M 贴图：</td>
<td>&nbsp;<script language=javascript>jk_em_type('s');</script></td>
</tr>
<tr<%response.write format_table(3,1)%>><td colspan=2 align=center height=60>&nbsp;&nbsp;<script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<%response.write forum_table3%>>
<td bgcolor='<%=web_var(web_color,6)%>'>快速回复：</td>
<td><input type=submit name=wsubmit value='快速发表我的回贴'>　&nbsp;<input type=button value='预览我的回复'>　&nbsp;<input type=reset value='清除重写'>　&nbsp;（按 Ctrl + Enter 可快速回复）
<%
  end if
else
  response.write "<div align=center><font class=red_2>这个贴子已被锁定！不能再对其进行回复</font></div>"
end if
%>
</td></tr></form></table>
<br>
<%
sql="update bbs_topic set counter=counter+1 where forum_id="&forumid&" and id="&viewid
conn.execute(sql)
'---------------------------------center end-------------------------------
call web_end(0)
%>