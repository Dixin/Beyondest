<!-- #include file="INCLUDE/config_forum.asp" -->
<% if not(isnumeric(forumid)) or not(isnumeric(viewid)) then call cookies_type("view_id") %>
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

call forum_first()
call forum_word()

dim ii,rssum,nummer,thepages,viewpage,page,pageurl,view_temp,iid,qid,table_bg,id,money_name,thetopictt
dim u_username,u_nname,u_sex,u_whe,u_qq,u_remark,u_bbs_counter,u_integral,u_emoney,u_power,u_popedom
dim fir_id,fir_topic,fir_top,counter,re_counter,fir_istop,fir_isgood,fir_islock,del_type
tit=forumname&"��������ӣ�"
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
  if (confirm("�˲�����ɾ��idΪ "+data2+" �Ļ�����\n\n���Ҫɾ����\nɾ�����޷��ָ���"))
    window.location = "forum_isaction.asp?isaction=del&forum_id="+data1+"&del_id="+data2
}
function forum_do_delete(data1,data2)
{
  if (confirm("�˲�����ɾ��idΪ "+data2+" �����ӣ�\n\n���Ҫɾ����\nɾ�����޷��ָ���"))
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

response.write forum_top("������� ���ظ���<font class=red>"&re_counter&"</font>&nbsp;�����<font class=red>"&counter+1&"</font>��")%>

<table border=0 width='98%' cellspacing=0 cellpadding=0><tr><td align=left width='15%'><a href='forum_write.asp?forum_id=<%=forumid%>'><img src='images/<%=web_var(web_config,5)%>/new_topic.gif' align=absMiddle border=0 title='�� <%=forumname%> �﷢���ҵ�����'></a></td><td align=right width='85%'></td></tr></table>



<%response.write kong&thetopictt%> 

<table boder=0 width='100%' cellspacing=0 cellpadding=0>
<tr><td width='80%'>&nbsp;���⣺<b><font class=end title='<%response.write fir_top%>'><%response.write code_html(fir_topic,1,30)%></font></b></td>
<td align=center width='20%'><table border=0 cellspacing=0 cellpadding=0><tr align=center>
  <td width=50><a href='javascript:;' onclick="javascript:document.location.reload()"><%response.write img_small("page_refresh")%></a></td>
  <td width=55><a href="javascript:window.external.AddFavorite('<%response.write web_var(web_config,2)&pageurl%>','<%response.write web_var(web_config,1)&" - "&forumname&"�����ӣ�"&fir_topic&"��"%>')"><%response.write img_small("page_fav")%></td>
  </tr></table>
  
</td></tr></table>

<%response.write view_temp&kong&format_table(1,2)
%>



<tr height=30<%response.write forum_table3%>>
<td width='75%'>&nbsp;��ҳ��<% response.write jk_pagecute(nummer,thepages,viewpage,pageurl,6,"#ff0000") %></td>
<td width='25%' align=center><% response.write forum_go() %></td>
</tr>
<tr height=30 align=center<%response.write format_table(3,1)%>><td>���������ͣ�<font class=blue><%
if fir_istop<>0 or fir_isgood<>0 or fir_islock<>0 then
  if fir_istop=1 then
    response.write "[ �̶� ]&nbsp;"
  elseif fir_istop=2 then
    response.write "[ �̶ܹ� ]&nbsp;"
  end if
  if fir_isgood<>0 then response.write "[ ���� ]&nbsp;"
  if fir_islock<>0 then response.write "[ ���� ]&nbsp;"
else
  response.write "[ ���� ]&nbsp;"
end if
response.write "</font>"

if format_user_power(login_username,login_mode,forumpower)="yes" then
%>&nbsp;��ز�����
<a href='forum_isaction.asp?isaction=is&forum_id=<% response.write forumid %>&view_id=<% response.write id %>&action=istop<%
select case fir_istop
case 1
  response.write "&cancel=yes' class=red_3>ȡ���̶�</a>&nbsp;��" & _
                 "<a href='forum_isaction.asp?isaction=is&forum_id="&forumid&"&view_id="&id&"&action=istops'>�̶ܹ�</a>"
case 2
  response.write "'>�̶�</a>&nbsp;��" & _
                 "<a href='forum_isaction.asp?isaction=is&forum_id="&forumid&"&view_id="&id&"&action=istops&cancel=yes' class=red_3>ȡ���̶ܹ�</a>"
case else
  response.write "'>�̶�</a>&nbsp;��" & _
                 "<a href='forum_isaction.asp?isaction=is&forum_id="&forumid&"&view_id="&id&"&action=istops'>�̶ܹ�</a>"
end select
%>&nbsp;��
<a href='forum_isaction.asp?isaction=is&forum_id=<% response.write forumid %>&view_id=<% response.write id %>&action=isgood<%
  if fir_isgood=0 then
    response.write "'>"
  else
    response.write "&cancel=yes' class=red_3>ȡ��"
  end if
%>����</a>&nbsp;��
<a href='forum_isaction.asp?isaction=is&forum_id=<% response.write forumid %>&view_id=<% response.write id %>&action=islock<%
  if fir_islock=0 then
    response.write "'>"
  else
    response.write "&cancel=yes' class=red_3>ȡ��"
  end if
%>����</a>&nbsp;��
<a href='forum_isaction.asp?isaction=delete&forum_id=<% response.write forumid %>&del_id=<% response.write id %>'>ɾ��</a>
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
&nbsp;�� ���ٻظ���<b><font class=red_3><% response.write fir_top %></b></font>
</td></tr>
<tr height=30<%response.write format_table(3,1)%>>
<td width='20%' align=center bgcolor='<%=web_var(web_color,6)%>'>�û���Ϣ��</td>
<td width='80%'>&nbsp;&nbsp;�û�����<input type=username name=username value='<%response.write login_username%>' size=18 maxlength=20>&nbsp;&nbsp;
���룺<input type=password name=password value='<%response.write login_password%>' size=18 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='user_main.asp?username=<%response.write server.htmlencode(login_username)%>'>�û�����</a> ]</font>&nbsp;&nbsp;&nbsp;&nbsp;
<font class=gray>[ <a href='login.asp?action=logout'>�˳���½</a> ]</font></td>
</tr>
<tr height=30<%response.write forum_table3%>>
<td align=center bgcolor='<%=web_var(web_color,6)%>'>������ţ�</td>
<td>&nbsp;&nbsp;<%response.write icon_type(9,3)%></td>
</tr>
<tr align=center<%response.write format_table(3,1)%>>
<td bgcolor='<%=web_var(web_color,6)%>'><table border=0><tr><td class=htd>�������ݣ�<%response.write redx%><br><%response.write word_remark%></td></tr></table></td>
<td><table border=0><tr><td><textarea name=jk_word rows=8 cols=95 title='�� Ctrl+Enter ��ֱ�ӷ���' onkeydown="javascript:frm_quicksubmit();"></textarea></td></tr></table></td>
</tr>
<script language=javascript src='style/em_type.js'></script>
<tr height=30<%response.write forum_table3%>>
<td align=center bgcolor='<%=web_var(web_color,6)%>'>E M ��ͼ��</td>
<td>&nbsp;<script language=javascript>jk_em_type('s');</script></td>
</tr>
<tr<%response.write format_table(3,1)%>><td colspan=2 align=center height=60>&nbsp;&nbsp;<script language=javascript>jk_em_type('b');</script></td></tr>
<tr align=center height=30<%response.write forum_table3%>>
<td bgcolor='<%=web_var(web_color,6)%>'>���ٻظ���</td>
<td><input type=submit name=wsubmit value='���ٷ����ҵĻ���'>��&nbsp;<input type=button value='Ԥ���ҵĻظ�'>��&nbsp;<input type=reset value='�����д'>��&nbsp;���� Ctrl + Enter �ɿ��ٻظ���
<%
  end if
else
  response.write "<div align=center><font class=red_2>��������ѱ������������ٶ�����лظ�</font></div>"
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