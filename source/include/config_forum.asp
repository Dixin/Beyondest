<!-- #include file="config.asp" -->
<!-- #include file="skin.asp" -->
<%
'*******************************************************************

'

'                     Beyondest.Com V3.6 Demo版

' 




'           网址：http://www.beyondest.com

' 

'*******************************************************************

dim forum_mode,forum_table1,forum_table2,forum_table3,forum_table4,ptnums,ffk
dim forumid,viewid,forumname,forumpower,forumtype,forumtopicnum,forumdatanum,word_size,word_remark
forum_table1=format_table(1,3)
forum_table2=format_table(3,6)
forum_table3=format_table(3,5)
forum_table4=format_table(3,1)
forumid=trim(request.querystring("forum_id"))
viewid=trim(request.querystring("view_id"))

ffk="fk4"
index_url="forum"
tit_fir=format_menu(index_url)
ptnums=web_var_num(web_setup,6,1)

'-------------------------------------初始化 1--------------------------------------
sub forum_first()
  sql="select forum_name,forum_power,forum_topic_num,forum_data_num,forum_type " & _
      "from bbs_forum where forum_id="&forumid&" and forum_hidden=0"
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call close_conn()
    call cookies_type("forum_id")
  end if
  forumname=rs("forum_name"):forumpower=rs("forum_power")
  forumtopicnum=rs("forum_topic_num"):forumdatanum=rs("forum_data_num"):forumtype=rs("forum_type")
  rs.close
  page_power=format_forum_type(forumtype,0)
end sub

'-------------------------------------论坛标头--------------------------------------
function forum_top(ft)
  forum_top=vbcrlf & ukong&forum_table1 & _
	    vbcrlf & "<tr "&forum_table2&">" & _
	    vbcrlf & "<td width='70%'>&nbsp;&nbsp;" & _
	    vbcrlf & "讨论区：<a href='forum_list.asp?forum_id="&forumid&"'>"&forumname&"</a> &nbsp;- &nbsp;"&ft&"<font class=gray>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;【<a href='forum_list.asp?forum_id="&forumid&"&action=isgood'>本版精华</a>】&nbsp;【<a href='forum_list.asp?forum_id="&forumid&"&action=manage'>版面管理</a>】</font></td>" & _
	    vbcrlf & "<td align=right>版主："&forum_power(forumpower,ptnums)&"&nbsp;" & _
	    vbcrlf & "</td>" & _
	    vbcrlf & "</tr></table>" & _
            vbcrlf & ""&ukong
end function

'-------------------------------------数据生成--------------------------------------
sub forum_word()
  word_size=web_var(web_num,6)
  word_remark=web_var(web_error,3)&"<br>长度<="&word_size&"KB"
end sub

'-------------------------------------论坛版主--------------------------------------
function forum_power(forum_admin,ft)
  dim forumadmin,k
  forum_power="<img src='images/small/forum_power.gif' title='论坛版主' align=absmiddle border=0>&nbsp;"
  if ft=0 then forum_power=forum_power&"<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}""><option>本版版主</option><option>--------</option>"
  if forum_admin<>"" and not isnull(forum_admin) then
    forumadmin=split(forum_admin, "|")
    for k=0 to ubound(forumadmin)
      if ft=0 then
        forum_power=forum_power&"<option value='user_view.asp?username=" & server.urlencode(forumadmin(k)) & "'>"&forumadmin(k)&"</option>"
      else
        forum_power=forum_power & "<a href='user_view.asp?username=" & server.urlencode(forumadmin(k)) & "' title='查看（版主）" & forumadmin(k) & " 的详细资料' target=_blank>" & forumadmin(k) & "</a>&nbsp;"
      end if
    next
    erase forumadmin
  else
    if ft=0 then
      forum_power=forum_power&"<option>还没呢</option>"
    else
      forum_power=forum_power & "<font class=gray>还没呢&nbsp;</font>"
    end if
  end if
  if ft=0 then forum_power=forum_power&"</select>"
end function

'-------------------------------------论坛等级--------------------------------------
function format_forum_type(fvars,ft)
  dim fdim,fvar:fvar=fvars-1:format_forum_type=""
  fdim=split(forum_type,"|")
  for i=0 to ubound(fdim)
    if ft=0 then
      if fvar=i then format_forum_type=left(fdim(i),instr(fdim(i),":")-1):exit for
    else
      if fvar=i then format_forum_type=right(fdim(i),len(fdim(i))-instr(fdim(i),":")):exit for
    end if
  next
  erase fdim
end function

'-----------------------------------主题转移操作------------------------------------
sub forum_moved(fid,vid)
  if not(isnumeric(fid)) or not(isnumeric(vid)) or login_mode<>format_power2(1,1) then response.write "<script language=javascript>alert(""转移主题失败：\n\n可能是您进行了不适合的操作！"");</script>":exit sub
  dim frs,fsql
  fsql="select forum_id from bbs_topic where id="&vid
  set frs=conn.execute(fsql)
  if frs.eof and frs.bof then
    frs.close:set frs=nothing:close_conn
    call cookies_type("view_id"):exit sub
  end if
  frs.close:set frs=nothing
  fsql="update bbs_topic set forum_id="&fid&" where id="&vid
  conn.execute(fsql)
  fsql="update bbs_data set forum_id="&fid&" where reply_id="&vid
  conn.execute(fsql)
  response.write "<script language=javascript>alert(""转移主题成功！"");</script>"
end sub

'-------------------------------------主题转移--------------------------------------
function forum_move(fmfid,fmid)
  dim rsclass,strsqlclass,rsboard,strsqlboard,fid
  forum_move=vbcrlf & "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & _
	   vbcrlf & "<option selected>将此主题转移至...</option>"
  strsqlclass="select class_id,class_name from bbs_class order by class_order"
  set rsclass=conn.execute(strsqlclass)
  if not(rsclass.bof and rsclass.eof) then
    do while not rsclass.eof
      forum_move=forum_move & vbcrlf & "<option class=bg_2>╋ "& rsclass("class_name") &"</option>"
      strsqlboard="select forum_id,forum_name from bbs_forum where class_id=" & rsclass("class_id") & " and forum_hidden=0 order by forum_order"
      set rsboard=conn.execute(strsqlboard)
      if rsboard.eof and rsboard.bof then
        forum_move=forum_move & vbcrlf & "<option>没有论坛</option>"
      else
        do while not rsboard.eof
          fid=rsboard("forum_id")
          forum_move=forum_move & vbcrlf & "<option"
          if int(fid)<>int(fmfid) then  forum_move=forum_move&" value='forum_list.asp?action=move&view_id="&fmid&"&forum_id=" &fid& "'"
          forum_move=forum_move&">　├" & rsboard("forum_name") & "</option>"
	  rsboard.movenext
        loop
      end if
      rsclass.movenext
    loop
  end if
  set rsclass=nothing:set rsboard=nothing
  forum_move=forum_move & vbcrlf & "</select>"
end function

'-------------------------------------论坛跳转--------------------------------------
function forum_go()
  dim rsclass,strsqlclass,rsboard,strsqlboard
  forum_go=vbcrlf & "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & _
	   vbcrlf & "<option selected>快速跳转论坛至...</option>"
  strsqlclass="select class_id,class_name from bbs_class order by class_order"
  set rsclass=conn.execute(strsqlclass)
  if not(rsclass.bof and rsclass.eof) then
    do while not rsclass.eof
      forum_go=forum_go & vbcrlf & "<option class=bg_2>╋ "& rsclass("class_name") &"</option>"
      strsqlboard="select forum_id,forum_name from bbs_forum where class_id=" & rsclass("class_id") & " and forum_hidden=0 order by forum_order"
      set rsboard=conn.execute(strsqlboard)
      if rsboard.eof and rsboard.bof then
        forum_go=forum_go & vbcrlf & "<option>没有论坛</option>"
      else
        do while not rsboard.eof
          forum_go=forum_go & vbcrlf & "<option value='forum_list.asp?forum_id=" &rsboard("forum_id")& "'>　├" & rsboard("forum_name") & "</option>"
	  rsboard.movenext
        loop
      end if
      rsclass.movenext
    loop
  end if
  set rsclass=nothing:set rsboard=nothing
  forum_go=forum_go & vbcrlf & "<option class=bg_2>————————</option>" & _
	   vbcrlf & "<option value='forum.asp' class=bg_1>"&tit_fir&"首页</option>" & _
	   vbcrlf & "<option class=bg_2>————————</option>" & _
	   vbcrlf & "<option value='forum_action.asp?action=new'>　♀ 论坛新贴</option>" & _
	   vbcrlf & "<option value='forum_action.asp?action=tim'>　♀ 回复新贴</option>" & _
	   vbcrlf & "<option value='user_action.asp?action=list'>　♀ 用户列表</option>" & _
	   vbcrlf & "<option value='help.asp?action=forum'>　♀ 论坛帮助</option>" & _
	   vbcrlf & "</select>"
end function

'-------------------------------------主题分页--------------------------------------
function index_pagecute(viewurl,replynum,pagecutenum,pagecutecolor)
  dim pagecutepage,pagecutei
  index_pagecute=""
  if replynum mod pagecutenum > 0 then
    pagecutepage=replynum\pagecutenum+1
  else
    pagecutepage=replynum\pagecutenum
  end if
  if pagecutepage>1 then
    for pagecutei=2 to 3
      if pagecutei>pagecutepage then exit for
      index_pagecute=index_pagecute & vbcrlf & "<a href='" & viewurl & "&page=" & pagecutei & "'><font color='" & pagecutecolor & "' title='第 " & pagecutei & " 页'>[" & pagecutei & "]</font></a>"
    next
    if pagecutepage>3 then
      if pagecutepage=4 then
        index_pagecute=index_pagecute & vbcrlf & "<a href='" & viewurl & "&page=4'><font color='" & pagecutecolor & "' title='第 4 页'>[4]</font></a>"
      else
        index_pagecute=index_pagecute & vbcrlf & "<font color='" & pagecutecolor & "'>… </font>" & "<a href='" & viewurl & "&page=" & pagecutepage & "'><font color='" & pagecutecolor & "' title='第 " & pagecutepage & " 页'>[" & pagecutepage & "]</font></a>"
      end if
    end if
  end if
  if len(index_pagecute)>1 then index_pagecute="<img src='images/small/page_head.gif' align=absMiddle alt='快速分页' border=0>"&index_pagecute
end function

'---------------------------------------main----------------------------------------
sub forum_down(dt)
  dim udim,ui,j,dts,sql,rs,l_username,forum_table4,online
  online=trim(request.querystring("online"))
  j=5:dts=0:forum_table4=format_table(3,1)
  if forum_mode="full" then j=8
  if online="open" or dt=1 then dts=1
  if online="close" then dts=0
  response.write forum_table1
%>
<tr<%response.write forum_table2%> height=25><td background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small("fk4") %>&nbsp;<font class=end><b>论坛图例</b></font></td></tr>
<tr<%response.write forum_table4%>><td align=left height=30>&nbsp;&nbsp;<% response.write ip_sys(0,0) %></td></tr>
<tr<%response.write forum_table4%>><td align=left height=30>&nbsp;&nbsp;<%response.write user_power_type(0)%></td></tr>
<tr<%response.write forum_table4%>><td align=left>
  <table border=0 width='100%'>
  <tr><td colspan=5>&nbsp;网站当前用户在线：<font class=red><%
sql="select count(l_id) from user_login where l_type=0"
set rs=conn.execute(sql)
response.write rs(0)
rs.close
response.write "</font> 人  [ <a href='?mode="&forum_mode&"&online="
if dts=0 then
  response.write "open'>打开"
else
  response.write "close'>关闭"
end if
%>在线列表</a> ] </td></tr>
<%if dts<>0 then%>
  <tr><td width='20%'></td><td width='20%'></td><td width='20%'></td><td width='20%'></td><td width='20%'></td></tr>
<%
  sql="select user_login.*,user_data.power from user_data inner join user_login on user_login.l_username=user_data.username where user_login.l_type=0 order by user_login.l_id"
  set rs=conn.execute(sql)
  do while not rs.eof
    response.write "<tr>"
    for ui=1 to 5
      if rs.eof then exit for
      l_username=rs("l_username")
      response.write "<td>&nbsp;"&img_small("icon_"&rs("power"))&"<a href='user_view.asp?username="&server.urlencode(l_username)&"' title='目前位置："&rs("l_where")&"<br>来访时间："&rs("l_tim_login")&"<br>活动时间："&rs("l_tim_end")&"<br>"&ip_types(rs("l_ip"),l_username,0)&"<br>"&view_sys(rs("l_sys"))&"' target=_blank>"&l_username&"</a></td>"
      rs.movenext
    next
    response.write "</tr>"
  loop
  rs.close
end if
set rs=nothing
%>
  </table>
</td></tr>
</table>
<table border=0 width='100%'>
<tr><td align=center height=50>
<%
udim=split(forum_type,"|")
for ui=0 to ubound(udim)
  response.write vbcrlf&"&nbsp;<img src='images/small/label_"&ui+1&".gif' border=0 align=absmiddle>&nbsp;"&right(udim(ui),len(udim(ui))-instr(udim(ui),":"))&"&nbsp;"
next
erase udim
%>
</td></tr>
</table>
<%
  response.write kong
end sub

sub forum_cast(nh,nj,n_num,c_num)
  dim temp1,njj,topic,tbb:njj=""
  if nj<>"" then njj=img_small(nj)
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&n_num&" id,topic,username,tim from bbs_cast where sort='forum' order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("topic")
    temp1=temp1&"<tr><td class=bw height="&space_mod&">"&njj&"<a href='update.asp?action=forum&id="&rs("id")&"' target=_blank title='公告标题："&code_html(topic,1,0)&"<br>管 理 员："&rs("username")&"<br>发布时间："&time_type(rs("tim"),88)&"'>"&code_html(topic,1,c_num)&"</a></td></tr>"
    rs.movenext
  loop
  temp1=temp1&"</table>"
  response.write format_barc("<font class=end><b>论坛公告</b></font>",temp1,2,0,7)
end sub

sub forum_new(nh,nj,fid,n_num,c_num,tb)
  dim temp1,njj,topic,tbb:njj="":tbb=""
  if nj<>"" then njj=img_small(nj)
  if tb=1 then tbb=" target=_blank"
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&n_num&" id,forum_id,topic,tim,username,re_username from bbs_topic"
  if fid>0 then sql=sql&" where forum_id="&fid
  sql=sql&" order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("topic")
    temp1=temp1&"<tr><td class=bw height="&space_mod&">"&njj&"<a href='forum_view.asp?forum_id="&rs("forum_id")&"&view_id="&rs("id")&"'"&tbb&" title='贴子主题："&code_html(topic,1,c_num)&"<br>发 贴 人："&rs("username")&"<br>发贴时间："&time_type(rs("tim"),88)&"<br>最后回复："&rs("re_username")&"'>"&code_html(topic,1,c_num)&"</a></td></tr>"
    rs.movenext
  loop
  temp1=temp1&"</table>"
  response.write format_barc("<font class=end><b>论坛新贴</b></font>",temp1,2,0,8)
end sub

function forum_main(mh)
  dim rsclass,strsqlclass,rsforum,strsqlforum,rstopic,sqltopic,topics,classid,forumid,forumname,forum_type,forum_new_info,forum_pic,new_info_dim,forumpic
  strsqlclass="select class_id,class_name from bbs_class order by class_order"
  set rsclass=conn.execute(strsqlclass)
  do while not rsclass.eof
    classid=rsclass("class_id")
    response.write vbcrlf & forum_table1 & "<tr"&forum_table2&"><td height=25 colspan=4 background=images/"&web_var(web_config,5)&"/bar_3_bg.gif>&nbsp;" & img_small(mh) & vbcrlf & "<font class=end><b>" & rsclass("class_name") & "</b></font></td></tr>"
    strsqlforum="select forum_id,forum_name,forum_type,forum_new_info,forum_topic_num,forum_data_num,forum_power,forum_remark,forum_pic " & _
	        "from bbs_forum where class_id=" & classid & " and forum_hidden=0 order by forum_order,forum_id desc"
    set rsforum=conn.execute(strsqlforum)
    do while not rsforum.eof
      forumid=rsforum("forum_id"):forumname=rsforum("forum_name")
      forum_type=rsforum("forum_type")
      forum_new_info=rsforum("forum_new_info")
      forum_pic=rsforum("forum_pic")

      if len(forum_new_info)>3 then
        new_info_dim=split(forum_new_info,"|")
        new_info_dim(0)=format_user_view(new_info_dim(0),1,"")
        if isdate(new_info_dim(1)) then
          new_info_dim(1)=time_type(new_info_dim(1),8)
        end if
        new_info_dim(3)="<a href='forum_view.asp?forum_id="&forumid&"&view_id="&new_info_dim(2)&"' title='"&code_html(new_info_dim(3),0,0)&"'>"&code_html(new_info_dim(3),0,8)&"</a>"
      else
        redim new_info_dim(3)
      end if
      if len(forum_pic)>1 then
        if left(forum_pic,1)="$" then forum_pic="images/forum/"&right(forum_pic,len(forum_pic)-1)
        forum_pic="<td align=right><img src='"&forum_pic&"' border=0></td>"
      else
        forum_pic=""
      end if
      response.write vbcrlf&"<tr"&format_table(3,1)&"><td width='10%' rowspan=2 align=center><img src='images/small/label_"&forum_type&".gif' border=0></td>" & _
		     vbcrlf&"<td width='24%' align=center height=20 "&forum_table2&"><a href='forum_list.asp?forum_id="&forumid&"'>『 " & forumname & " 』</a></td>" & _
		     vbcrlf&"<td width='38%'"&forum_table2&">" & _
		     vbcrlf&"  <table border=0 width='100%'><tr align=center>" & _
		     vbcrlf&"  <td width='45%'>论坛主题数&nbsp;&nbsp;<font class=red_3>"&rsforum("forum_topic_num")&"</font></td>" & _
		     vbcrlf&"  <td width='45%'>论坛贴子数&nbsp;&nbsp;<font class=red_3>"&rsforum("forum_data_num")&"</font></td>" & _
		     vbcrlf&"  <td width='10%'></td><td width='16%'><a href='forum_write.asp?forum_id="&forumid&"'><img src='images/small/mini_write.gif' align=absmiddle title='发表主题' border=0></a></td></tr></table>" & _
		     vbcrlf&"</td>" & _
		     vbcrlf&"<td width='30%'"&forum_table2&">版主："&forum_power(rsforum("forum_power"),ptnums)&"</td></tr>" & _
		     vbcrlf&"<tr"&format_table(3,1)&"><td colspan=2 align=center><table border=0 width='99%'><tr><td class=htd>"&code_html(rsforum("forum_remark"),2,0)&"</td>"&forum_pic&"</tr></table></td>" & _
		     vbcrlf&"<td align=left valign=top class=htd>新贴："&new_info_dim(3)&"<br>作者："&new_info_dim(0)&"<br>时间："&new_info_dim(1)&"</td></tr>"
      erase new_info_dim

      rsforum.movenext
    loop
    rsclass.movenext
    response.write "</table>" & kong
  loop
  set rsclass=nothing:set rsforum=nothing
end function

'------------------------------------forum_list-------------------------------------
sub forum_view()
  select case int(istop)
  case 1
    folder_type="istop"
  case 2
    folder_type="istops"
  case else
    if int(isgood)=1 then
      folder_type="isgood"
    else
      if int(islock)=1 then
        folder_type="islock"
      elseif int(re_counter)>=10 then
        folder_type="ishot"
      end if
    end if
  end select

  view_url="forum_view.asp?forum_id="&forumnid&"&view_id="&id
  if int(re_counter)>0 then
    topic_head="<img loaded=no src='images/small/fk_plus.gif' border=0 id=followImg"&id&" style=""cursor:hand;"" onclick=""load_tree("&forumnid&","&id&")"" title='展开贴子列表'>"
  else
    topic_head="<img src='images/small/fk_minus.gif' border=0 id=followImg"&id&">"
  end if
  
  response.write vbcrlf&"<tr align=center"&format_table(3,1)&">" & _
  		 vbcrlf&"<td bgcolor="&web_var(web_color,5)&"><img src='images/small/"&folder_type&".gif' border=0></td>" & _
  		 vbcrlf&"<td bgcolor="&web_var(web_color,6)&">"
  if action="manage" then
    response.write "<input type=checkbox name=del_id value='"&id&"' class=bg_3>"
  else
    response.write "<img src='images/icon/"&icon&".gif' border=0>"
  end if
  response.write "</td>" & _
  		 vbcrlf&"<td align=left>"&topic_head&"<a href='"&view_url&"' title='主题："&code_html(topic,1,0)&"<br>发贴时间："&tim&"<br>最后回复："&re_username&"<br>回复时间："&re_tim&"'>"&code_html(topic,0,22)&"</a>&nbsp;"&index_pagecute(view_url,re_counter+1,web_var(web_num,3),"#cc3300")&"</td>" & _
  		 vbcrlf&"<td bgcolor="&web_var(web_color,6)&">"&format_user_view(username,1,"")&"</td>" & _
  		 vbcrlf&"<td><a href='"&view_url&"' target=_blank><img src='images/small/new_win.gif' alt='打开新窗口浏览此贴' border=0 width=13 height=11></a></td>" & _
  		 vbcrlf&"<td bgcolor="&web_var(web_color,6)&">"&re_counter&"<font class=gray>/</font>"&counter&"</td>" & _
  		 vbcrlf&"<td align=left><font class=timtd>"&time_type(re_tim,6)&"</font><font class=red>│</font>"&format_user_view(re_username,1,"")&"</td>" & _
  		 vbcrlf&"</tr>" & _
  		 vbcrlf&"<tr"&format_table(3,1)&" style=""display:none"" id=follow"&id&" height=30><td colspan=2>&nbsp;</td><td colspan=6 id=followTd"&id&" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:"&web_var(web_color,5)&";color:"&web_var(web_color,7)&";padding:2px"" onclick=""load_tree("&forumnid&","&id&")"">正在读取关于本主题的跟贴，请稍侯……</div></td></tr>"
  del_temp=del_temp+1
end sub

'------------------------------------forum_view-------------------------------------
function view_type()
  dim up:up=int(popedom_format(u_popedom,42))
  table_bg=format_table(3,1)
  if ii mod 2=0 then table_bg=forum_table3
  if var_null(u_whe)<>"" then
    u_whe="来自："&u_whe&"<br>"
  end if
  if var_null(u_nname)<>"" then
    u_nname="头衔："&u_nname&"<br>"
  end if
  
  view_type=vbcrlf & "<tr align=center valign=top"&table_bg&"><td width='20%' bgcolor='"&web_var(web_color,6)&"'>" & _
	    vbcrlf & "<table border=0 width='94%'><tr><td align=center height=30><table border=0><tr><td><font class=blue><b>"&u_username&"</b></font></td><td>&nbsp;"&user_view_sex(u_sex)&"</td></tr></table></td></tr>" & _
	    vbcrlf & "<tr><td align=center height=96><img src='images/face/"&rs("u_face")&".gif' border=0></td></tr>" & _
	    vbcrlf & "<tr><td height=15><img src='images/star/star_"&user_star(u_integral,u_power,1)&".gif' border=0></td></tr>" & _
	    vbcrlf & "<tr><td>等级："&user_view_power(u_power,0)&user_star(u_integral,u_power,2)&"<br>"&u_nname&"发贴："&u_bbs_counter&"<br>积分："&u_integral&"<br>"&u_whe&"注册："&formatdatetime(rs("u_tim"),2)&"</td></tr>" & _
	    vbcrlf & "</table></td><td width='80%' height='100%'>" & _
	    vbcrlf & "<table border=0 width='99%' cellspacing=2 cellpadding=0 height='100%'><tr height=25><td width='85%'>" & _
	    vbcrlf & "<a target=_blank href='user_view.asp?username="&server.urlencode(u_username)&"'><img src='images/small/forum_profile.gif' title='查看 "&u_username&" 的详细信息' border=0></a>&nbsp;" & _
	    vbcrlf & "<a target=_blank href='user_friend.asp?action=add&add_username="&server.urlencode(u_username)&"'><img src='images/small/forum_friend.gif' title='将 "&u_username&" 加为我的好友' border=0></a>&nbsp;" & _
	    vbcrlf & "<a target=_blank href='user_message.asp?action=write&accept_uaername="&server.urlencode(u_username)&"'><img src='images/small/forum_message.gif' title='给 "&u_username&" 发短信' border=0></a>&nbsp;" & _
	    vbcrlf & "<a href='forum_edit.asp?forum_id="&forumid&"&edit_id="&qid&"'><img src='images/small/forum_edit.gif' title='编辑这个贴子' border=0></a>&nbsp;"
  if int(fir_islock)<>1 then
    view_type=view_type&vbcrlf&"<a href='forum_reply.asp?forum_id="&forumid&"&quote=yes&view_id="&qid&"'><img src='images/small/forum_quote.gif' title='引用并回复这个贴子' border=0></a>&nbsp;" & _
	      vbcrlf&"<a href='forum_reply.asp?forum_id="&forumid&"&view_id="&qid&"'><img src='images/small/forum_reply.gif' title='回复这个贴子' border=0></a>"
  else
    view_type=view_type&vbcrlf&"<img src='images/small/forum_reply.gif' title='这个贴子已被锁定' border=0>"
  end if
  view_type=view_type&vbcrlf & "</td><td width='15%' align=center>第 <font class=red_3><b>"&ii+(viewpage-1)*nummer&"</b></font> 楼</td></tr>" & _
	    vbcrlf & "<tr><td colspan=2 height=1 bgcolor="&web_var(web_color,3)&"></td></tr>"
	    
  if up=0 then
    view_type=view_type&vbcrlf & "<tr><td colspan=2 valign=top align=center>" & _
	      vbcrlf & "<table border=0 width='98%' class=tf><tr><td height=30>" & _
	      vbcrlf & "<img src='images/icon/"&rs("icon")&".gif' align=absMiddle border=0>&nbsp;<font class=red_3><b>"&code_html(rs("topic"),1,0)&"</b></font></td></tr>" & _
	      vbcrlf & "<tr><td class=bw><font class=htd>"&code_jk(rs("word"))&"</font></td></tr>" & _
	      vbcrlf & "</table></td></tr>" & _
	      vbcrlf & "<tr><td colspan=2 height=20 align=right>"&img_small("signature")&"</td></tr>" & _
	      vbcrlf & "<tr><td colspan=2 height=30 align=center valign=top><table border=0 width='96%' class=tf><tr><td class=bw><font class=htd>"&u_remark&"</font></a></td></tr></table></td></tr>"
  else
    view_type=view_type&vbcrlf & "<tr><td colspan=2 valign=top><br><table border=0><tr><td class=htd><font class=red_2>========================<br>&nbsp;&nbsp;该用户的论坛发言已暂时被管理屏蔽！<br>========================</font></td></tr></table></td></tr>"
  end if
  view_type=view_type&vbcrlf & "<tr><td height=25 colspan=2><table border=0 width='100%'><tr><td>"&img_small("forum_tim")&"<font class=gray>本贴发表时间："&rs("tim")&"</font></td><td align=right>"&ip_types(rs("ip"),u_username,1)&"　<img src='images/small/sys.gif' align=absMiddle title='"&view_sys(rs("sys"))&"' border=0>　<a href=""javascript:"&del_type&"('"&forumid&"','"&iid&"');""><img src='images/small/forum_del.gif' align=absMiddle border=0></a></td></tr></table></td></tr>" & _
	    vbcrlf & "</table>"
end function

function user_view_sex(us)
  if us=false then
    user_view_sex="<img src='images/small/forum_girl.gif' align=absmiddle title='青春女孩' border=0>":exit function
  else
    user_view_sex="<img src='images/small/forum_boy.gif' align=absmiddle title='阳光男孩' border=0>":exit function
  end if
end function

function user_view_power(uvp,ut)
  user_view_power=img_small("icon_"&uvp)
  if ut=1 then user_view_power=user_view_power&"<font class=red_3>"&format_power(uvp,1)&"</font>"
end function
%>