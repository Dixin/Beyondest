<!-- #include file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/conn.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim nsort,sql2,rs2,del_temp,data_name,cid,sid,nid,ncid,nsid,id,left_type,now_id,nummer,sqladd,page,rssum,thepages,viewpage,pageurl,csid
tit=vbcrlf & "<a href='?'>��ҵ��̬</a>&nbsp;��&nbsp;" & _
    vbcrlf & "<a href='?action=add'>��������</a>&nbsp;��&nbsp;" & _
    vbcrlf & "<a href='admin_nsort.asp?nsort=news'>���ŷ���</a>"
response.write header(12,tit)
pageurl="?action="&action&"&":nsort="news":data_name="news":sqladd="":nummer=15
call admin_cid_sid()

if trim(request("del_ok"))="ok" then
  response.write del_select(trim(request.form("del_id")))
end if

id=trim(request.querystring("id"))
if (action="hidden" or action="istop") and isnumeric(id) then
  sql="select "&action&" from "&data_name&" where id="&id
  set rs=conn.execute(sql)
  if not(rs.eof and rs.bof) then
    if rs(action)=true then
      sql="update "&data_name&" set "&action&"=0 where id="&id
    else
      sql="update "&data_name&" set "&action&"=1 where id="&id
    end if
    conn.execute(sql)
  end if
  rs.close:action=""
end if

select case action
case "add"
  call news_add()
case "edit"
  if not(isnumeric(id)) then
    call news_main()
  else
    set rs=server.createobject("adodb.recordset")
    sql="select * from "&data_name&" where id="&id
    rs.open sql,conn,1,3
    call news_edit()
  end if
case else
  call news_main()
end select

call close_conn()
response.write ender()

sub news_edit()
  dim rs3,sql3,topic,comto,istop,word,ispic,pic,keyes
  if trim(request.querystring("edit"))="chk" then
    topic=code_admin(request.form("topic"))
    csid=trim(request.form("csid"))
    comto=code_admin(request.form("comto"))
    keyes=code_admin(request.form("keyes"))
    istop=trim(request.form("istop"))
    word=request.form("word")
    ispic=trim(request.form("ispic"))
    pic=trim(request.form("pic"))
    if len(csid)<1 then
      response.write "<font class=red_2>��ѡ���������ͣ�</font><br><br>"&go_back
    elseif len(topic)<1 or len(word)<10 then
      response.write "<font class=red_2>���ű�������ݲ���Ϊ�գ�</font><br><br>"&go_back
    else
      call chk_cid_sid()
      rs("c_id")=cid
      rs("s_id")=sid
      if trim(request.form("username_my"))="yes" then rs("username")=login_username
      rs("topic")=topic
      rs("comto")=comto
      rs("keyes")=keyes
      rs("word")=word
      if istop="yes" then
        rs("istop")=true
      else
        rs("istop")=false
      end if
      if ispic="yes" then
        rs("ispic")=true
      else
        rs("ispic")=false
      end if
      if trim(request.form("hidden"))="yes" then
        rs("hidden")=false
      else
        rs("hidden")=true
      end if
      rs("pic")=pic
      if isnumeric(trim(request.form("counter"))) then rs("counter")=trim(request.form("counter"))
      rs.update
      rs.close:set rs=nothing
      call upload_note(data_name,id)
      response.write "<font class=red>�ѳɹ��޸���һƪ���ţ�</font><br><br><a href='?c_id="&cid&"&s_id="&sid&"'>�������</a><br><br>"
    end if
  else
%><table border=0 width='98%' cellspacing=0 cellpadding=1>
<form name='add_frm' action='<%response.write pageurl%>c_id=<%response.write cid%>&s_id=<%response.write sid%>&id=<%response.write id%>&edit=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>���ű��⣺</td><td width='85%'><input type=text size=70 name=topic value='<%response.write rs("topic")%>' maxlength=100><%=redx%></td></tr>
  <tr><td align=center>�������ͣ�</td><td><%call chk_csid(cid,sid)%>&nbsp;&nbsp;&nbsp;������<input type=text size=20 name=comto value='<%response.write rs("comto")%>' maxlength=10>&nbsp;&nbsp;&nbsp;<input type=checkbox name=username_my value='yes'>&nbsp;<font alt='�����ˣ�<%response.write rs("username")%>'>�޸ķ�����Ϊ��</font></td></tr>
<%
    pic=rs("pic"):ispic=pic
    if Instr(ispic,"/")>0 then ispic=right(ispic,len(ispic)-Instr(ispic,"/"))
    if Instr(ispic,".")>0 then ispic=left(ispic,Instr(ispic,".")-1)
    if len(ispic)<1 then ispic="n"&upload_time(now_time)
%>  <tr><td align=center>�� �� �֣�</td><td><input type=text size=20 name=keyes value='<%response.write rs("keyes")%>' maxlength=20>&nbsp;&nbsp;&nbsp;�Ƽ���<input type=checkbox name=istop<%if rs("istop")=true then response.write " checked"%> value='yes'>&nbsp;ѡΪ�Ƽ���ʾ&nbsp;&nbsp;&nbsp;���أ�<input type=checkbox name=hidden<%if rs("hidden")=false then response.write " checked"%> value='yes'>&nbsp;ѡΪ������ʾ</td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","word","&nbsp;&nbsp;")%></td></tr>
  <tr><td align=center valign=top><br>�������ݣ�</td><td><textarea name=word rows=15 cols=70><%response.write rs("word")%></textarea></td></tr>
  <tr><td align=center>ͼƬ���ţ�</td><td><input type=checkbox name=ispic<%if rs("ispic")=true then response.write " checked"%> value='yes'>&nbsp;ѡΪͼƬ����&nbsp;&nbsp;&nbsp;ͼƬ��<input type=test name=pic value='<%response.write pic%>' size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=<%response.write ispic%>&uptext=pic' target=upload_frame>�ϴ�ͼƬ</a>&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=n&uptext=word' target=upload_frame>�ϴ�������</a></td></tr>
  <tr><td align=center>�ϴ�ͼƬ��</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=news&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' �� �� �� �� '></td></tr>
</form></table><%
  end if
end sub

sub news_add()
  if trim(request.querystring("add"))="chk" then
    dim topic,comto,istop,word,ispic,pic,keyes
    topic=code_admin(request.form("topic"))
    csid=trim(request.form("csid"))
    comto=code_admin(request.form("comto"))
    keyes=code_admin(request.form("keyes"))
    istop=trim(request.form("istop"))
    word=request.form("word")
    ispic=trim(request.form("ispic"))
    pic=trim(request.form("pic"))
    if len(csid)<1 then
      response.write "<font class=red_2>��ѡ���������ͣ�</font><br><br>"&go_back
    elseif len(topic)<1 or len(word)<10 then
      response.write "<font class=red_2>���ű�������ݲ���Ϊ�գ�</font><br><br>"&go_back
    else
      call chk_cid_sid()
      set rs=server.createobject("adodb.recordset")
      sql="select * from "&data_name
      rs.open sql,conn,1,3
      rs.addnew
      rs("c_id")=cid
      rs("s_id")=sid
      rs("username")=login_username
      rs("hidden")=true
      rs("topic")=topic
      rs("comto")=comto
      rs("keyes")=keyes
      rs("word")=word
      if istop="yes" then
        rs("istop")=true
      else
        rs("istop")=false
      end if
      if ispic="yes" then
        rs("ispic")=true
      else
        rs("ispic")=false
      end if
      rs("pic")=pic
      rs("tim")=now_time
      rs("counter")=0
      rs.update
      rs.close:set rs=nothing
      call upload_note(data_name,first_id(data_name))
      response.write "<font class=red>�ѳɹ�������һƪ���ţ�</font><br><br><a href='?c_id="&cid&"&s_id="&sid&"'>�������</a><br><br>"
    end if
  else
%><table border=0 width='98%' cellspacing=0 cellpadding=1>
<form name='add_frm' action='<%response.write pageurl%>add=chk' method=post>
<input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>���ű��⣺</td><td width='85%'><input type=text size=70 name=topic maxlength=100><%=redx%></td></tr>
  <tr><td align=center>�������ͣ�</td><td><%call chk_csid(cid,sid)%>&nbsp;&nbsp;&nbsp;&nbsp;������<input type=text size=30 name=comto maxlength=10></td></tr>
  <tr><td align=center>�� �� �֣�</td><td><input type=text size=20 name=keyes maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;�Ƽ���<input type=checkbox name=istop value='yes'>&nbsp;ѡ��Ϊ������ҳ��ʾ</td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","word","&nbsp;&nbsp;")%></td></tr>
  <tr><td valign=top align=center><br>�������ݣ�</td><td><textarea name=word rows=15 cols=70></textarea></td></tr>
<%ispic="n"&upload_time(now_time)%>
  <tr><td align=center>ͼƬ���ţ�</td><td><input type=checkbox name=ispic value='yes'>&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ��<input type=test name=pic size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=<%response.write ispic%>&uptext=pic' target=upload_frame>�ϴ�ͼƬ</a>&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=n&uptext=word' target=upload_frame>�ϴ�������</a></td></tr>
  <tr><td align=center>�ϴ�ͼƬ��</td><td><iframe frameborder=0 name=upload_frame width='100%' height=30 scrolling=no src='upload.asp?uppath=news&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' �� �� �� �� '></td></tr>
</form></table><%
  end if
end sub

sub news_main()
%>
<script language=javascript src='STYLE/admin_del.js'></script>
<table border=0 width='100%' cellpadding=2>
  <tr valign=top height=350>
    <td width='25%' class=htd><br><%call left_sort()%></td>
    <td width='75%' align=center>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<form name=del_form action='<%=pageurl%>del_ok=ok' method=post>
<tr><td width='6%'></td><td width='81%'></td><td width='13%'></td></tr>
<%
  call sql_cid_sid()
  sql="select id,c_id,s_id,topic,istop,hidden from "&data_name&sqladd&" order by id desc"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
    rssum=0
  else
    rssum=rs.recordcount
  end if
  call format_pagecute()
  del_temp=nummer
  if rssum=0 then del_temp=0
  if int(page)=int(thepages) then
    del_temp=rssum-nummer*(thepages-1)
  end if
%>
<tr><td colspan=3 align=center height=25>
����<font class=red><%response.write rssum%></font>ƪ���š�<%response.write "<a href='?action=add&c_id="&cid&"&s_id="&sid&"'>�������</a>"%>
��<input type=checkbox name=del_all value=1 onClick=selectall('<%response.write del_temp%>')> ѡ�����С�<input type=submit value='ɾ����ѡ' onclick=""return suredel('<%response.write del_temp%>');"">
</td></tr>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<%
  if int(viewpage)<>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    now_id=rs("id"):ncid=rs("c_id"):nsid=rs("s_id")
    response.write news_center()
    rs.movenext
  next
  rs.close:set rs=nothing
%></form>
<tr><td colspan=3 height=1 bgcolor=#ededede></td></tr>
<tr><td colspan=3 height=25>ҳ�Σ�<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font>
��ҳ��<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%>
</td></tr>
</table>
    </td>
  </tr>
</table>
<%
end sub

function news_center()
  news_center=VbCrLf & "<tr"&mtr&">" & _
		 VbCrLf & "<td>" & i+(viewpage-1)*nummer & ". </td><td>" & _
		 VbCrLf & "<a href='?action=edit&c_id="&ncid&"&s_id="&nsid&"&id=" & now_id & "'>" & cuted(rs("topic"),30) & "</a>" & _
		 "</td><td align=right><a href='?action=hidden&c_id="&cid&"&s_id="&sid&"&id="&now_id&"&page="&viewpage&"'>"
  if rs("hidden")=true then
    news_center=news_center&"��"
  else
    news_center=news_center&"<font class=red_2>��</font>"
  end if
  news_center=news_center&"</a> <a href='?action=istop&c_id="&cid&"&s_id="&sid&"&id="&now_id&"&page="&viewpage&"'>"
  if rs("istop")=true then
    news_center=news_center&"<font class=red>��</font>"
  else
    news_center=news_center&"��"
  end if
  news_center=news_center&"</a><input type=checkbox name=del_id value='"&now_id&"'></td></tr>"
end function
%>