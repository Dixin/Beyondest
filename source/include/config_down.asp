<!-- #include file="config.asp" -->
<!-- #include file="config_nsort.asp" -->
<!-- #include file="skin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim atb,nid,sqladd,name
atb=" target=_blank":sk_bar=12:sk_class="end"
index_url="down":n_sort="down"
tit_fir=format_menu(index_url)

sub down_class_sort(t1,t2)
  response.write class_sortp(n_sort,index_url,t1,t2)
end sub

sub down_intro(introid,introsn)
  dim tempix,sqlx,theintrox,thepicx,rsx
  tempix="<table border=0 width='100%' cellspacing=0 cellpadding=12><tr><td width='40%' align=center valign=top>"
  sqlx="select intro,pic from jk_sort where s_id="&introid
  set rsx=conn.execute(sqlx)
  theintrox=rsx(0)
  thepicx=rsx(1)
  tempix=tempix&"<img src=images/down/"&thepicx&".jpg></td><td>"&kong&"<font class=big><b>"&introsn&"</b></font>"&kong&"&nbsp;&nbsp;&nbsp;&nbsp;"&code_jk(theintrox)&"</td></tr></table>"
  rsx.close:set rsx=nothing
  response.write tempix

  
end sub

sub down_class_sortt(t1,t2)
  response.write format_barc("<font class="&sk_class&"><b>ר���б�</b></font>",class_sort(n_sort,index_url,t1,t2),3,0,6)
end sub

sub down_new_hot(n_jt,nnhead,nmore,nsql,nt,n_num,n_m,c_num,et,tt)
  dim rs,sql,di,temp1,tim,counter,nhead:nhead=nnhead
  if n_jt<>"" then n_jt=img_small(n_jt)
  temp1=vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&n_num+n_m&" id,name,username,tim,counter from down where hidden=1"&nsql
  select case nt
  case "hot"
    sql=sql&" order by counter desc,id desc"
    if nhead="" then nhead="��������"
  case "good"
    sql=sql&" and types=5 order by id desc"
    if nhead="" then nhead="�����Ƽ�"
  case else
    sql=sql&" order by id desc"
    if nhead="" then nhead="���ڸ���"
  end select
  set rs=conn.execute(sql)
  for di=1 to n_m
    if rs.eof or rs.bof then exit for
    rs.movenext
  next
  'if n_m>0 then rs.move(n_m)
  do while not rs.eof
    name=rs("name"):tim=rs("tim"):counter=rs("counter")
    temp1=temp1&vbcrlf&"<tr><td height="&space_mod&" class=bw>"&n_jt&"<a href='down_view.asp?id="&rs("id")&"'"&atb&" title='�������ƣ�"&code_html(name,1,0)&"<br>�� �� �ˣ�"&rs("username")&"<br>�����˴Σ�"&counter&"<br>����ʱ�䣺"&time_type(tim,88)&"'>"&code_html(name,1,c_num)&"</a>"
    if tt>0 then temp1=temp1&format_end(et,time_type(tim,tt)&",<font class=blue>"&counter&"</font>")
    temp1=temp1&"</td></tr>"
    rs.movenext
  loop
  rs.close:set rs=nothing
  temp1=temp1&vbcrlf&"</table>"
  response.write kong&format_barc("<font class="&sk_class&"><b>"&nhead&"</b></font>",temp1,2,0,8)
end sub

sub down_new_hotr(n_jt,nnhead,nmore,nsql,nt,n_num,n_m,c_num,et,tt)
  dim rs,sql,di,temp1,tim,counter,nhead:nhead=nnhead
  if n_jt<>"" then n_jt=img_small(n_jt)
  if n_jt="" then n_jt=img_small("jt0")
  temp1=vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&n_num+n_m&" id,name,username,tim,counter,order from down where hidden=1"&nsql
  select case nt
  case "hot"
    sql=sql&" order by counter desc,id desc"
    if nhead="" then nhead="��������"
  case "good"
    sql=sql&" and types=5 order by id desc"
    if nhead="" then nhead="�����Ƽ�"
  case else
    sql=sql&" order by [order],id"
    if nhead="" then nhead="���ڸ���"
  end select
  set rs=conn.execute(sql)
  for di=1 to n_m
    if rs.eof or rs.bof then exit for
    rs.movenext
  next
  'if n_m>0 then rs.move(n_m)
  do while not rs.eof
    name=rs("name"):tim=rs("tim"):counter=rs("counter")
    temp1=temp1&vbcrlf&"<tr><td height="&space_mod&" class=bw>"&n_jt&"<a href='down_view.asp?id="&rs("id")&"'"&atb&" title='�������ƣ�"&code_html(name,1,0)&"<br>�� �� �ˣ�"&rs("username")&"<br>�����˴Σ�"&counter&"<br>����ʱ�䣺"&time_type(tim,88)&"'>"&code_html(name,1,c_num)&"</a>"
    if tt>0 then temp1=temp1&format_end(et,time_type(tim,tt)&",<font class=blue>"&counter&"</font>")
    temp1=temp1&"</td></tr>"
    rs.movenext
  loop
  rs.close:set rs=nothing
  temp1=temp1&vbcrlf&"</table>"
  response.write format_barc("<font class="&sk_class&"><b>"&nhead&"</b></font>",temp1,3,0,8)
end sub


sub down_new_hotrn(n_jt,nnhead,nmore,nsql,nt,n_num,n_m,c_num,et,tt)
  dim rs,sql,di,temp1,tim,counter,nhead:nhead=nnhead
  if n_jt<>"" then n_jt=img_small(n_jt)
  temp1=vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&n_num+n_m&" id,name,username,tim,counter from down where hidden=1"&nsql
  select case nt
  case "hot"
    sql=sql&" order by counter desc,id desc"
    if nhead="" then nhead="��������"
  case "good"
    sql=sql&" and types=5 order by id desc"
    if nhead="" then nhead="�����Ƽ�"
  case else
    sql=sql&" order by id desc"
    if nhead="" then nhead="���ڸ���"
  end select
  set rs=conn.execute(sql)
  for di=1 to n_m
    if rs.eof or rs.bof then exit for
    rs.movenext
  next
  'if n_m>0 then rs.move(n_m)
  do while not rs.eof
    name=rs("name"):tim=rs("tim"):counter=rs("counter")
    temp1=temp1&vbcrlf&"<tr><td height="&space_mod&" class=bw>"&n_jt&"<a href='down_view.asp?id="&rs("id")&"'"&atb&" title='�������ƣ�"&code_html(name,1,0)&"<br>�� �� �ˣ�"&rs("username")&"<br>�����˴Σ�"&counter&"<br>����ʱ�䣺"&time_type(tim,88)&"'>"&code_html(name,1,c_num)&"</a>"
    if tt>0 then temp1=temp1&format_end(et,time_type(tim,tt)&",<font class=blue>"&counter&"</font>")
    temp1=temp1&"</td></tr>"
    rs.movenext
  loop
  rs.close:set rs=nothing
  temp1=temp1&vbcrlf&"</table>"
  response.write format_barc("<font class="&sk_class&"><b>"&nhead&"</b></font>",temp1,1,1,8)
end sub



sub down_pic(nnhead,dsql,nt,n_num,c_num)
  dim rs,sql,temp1,nhead:nhead=nnhead
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2><tr align=center valign=top>"
  sql="select top "&n_num&" id,name,tim,pic from down where hidden=1"&dsql
  select case nt
  case "hot"
    sql=sql&" order by counter desc,id desc"
    if nhead="" then nhead="�ȵ�����"
  case "good"
    sql=sql&" and types=5 order by id desc"
    if nhead="" then nhead="��Ʒ�Ƽ�"
  case else
    sql=sql&" order by id desc"
    if nhead="" then nhead="��������"
  end select
  set rs=conn.execute(sql)
  do while not rs.eof
    name=rs("name"):nid=rs("id")
    temp1=temp1&vbcrlf&"<td width='"&int(100\n_num)&"%'><table border=0 cellspacing=0 cellpadding=2 width='100%' class=tf><tr><td align=center><a href='down_view.asp?id="&nid&"'"&atb&"><img src='images/down/"&rs("pic")&"' border=0 ></a></td></tr>" & _
	  vbcrlf&"<tr><td align=center class=bw><a href='down_view.asp?id="&nid&"'"&atb&" class=red_3><b>"&code_html(name,1,0)&"</b></a></td></tr></table></td>"
    rs.movenext
  loop
  if temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2><tr align=center valign=top>" then temp1=temp1&"<td>��</td>"
  rs.close:set rs=nothing
  temp1=temp1&"</tr></table>"
  response.write format_barc("<font class="&sk_class&"><b>"&nhead&"</b></font>",temp1,3,0,5)
end sub

sub down_remark(njt)
  dim temp1
  temp1=vbcrlf&"<table border=0 width='98%' align=center>" & _
	vbcrlf&"<tr><td>"&img_small(njt)&"��վ�Ƽ�ʹ�� <a href='file/soft/flashget.rar'>���ʿ쳵</a> �������֣�һ������������ء�</td></tr>" & _
	vbcrlf&"<tr><td>"&img_small(njt)&"��������ֱ�վ���κ�������������⣬��<a href='gbook.asp?action=write'"&atb&">����֪ͨ��</a>��лл��</td></tr>" & _
	vbcrlf&"<tr><td>"&img_small(njt)&"��վ������ļ����� <a href='"&web_var(web_down,5)&"/soft/winrar.exe'>WinRAR</a> ѹ�������ڴ��������°汾��</td></tr>" & _
	vbcrlf&"<tr><td class=red>"&img_small(njt)&"��������ӱ�վ�ļ�����ע�����ԣ�<a href='"&web_var(web_config,2)&"'"&atb&">"&web_var(web_config,1)&"</a>��лл����֧�֣�</td></tr>" & _
	vbcrlf&"<tr><td>"&img_small(njt)&"��վ�ṩ���������ؽ���������������Ȩ���뼰ʱ <a href='gbook.asp?action=write'"&atb&">֪ͨ��</a> ��<font color='#ff0000'>ϣ�����֧�����档</font></td></tr>" & _
	vbcrlf&"<tr><td>"&img_small(njt)&"��ӭ��ҵ���վ <a href='forum.asp'>��̳</a> ����ͽ������ļ��⡣��л���ķ��ʣ�</td></tr>" & _
	vbcrlf&"</table>"
  response.write format_barc("<font class="&sk_class&"><b>��������˵��</b></font>",temp1,4,1,"")
end sub

sub down_tool()
  dim temp1
  temp1=vbcrlf&"<table border=0 cellspacing=0 cellpadding=2><tr><td height=5></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_winrar.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/winrar.exe'>WinRAR</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_qq.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/qq.rar'>QQ2004(ȥ�����IP)</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_winamp.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/winamp.rar'>Winamp</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_realone.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/realoneplayer.rar'>RealOnePlayer</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_wmp.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/wmp2k.rar'>Windows Midia Player(2k&98)</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_wmp.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/wmpxp.rar'>Windows Midia Player(xp)</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_flashget.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/flashget.rar'>Flashget</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_cuteftp.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/flashfxp.rar'>FlashFXP</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_wopti.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/wom.rar'>Windows�Ż���ʦ</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_norton.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/norton.rar'>Norton Antivirus 2004</a></td></tr>" & _
	vbcrlf&"<tr><td><img src='images/down/tool_norton.gif' border=0 align=absmiddle>&nbsp;<a href='"&web_var(web_down,5)&"/soft/nortonsp.rar'>Norton���²�����</a></td></tr>" & _
	vbcrlf&"</table>"
  response.write format_barc("<font class="&sk_class&"><b>���ù���</b></font>",temp1,1,1,1)
end sub

sub down_atat()
  dim temp1,num1,num2,num3,sq,rs
  sql="select count(id) from down where hidden=1 and tim>=#"&formatdatetime(formatdatetime(now_time,2))&"#"
  set rs=conn.execute(sql)
  num1=rs(0)
  rs.close
  sql="select num_down from configs where id=1"
  'sql="select count(id) from down where hidden=1"
  set rs=conn.execute(sql)
  num2=rs(0)
  rs.close
  sql="select sum(counter) from down where hidden=1"
  set rs=conn.execute(sql)
  num3=rs(0)
  rs.close:set rs=nothing
  temp1=vbcrlf&"<table border=0 cellspacing=0 cellpadding=3><tr><td height=5></td></tr>" & _
	vbcrlf&"<tr><td>���ո��£�<font class=red>"&num1&"</font>������</td></tr>" & _
	vbcrlf&"<tr><td>����������<font class=red>"&num2&"</font>������</td></tr>" & _
	vbcrlf&"<tr><td>�����أ�<font class=red>"&num3&"</font>�˴�</td></tr>" & _
	vbcrlf&"<tr><td>[ <a href='down_list.asp'>�� ������ַ���</a> ]</td></tr>" & _
	vbcrlf&"<tr><td>[ <a href='gbook.asp?action=write'>�� �������ӱ���</a> ]</td></tr>" & _
	vbcrlf&"<tr><td>"&put_type("down")&"</td></tr>" & _
	vbcrlf&"</table>"
  response.write format_barc("<font class="&sk_class&"><b>��Ŀͳ��</b></font>",temp1,2,0,5)
end sub

sub down_main()
  dim rs2,sql2
  if cid=0 then
    sql2="select c_id,c_name from jk_class where nsort='"&n_sort&"' order by c_order"
    set rs2=conn.execute(sql2)
    do while not rs2.eof
      nid=rs2("c_id"):sqladd=" and c_id="&nid
%>
<tr align=center valign=top>
<td width='60%'><%call down_new_hotr("jt0","<a href='down_list.asp?c_id="&nid&"'><font class="&sk_class&">"&rs2("c_name")&"</font></a>","<a href='down_list.asp?c_id="&nid&"&action=more'><font class="&sk_class&">����...</font></a>",sqladd,"new",15,0,20,1,8)%></td>
<td width=1 bgcolor='<%=web_var(web_color,3)%>'></td>
<td bgcolor='<%=web_var(web_color,1)%>'><%
call down_new_hotr("","��������","",sqladd,"hot",5,0,11,1,0)
call down_pic("վ���Ƽ�",sqladd,"good",1,10)
%></td>
</tr>
<%
      rs2.movenext
    loop
    rs2.close:set rs2=nothing
  else
    if sid=0 then
      sql2="select s_id,s_name from jk_sort where c_id="&cid&" order by s_order"
      set rs2=conn.execute(sql2)
      response.write "<tr height=1><td colspan=3 align=center>"&format_img("rdown.jpg")&"</td></tr>"
      do while not rs2.eof
        nid=rs2("s_id"):sqladd=" and c_id="&cid&" and s_id="&nid
%>
<tr height=1><td colspan=3 bgcolor="<%response.write web_var(web_color,3)%>"></td></tr>
<tr align=center><td colspan=3>
<%call down_intro(nid,rs2("s_name"))%>
</td></tr>
<tr align=center valign=top>
<td width=400><%call down_new_hotr("jt0","<a href='down_list.asp?c_id="&cid&"&s_id="&nid&"'><font class="&sk_class&">"&rs2("s_name")&"</font></a>","<a href='down_list.asp?c_id="&cid&"&s_id="&nid&"&action=more'><font class="&sk_class&">����...</font></a>",sqladd,"new",40,0,20,1,8)%></td>
<td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td><%
call down_new_hotrn("jt0","��������","",sqladd,"hot",40,0,11,1,0)
'call down_pic("վ���Ƽ�",sqladd,"good",1,10)
%></td>
</tr>
<%
        rs2.movenext
      loop
      rs2.close:set rs2=nothing
    else
      sql2="select jk_class.c_name,jk_sort.s_name from jk_sort inner join jk_class on jk_sort.c_id=jk_class.c_id where jk_sort.c_id="&cid&" and jk_sort.s_id="&sid
      set rs2=conn.execute(sql2)
      if rs2.eof and rs2.bof then
        rs2.close:set rs2=nothing
        cid=0:sid=0
        call down_main():exit sub
      end if
      sqladd=" and c_id="&cid&" and s_id="&sid
%>
<tr align=center>
<td colspan=3><%call down_intro(sid,rs2("s_name"))%></td>
</tr>
<tr align=center>
<td colspan=3><%call down_pic("վ���Ƽ�",sqladd,"good",5,20)%></td>
</tr>
<tr align=center valign=top>
<td width=400><%call down_new_hotr("jt0","<a href='down_list.asp?c_id="&cid&"'><font class="&sk_class&">"&rs2("c_name")&"</font></a> �� <a href='down_list.asp?c_id="&cid&"&s_id="&sid&"'><font class="&sk_class&">"&rs2("s_name")&"</font></a>","<a href='down_list.asp?c_id="&cid&"&s_id="&sid&"&action=more'><font class="&sk_class&">����...</font></a>",sqladd,"new",40,0,20,1,8)%></td>
<td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td>
<td><%
call down_new_hotrn("jt0","��������","",sqladd,"hot",40,0,11,1,0)
%></td>
</tr>
<%
      rs2.close:set rs2=nothing
    end if
  end if
end sub

sub down_more(c_num,tt)
  dim temp1,tim,cnum,sql2,mhead,name,c1,c2,sql,rs,cname,sname
  c1=web_var(web_color,6):c2=web_var(web_color,1)
  pageurl="?action=more&"
  keyword=code_form(request.querystring("keyword"))
  sea_type=trim(request.querystring("sea_type"))
  if sea_type<>"username" then sea_type="name"
  call cid_sid_sql(2,sea_type)
  
  temp1=vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=4><tr><td colspan=5 height=5></td></tr>" & _
	vbcrlf&"<tr align=left height=20 valign=bottom>" & _
	vbcrlf&"<td width='6%'>���</td>" & _
	vbcrlf&"<td width='44%'>��������</td>" & _
	vbcrlf&"<td width='28%'>��������</td>" & _
	vbcrlf&"<td width='12%'>�Ƽ��ȼ�</td>" & _
	vbcrlf&"<td width='10%'>���ش���</td>" & _
	vbcrlf&"</tr>" & _
	vbcrlf&"<tr><td colspan=5 height=1 background='images/bg_dian.gif'></td></tr>"
  sql="select id,name,username,tim,counter,types from down where hidden=1 "&sqladd
  if cid>0 then
    sql=sql&" and c_id="&cid
    if sid>0 then
      sql=sql&" and s_id="&sid
      sql2="select jk_class.c_name,jk_sort.s_name from jk_sort inner join jk_class on jk_sort.c_id=jk_class.c_id where jk_sort.c_id="&cid&" and jk_sort.s_id="&sid
    else
      sql2="select c_name from jk_class where c_id="&cid
    end if
  end if
  sql=sql&" order by id desc"
  
  if cid>0 then
    set rs=conn.execute(sql2)
    if rs.eof and rs.bof then rs.close:set rs=nothing:call down_main():exit sub
    cname=code_html(rs("c_name"),1,0)
    if sid>0 then sname=code_html(rs("s_name"),1,0)
    rs.close
  else
    cname="�������"
  end if
  mhead="<a href='down_list.asp?c_id="&cid&"'><b><font class="&sk_class&">"&cname&"</font></b></a>"
  if cid>0 and sid>0 then mhead=mhead&"&nbsp;<font class="&sk_class&">��</font>&nbsp;<a href='down_list.asp?c_id="&cid&"&s_id="&sid&"'><b><font class="&sk_class&">"&sname&"</font></b></a>"
  
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
    rssum=0
  else
    rssum=rs.recordcount
  end if
  call format_pagecute()
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    name=rs("name"):tim=rs("tim")
    temp1=temp1&vbcrlf&"<tr onmouseover=""javascript:this.bgColor='"&c1&"';"" onmouseout=""javascript:this.bgColor='';""><td>"&i+(viewpage-1)*nummer&".</td>" & _
	  vbcrlf&"<td><a href='down_view.asp?id="&rs("id")&"'"&atb&" title='�������ƣ�"&code_html(name,1,0)&"<br>�� �� �ˣ�"&rs("username")&"<br>����ʱ�䣺"&tim&"'>"&code_html(name,1,c_num)&"</a></td>" & _
	  vbcrlf&"<td>"&time_type(tim,tt)&"</td>" & _
	  vbcrlf&"<td><img src='images/down/star"&rs("types")&".gif' border=0></td>" & _
	  vbcrlf&"<td align=center class=blue>"&rs("counter")&"</td></tr>" & _
	  vbcrlf&"<tr><td colspan=5 height=1 background='images/bg_dian.gif'></td></tr>"
    rs.movenext
  next
  rs.close:set rs=nothing
  temp1=temp1&vbcrlf&"<tr><td colspan=5 height=25 valign=bottom>" & _
	vbcrlf&"����&nbsp;<font class=red>"&rssum&"</font>&nbsp;���ļ�&nbsp;" & _
	vbcrlf&"ҳ�Σ�<font class=red>"&viewpage&"</font>/<font class=red>"&thepages&"</font>&nbsp;" & _
	vbcrlf&"��ҳ��"&jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000")& _
	vbcrlf&"</td></tr></table>"
  response.write "<tr><td colspan=3 align=center>"&format_barc(mhead,temp1,3,0,11)&"</td></tr>"
end sub

sub down_sea()
  dim temp1,nid,nid2,rs,sql,rs2,sql2
  temp1=vbcrlf&"<table border=0 cellspacing=0 cellpadding=0 align=center>" & _
        vbcrlf&"<script language=javascript><!--" & _
        vbcrlf&"function down_sea()" & _
        vbcrlf&"{" & _
        vbcrlf&"  if (down_sea_frm.keyword.value==""������ؼ���"")" & _
        vbcrlf&"  {" & _
        vbcrlf&"    alert(""������������ǰ������Ҫ��ѯ�� �ؼ��� ��"");" & _
        vbcrlf&"    down_sea_frm.keyword.focus();" & _
        vbcrlf&"    return false;" & _
        vbcrlf&"  }" & _
        vbcrlf&"}" & _
        vbcrlf&"--></script>" & _
        vbcrlf&"<form name=down_sea_frm action='down_list.asp' method=get onsubmit=""return down_sea()"">" & _
        vbcrlf&"<input type=hidden name=action value='more'><tr><td height=3></td></tr>" & _
        vbcrlf&"<tr><td>" & _
        vbcrlf&"  <table border=0><tr><td colspan=2><input type=text name=keyword value='������ؼ���' onfocus=""if (value =='������ؼ���'){value =''}"" onblur=""if (value ==''){value='������ؼ���'}"" size=20 maxlength=20></td></tr>" & _


        vbcrlf&"  </table>" & _
        vbcrlf&"</td></tr><tr><td>" & _
        vbcrlf&"  <table border=0><tr>" & _
        vbcrlf&"  <td><select name=c_id sizs=1><option value=''>ȫ�����</option>"
  sql="select c_id,c_name from jk_class where nsort='"&n_sort&"' order by c_order,c_id"
  set rs=conn.execute(sql)
  do while not rs.eof
    nid=int(rs(0))
    temp1=temp1&vbcrlf&"<option value='"&nid&"' class=bg_2"
    if cid=nid then temp1=temp1&" selected"
    temp1=temp1&">"&rs(1)&"</option>"
    sql2="select s_id,s_name from jk_sort where c_id="&nid&" order by s_order,s_id"
    set rs2=conn.execute(sql2)
    do while not rs2.eof
      nid2=rs2(0)
      temp1=temp1&vbcrlf&"<option value='"&nid&"&s_id="&nid2&"'"
      if sid=nid2 then temp1=temp1&" selected"
      temp1=temp1&">��"&rs2(1)&"</option>"
      rs2.movenext
    loop
    rs2.close:set rs2=nothing
    rs.movenext
  loop
  rs.close:set rs=nothing
  temp1=temp1&vbcrlf&"</select></td>" & _
        vbcrlf&"  <td></td></tr>" & _
        vbcrlf&"  <tr height=25><td><select name=sea_type size=1><option value='name'>��������</option><option value='username'>������</option></select></td>" & _
        vbcrlf&"  <td align=left><input type=image src='images/small/search_go.gif' border=0 height=25 width=40></td>" & _
        vbcrlf&"  </tr></table>" & _
        vbcrlf&"</td></tr>" & _
        vbcrlf&"</form><tr><td height=1></td></tr></table>"
  response.write kong&format_barc("<font class="&sk_class&"><b>��������</b></font>",temp1,2,0,4)
end sub
%>