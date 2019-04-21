<!-- #include file="include/config_user.asp" -->
<!-- #include file="include/config_upload.asp" -->
<!-- #include file="include/config_frm.asp" -->
<!-- #include file="include/config_put.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim cid,sid,ispic,pic,nsort,data_name,nid,rs2,sql2,now_id,add_integral,ddim,csid
add_integral=web_varn(web_num,15)

select case action
case "article"
  tit="发表文章"
case "down"
  tit="添加音乐"
case "gallery"
  tit="上传文件"
case "website"
  tit="推荐网站"
case else
  action="news"
  tit="发布新闻"
end select

call web_head(2,0,0,0,0)

if int(popedom_format(login_popedom,41)) then call close_conn():call cookies_type("locked")
'------------------------------------left----------------------------------
call left_user()
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong&table1
%>
<tr<%response.write table2%> height=25><td class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small(us)%>&nbsp;&nbsp;<b>查看我所发表的相关信息</b></td></tr>
<tr<%response.write table3%>><td align=center height=30>
<%response.write img_small("jt1")%><a href='?action=news'<%if action="news" then response.write "class=red_3"%>>发布我的新闻</a>　&nbsp;
<%response.write img_small("jt1")%><a href='?action=article'<%if action="article" then response.write "class=red_3"%>>发表我的文章</a>　&nbsp;
<%response.write img_small("jt1")%><a href='?action=down'<%if action="down" then response.write "class=red_3"%>>添加我的音乐</a>　&nbsp;
<%response.write img_small("jt1")%><a href='?action=gallery'<%if action="gallery" then response.write "class=red_3"%>>上传我的文件</a>　&nbsp;
<%response.write img_small("jt1")%><a href='?action=website'<%if action="website" then response.write "class=red_3"%>>我要推荐网站</a>
</td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='94%'>
  <tr><td class=htd>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;您在发表相关信息后，您的积分将自动增加积分<font class=red><%response.write add_integral%></font>分。<font class=red>请勿恶意乱发！</font></td></tr>
  </table>
</td></tr>
</table>
<%
response.write ukong&table1
%>
<tr<%response.write table2%> height=25><td class=end  background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;<%response.write img_small(us)%>&nbsp;&nbsp;<b><%response.write tit%></b></td></tr>
<tr<%response.write table3%>><td align=center height=350>
<%
select case action
case "article"
  if int(mid(web_var(web_config,9),2,1))=0 then
    call put_close()
  else
    data_name=action
    nsort="art"
    call put_article()
  end if
case "down"
  if int(mid(web_var(web_config,9),3,1))=0 then
    call put_close()
  else
    data_name=action
    nsort="down"
    call put_down()
  end if
case "gallery"
  if int(mid(web_var(web_config,9),4,1))=0 then
    call put_close()
  else
    data_name=action
    nsort="gall"
    call put_gallery()
  end if
case "website"
  if int(mid(web_var(web_config,9),5,1))=0 then
    call put_close()
  else
    data_name=action
    nsort="web"
    call put_website()
  end if
case else
  if int(mid(web_var(web_config,9),1,1))=0 then
    call put_close()
  else
    data_name=action
    nsort="news"
    call put_news()
  end if
end select
%>
</td></tr>
</table>
<br>
<%
'---------------------------------center end-------------------------------
call web_end(0)

sub put_close()
  response.write "<font class=red_2>对不起！本站暂时关闭用户 <font class=blue>"&tit&"</font> 的功能。</font><br><br>如有需要，请与管理员联系。谢谢！"
end sub

sub put_website()
  if trim(request.form("put"))="yes" then
    dim name,url,isgood,country,lang,remark
    name=code_form(request.form("name"))
    csid=trim(request.form("csid"))
    url=code_form(request.form("url"))
    isgood=trim(request.form("isgood"))
    remark=request.form("remark")
    country=trim(request.form("country"))
    lang=trim(request.form("lang"))
    pic=trim(request.form("picg"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择网站类型！</font><br><br>"&go_back
    elseif len(name)<1 or len(url)<1 then
      response.write "<font class=red_2>网站名称和地址不能为空！</font><br><br>"&go_back
    elseif len(remark)>250 then
      response.write "<font class=red_2>网站说明不能长于250个字符！</font><br><br>"&go_back
    else
      call chk_cid_sid()
      set rs=server.createobject("adodb.recordset")
      sql="select * from "&data_name
      rs.open sql,conn,1,3
      rs.addnew
      rs("c_id")=cid
      rs("s_id")=sid
      rs("username")=login_username
      rs("hidden")=false
      rs("name")=name
      rs("url")=url
      rs("country")=country
      rs("lang")=lang
      rs("remark")=remark
      if isgood="yes" then
        rs("isgood")=true
      else
        rs("isgood")=false
      end if
      rs("username")=login_username
      if len(pic)<3 then
        rs("pic")="no_pic.gif"
      else
        rs("pic")=pic
      end if
      rs("tim")=now_time
      rs("counter")=0
      rs.update
      rs.close:set rs=nothing
      call user_integral("add",add_integral,login_username)
      call upload_note(action,first_id(action))
      response.write "<font class=red>已成功推荐了一个网站！</font><br><br>请等待管理员审核通过……<br><br>"
    end if
  else
%><table border=0 cellspacing=0 cellpadding=3>
<form name=add_frm action='?action=<%response.write action%>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%'>网站名称：</td><td width='85%'><input type=text size=70 name=name maxlength=50><%=redx%></td></tr>
  <tr><td>网站类型：</td><td><%call chk_csid(cid,sid)%></td></tr>
  <tr><td>网站地址：</td><td><input type=text size=70 name=url value='http://' maxlength=100><%=redx%></td></tr>
  <tr><td>国家地区：</td><td><select name=country size=1>
<option>中国</option>
<option>香港</option>
<option>台湾</option>
<option>美国</option>
<option>英国</option>
<option>日本</option>
<option>韩国</option>
<option>加拿大</option>
<option>澳大利亚</option>
<option>新西兰</option>
<option>俄罗斯</option>
<option>意大利</option>
<option>法国</option>
<option>西班牙</option>
<option>德国</option>
<option>其它国家</option>
</select>&nbsp;&nbsp;&nbsp;&nbsp;站点语言：<select name=lang size=1>
<option>简体中文</option>
<option>繁体中文</option>
<option>English</option>
<option>其它语言</option>
</select>&nbsp;&nbsp;&nbsp;推荐：<input type=checkbox name=isgood value='yes'></td></tr>
<% ispic="w"&upload_time(now_time) %>
  <tr><td>图片地址：</td><td><input type=test name=pic size=70 maxlength=100></td></tr>
  <tr><td>上传图片：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=60 scrolling=no src='upload.asp?uppath=website&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr><td valign=top class=htd><br>网站内容：<br><=250B</td><td><textarea name=remark rows=5 cols=70></textarea></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' 推 荐 网 站 '>　　<input type=reset value='重新填写'></td></tr>
</form></table><%
  end if
end sub

sub put_news()
  if trim(request.form("put"))="yes" then
    dim topic,comto,istop,word,ispic,pic,keyes
    topic=code_form(request.form("topic"))
    csid=trim(request.form("csid"))
    comto=code_form(request.form("comto"))
    keyes=code_form(request.form("keyes"))
    istop=trim(request.form("istop"))
    word=request.form("word")
    ispic=trim(request.form("ispic"))
    pic=trim(request.form("pic"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择新闻类型！</font><br><br>"&go_back
    elseif len(topic)<1 or len(word)<10 then
      response.write "<font class=red_2>新闻标题和内容不能为空！</font><br><br>"&go_back
    else
      call chk_cid_sid()
      set rs=server.createobject("adodb.recordset")
      sql="select * from "&data_name
      rs.open sql,conn,1,3
      rs.addnew
      rs("c_id")=cid
      rs("s_id")=sid
      rs("username")=login_username
      rs("hidden")=false
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
      call user_integral("add",add_integral,login_username)
      call upload_note(action,first_id(action))
      response.write "<font class=red>已成功发布了一篇新闻！</font><br><br>请等待管理员审核通过……<br><br>"
    end if
  else
%><table border=0 cellspacing=0 cellpadding=3 align=center>
<form name=add_frm action='?action=<%response.write action%>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>新闻标题：</td><td width='85%'><input type=text size=70 name=topic maxlength=100><%=redx%></td></tr>
  <tr><td align=center>新闻类别：</td><td><%call chk_csid(cid,sid)%>&nbsp;&nbsp;&nbsp;&nbsp;出处：<input type=text size=30 name=comto maxlength=10></td></tr>
  <tr><td align=center>关 键 字：</td><td><input type=text size=20 name=keyes maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;推荐：<input type=checkbox name=istop value='yes'>&nbsp;选上为新闻首页显示</td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","word","&nbsp;&nbsp;")%></td></tr>
  <tr><td valign=top align=center><br>新闻内容：</td><td><textarea name=word rows=15 cols=70></textarea></td></tr>
<%ispic="n"&upload_time(now_time)%>
  <tr><td align=center>图片新闻：</td><td><input type=checkbox name=ispic value='yes'>&nbsp;&nbsp;&nbsp;&nbsp;图片：<input type=test name=pic size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=<%response.write ispic%>&uptext=pic' target=upload_frame>上传图片</a>&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=n&uptext=word' target=upload_frame>上传至内容</a></td></tr>
  <tr><td align=center>上传图片：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=60 scrolling=no src='upload.asp?uppath=news&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr><td colspan=2 align=center height=30><input type=submit value='发 布 我 的 新 闻'>　　<input type=reset value='重新填写'></td></tr>
</form></table><%
  end if
end sub

sub put_article()
  if trim(request.form("put"))="yes" then
    dim topic
    topic=code_form(request.form("topic"))
    csid=trim(request.form("csid"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择文章类型！</font><br><br>"&go_back
    elseif topic="" then
      response.write "<font class=red_2>文章标题不能为空！</font><br><br>"&go_back
    else
      call chk_cid_sid()
      set rs=server.createobject("adodb.recordset")
      sql="select * from "&data_name
      rs.open sql,conn,1,3
      rs.addnew
      rs("c_id")=cid
      rs("s_id")=sid
      rs("username")=login_username
      rs("hidden")=false
      rs("topic")=topic
      rs("word")=request.form("word")
      if isnumeric(trim(request.form("emoney"))) then
        rs("emoney")=trim(request.form("emoney"))
      else
        rs("emoney")=0
      end if
      rs("author")=code_admin(request.form("author"))
      rs("power")=replace(replace(trim(request.form("power"))," ",""),",",".")
      rs("keyes")=code_admin(request.form("keyes"))
      if trim(request.form("istop"))="yes" then
        rs("istop")=1
      else
        rs("istop")=0
      end if
      rs("tim")=now_time
      rs("counter")=0
      rs.update
      rs.close:set rs=nothing
      call user_integral("add",add_integral,login_username)
      call upload_note(action,first_id(action))
      response.write "<font class=red>已成功发布了一篇文章！</font><br><br>请等待管理员审核通过……<br><br>"
    end if
  else
%><table border=0 width='100%' cellspacing=0 cellpadding=2 align=center>
<form name=add_frm action='?action=<%response.write action%>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>文章标题：</td><td width='85%'><input type=text size=70 name=topic maxlength=40><%=redx%></td></tr>
  <tr><td align=center>文章类型：</td><td><%call chk_csid(cid,sid):call chk_emoney(0)%></td></tr>
  <tr><td align=center>浏览权限：</td><td><%call chk_power("",1)%></td></tr>
  <tr><td align=center>文章作者：</td><td><input type=text size=12 name=author maxlength=20>&nbsp;&nbsp;关键字：<input type=text name=keyes size=12 maxlength=20>&nbsp;&nbsp;推荐：<input type=checkbox name=istop value='yes'></td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","word","&nbsp;&nbsp;")%></td></tr>
  <tr><td valign=top align=center><br>文章内容：</td><td><textarea name=word rows=15 cols=70></textarea></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=90 scrolling=no src='upload.asp?uppath=article&upname=a&uptext=word'></iframe></td></tr>
  <tr><td></td><td height=30><input type=submit value='发 布 我 的 文 章'>　　<input type=reset value='重新填写'></td></tr>
</form></table><%
  end if
end sub

sub put_down()
  if trim(request.form("put"))="yes" then
    dim name,sizes,url,url2,homepage,remark,types,keyes,pic
    csid=trim(request.form("csid"))
    name=code_form(request.form("name"))
    sizes=code_form(request.form("sizes"))
    url=code_form(request.form("url"))
    url2=code_form(request.form("url2"))
    homepage=code_form(request.form("homepage"))
    keyes=code_form(request.form("keyes"))
    remark=request.form("remark")
    pic=request.form("pic")
    if len(pic)<3 then pic="no_pic.gif"
    types=request.form("types")
    if len(csid)<1 or var_null(name)="" or var_null(url)="" then
      response.write("<font class=red_2>音乐的类型、名称和下载地址不能为空！</font><br><br>"&go_back)
    else
      call chk_cid_sid()
      sql="select * from down"
      set rs=server.createobject("adodb.recordset")
      rs.open sql,conn,1,3
      rs.addnew
      rs("c_id")=cid
      rs("s_id")=sid
      rs("username")=login_username
      rs("hidden")=false
      rs("name")=name
      rs("sizes")=sizes
      if isnumeric(trim(request.form("emoney"))) then
        rs("emoney")=trim(request.form("emoney"))
      else
        rs("emoney")=0
      end if
      rs("genre")=trim(request.form("genre"))
      rs("os")=replace(trim(request.form("os"))," ","")
      rs("power")=replace(replace(trim(request.form("power"))," ",""),",",".")
      rs("url")=url
      rs("url2")=url2
      rs("homepage")=homepage
      rs("remark")=remark
      rs("keyes")=keyes
      rs("pic")=pic
      rs("tim")=now_time
      rs("counter")=0
      rs("types")=types
      rs.update
      rs.close:set rs=nothing
      call user_integral("add",add_integral,login_username)
      call upload_note(action,first_id(action))
      response.write "<font class=red>已成功添加了一个文件！</font><br><br>请等待管理员审核通过……<br><br>"
  end if
else
%>
<table border=0 width=560 cellspacing=0 cellpadding=2>
<form name=add_frm action='?action=<%response.write action%>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>音乐名称：</td><td width='85%'><input type=text name=name size=70 maxlength=40><% response.write redx %></td></tr>
  <tr><td align=center>音乐类别：</td><td><%call chk_csid(cid,sid):call chk_emoney(0)%></td></tr>
  <tr><td align=center>下载权限：</td><td><%call chk_power("",1)%></td></tr>
  <tr><td align=center>文件大小：</td><td><input type=text name=sizes value='KB' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;推荐等级：<select name=types size=1>
<option value='0'>没有等级</option>
<option value='1'>一星级</option>
<option value='2'>二星级</option>
<option value='3'>三星级</option>
<option value='4'>四星级</option>
<option value='5'>五星级</option>
</select>&nbsp;&nbsp;&nbsp;音乐类型：<select name=genre size=1><%
  ddim=split(web_var(web_down,4),":")
  for i=0 to ubound(ddim)
    response.write vbcrlf&"<option>"&ddim(i)&"</option>"
  next
  erase ddim
%></select></td></tr>
  <tr><td align=center>播放软件：</td><td><%
  ddim=split(web_var(web_down,3),":")
  for i=0 to ubound(ddim)
    response.write "<input type=checkbox name=os value='"&ddim(i)&"' class=bg_1>"&ddim(i)
  next
  erase ddim
%></td></tr>
  <tr><td align=center>本站下载：</td><td><input type=text name=url size=70 maxlength=200><% response.write redx %></td></tr>
  <tr><td align=center>镜像下载：</td><td><input type=text name=url2 value='http://' size=70 maxlength=200></td></tr>
  <tr><td align=center>文件来自：</td><td><input type=text name=homepage value='http://' size=50 maxlength=50></td></tr>
  <tr height=35<%response.write format_table(3,1)%>><td align=center><%call frm_ubb_type()%></td><td><%call frm_ubb("add_frm","remark","&nbsp;&nbsp;")%></td></tr>
  <tr><td valign=top align=center><br>音乐备注</td><td><textarea rows=6 name=remark cols=70></textarea></td></tr>
<%ispic="d"&upload_time(now_time)%>
  <tr><td align=center>关 键 字：</td><td><input type=text name=keyes size=12 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;图片：<input type=text name=pic size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=<%response.write ispic%>&uptext=pic' target=upload_frame>上传图片</a>&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=d&uptext=remark' target=upload_frame>上传至内容</a></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=90 scrolling=no src='upload.asp?uppath=down&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr height=30><td></td><td><input type=submit value=' 添 加 我 的 音 乐 '></td></tr>
</form></table><%
  end if
end sub

sub put_gallery()
  dim name,pic,ispic,types
  if trim(request.form("put"))="yes" then
    name=code_form(request.form("name"))
    csid=trim(request.form("csid"))
    pic=code_form(request.form("pic"))
    types=trim(request.form("types"))
    if len(csid)<1 then
      response.write "<font class=red_2>请选择图片分类！</font><br><br>"&go_back
    elseif len(name)<1 then
      response.write "<font class=red_2>图片名称说明不能为空！</font><br><br>"&go_back
    elseif len(pic)<8 then
      response.write "<font class=red_2>请上传图片或输入图片的地址！</font><br><br>"&go_back
    else
      call chk_cid_sid()
      set rs=server.createobject("adodb.recordset")
      sql="select * from "&data_name
      rs.open sql,conn,1,3
      rs.addnew
      rs("c_id")=cid
      rs("s_id")=sid
      rs("username")=login_username
      rs("types")=types
      rs("name")=name
      if len(code_admin(request.form("spic")))<3 then
        rs("spic")="no_pic.gif"
      else
        rs("spic")=code_admin(request.form("spic"))
      end if
      rs("pic")=pic
      rs("remark")=left(request.form("remark"),250)
      rs("power")=replace(replace(trim(request.form("power"))," ",""),",",".")
      if isnumeric(trim(request.form("emoney"))) then
        rs("emoney")=trim(request.form("emoney"))
      else
        rs("emoney")=0
      end if
      if trim(request.form("istop"))="yes" then
        rs("istop")=1
      else
        rs("istop")=0
      end if
      rs("counter")=0
      rs("tim")=now_time
      rs("hidden")=false
      rs.update
      rs.close:set rs=nothing
      call user_integral("add",add_integral,login_username)
      call upload_note(action,first_id(action))
      response.write "<font class=red>已成功添加了一张图片！</font><br><br>请等待管理员审核通过……<br><br>"
    end if
  else
%><table border=0 cellspacing=0 cellpadding=3>
<form name=add_frm action='?action=<%response.write action%>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>文件名称：</td><td width='85%'><input type=text size=70 name=name maxlength=50><%=redx%></td></tr>
  <tr><td align=center>文件分类：</td><td><%call chk_csid(cid,sid)%>&nbsp;&nbsp;文件类型：<select name=types size=1>
<option value='paste'<%if types="paste" then response.write " selected"%>>贴图</option>
<option value='flash'<%if types="flash" then response.write " selected"%>>FLASH</option>
<option value='film'<%if types="film" then response.write " selected"%>>视频</option>
<option value='logo'<%if types="logo" then response.write " selected"%>>LOGO</option>
<option value='baner'<%if types="baner" then response.write " selected"%>>BANNER</option>
</select><%response.write redx%>&nbsp;&nbsp;<%call chk_emoney(0)%></td></tr>
  <tr><td align=center>浏览权限：</td><td><%call chk_power("",1)%></td></tr>
<%ispic="gs"&upload_time(now_time)%>
  <tr><td align=center>小 图 片：</td><td><input type=test name=spic size=70 maxlength=100></td></tr>
  <tr><td align=center>上传图片：</td><td><iframe frameborder=0 name=upload_frames width='100%' height=60 scrolling=no src='upload.asp?uppath=gallery&upname=<%response.write ispic%>&uptext=spic'></iframe></td></tr>
<%ispic="g"&upload_time(now_time)%>
  <tr><td align=center>文件地址：</td><td><input type=test name=pic size=70 maxlength=100><%response.write redx%></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=90 scrolling=no src='upload.asp?uppath=gallery&upname=<%response.write ispic%>&uptext=pic'></iframe></td></tr>
  <tr><td align=center>文件说明：<br><br><=250字符</td><td><textarea name=remark rows=5 cols=70></textarea></td></tr>
  <tr><td colspan=2 align=center height=30><input type=submit value=' 上 传 我 的 文 件 '>　　<input type=reset value='重新填写'></td></tr>
</form></table><%
  end if
end sub
%>