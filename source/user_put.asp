<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="INCLUDE/config_upload.asp" -->
<!-- #include file="INCLUDE/config_frm.asp" -->
<!-- #include file="INCLUDE/config_put.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim cid,sid,ispic,pic,nsort,data_name,nid,rs2,sql2,now_id,add_integral,ddim,csid
add_integral = web_varn(web_num,15)

Select Case action
    Case "article"
        tit    = "发表文章"
    Case "down"
        tit    = "添加音乐"
    Case "gallery"
        tit    = "上传文件"
    Case "website"
        tit    = "推荐网站"
    Case Else
        action = "news"
        tit    = "发布新闻"
End Select

Call web_head(2,0,0,0,0)

If Int(popedom_format(login_popedom,41)) Then Call close_conn():Call cookies_type("locked")
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong & table1 %>
<tr<% Response.Write table2 %> height=25><td class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small(us) %>&nbsp;&nbsp;<b>查看我所发表的相关信息</b></td></tr>
<tr<% Response.Write table3 %>><td align=center height=30>
<% Response.Write img_small("jt1") %><a href='?action=news'<% If action = "news" Then Response.Write "class=red_3" %>>发布我的新闻</a>　&nbsp;
<% Response.Write img_small("jt1") %><a href='?action=article'<% If action = "article" Then Response.Write "class=red_3" %>>发表我的文章</a>　&nbsp;
<% Response.Write img_small("jt1") %><a href='?action=down'<% If action = "down" Then Response.Write "class=red_3" %>>添加我的音乐</a>　&nbsp;
<% Response.Write img_small("jt1") %><a href='?action=gallery'<% If action = "gallery" Then Response.Write "class=red_3" %>>上传我的文件</a>　&nbsp;
<% Response.Write img_small("jt1") %><a href='?action=website'<% If action = "website" Then Response.Write "class=red_3" %>>我要推荐网站</a>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='94%'>
  <tr><td class=htd>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;您在发表相关信息后，您的积分将自动增加积分<font class=red><% Response.Write add_integral %></font>分。<font class=red>请勿恶意乱发！</font></td></tr>
  </table>
</td></tr>
</table>
<%
Response.Write ukong & table1 %>
<tr<% Response.Write table2 %> height=25><td class=end  background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small(us) %>&nbsp;&nbsp;<b><% Response.Write tit %></b></td></tr>
<tr<% Response.Write table3 %>><td align=center height=350>
<%

Select Case action
    Case "article"

        If Int(Mid(web_var(web_config,9),2,1)) = 0 Then
            Call put_close()
        Else
            data_name = action
            nsort     = "art"
            Call put_article()
        End If

    Case "down"

        If Int(Mid(web_var(web_config,9),3,1)) = 0 Then
            Call put_close()
        Else
            data_name = action
            nsort     = "down"
            Call put_down()
        End If

    Case "gallery"

        If Int(Mid(web_var(web_config,9),4,1)) = 0 Then
            Call put_close()
        Else
            data_name = action
            nsort     = "gall"
            Call put_gallery()
        End If

    Case "website"

        If Int(Mid(web_var(web_config,9),5,1)) = 0 Then
            Call put_close()
        Else
            data_name = action
            nsort     = "web"
            Call put_website()
        End If

    Case Else

        If Int(Mid(web_var(web_config,9),1,1)) = 0 Then
            Call put_close()
        Else
            data_name = action
            nsort     = "news"
            Call put_news()
        End If

End Select %>
</td></tr>
</table>
<br>
<%
'---------------------------------center end-------------------------------
Call web_end(0)

Sub put_close()
    Response.Write "<font class=red_2>对不起！本站暂时关闭用户 <font class=blue>" & tit & "</font> 的功能。</font><br><br>如有需要，请与管理员联系。谢谢！"
End Sub

Sub put_website()

    If Trim(Request.form("put")) = "yes" Then
        Dim name,url,isgood,country,lang,remark
        name    = code_form(Request.form("name"))
        csid    = Trim(Request.form("csid"))
        url     = code_form(Request.form("url"))
        isgood  = Trim(Request.form("isgood"))
        remark  = Request.form("remark")
        country = Trim(Request.form("country"))
        lang    = Trim(Request.form("lang"))
        pic     = Trim(Request.form("picg"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>请选择网站类型！</font><br><br>" & go_back
        ElseIf Len(name) < 1 Or Len(url) < 1 Then
            Response.Write "<font class=red_2>网站名称和地址不能为空！</font><br><br>" & go_back
        ElseIf Len(remark) > 250 Then
            Response.Write "<font class=red_2>网站说明不能长于250个字符！</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            Set rs = Server.CreateObject("adodb.recordset")
            sql    = "select * from " & data_name
            rs.open sql,conn,1,3
            rs.addnew
            rs("c_id")     = cid
            rs("s_id")     = sid
            rs("username")     = login_username
            rs("hidden")     = False
            rs("name")     = name
            rs("url")     = url
            rs("country")     = country
            rs("lang")     = lang
            rs("remark")     = remark

            If isgood = "yes" Then
                rs("isgood") = True
            Else
                rs("isgood") = False
            End If

            rs("username")     = login_username

            If Len(pic) < 3 Then
                rs("pic") = "no_pic.gif"
            Else
                rs("pic") = pic
            End If

            rs("tim")     = now_time
            rs("counter")     = 0
            rs.update
            rs.Close:Set rs = Nothing
            Call user_integral("add",add_integral,login_username)
            Call upload_note(action,first_id(action))
            Response.Write "<font class=red>已成功推荐了一个网站！</font><br><br>请等待管理员审核通过……<br><br>"
        End If

    Else %><table border=0 cellspacing=0 cellpadding=3>
<form name=add_frm action='?action=<% Response.Write action %>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%'>网站名称：</td><td width='85%'><input type=text size=70 name=name maxlength=50><% = redx %></td></tr>
  <tr><td>网站类型：</td><td><% Call chk_csid(cid,sid) %></td></tr>
  <tr><td>网站地址：</td><td><input type=text size=70 name=url value='http://' maxlength=100><% = redx %></td></tr>
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
<% ispic = "w" & upload_time(now_time) %>
  <tr><td>图片地址：</td><td><input type=test name=pic size=70 maxlength=100></td></tr>
  <tr><td>上传图片：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=60 scrolling=no src='upload.asp?uppath=website&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr><td valign=top class=htd><br>网站内容：<br><=250B</td><td><textarea name=remark rows=5 cols=70></textarea></td></tr>
  <tr><td colspan=2 align=center height=25><input type=submit value=' 推 荐 网 站 '>　　<input type=reset value='重新填写'></td></tr>
</form></table><%
    End If

End Sub

Sub put_news()

    If Trim(Request.form("put")) = "yes" Then
        Dim topic,comto,istop,word,ispic,pic,keyes
        topic = code_form(Request.form("topic"))
        csid  = Trim(Request.form("csid"))
        comto = code_form(Request.form("comto"))
        keyes = code_form(Request.form("keyes"))
        istop = Trim(Request.form("istop"))
        word  = Request.form("word")
        ispic = Trim(Request.form("ispic"))
        pic   = Trim(Request.form("pic"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>请选择新闻类型！</font><br><br>" & go_back
        ElseIf Len(topic) < 1 Or Len(word) < 10 Then
            Response.Write "<font class=red_2>新闻标题和内容不能为空！</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            Set rs = Server.CreateObject("adodb.recordset")
            sql    = "select * from " & data_name
            rs.open sql,conn,1,3
            rs.addnew
            rs("c_id")     = cid
            rs("s_id")     = sid
            rs("username")     = login_username
            rs("hidden")     = False
            rs("topic")     = topic
            rs("comto")     = comto
            rs("keyes")     = keyes
            rs("word")     = word

            If istop = "yes" Then
                rs("istop") = True
            Else
                rs("istop") = False
            End If

            If ispic = "yes" Then
                rs("ispic") = True
            Else
                rs("ispic") = False
            End If

            rs("pic")     = pic
            rs("tim")     = now_time
            rs("counter")     = 0
            rs.update
            rs.Close:Set rs = Nothing
            Call user_integral("add",add_integral,login_username)
            Call upload_note(action,first_id(action))
            Response.Write "<font class=red>已成功发布了一篇新闻！</font><br><br>请等待管理员审核通过……<br><br>"
        End If

    Else %><table border=0 cellspacing=0 cellpadding=3 align=center>
<form name=add_frm action='?action=<% Response.Write action %>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>新闻标题：</td><td width='85%'><input type=text size=70 name=topic maxlength=100><% = redx %></td></tr>
  <tr><td align=center>新闻类别：</td><td><% Call chk_csid(cid,sid) %>&nbsp;&nbsp;&nbsp;&nbsp;出处：<input type=text size=30 name=comto maxlength=10></td></tr>
  <tr><td align=center>关 键 字：</td><td><input type=text size=20 name=keyes maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;推荐：<input type=checkbox name=istop value='yes'>&nbsp;选上为新闻首页显示</td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","word","&nbsp;&nbsp;") %></td></tr>
  <tr><td valign=top align=center><br>新闻内容：</td><td><textarea name=word rows=15 cols=70></textarea></td></tr>
<% ispic = "n" & upload_time(now_time) %>
  <tr><td align=center>图片新闻：</td><td><input type=checkbox name=ispic value='yes'>&nbsp;&nbsp;&nbsp;&nbsp;图片：<input type=test name=pic size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=<% Response.Write ispic %>&uptext=pic' target=upload_frame>上传图片</a>&nbsp;&nbsp;<a href='upload.asp?uppath=news&upname=n&uptext=word' target=upload_frame>上传至内容</a></td></tr>
  <tr><td align=center>上传图片：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=60 scrolling=no src='upload.asp?uppath=news&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr><td colspan=2 align=center height=30><input type=submit value='发 布 我 的 新 闻'>　　<input type=reset value='重新填写'></td></tr>
</form></table><%
    End If

End Sub

Sub put_article()

    If Trim(Request.form("put")) = "yes" Then
        Dim topic
        topic = code_form(Request.form("topic"))
        csid  = Trim(Request.form("csid"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>请选择文章类型！</font><br><br>" & go_back
        ElseIf topic = "" Then
            Response.Write "<font class=red_2>文章标题不能为空！</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            Set rs = Server.CreateObject("adodb.recordset")
            sql    = "select * from " & data_name
            rs.open sql,conn,1,3
            rs.addnew
            rs("c_id")     = cid
            rs("s_id")     = sid
            rs("username")     = login_username
            rs("hidden")     = False
            rs("topic")     = topic
            rs("word")     = Request.form("word")

            If IsNumeric(Trim(Request.form("emoney"))) Then
                rs("emoney") = Trim(Request.form("emoney"))
            Else
                rs("emoney") = 0
            End If

            rs("author")     = code_admin(Request.form("author"))
            rs("power")     = Replace(Replace(Trim(Request.form("power"))," ",""),",",".")
            rs("keyes")     = code_admin(Request.form("keyes"))

            If Trim(Request.form("istop")) = "yes" Then
                rs("istop") = 1
            Else
                rs("istop") = 0
            End If

            rs("tim")     = now_time
            rs("counter")     = 0
            rs.update
            rs.Close:Set rs = Nothing
            Call user_integral("add",add_integral,login_username)
            Call upload_note(action,first_id(action))
            Response.Write "<font class=red>已成功发布了一篇文章！</font><br><br>请等待管理员审核通过……<br><br>"
        End If

    Else %><table border=0 width='100%' cellspacing=0 cellpadding=2 align=center>
<form name=add_frm action='?action=<% Response.Write action %>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>文章标题：</td><td width='85%'><input type=text size=70 name=topic maxlength=40><% = redx %></td></tr>
  <tr><td align=center>文章类型：</td><td><% Call chk_csid(cid,sid):Call chk_emoney(0) %></td></tr>
  <tr><td align=center>浏览权限：</td><td><% Call chk_power("",1) %></td></tr>
  <tr><td align=center>文章作者：</td><td><input type=text size=12 name=author maxlength=20>&nbsp;&nbsp;关键字：<input type=text name=keyes size=12 maxlength=20>&nbsp;&nbsp;推荐：<input type=checkbox name=istop value='yes'></td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","word","&nbsp;&nbsp;") %></td></tr>
  <tr><td valign=top align=center><br>文章内容：</td><td><textarea name=word rows=15 cols=70></textarea></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=90 scrolling=no src='upload.asp?uppath=article&upname=a&uptext=word'></iframe></td></tr>
  <tr><td></td><td height=30><input type=submit value='发 布 我 的 文 章'>　　<input type=reset value='重新填写'></td></tr>
</form></table><%
    End If

End Sub

Sub put_down()

    If Trim(Request.form("put")) = "yes" Then
        Dim name,sizes,url,url2,homepage,remark,types,keyes,pic
        csid     = Trim(Request.form("csid"))
        name     = code_form(Request.form("name"))
        sizes    = code_form(Request.form("sizes"))
        url      = code_form(Request.form("url"))
        url2     = code_form(Request.form("url2"))
        homepage = code_form(Request.form("homepage"))
        keyes    = code_form(Request.form("keyes"))
        remark   = Request.form("remark")
        pic      = Request.form("pic")
        If Len(pic) < 3 Then pic = "no_pic.gif"
        types    = Request.form("types")

        If Len(csid) < 1 Or var_null(name) = "" Or var_null(url) = "" Then
            Response.Write("<font class=red_2>音乐的类型、名称和下载地址不能为空！</font><br><br>" & go_back)
        Else
            Call chk_cid_sid()
            sql    = "select * from down"
            Set rs = Server.CreateObject("adodb.recordset")
            rs.open sql,conn,1,3
            rs.addnew
            rs("c_id")     = cid
            rs("s_id")     = sid
            rs("username")     = login_username
            rs("hidden")     = False
            rs("name")     = name
            rs("sizes")     = sizes

            If IsNumeric(Trim(Request.form("emoney"))) Then
                rs("emoney") = Trim(Request.form("emoney"))
            Else
                rs("emoney") = 0
            End If

            rs("genre")     = Trim(Request.form("genre"))
            rs("os")     = Replace(Trim(Request.form("os"))," ","")
            rs("power")     = Replace(Replace(Trim(Request.form("power"))," ",""),",",".")
            rs("url")     = url
            rs("url2")     = url2
            rs("homepage")     = homepage
            rs("remark")     = remark
            rs("keyes")     = keyes
            rs("pic")     = pic
            rs("tim")     = now_time
            rs("counter")     = 0
            rs("types")     = types
            rs.update
            rs.Close:Set rs = Nothing
            Call user_integral("add",add_integral,login_username)
            Call upload_note(action,first_id(action))
            Response.Write "<font class=red>已成功添加了一个文件！</font><br><br>请等待管理员审核通过……<br><br>"
        End If

    Else %>
<table border=0 width=560 cellspacing=0 cellpadding=2>
<form name=add_frm action='?action=<% Response.Write action %>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>音乐名称：</td><td width='85%'><input type=text name=name size=70 maxlength=40><% Response.Write redx %></td></tr>
  <tr><td align=center>音乐类别：</td><td><% Call chk_csid(cid,sid):Call chk_emoney(0) %></td></tr>
  <tr><td align=center>下载权限：</td><td><% Call chk_power("",1) %></td></tr>
  <tr><td align=center>文件大小：</td><td><input type=text name=sizes value='KB' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;推荐等级：<select name=types size=1>
<option value='0'>没有等级</option>
<option value='1'>一星级</option>
<option value='2'>二星级</option>
<option value='3'>三星级</option>
<option value='4'>四星级</option>
<option value='5'>五星级</option>
</select>&nbsp;&nbsp;&nbsp;音乐类型：<select name=genre size=1><%
        ddim = Split(web_var(web_down,4),":")

        For i = 0 To UBound(ddim)
            Response.Write vbcrlf & "<option>" & ddim(i) & "</option>"
        Next

        Erase ddim %></select></td></tr>
  <tr><td align=center>播放软件：</td><td><%
        ddim = Split(web_var(web_down,3),":")

        For i = 0 To UBound(ddim)
            Response.Write "<input type=checkbox name=os value='" & ddim(i) & "' class=bg_1>" & ddim(i)
        Next

        Erase ddim %></td></tr>
  <tr><td align=center>本站下载：</td><td><input type=text name=url size=70 maxlength=200><% Response.Write redx %></td></tr>
  <tr><td align=center>镜像下载：</td><td><input type=text name=url2 value='http://' size=70 maxlength=200></td></tr>
  <tr><td align=center>文件来自：</td><td><input type=text name=homepage value='http://' size=50 maxlength=50></td></tr>
  <tr height=35<% Response.Write format_table(3,1) %>><td align=center><% Call frm_ubb_type() %></td><td><% Call frm_ubb("add_frm","remark","&nbsp;&nbsp;") %></td></tr>
  <tr><td valign=top align=center><br>音乐备注</td><td><textarea rows=6 name=remark cols=70></textarea></td></tr>
<% ispic = "d" & upload_time(now_time) %>
  <tr><td align=center>关 键 字：</td><td><input type=text name=keyes size=12 maxlength=20>&nbsp;&nbsp;&nbsp;&nbsp;图片：<input type=text name=pic size=30 maxlength=100>&nbsp;&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=<% Response.Write ispic %>&uptext=pic' target=upload_frame>上传图片</a>&nbsp;&nbsp;<a href='upload.asp?uppath=down&upname=d&uptext=remark' target=upload_frame>上传至内容</a></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=90 scrolling=no src='upload.asp?uppath=down&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr height=30><td></td><td><input type=submit value=' 添 加 我 的 音 乐 '></td></tr>
</form></table><%
    End If

End Sub

Sub put_gallery()
    Dim name,pic,ispic,types

    If Trim(Request.form("put")) = "yes" Then
        name  = code_form(Request.form("name"))
        csid  = Trim(Request.form("csid"))
        pic   = code_form(Request.form("pic"))
        types = Trim(Request.form("types"))

        If Len(csid) < 1 Then
            Response.Write "<font class=red_2>请选择图片分类！</font><br><br>" & go_back
        ElseIf Len(name) < 1 Then
            Response.Write "<font class=red_2>图片名称说明不能为空！</font><br><br>" & go_back
        ElseIf Len(pic) < 8 Then
            Response.Write "<font class=red_2>请上传图片或输入图片的地址！</font><br><br>" & go_back
        Else
            Call chk_cid_sid()
            Set rs = Server.CreateObject("adodb.recordset")
            sql    = "select * from " & data_name
            rs.open sql,conn,1,3
            rs.addnew
            rs("c_id")     = cid
            rs("s_id")     = sid
            rs("username")     = login_username
            rs("types")     = types
            rs("name")     = name

            If Len(code_admin(Request.form("spic"))) < 3 Then
                rs("spic") = "no_pic.gif"
            Else
                rs("spic") = code_admin(Request.form("spic"))
            End If

            rs("pic")     = pic
            rs("remark")     = Left(Request.form("remark"),250)
            rs("power")     = Replace(Replace(Trim(Request.form("power"))," ",""),",",".")

            If IsNumeric(Trim(Request.form("emoney"))) Then
                rs("emoney") = Trim(Request.form("emoney"))
            Else
                rs("emoney") = 0
            End If

            If Trim(Request.form("istop")) = "yes" Then
                rs("istop") = 1
            Else
                rs("istop") = 0
            End If

            rs("counter") = 0
            rs("tim") = now_time
            rs("hidden") = False
            rs.update
            rs.Close:Set rs = Nothing
            Call user_integral("add",add_integral,login_username)
            Call upload_note(action,first_id(action))
            Response.Write "<font class=red>已成功添加了一张图片！</font><br><br>请等待管理员审核通过……<br><br>"
        End If

    Else %><table border=0 cellspacing=0 cellpadding=3>
<form name=add_frm action='?action=<% Response.Write action %>' method=post>
<input type=hidden name=put value='yes'><input type=hidden name=upid value=''>
  <tr><td width='15%' align=center>文件名称：</td><td width='85%'><input type=text size=70 name=name maxlength=50><% = redx %></td></tr>
  <tr><td align=center>文件分类：</td><td><% Call chk_csid(cid,sid) %>&nbsp;&nbsp;文件类型：<select name=types size=1>
<option value='paste'<% If types = "paste" Then Response.Write " selected" %>>贴图</option>
<option value='flash'<% If types = "flash" Then Response.Write " selected" %>>FLASH</option>
<option value='film'<% If types = "film" Then Response.Write " selected" %>>视频</option>
<option value='logo'<% If types = "logo" Then Response.Write " selected" %>>LOGO</option>
<option value='baner'<% If types = "baner" Then Response.Write " selected" %>>BANNER</option>
</select><% Response.Write redx %>&nbsp;&nbsp;<% Call chk_emoney(0) %></td></tr>
  <tr><td align=center>浏览权限：</td><td><% Call chk_power("",1) %></td></tr>
<% ispic = "gs" & upload_time(now_time) %>
  <tr><td align=center>小 图 片：</td><td><input type=test name=spic size=70 maxlength=100></td></tr>
  <tr><td align=center>上传图片：</td><td><iframe frameborder=0 name=upload_frames width='100%' height=60 scrolling=no src='upload.asp?uppath=gallery&upname=<% Response.Write ispic %>&uptext=spic'></iframe></td></tr>
<% ispic = "g" & upload_time(now_time) %>
  <tr><td align=center>文件地址：</td><td><input type=test name=pic size=70 maxlength=100><% Response.Write redx %></td></tr>
  <tr><td align=center>上传文件：</td><td><iframe frameborder=0 name=upload_frame width='100%' height=90 scrolling=no src='upload.asp?uppath=gallery&upname=<% Response.Write ispic %>&uptext=pic'></iframe></td></tr>
  <tr><td align=center>文件说明：<br><br><=250字符</td><td><textarea name=remark rows=5 cols=70></textarea></td></tr>
  <tr><td colspan=2 align=center height=30><input type=submit value=' 上 传 我 的 文 件 '>　　<input type=reset value='重新填写'></td></tr>
</form></table><%
    End If

End Sub %>