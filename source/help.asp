<!-- #include file="include/config_other.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

index_url = "user_main"
tit       = "帮助中心"
tit_fir   = format_menu(index_url)

Call web_head(0,1,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Call help_left("jt12")
Response.Write left_action("jt13",4)
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong

Select Case action
    Case "about"
        Call help_about()
    Case "register"
        Call help_register()
    Case "put"
        Call help_put()
    Case "mail"
        Call help_mail()
    Case "forum"
        Call help_forum()
    Case "ubb"
        Call help_ubb()
    Case Else
        Call help_about()
        'call help_main()
End Select

Response.Write kong
'---------------------------------center end-------------------------------
Call web_end(1)

Sub help_left(ljt)
    If ljt <> "" Then ljt = img_small(ljt)
    tit = vbcrlf & "<table border=0 width='96%' cellpadding=0 cellspacing=6 align=center>" & _
    vbcrlf & "<tr><td width='50%'></td><td width='50%'></td></tr>" & _
    vbcrlf & "<tr><td>" & ljt & "<a href='?action=about'>关于我们</a></td><td>" & ljt & "<a href='?action=register'>注册说明</a></td></tr>" & _
    vbcrlf & "<tr><td>" & ljt & "<a href='?action=put'>发布信息</a></td><td>" & ljt & "<a href='?action=mail'>站内短信</a></td></tr>" & _
    vbcrlf & "<tr><td>" & ljt & "<a href='?action=forum'>论坛帮助</a></td><td>" & ljt & "<a href='?action=ubb'>UBB语法</a></td></tr>" & _
    vbcrlf & "</table>"
    Call left_type(tit,"help",1)
End Sub

Sub help_main()
    Response.Write table1 %>
<tr<% Response.Write table2 %>><td>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>帮助中心</b></font></td></tr>
<tr<% Response.Write table3 %>><td class=htd></td></tr>
<tr<% Response.Write table3 %>><td align=center>

</td></tr>
</table>
<%
End Sub

Sub help_register()
    Response.Write table1 %>
<tr<% Response.Write table2 %>><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>注册说明</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <br>
  <table border=0 width='94%'>
  <tr><td class=htd>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;欢迎您加入本站（<a href='<% = web_var(web_config,2) %>'><% = web_var(web_config,1) %></a>）参加交流和讨论，<a href='<% = web_var(web_config,2) %>'><% = web_var(web_config,1) %></a>为完全非赢利性、商业性的网站，<font color="#FF0000">我们的目的是推广Beyond的音乐，宣传Beyond的精神，研究相关的技术和艺术问题。我们承诺以更好地为广大歌迷朋友提供各种方便和服务为宗旨，因此我们的所有服务都是免费的，我们决不以任何理由向用户收取任何费用，决不向用户出售任何商品。</font><br><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;为维护网上公共秩序，请您自觉遵守以下条款<br><br>
　一、不得利用本站危害国家安全、泄露国家秘密，不得侵犯国家、社会、集体和公民的合法权益，不得利用本站制作、复制和传播下列信息： <br>
　　（一）煽动抗拒、破坏宪法和法律实施的；<br>
　　（二）煽动颠覆国家政权，推翻社会主义制度的；<br>
　　（三）煽动分裂国家、破坏国家统一的；<br>
　　（四）煽动民族仇恨、民族歧视，破坏民族团结的；<br>
　　（五）捏造或者歪曲事实，散布谣言，扰乱社会秩序的；<br>
　　（六）宣扬封建迷信、淫秽、色情、赌博、暴力、凶杀、恐怖、教唆犯罪的；<br>
　　（七）公然侮辱他人或者捏造事实诽谤他人的，或者进行其他恶意攻击的；<br>
　　（八）损害本站信誉的；<br>
　　（九）恶意使用污言秽语的；<br>
　　（十）进行商业性质的行为的。<br>
　二、互相尊重，对自己的言论和行为负责。<br>
　三、尊重我们的劳动成果。<br>
　　（一）转载本站资料请注明出处；<br>
　　（二）请勿将本站提供的资料用于商业用途。<br>

  </td></tr>
  <form name=form_reg action='login.asp?action=register' method=post>
  <input type=hidden name=reg_action value='reg_main'>
  <tr><td align=center height=30><input type=submit value='我已阅读并同意以上条款'></td></tr>
  </form>
  </table>
  <br>
</td></tr>
</table>
<%
End Sub

Sub help_put()
    Response.Write table1 %>
<tr<% Response.Write table2 %> height=25><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>发布信息</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr valign=top>
  <td width='30%'>
    <table border=0>
    <tr><td class=htd>　　您可以在本站发布一些如文章或软件等图文信息，通过管理员审核之后就可以在网站里显示出来和大家分享了，我们欢迎并感谢您为本站提供各种资料。<br>　　请在发布信息前先<a href='login.asp?action=register'>注册</a>并且<a href='login.asp'>登陆</a>本站，这样您就能正常的进行信息发布。</td></tr>
    </table>
  </td>
  <td width='3%'></td>
  <td width='67%'>
    <table border=0>
    <tr><td height=1 width='5%'></td><td width='95%'></td></tr>
    <tr><td colspan=2 height=20><% Response.Write img_small("jt1") %><a href='user_put.asp?action=news'>发布我的新闻</a></td></tr>
    <tr><td></td><td class=htd>发布关于Beyond的新闻，具体内容有：标题、内容、出处、关键字、图片（可上传）等，须管理员审核。</td></tr>
    <tr><td height=5></td></tr>
    <tr><td colspan=2 height=20><% Response.Write img_small("jt1") %><a href='user_put.asp?action=article'>发表我的文章</a></td></tr>
    <tr><td></td><td class=htd>发布关于Beyond的文章，具体内容有：标题、类型和内容等，须管理员审核。</td></tr>
    <tr><td height=5></td></tr>
    <tr><td colspan=2 height=20><% Response.Write img_small("jt1") %><a href='user_put.asp?action=down'>发布我的音乐</a></td></tr>
    <tr><td></td><td class=htd>和大家分享Beyond的精彩！具体内容有：名称、类型、大小、推荐等级、播放软件、说明、关键字、图片（可上传）等，须管理员审核。</td></tr>
    <tr><td height=5></td></tr>
    <tr><td colspan=2 height=20><% Response.Write img_small("jt1") %><a href='user_put.asp?action=gallery'>上传我的图片</a></td></tr>
    <tr><td></td><td class=htd>上传Beyond的图片或FLASH，具体内容有：名称、类型、说明、图片（可上传）等，须管理员审核。</td></tr>
    <tr><td height=5></td></tr>
    <tr><td colspan=2 height=20><% Response.Write img_small("jt1") %><a href='user_put.asp?action=website'>我要推荐网站</a></td></tr>
    <tr><td></td><td class=htd>推荐相关的网站或网页，具体内容有：名称、类型、地址、国家地区、站点语言、说明、图片（可上传）等，须管理员审核。</td></tr>
    </table>
  </td></tr>
  </table>
</td></tr>
</table>
<%
End Sub

Sub help_mail()
    Response.Write table1 %>
<tr<% Response.Write table2 %>><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>站内短信</b></font></td></tr>
<tr<% Response.Write table3 %>><td class=htd align=center height=30>站内短信可使你自如，安全地收发私人信息。不会被他人监听或查看到！</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='80%'>
  <tr><td width='5%'></td><td width='95%'></td></tr>
  <tr><td colspan=2 class=blue><% Response.Write img_small("jt0") %>发送站内短信</td></tr>
  <tr><td></td><td class=htd>首先请进入<a href='user_main.asp'>“用户中心”</a>点击“
  站内短信”，再在其中点击“发送消息”按钮，输入收件人的名字和信息主题。如果版块支持心情图释，所输入的图释代码将会自动转化为相应图片。注意，在按发送键前请确保已填写完所有的项目。</td></tr>
  <tr><td colspan=2 class=blue><% Response.Write img_small("jt0") %>收件箱</td></tr>
  <tr><td></td><td class=htd>您的收件箱中存放所有发给你的私人信息，您可以阅读或是删除它们。</td></tr>
  <tr><td colspan=2 class=blue><% Response.Write img_small("jt0") %>发件箱</td></tr>
  <tr><td></td><td class=htd>发件箱中则存放有您所发送过的全部消息记录，以使你清楚向谁发送过什么消息。除了阅读外，您删除它们！</td></tr>
  <tr><td colspan=2 class=blue><% Response.Write img_small("jt0") %>特别提醒</td></tr>
  <tr><td></td><td class=htd>请不要用此信使发送无聊或是使人不愉快的消息，尊重他人，也是尊重自己！</td></tr>
  </table>
</td></tr>
</table>
<%
End Sub

Sub help_forum()
    Response.Write table1 %>
<tr<% Response.Write table2 %>><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>论坛帮助</b></font></td></tr>
<tr<% Response.Write table4 %>><td class=htd align=center bgcolor=<% = web_var(web_color,6) %>><font class=red_3>本社区注册、发贴、回帖、删除帖子等操作对用户分值的影响如下说明所示：</font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='95%'>
  <tr><td colspan=2 class=btd>&nbsp;<font class=blue>（一）贴子</font></td></tr>
  <tr><td width='5%'></td><td width='95%'>注册初始贴子：<font class=red>0</font>&nbsp;&nbsp;发帖增加贴子：<font class=red>1</font>&nbsp;回帖增加贴子：<font class=red>1</font>&nbsp;&nbsp;删除增加贴子：<font class=red>1</font></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='95%'>
  <tr><td colspan=2 class=btd>&nbsp;<font class=blue>（二）积分</font></td></tr>
  <tr><td width='5%'></td><td width='95%'>注册初始积分：<font class=red>0</font>&nbsp;&nbsp;发帖增加积分：<font class=red>2</font>&nbsp;&nbsp;回帖增加积分：<font class=red>1</font>&nbsp;&nbsp;删除减少积分：主贴&nbsp;<font class=red>3</font>&nbsp;&nbsp;回贴&nbsp;<font class=red>2</font></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='95%'>
  <tr><td colspan=2 class=btd>&nbsp;<font class=blue>（三）金钱</font></td></tr>
  <tr><td width='5%'></td><td width='95%'>注册初始金钱：<font class=red>0</font>&nbsp;&nbsp;&nbsp;&nbsp;<font class=gray>其余待定……</font></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table4 %>><td class=htd align=center bgcolor=<% = web_var(web_color,6) %>><font class=red_3>本社区用户积分（等级）图例选项如下说明所示：</font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0>
  <tr align=center>
  <td height=30 width=100><% Response.Write img_small("icon_admin") %>管理员</td>
  <td width=100><img src='IMAGES/STAR/star_admin.gif'></td>
  <td width=100><% Response.Write img_small("icon_super") %>论坛版主</td>
  <td width=100><img src='IMAGES/STAR/star_super.gif'></td>
  </tr>
  </table>
  <table border=0>
  <tr align=center>
  <td height=30 width=50 align=left>级别</td>
  <td><% Response.Write img_small("icon_user") %>普通/<% Response.Write img_small("icon_puser") %>会员用户</td>
  <td><% Response.Write img_small("icon_vip") %>VIP用户</td>
  <td width=80>等级名称</td>
  <td>所需积分</td>
  </tr>
<%
    Dim sdim
    Dim sn
    Dim su:su = 0
    sdim = Split(user_grade,"|")

    For sn = 0 To UBound(sdim) %>
  <tr>
  <td><% Response.Write sn %>级</td>
  <td><img src='images/star/star_<% Response.Write sn %>.gif'></td>
  <td><img src='images/star/star_p<% Response.Write sn %>.gif'></td>
  <td align=center><% Response.Write Right(sdim(sn),Len(sdim(sn)) - InStr(sdim(sn),":")) %></td>
  <td><%

        If sn = Int(UBound(sdim)) Then
            Response.Write Left(sdim(sn),InStr(sdim(sn),":") - 1) & "分以上"
        Else
            Response.Write Left(sdim(sn),InStr(sdim(sn),":") - 1) & "-" & (Left(sdim(sn + 1),InStr(sdim(sn + 1),":") - 1) - 1)
        End If %></td>
  </tr>
<%
    Next

    Erase sdim %>
  </table>
</td></tr>
<tr<% Response.Write table4 %>><td class=htd align=center bgcolor=<% = web_var(web_color,6) %>><font class=red_3>本社区用户发贴和个人签名可用与不可能选项如下说明所示：</font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='90%'>
  <tr><td class=htd><% Response.Write web_var(web_error,3) & "<br>小于" & web_var(web_num,6) & "KB" %></td></tr>
  </table>
</td></tr>
</table>
<% Response.Write kong & table1 %>
<tr<% Response.Write table2 %>><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>论坛图例</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center height=30><% Response.Write ip_sys(0,0) %></td></tr>
<tr<% Response.Write table3 %>><td align=center height=30><% Response.Write user_power_type(0) %></td></tr>
<tr<% Response.Write table3 %>><td align=center height=30>
<% Response.Write img_small("isok") %>&nbsp;开放的主题&nbsp;&nbsp;
<% Response.Write img_small("ishot") %>&nbsp;回复超过10贴&nbsp;&nbsp;
<% Response.Write img_small("islock") %>&nbsp;锁定的主题&nbsp;&nbsp;
<% Response.Write img_small("istop") %>&nbsp;固定顶端的主题&nbsp;&nbsp;
<% Response.Write img_small("isgood") %>&nbsp;精华帖子
</td></tr>
</table>
<%
End Sub

Sub help_ubb()
    Response.Write table1 %>
<tr<% Response.Write table2 %>><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>UBB语法</b></font></td></tr>
<tr<% Response.Write table3 %>><td class=htd>　　以下为本站使用的UBB语法的具体使用说明，因为需要而进行了一些改进。UBB标签就是不允许使用HTML语法的情况下，通过特殊转换程序，以至可以支持少量常用的、无危害性的HTML效果显示。</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td class=htd>
<li><font color=red>[B]</font><B>文字</B><font color=red>[/B]</font>：在文字的位置可以任意加入您需要的字符，显示为粗体效果。</li>
<li><font color=red>[I]</font><I>文字</I><font color=red>[/I]</font>：在文字的位置可以任意加入您需要的字符，显示为斜体效果。</li>
<li><font color=red>[U]</font><U>文字</U><font color=red>[/U]</font>：在文字的位置可以任意加入您需要的字符，显示为下划线效果。</li>
<li><font color=red>[ALIGN=center]</font>文字<font color=red>[/ALIGN]</font>：在文字的位置可以任意加入您需要的字符，center位置center表示居中，left表示居左，right表示居右。</li>
<li><font color=red>[COLOR=颜色代码]</font>文字<font color=red>[/COLOR]</font>：输入您的颜色代码，在标签的中间插入文字可以实现文字颜色改变。</li>
<li><font color=red>[SIZE=数字]</font>文字<font color=red>[/SIZE]</font>：输入您的字体大小，在标签的中间插入文字可以实现文字大小改变。</li>
<li><font color=red>[FACE=字体]</font>文字<font color=red>[/FACE]</font>：输入您需要的字体，在标签的中间插入文字可以实现文字字体转换。</li>
<li><font color=red>[FLY]</font>飞翔的文字<font color=red>[/FLY]</font>：在标签的中间插入文字可以实现文字飞翔效果，类似跑马灯。</li>
<li><font color=red>[MOVE]</font>移动的文字<font color=red>[/MOVE]</font>：在标签的中间插入文字可以实现文字移动效果，为来回飘动。</li>
<li><font color=red>[GLOW=255,red,2]</font>文字<font color=red>[/GLOW]</font>：在标签的中间插入文字可以实现文字发光特效，glow内属性依次为宽度、颜色和边界大小。</li>
<li><font color=red>[SHADOW=255,red,2]</font>文字<font color=red>[/SHADOW]</font>：在标签的中间插入文字可以实现文字阴影特效，shadow内属性依次为宽度、颜色和边界大小。</li>
<li><font color=red>[URL]</font><A href="<% Response.Write web_var(web_config,2) %>"><% Response.Write web_var(web_config,2) %></A><font color=red>[/URL]</font></li>
<li><font color=red>[URL=<% Response.Write web_var(web_config,2) %>]</font><A href="<% Response.Write web_var(web_config,2) %>"><% Response.Write web_var(web_config,1) %></A><font color=red>[/URL]</font>：有两种方法可以加入超级连接，可以连接具体地址或者文字连接。</li>
<li><font color=red>[EMAIL]</font><A href="mailto:plinq@live.com">plinq@live.com</A><font color=red>[/EMAIL]</font></li>
<li><font color=red>[EMAIL=plinq@live.com]</font><A href="mailto:plinq@live.com">笼民</A><font color=red>[/EMAIL]</font>：有两种方法可以加入邮件连接，可以连接具体地址或者文字连接。</li>
<li><font color=red>[IMG]images/logo.gif[/IMG]</font> ：在标签的中间插入图片地址可以实现插图效果。
<li><font color=red>[DOWNLOAD]http://beyondest.com/music/test.rar[/DOWNLOAD]</font>：在标签的中间插入提供下载的文件地址可以实现文件下载效果。

<li><font color=red>[FLASH=宽度,高度]</font>Flash连接地址<font color=red>[/FLASH]</font>：在标签的中间插入Flash图片地址可以实现插入Flash。</li>
<li><font color=red>[CODE]</font>文字<font color=red>[/CODE]</font>：在标签中写入文字可实现html中编号效果。</li>
<li><font color=red>[OTE]</font>引用<font color=red>[/QUOTE]</font>：在标签的中间插入文字可以实现HTMl中引用文字效果。</li>
<li><font color=red>[RM=宽度,高度]</font>http://<font color=red>[/RM]</font>：为插入realplayer格式的rm文件，中间的数字为宽度和长度。</li>
<li><font color=red>[MP=宽度,高度]</font>http://<font color=red>[/MP]</font>：为插入为midia player格式的文件，中间的数字为宽度和长度。</li>
<li><font color=red>[DIR=宽度,高度]</font>http://<font color=red>[/DIR]</font>：为插入shockwave格式文件，中间的数字为宽度和长度。</li>
<li><font color=red>[QT=500,350]</font>http://<font color=red>[/QT]</font>：为插入为Quick time格式的文件，中间的数字为宽度和长度。</li>
  </td></tr>
  </table>
</td></tr>
<tr<% Response.Write table2 %>><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>EM 贴图</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0>
  <tr><td width=80></td><td></td></tr>
  <tr align=center>
  <td>小 EM 贴图<br>（1-8）</td>
  <td>
    <table border=0>
    <tr align=center><%

    For i = 1 To 8
        Response.Write vbcrlf & "    <td width=50><img src='images/icon/em" & i & ".gif' border=0></td>"
    Next %></tr>
    <tr align=center><%

    For i = 1 To 8
        Response.Write vbcrlf & "    <td>[em" & i & "]</td>"
    Next %></tr>
    </table>
  </td>
  </tr>
  <tr><td colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr align=center>
  <td>小 EM 贴图<br>（9-16）</td>
  <td>
    <table border=0>
    <tr align=center><%

    For i = 9 To 16
        Response.Write vbcrlf & "    <td width=50><img src='images/icon/em" & i & ".gif' border=0></td>"
    Next %></tr>
    <tr align=center><%

    For i = 9 To 16
        Response.Write vbcrlf & "    <td>[em" & i & "]</td>"
    Next %></tr>
    </table>
  </td>
  </tr>
  <tr><td colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr align=center>
  <td>大 EM 贴图<br>（1-7）</td>
  <td>
    <table border=0>
    <tr><%

    For i = 1 To 7
        Response.Write vbcrlf & "    <td width=60><img src='images/icon/emb" & i & ".gif' border=0></td>"
    Next %></tr>
    <tr><%

    For i = 1 To 7
        Response.Write vbcrlf & "    <td>[emb" & i & "]</td>"
    Next %></tr>
    </table>
  </td>
  </tr>
  <tr><td colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr>
  <td align=center>大 EM 贴图<br>（8-13）</td>
  <td>
    <table border=0>
    <tr><%

    For i = 8 To 13
        Response.Write vbcrlf & "    <td width=60><img src='images/icon/emb" & i & ".gif' border=0></td>"
    Next %></tr>
    <tr><%

    For i = 8 To 13
        Response.Write vbcrlf & "    <td>[emb" & i & "]</td>"
    Next %></tr>
    </table>
  </td>
  </tr>
  </table>
</td></tr>
</table>
<%
End Sub

Sub help_about()
    Response.Write table1 %>
<tr<% Response.Write table2 %>><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>关于我们</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr>
  <td width='100%' class=htd>　　本站（<a href='<% = web_var(web_config,2) %>'><% = web_var(web_config,1) %></a>）正式创建于1999年12月，属于我国比较早的一批Beyond网站。当时最初的版本为纯静态的html，并将重点放在了视觉设计上，经过几次较大规模的升级与改进，我们的越来越注重技术的提高。由于经济条件和学业的原因，从2000年下半年起本站的发展一直局限在很小的范围内，并最终非常勉强地维持到了2003年。但是经过4年的资料收集和技术提高，我们的不懈努力和积累终于使<a href='<% = web_var(web_config,2) %>'><% = web_var(web_config,1) %></a>初具规模。本站现采用了asp+access技术，今后将向以技术为主的综合型网站发展。<br>
　　<a href='<% = web_var(web_config,2) %>'><% = web_var(web_config,1) %></a>长期以来始终坚持“自由”的精神，是完全非赢利性、非商业性的网站。我们以更好地为广大歌迷朋友提供各种方便和服务为宗旨，因此我们的所有网站内容和服务和都是免费的，我们决不以任何理由向用户收取任何费用，决不向用户出售任何商品。我们所做的一切，目的仅仅是推广Beyond的音乐，宣传Beyond的精神，研究相关的技术和艺术问题。
</td>
  </tr>
  </table>
</td></tr>
</table>

<br>

<% Response.Write table1 %>
<tr<% Response.Write table2 %>><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk0") %>&nbsp;<font class=end><b>关于我</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr>
        <td width='100%' class=htd><br><center><img src=images/yandixin.jpg></center>
          <br>　　有的人做网站，是因为社会有这样一种商品需求；有的人做网站，是因为他有一些话要说。但是真正的勇士，敢于直面惨淡的人生。<br>
<br>　　每每在那些失眠的夜里，恍惚迷漓中听一些老歌，想起和那个歌手一样离去的人和事，想起那些如火如荼、挥斥方遒的时节……，一阵伤怀：今夜，只有我还在！音乐！音乐！！！带着它的锐利，穿透时空，穿透我的心脏！<br>
<br>


　　听好的音乐，创造好的生活！！！




</td>
  </tr>
  </table>
</td></tr>
</table>
<%
End Sub %>