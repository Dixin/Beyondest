<!-- #include file="include/config_forum.asp" -->
<% If Not(IsNumeric(forumid)) Then Call cookies_type("forum_id") %>
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Call forum_first()

Select Case action
    Case "manage"
        tit = forumname & " [版面管理]"
    Case "isgood"
        tit = forumname & " [本版精华]"
    Case "move"
    Case Else
        action = ""
        tit    = forumname
End Select

Call web_head(0,0,2,0,0)
'-----------------------------------center---------------------------------
Dim rssum
Dim nummer
Dim thepages
Dim viewpage
Dim page
Dim pageurl
Dim page_cute_num
Dim view_url
Dim topic_head
Dim del_temp
Dim keyword
Dim sea_type
Dim sea_true
Dim sea_write
Dim forum_temp
Dim id
Dim username
Dim icon
Dim topic
Dim tim
Dim counter
Dim re_counter
Dim re_username
Dim re_tim
Dim istop
Dim islock
Dim isgood
Dim folder_type
Dim forumnid

If action = "move" Then Call forum_moved(forumid,Trim(Request.querystring("view_id")))

If action = "manage" Then
    If format_user_power(login_username,login_mode,forumpower) <> "yes" Then action = ""
End If

pageurl       = "forum_list.asp?forum_id=" & forumid & "&action=" & action & "&"
rssum         = 0:thepages = 0:viewpage = 0:nummer = web_var(web_num,2):page_cute_num = web_var(web_num,3)
del_temp      = 0:forum_temp = ""
keyword       = code_form(Request.querystring("keyword"))
sea_type      = Trim(Request.querystring("sea_type"))

If (sea_type = "topic" Or sea_type = "username" Or sea_type = "re_username") And Len(keyword) > 0 Then
    sea_true  = "yes"
    sea_write = ".搜索"
    pageurl   = pageurl & "sea_type=" & sea_type & "&keyword=" & Server.htmlencode(keyword) & "&"
Else
    sea_true  = "no"
    sea_write = ""
End If

Select Case action
    Case "manage"
        Response.Write forum_top("帖子列表 [版面管理" & sea_write & "]")
    Case "isgood"
        Response.Write forum_top("帖子列表 [精华列表" & sea_write & "]")
    Case Else

        If sea_true = "yes" Then
            Response.Write forum_top("帖子列表 [搜索结果]")
        Else
            Response.Write forum_top("帖子列表 （主题：<font class=red>" & forumtopicnum & "</font>）")
        End If

End Select %>
<script language=javascript>
<!--
function load_tree(f_id,v_id){
  var targetImg =eval("document.all.followImg" + v_id);
  var targetDiv =eval("document.all.follow" + v_id);
  if (targetImg.src.indexOf("nofollow")!=-1){return false;}
    if ("object"==typeof(targetImg)){
      if (targetDiv.style.display!='block'){
        targetDiv.style.display="block";
        targetImg.src="images/small/fk_minus.gif";
        if (targetImg.loaded=="no"){
          document.frames["hiddenframe"].location.replace("forum_loadtree.asp?forum_id="+f_id+"&view_id="+v_id);
        }
      }else{
      targetDiv.style.display="none";
      targetImg.src="images/small/fk_plus.gif";
    }
  }
}
-->
</script>
<iframe width=0 height=0 src='about:blank' id=hiddenframe></iframe>
<table border=0 width='98%'><tr><td align=left width='15%'><a href='forum_write.asp?forum_id=<% = forumid %>'><img src='images/<% = web_var(web_config,5) %>/new_topic.gif' align=absMiddle border=0 title='在 <% = forumname %> 里发表我的新贴'></a></td><td align=right width='85%'><table border=0><form action='?' method=get><input type=hidden name=forum_id value='<% = forumid %>'><input type=hidden name=action value='<% = action %>'><input type=hidden name=page value='<% = viewpage %>'><tr><td>论坛搜索：</td><td><select name=sea_type size=1><option value='topic'>按主题</option><option value='username'>按作者</option><option value='re_username'>按回复人</option></select></td><td><input type=text name=keyword size=20 maxlength=20></td><td>&nbsp;<input type=submit value='搜 索'></td></tr></table></td></tr></table>
<% Response.Write forum_table1 %>
<tr align=center<% Response.Write forum_table2 %> height=25 >
<td rowap width='4%' class=end  background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>图</td>
<td rowap width='3%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>例</td>
<td rowap width='48%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>主题（每页&nbsp;<% Response.Write nummer %>&nbsp;贴&nbsp;&nbsp;点击&nbsp;<img src='IMAGES/SMALL/FK_PLUS.GIF' align=absMiddle border=0>&nbsp;可展开贴子列表）</td>
<td rowap width='12%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>作者</td>
<td rowap width='4%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;</td>
<td rowap width='7%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>人气</td>
<td rowap width='22%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>最后回复信息</td>
</tr>
<%

If action = "manage" Then
    Response.Write "<form name=del_form action='forum_isaction.asp?isaction=delete&forum_id=" & forumid & "' method=post>"
End If

If action = "isgood" Then
    sql           = "select count(id) from bbs_topic where forum_id=" & forumid & " and isgood=1"
    Set rs        = conn.execute(sql)
    forumtopicnum = rs(0)
    rs.Close
End If

If sea_true = "yes" Then
    sql           = "select count(id) from bbs_topic where forum_id=" & forumid & " and " & sea_type & " like '%" & keyword & "%'"
    Set rs        = conn.execute(sql)
    forumtopicnum = rs(0)
    rs.Close
End If

sql     = "select * from bbs_topic where (forum_id=" & forumid & " or istop=2) "

If action = "isgood" Then
    sql = sql & "and isgood=1"
End If

If sea_true = "yes" Then
    sql   = sql & "and " & sea_type & " like '%" & keyword & "%'"
End If

sql       = sql & " order by istop desc,re_tim desc,id desc"
Set rs    = conn.execute(sql)

If rs.eof And rs.bof Then
    rssum = 0
    Response.Write "<tr><td colspan=8 align=center height=50>本论坛暂时没有贴子。</td></tr>"
Else
    rssum = forumtopicnum		'rs.recordcount
End If

If Int(rssum) > 0 Then
    Call format_pagecute()

    If Int(viewpage) > 1 Then
        rs.move (viewpage - 1)*nummer
    End If

    For i = 1 To nummer
        If rs.eof Then Exit For
        folder_type = "isok"
        forumnid = rs("forum_id")
        id = rs("id")
        username = rs("username")
        topic = rs("topic")
        icon = rs("icon")
        tim = rs("tim")
        counter = rs("counter")
        re_counter = rs("re_counter")
        re_username = rs("re_username")
        re_tim = rs("re_tim")
        istop = rs("istop")
        islock = rs("islock")
        isgood = rs("isgood")

        Call forum_view()

        rs.movenext
    Next

End If

rs.Close:Set rs = Nothing

Response.Write "</table>"

If Int(thepages) < 1 Then page_cute_num = 1

Response.Write kong & forum_table1 %>
<tr height=25<% Response.Write forum_table3 %>>
<td width='35%'>
主题：<font class=red><% Response.Write forumtopicnum %></font>&nbsp;
<%

If action <> "isgood" And sea_true <> "yes" Then
    Response.Write "贴子总数：<font class=red>" & forumdatanum & "</font>&nbsp;"
End If %>
页次：<font class=red><% Response.Write viewpage & "</font>/<font class=red>" & thepages %></font><td align=center>
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,page_cute_num,"#ff0000") %></td>
</td><td align=center width='25%'><% Response.Write forum_go() %></td>
</tr>
<% If action = "manage" Then %>
<script language=javascript src='STYLE/admin_del.js'></script>
<tr<% Response.Write format_table(3,1) %>><td height=25 align=center colspan=3>版面管理：　<input type=checkbox name=del_all value=1 onClick="selectall('<% Response.Write del_temp %>')" class=bg_1> 选中所有　<input type=submit value='删除所选' onclick="return suredel('<% Response.Write del_temp %>');"></td></tr>
</form>
<% End If %>
</table>
<table border=0 width='95%'>

<tr><td align=center colspan=2 height=30>
<% Response.Write web_var(web_config,1) %>论坛主题图例：&nbsp;
<% Call is_type() %>
</td></tr>
</table>
<%

'---------------------------------center end-------------------------------
Call web_end(0) %>