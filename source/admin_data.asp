<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit = "<a href='admin_update.asp?'>网站更新</a> ┋ " & _
"<a href='admin_data.asp'>数据更新</a> ┋ " & _
"<a href='admin_update.asp?nsort=news'>更新公告</a> ┋ " & _
"<a href='admin_update.asp?nsort=forum'>论坛公告</a> ┋ " & _
"<a href='admin_update.asp?action=add'>新增公告</a>"
Response.Write header(7,tit)

Dim actions:actions = Trim(Request.querystring("actions"))

Select Case action
    Case "update_config"
        Call update_config()
    Case "update_forum"
        Call update_forum()
    Case "clear_notes"
        Call clear_notes()
    Case "clear_message"
        Call clear_message()
End Select

Sub update_config()
    Dim rs,sql,num_topic,num_data,num_reg,new_username,num_news,num_article,num_down,num_flash,num_film,num_photo,num_desktop
    num_reg = 0:num_topic = 0:num_data = 0:num_news = 0:num_article = 0:num_down = 0:num_flash = 0:num_film = 0:num_photo = 0:num_desktop = 0
    Set rs  = Server.CreateObject("adodb.recordset")
    sql     = "select username from user_data order by id desc"
    rs.open sql,conn,1,1

    If Not(rs.eof And rs.bof) Then
        num_reg      = Int(rs.recordcount)
        new_username = rs("username")
    End If

    rs.Close

    sql    = "select count(id) from bbs_topic"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_topic = Int(rs(0))
    rs.Close

    sql    = "select count(id) from bbs_data"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_data = Int(rs(0))
    rs.Close

    sql    = "select count(id) from news where hidden=1"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_news = Int(rs(0))
    rs.Close

    sql    = "select count(id) from article where hidden=1"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_article = Int(rs(0))
    rs.Close

    sql    = "select count(id) from down where hidden=1"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_down = Int(rs(0))
    rs.Close

    sql    = "select count(id) from gallery where hidden=1 and types='flash'"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_flash = Int(rs(0))
    rs.Close

    sql    = "select count(id) from gallery where hidden=1 and types='film'"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_film = Int(rs(0))
    rs.Close

    sql    = "select count(id) from gallery where hidden=1 and types='baner'"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_photo = Int(rs(0))
    rs.Close

    sql    = "select count(id) from gallery where hidden=1 and types='paste'"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_desktop = Int(rs(0))
    rs.Close

    sql = "update configs set num_topic=" & num_topic & ",num_data=" & num_data & ",num_reg=" & num_reg & ",new_username='" & new_username & "',num_news=" & num_news & ",num_article=" & num_article & ",num_down=" & num_down & ",num_flash=" & num_flash & ",num_film=" & num_film & ",num_photo=" & num_photo & ",num_desktop=" & num_desktop & " where id=1"
    conn.execute(sql)

    Response.Write "<script language=javascript>alert(""成功更新了网站统计数据！"");</script>"
End Sub

Sub update_forum()
    Dim rsf,sqlf,rssum,i,rs,sql,forumid,t1,t2,t3
    sqlf        = "select * from bbs_forum order by forum_id"
    Set rsf     = conn.execute(sqlf)

    Do While Not rsf.eof
        forumid = rsf("forum_id")
        Set rs  = Server.CreateObject("adodb.recordset")
        sql     = "select * from bbs_topic where forum_id=" & forumid & " order by id desc"
        rs.open sql,conn,1,1

        If rs.eof And rs.bof Then
            t1 = 0
            t2 = "|||"
        Else
            t1 = rs.recordcount
            t2 = rs("username") & "|" & rs("tim") & "|" & rs("id") & "|" & rs("topic")
            t2 = Replace(t2,"'","")
        End If

        rs.Close:Set rs = Nothing

        sql    = "select count(*) from bbs_data where forum_id=" & forumid
        Set rs = conn.execute(sql)
        t3     = rs(0)
        rs.Close:Set rs = Nothing
        If Int(t3) < 1 Then t3 = 0

        sql = "update bbs_forum set forum_topic_num=" & t1 & ",forum_new_info='" & t2 & "',forum_data_num=" & t3 & " where forum_id=" & forumid
        conn.execute(sql)
        rsf.movenext
    Loop

    rsf.Close:Set rsf = Nothing

    Response.Write "<script language=javascript>alert(""成功更新了分论坛数据！"");</script>"
End Sub

Sub clear_notes()
    Dim clear_msg

    Select Case actions
        Case "week"
            sql       = "delete from notes where DateDiff('d',tim,'" & now_time & "')>7"
            clear_msg = "一周前"
        Case "all"
            sql       = "delete from notes"
            clear_msg = "所有"
        Case Else
            sql       = "delete from notes where DateDiff('d',tim,'" & now_time & "')>30"
            clear_msg = "一个月前"
    End Select

    conn.execute(sql)
    Response.Write "<script language=javascript>alert(""成功清理了" & clear_msg & "的消费数据！"");</script>"
End Sub

Sub clear_message()
    Dim clear_msg

    Select Case actions
        Case "week"
            sql = "delete from user_mail where DateDiff('d',tim,'" & now_time & "')>7"
            clear_msg = "一周前"
        Case "all"
            sql = "delete from user_mail"
            clear_msg = "所有"
        Case Else
            sql = "delete from user_mail where DateDiff('d',tim,'" & now_time & "')>30"
            clear_msg = "一个月前"
    End Select

    conn.execute(sql)
    Response.Write "<script language=javascript>alert(""成功清理了" & clear_msg & "的用户短信息！"");</script>"
End Sub %>
<table border=1 cellspacing=0 cellpadding=2 width=500 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<tr height=50 align=center>
<td width='20%'><font class=red_2>注意事项</font></td>
<td width='80%'>下面有的操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行！</td>
</tr>
<tr align=center height=80>
<td><font class=red_3>更新网站总数据</font></td>
<td class=htd>点击下面的按钮将重新计算整个论网站的基本信息，包括新闻、文章、软件、论坛等，建议每隔一段时间运行一次。<br>
<input type=button value='立即更新网站统计数据' onclick=update_config() class=red></td>
</tr>
<tr align=center height=80>
<td><font class=red_3>更新分论坛数据</font></td>
<td class=htd>点击下面的按钮将重新计算每个论坛的帖子主题、回复数量，最新主题、回复、时间等信息，建议每隔一段时间运行一次。<br>
<input type=button value='立即更新分论坛数据' onclick=update_forum() class=red></td>
</tr>
<tr align=center height=80>
<td><font class=red_3>清理消费数据</font></td>
<td class=htd>点击下面的按钮将重新清理用户消费货币的数据信息，建议每隔一段时间运行一次。<br>
<input type=button value='清除一个月前的数据' onclick="javascript:clear_notes('month');" class=red style="width:140px;">&nbsp;&nbsp;
<input type=button value='清除一周前的数据' onclick="javascript:clear_notes('week');" class=red style="width:120px;">&nbsp;&nbsp;
<input type=button value='全部清空' onclick="javascript:clear_notes('all');" class=red>
</td>
</tr>
<tr align=center height=80>
<td><font class=red_3>清理用户短信息</font></td>
<td class=htd>点击下面的按钮将重新清理用户短信息的数据信息，建议每隔一段时间运行一次。<br>
<input type=button value='清除一个月前的短信' onclick="javascript:clear_message('month');" class=red style="width:140px;">&nbsp;&nbsp;
<input type=button value='清除一周前的短信' onclick="javascript:clear_message('week');" class=red style="width:120px;">&nbsp;&nbsp;
<input type=button value='全部清空' onclick="javascript:clear_message('all');" class=red>
</td>
</tr>
</table>
<script language=JavaScript>
<!--
function update_config()
{
  if (confirm("此操作将 更新分论坛数据！\n\n真的要进行吗？\n\n更新后将无法恢复！"))
    window.location="?action=update_config"
}

function update_forum()
{
  if (confirm("此操作将 更新网站统计数据！\n\n真的要进行吗？\n\n更新后将无法恢复！"))
    window.location="?action=update_forum"
}

function clear_notes(cv)
{
  if (confirm("此操作将 清理消费数据！\n\n真的要进行吗？\n\n更新后将无法恢复！"))
    window.location="?action=clear_notes&actions="+cv
}

function clear_message(cv)
{
  if (confirm("此操作将 清理用户短信息！\n\n真的要进行吗？\n\n更新后将无法恢复！"))
    window.location="?action=clear_message&actions="+cv
}
//-->
</script>
<%
Call close_conn()
Response.Write ender() %>