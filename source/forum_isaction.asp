<!-- #include file="INCLUDE/config_forum.asp" -->
<% If Not(IsNumeric(forumid)) Then Call cookies_type("view_id") %>
<!-- #include file="INCLUDE/config_upload.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Call forum_first()
Call web_head(2,2,0,0,0)

If action = "istops" Then
    If format_user_power(login_username,login_mode,"") <> "yes" Then close_conn():Call cookies_type("power")
Else
    If format_user_power(login_username,login_mode,forumpower) <> "yes" Then close_conn():Call cookies_type("power")
End If

Dim isaction,delid
isaction = Trim(Request.querystring("isaction"))

Select Case isaction
    Case "del"
        Call is_del()
    Case "delete"
        Call is_delete()
    Case Else
        Call is_action()
End Select

Call close_conn()

Sub is_action()

    If Not(IsNumeric(viewid)) And (action <> "isgood" And action <> "islock" And action <> "istop" And action <> "istops") Then
        Call close_conn()
        Call cookies_type("del_id")
    End If

    Dim ismsg,ist,upss

    Select Case action
        Case "isgood"
            ist = "����"
        Case "islock"
            ist = "����"
        Case "istop"
            ist = "�̶�"
        Case "istops"
            ist = "�̶ܹ�"
    End Select

    If Trim(Request.querystring("cancel")) = "yes" Then
        If action = "istops" Then  action = "istop"
        upss  = 0
        ismsg = "�ѳɹ��Ķ����⣨ID��" & viewid & "��ȡ�� " & ist & " ��"
    Else

        If action = "istops" Then
            action = "istop"
            upss   = 2
        Else
            upss   = 1
        End If

        ismsg      = "�ѳɹ��Ľ����⣨ID��" & viewid & "����Ϊ " & ist & " ��"
    End If

    sql = "update bbs_topic set " & action & "=" & upss & " where id=" & viewid
    conn.execute(sql)

    Response.Write "<script language=javascript>" & _
    vbcrlf & "alert(""" & ismsg & "\n\n������ء�"");" & _
    vbcrlf & "location='forum_list.asp?forum_id=" & forumid & "'" & _
    vbcrlf & "</script>"
    'response.redirect "forum_list.asp?forum_id="&forumid
End Sub

Sub is_del()
    delid = Trim(Request.querystring("del_id"))

    If Not(IsNumeric(delid)) Then
        Call close_conn()
        Call cookies_type("del_id")
    End If

    Dim reid,username
    sql    = "select reply_id,username from bbs_data where forum_id=" & forumid & " and id=" & delid
    Set rs = conn.execute(sql)

    If rs.eof And rs.bof Then
        rs.Close:Set rs = Nothing
        Call close_conn()
        Call cookies_type("del_id")
    End If

    reid     = rs("reply_id")
    username = rs("username")
    rs.Close:Set rs = Nothing

    sql = "delete from bbs_data where id=" & delid
    conn.execute(sql)
    sql = "update bbs_topic set re_counter=re_counter-1 where id=" & reid
    conn.execute(sql)
    sql = "update bbs_forum set forum_data_num=forum_data_num-1 where forum_id=" & forumid
    conn.execute(sql)
    sql = "update configs set num_data=num_data-1 where id=1"
    conn.execute(sql)
    sql = "update user_data set bbs_counter=bbs_counter-1,integral=integral-2 where username='" & username & "'"
    conn.execute(sql)

    Response.Write "<script language=javascript>" & _
    vbcrlf & "alert(""�ɹ�ɾ����һ��������\n\n������ء�"");" & _
    vbcrlf & "location='forum_list.asp?forum_id=" & forumid & "'" & _
    vbcrlf & "</script>"
End Sub

Sub is_delete()
    delid = Trim(Request("del_id"))

    If Len(delid) < 1 Then
        Call close_conn()
        Call cookies_type("del_id")
    End If

    Dim del_dim,del_num,i,del_true,iok,ifail
    iok          = 0:ifail = 0
    delid        = Replace(delid," ","")
    del_dim      = Split(delid,",")
    del_num      = UBound(del_dim)

    For i = 0 To del_num
        del_true = forum_delete(del_dim(i))
        Call upload_del(index_url,del_dim(i))

        If del_true = "yes" Then
            iok   = iok + 1
        Else
            ifail = ifail + 1
        End If

    Next

    Erase del_dim
    Response.Write "<script language=javascript>" & _
    vbcrlf & "alert(""�ɹ�ɾ���� " & iok & " �����Ӽ��������\nɾ��ʧ�� " & ifail & " ����\n\n������ء�"");" & _
    vbcrlf & "location='forum_list.asp?forum_id=" & forumid & "'" & _
    vbcrlf & "</script>"
End Sub

Function forum_delete(did)
    Dim username,numd,sqladd
    did          = Trim(did)
    numd         = 1:sqladd = ""
    forum_delete = "yes"
    sql          = "select username from bbs_topic where forum_id=" & forumid & " and id=" & did
    Set rs       = conn.execute(sql)

    If rs.eof And rs.bof Then
        rs.Close:Set rs = Nothing
        forum_delete = "no":Exit Function
    End If

    username         = rs("username")
    rs.Close

    sql = "update user_data set bbs_counter=bbs_counter-1,integral=integral-3 where username='" & username & "'"
    conn.execute(sql)

    sql    = "select count(id) from bbs_data where forum_id=" & forumid & " and reply_id=" & did
    Set rs = conn.execute(sql)
    numd   = rs(0)
    rs.Close:Set rs = Nothing

    sql = "delete from bbs_data where reply_id=" & did
    conn.execute(sql)
    sql = "delete from bbs_topic where id=" & did
    conn.execute(sql)
    sql = "update bbs_forum set forum_topic_num=forum_topic_num-1,forum_data_num=forum_data_num-" & numd & " where forum_id=" & forumid
    conn.execute(sql)
    sql = "update configs set num_topic=num_topic-1,num_data=num_data-" & numd & " where id=1"
    conn.execute(sql)
End Function %>