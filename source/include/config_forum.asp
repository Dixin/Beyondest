<!-- #include file="config.asp" -->
<!-- #include file="skin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim forum_mode,forum_table1,forum_table2,forum_table3,forum_table4,ptnums,ffk
Dim forumid,viewid,forumname,forumpower,forumtype,forumtopicnum,forumdatanum,word_size,word_remark
forum_table1 = format_table(1,3)
forum_table2 = format_table(3,6)
forum_table3 = format_table(3,5)
forum_table4 = format_table(3,1)
forumid      = Trim(Request.querystring("forum_id"))
viewid       = Trim(Request.querystring("view_id"))

ffk          = "fk4"
index_url    = "forum"
tit_fir      = format_menu(index_url)
ptnums       = web_var_num(web_setup,6,1)

'-------------------------------------��ʼ�� 1--------------------------------------
Sub forum_first()
    sql    = "select forum_name,forum_power,forum_topic_num,forum_data_num,forum_type " & _
    "from bbs_forum where forum_id=" & forumid & " and forum_hidden=0"
    Set rs = conn.execute(sql)

    If rs.eof And rs.bof Then
        rs.Close:Set rs = Nothing
        Call close_conn()
        Call cookies_type("forum_id")
    End If

    forumname     = rs("forum_name"):forumpower = rs("forum_power")
    forumtopicnum = rs("forum_topic_num"):forumdatanum = rs("forum_data_num"):forumtype = rs("forum_type")
    rs.Close
    page_power    = format_forum_type(forumtype,0)
End Sub

'-------------------------------------��̳��ͷ--------------------------------------
Function forum_top(ft)
    forum_top = vbcrlf & ukong & forum_table1 & _
    vbcrlf & "<tr " & forum_table2 & ">" & _
    vbcrlf & "<td width='70%'>&nbsp;&nbsp;" & _
    vbcrlf & "��������<a href='forum_list.asp?forum_id=" & forumid & "'>" & forumname & "</a> &nbsp;- &nbsp;" & ft & "<font class=gray>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��<a href='forum_list.asp?forum_id=" & forumid & "&action=isgood'>���澫��</a>��&nbsp;��<a href='forum_list.asp?forum_id=" & forumid & "&action=manage'>�������</a>��</font></td>" & _
    vbcrlf & "<td align=right>������" & forum_power(forumpower,ptnums) & "&nbsp;" & _
    vbcrlf & "</td>" & _
    vbcrlf & "</tr></table>" & _
    vbcrlf & "" & ukong
End Function

'-------------------------------------��������--------------------------------------
Sub forum_word()
    word_size   = web_var(web_num,6)
    word_remark = web_var(web_error,3) & "<br>����<=" & word_size & "KB"
End Sub

'-------------------------------------��̳����--------------------------------------
Function forum_power(forum_admin,ft)
    Dim forumadmin,k
    forum_power = "<img src='images/small/forum_power.gif' title='��̳����' align=absmiddle border=0>&nbsp;"
    If ft = 0 Then forum_power = forum_power & "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}""><option>�������</option><option>--------</option>"

    If forum_admin <> "" And Not IsNull(forum_admin) Then
        forumadmin = Split(forum_admin, "|")

        For k = 0 To UBound(forumadmin)

            If ft = 0 Then
                forum_power = forum_power & "<option value='user_view.asp?username=" & Server.urlencode(forumadmin(k)) & "'>" & forumadmin(k) & "</option>"
            Else
                forum_power = forum_power & "<a href='user_view.asp?username=" & Server.urlencode(forumadmin(k)) & "' title='�鿴��������" & forumadmin(k) & " ����ϸ����' target=_blank>" & forumadmin(k) & "</a>&nbsp;"
            End If

        Next

        Erase forumadmin
    Else

        If ft = 0 Then
            forum_power = forum_power & "<option>��û��</option>"
        Else
            forum_power = forum_power & "<font class=gray>��û��&nbsp;</font>"
        End If

    End If

    If ft = 0 Then forum_power = forum_power & "</select>"
End Function

'-------------------------------------��̳�ȼ�--------------------------------------
Function format_forum_type(fvars,ft)
    Dim fdim,fvar:fvar = fvars - 1:format_forum_type = ""
    fdim = Split(forum_type,"|")

    For i = 0 To UBound(fdim)

        If ft = 0 Then
            If fvar = i Then format_forum_type = Left(fdim(i),InStr(fdim(i),":") - 1):Exit For
        Else
            If fvar = i Then format_forum_type = Right(fdim(i),Len(fdim(i)) - InStr(fdim(i),":")):Exit For
        End If

    Next

    Erase fdim
End Function

'-----------------------------------����ת�Ʋ���------------------------------------
Sub forum_moved(fid,vid)

    If Not(IsNumeric(fid)) Or Not(IsNumeric(vid)) Or login_mode <> format_power2(1,1) Then Response.Write "<script language=javascript>alert(""ת������ʧ�ܣ�\n\n�������������˲��ʺϵĲ�����"");</script>":Exit Sub
        Dim frs,fsql
        fsql    = "select forum_id from bbs_topic where id=" & vid
        Set frs = conn.execute(fsql)

        If frs.eof And frs.bof Then
            frs.Close:Set frs = Nothing:close_conn

            Call cookies_type("view_id"):Exit Sub
            End If

            frs.Close:Set frs = Nothing
            fsql = "update bbs_topic set forum_id=" & fid & " where id=" & vid
            conn.execute(fsql)
            fsql = "update bbs_data set forum_id=" & fid & " where reply_id=" & vid
            conn.execute(fsql)
            Response.Write "<script language=javascript>alert(""ת������ɹ���"");</script>"
        End Sub

        '-------------------------------------����ת��--------------------------------------

        Function forum_move(fmfid,fmid)
            Dim rsclass,strsqlclass,rsboard,strsqlboard,fid
            forum_move  = vbcrlf & "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & _
            vbcrlf & "<option selected>��������ת����...</option>"
            strsqlclass = "select class_id,class_name from bbs_class order by class_order"
            Set rsclass = conn.execute(strsqlclass)

            If Not(rsclass.bof And rsclass.eof) Then

                Do While Not rsclass.eof
                    forum_move     = forum_move & vbcrlf & "<option class=bg_2>�� " & rsclass("class_name") & "</option>"
                    strsqlboard    = "select forum_id,forum_name from bbs_forum where class_id=" & rsclass("class_id") & " and forum_hidden=0 order by forum_order"
                    Set rsboard    = conn.execute(strsqlboard)

                    If rsboard.eof And rsboard.bof Then
                        forum_move = forum_move & vbcrlf & "<option>û����̳</option>"
                    Else

                        Do While Not rsboard.eof
                            fid        = rsboard("forum_id")
                            forum_move = forum_move & vbcrlf & "<option"
                            If Int(fid) <> Int(fmfid) Then  forum_move = forum_move & " value='forum_list.asp?action=move&view_id=" & fmid & "&forum_id=" & fid & "'"
                            forum_move = forum_move & ">����" & rsboard("forum_name") & "</option>"
                            rsboard.movenext
                        Loop

                    End If

                    rsclass.movenext
                Loop

            End If

            Set rsclass = Nothing:Set rsboard = Nothing
            forum_move  = forum_move & vbcrlf & "</select>"
        End Function

        '-------------------------------------��̳��ת--------------------------------------

        Function forum_go()
            Dim rsclass,strsqlclass,rsboard,strsqlboard
            forum_go    = vbcrlf & "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & _
            vbcrlf & "<option selected>������ת��̳��...</option>"
            strsqlclass = "select class_id,class_name from bbs_class order by class_order"
            Set rsclass = conn.execute(strsqlclass)

            If Not(rsclass.bof And rsclass.eof) Then

                Do While Not rsclass.eof
                    forum_go     = forum_go & vbcrlf & "<option class=bg_2>�� " & rsclass("class_name") & "</option>"
                    strsqlboard  = "select forum_id,forum_name from bbs_forum where class_id=" & rsclass("class_id") & " and forum_hidden=0 order by forum_order"
                    Set rsboard  = conn.execute(strsqlboard)

                    If rsboard.eof And rsboard.bof Then
                        forum_go = forum_go & vbcrlf & "<option>û����̳</option>"
                    Else

                        Do While Not rsboard.eof
                            forum_go = forum_go & vbcrlf & "<option value='forum_list.asp?forum_id=" & rsboard("forum_id") & "'>����" & rsboard("forum_name") & "</option>"
                            rsboard.movenext
                        Loop

                    End If

                    rsclass.movenext
                Loop

            End If

            Set rsclass = Nothing:Set rsboard = Nothing
            forum_go    = forum_go & vbcrlf & "<option class=bg_2>����������������</option>" & _
            vbcrlf & "<option value='forum.asp' class=bg_1>" & tit_fir & "��ҳ</option>" & _
            vbcrlf & "<option class=bg_2>����������������</option>" & _
            vbcrlf & "<option value='forum_action.asp?action=new'>���� ��̳����</option>" & _
            vbcrlf & "<option value='forum_action.asp?action=tim'>���� �ظ�����</option>" & _
            vbcrlf & "<option value='user_action.asp?action=list'>���� �û��б�</option>" & _
            vbcrlf & "<option value='help.asp?action=forum'>���� ��̳����</option>" & _
            vbcrlf & "</select>"
        End Function

        '-------------------------------------�����ҳ--------------------------------------

        Function index_pagecute(viewurl,replynum,pagecutenum,pagecutecolor)
            Dim pagecutepage,pagecutei
            index_pagecute   = ""

            If replynum Mod pagecutenum > 0 Then
                pagecutepage = replynum\pagecutenum + 1
            Else
                pagecutepage = replynum\pagecutenum
            End If

            If pagecutepage > 1 Then

                For pagecutei = 2 To 3
                    If pagecutei > pagecutepage Then Exit For
                    index_pagecute = index_pagecute & vbcrlf & "<a href='" & viewurl & "&page=" & pagecutei & "'><font color='" & pagecutecolor & "' title='�� " & pagecutei & " ҳ'>[" & pagecutei & "]</font></a>"
                Next

                If pagecutepage > 3 Then

                    If pagecutepage = 4 Then
                        index_pagecute = index_pagecute & vbcrlf & "<a href='" & viewurl & "&page=4'><font color='" & pagecutecolor & "' title='�� 4 ҳ'>[4]</font></a>"
                    Else
                        index_pagecute = index_pagecute & vbcrlf & "<font color='" & pagecutecolor & "'>�� </font>" & "<a href='" & viewurl & "&page=" & pagecutepage & "'><font color='" & pagecutecolor & "' title='�� " & pagecutepage & " ҳ'>[" & pagecutepage & "]</font></a>"
                    End If

                End If

            End If

            If Len(index_pagecute) > 1 Then index_pagecute = "<img src='images/small/page_head.gif' align=absMiddle alt='���ٷ�ҳ' border=0>" & index_pagecute
        End Function

        '---------------------------------------main----------------------------------------

        Sub forum_down(dt)
            Dim udim,ui,j,dts,sql,rs,l_username,forum_table4,online
            online = Trim(Request.querystring("online"))
            j      = 5:dts = 0:forum_table4 = format_table(3,1)
            If forum_mode = "full" Then j = 8
            If online = "open" Or dt = 1 Then dts = 1
            If online = "close" Then dts = 0
            Response.Write forum_table1 %>
<tr<% Response.Write forum_table2 %> height=25><td background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small("fk4") %>&nbsp;<font class=end><b>��̳ͼ��</b></font></td></tr>
<tr<% Response.Write forum_table4 %>><td align=left height=30>&nbsp;&nbsp;<% Response.Write ip_sys(0,0) %></td></tr>
<tr<% Response.Write forum_table4 %>><td align=left height=30>&nbsp;&nbsp;<% Response.Write user_power_type(0) %></td></tr>
<tr<% Response.Write forum_table4 %>><td align=left>
  <table border=0 width='100%'>
  <tr><td colspan=5>&nbsp;��վ��ǰ�û����ߣ�<font class=red><%
            sql    = "select count(l_id) from user_login where l_type=0"
            Set rs = conn.execute(sql)
            Response.Write rs(0)
            rs.Close
            Response.Write "</font> ��  [ <a href='?mode=" & forum_mode & "&online="

            If dts = 0 Then
                Response.Write "open'>��"
            Else
                Response.Write "close'>�ر�"
            End If %>�����б�</a> ] </td></tr>
<% If dts <> 0 Then %>
  <tr><td width='20%'></td><td width='20%'></td><td width='20%'></td><td width='20%'></td><td width='20%'></td></tr>
<%
                sql    = "select user_login.*,user_data.power from user_data inner join user_login on user_login.l_username=user_data.username where user_login.l_type=0 order by user_login.l_id"
                Set rs = conn.execute(sql)

                Do While Not rs.eof
                    Response.Write "<tr>"

                    For ui = 1 To 5
                        If rs.eof Then Exit For
                        l_username = rs("l_username")
                        Response.Write "<td>&nbsp;" & img_small("icon_" & rs("power")) & "<a href='user_view.asp?username=" & Server.urlencode(l_username) & "' title='Ŀǰλ�ã�" & rs("l_where") & "<br>����ʱ�䣺" & rs("l_tim_login") & "<br>�ʱ�䣺" & rs("l_tim_end") & "<br>" & ip_types(rs("l_ip"),l_username,0) & "<br>" & view_sys(rs("l_sys")) & "' target=_blank>" & l_username & "</a></td>"
                        rs.movenext
                    Next

                    Response.Write "</tr>"
                Loop

                rs.Close
            End If

            Set rs = Nothing %>
  </table>
</td></tr>
</table>
<table border=0 width='100%'>
<tr><td align=center height=50>
<%
            udim   = Split(forum_type,"|")

            For ui = 0 To UBound(udim)
                Response.Write vbcrlf & "&nbsp;<img src='images/small/label_" & ui + 1 & ".gif' border=0 align=absmiddle>&nbsp;" & Right(udim(ui),Len(udim(ui)) - InStr(udim(ui),":")) & "&nbsp;"
            Next

            Erase udim %>
</td></tr>
</table>
<%
            Response.Write kong
        End Sub

        Sub forum_cast(nh,nj,n_num,c_num)
            Dim temp1,njj,topic,tbb:njj = ""
            If nj <> "" Then njj = img_small(nj)
            temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
            sql       = "select top " & n_num & " id,topic,username,tim from bbs_cast where sort='forum' order by id desc"
            Set rs    = conn.execute(sql)

            Do While Not rs.eof
                topic = rs("topic")
                temp1 = temp1 & "<tr><td class=bw height=" & space_mod & ">" & njj & "<a href='update.asp?action=forum&id=" & rs("id") & "' target=_blank title='������⣺" & code_html(topic,1,0) & "<br>�� �� Ա��" & rs("username") & "<br>����ʱ�䣺" & time_type(rs("tim"),88) & "'>" & code_html(topic,1,c_num) & "</a></td></tr>"
                rs.movenext
            Loop

            temp1 = temp1 & "</table>"
            Response.Write format_barc("<font class=end><b>��̳����</b></font>",temp1,2,0,7)
        End Sub

        Sub forum_new(nh,nj,fid,n_num,c_num,tb)
            Dim temp1,njj,topic,tbb:njj = "":tbb = ""
            If nj <> "" Then njj = img_small(nj)
            If tb = 1 Then tbb = " target=_blank"
            temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
            sql       = "select top " & n_num & " id,forum_id,topic,tim,username,re_username from bbs_topic"
            If fid > 0 Then sql = sql & " where forum_id=" & fid
            sql       = sql & " order by id desc"
            Set rs    = conn.execute(sql)

            Do While Not rs.eof
                topic = rs("topic")
                temp1 = temp1 & "<tr><td class=bw height=" & space_mod & ">" & njj & "<a href='forum_view.asp?forum_id=" & rs("forum_id") & "&view_id=" & rs("id") & "'" & tbb & " title='�������⣺" & code_html(topic,1,c_num) & "<br>�� �� �ˣ�" & rs("username") & "<br>����ʱ�䣺" & time_type(rs("tim"),88) & "<br>���ظ���" & rs("re_username") & "'>" & code_html(topic,1,c_num) & "</a></td></tr>"
                rs.movenext
            Loop

            temp1 = temp1 & "</table>"
            Response.Write format_barc("<font class=end><b>��̳����</b></font>",temp1,2,0,8)
        End Sub

        Function forum_main(mh)
            Dim rsclass,strsqlclass,rsforum,strsqlforum,rstopic,sqltopic,topics,classid,forumid,forumname,forum_type,forum_new_info,forum_pic,new_info_dim,forumpic
            strsqlclass            = "select class_id,class_name from bbs_class order by class_order"
            Set rsclass            = conn.execute(strsqlclass)

            Do While Not rsclass.eof
                classid            = rsclass("class_id")
                Response.Write vbcrlf & forum_table1 & "<tr" & forum_table2 & "><td height=25 colspan=4 background=images/" & web_var(web_config,5) & "/bar_3_bg.gif>&nbsp;" & img_small(mh) & vbcrlf & "<font class=end><b>" & rsclass("class_name") & "</b></font></td></tr>"
                strsqlforum        = "select forum_id,forum_name,forum_type,forum_new_info,forum_topic_num,forum_data_num,forum_power,forum_remark,forum_pic " & _
                "from bbs_forum where class_id=" & classid & " and forum_hidden=0 order by forum_order,forum_id desc"
                Set rsforum        = conn.execute(strsqlforum)

                Do While Not rsforum.eof
                    forumid        = rsforum("forum_id"):forumname = rsforum("forum_name")
                    forum_type     = rsforum("forum_type")
                    forum_new_info = rsforum("forum_new_info")
                    forum_pic      = rsforum("forum_pic")

                    If Len(forum_new_info) > 3 Then
                        new_info_dim        = Split(forum_new_info,"|")
                        new_info_dim(0)     = format_user_view(new_info_dim(0),1,"")

                        If IsDate(new_info_dim(1)) Then
                            new_info_dim(1) = time_type(new_info_dim(1),8)
                        End If

                        new_info_dim(3)     = "<a href='forum_view.asp?forum_id=" & forumid & "&view_id=" & new_info_dim(2) & "' title='" & code_html(new_info_dim(3),0,0) & "'>" & code_html(new_info_dim(3),0,8) & "</a>"
                    Else
                        ReDim new_info_dim(3)
                    End If

                    If Len(forum_pic) > 1 Then
                        If Left(forum_pic,1) = "$" Then forum_pic = "images/forum/" & Right(forum_pic,Len(forum_pic) - 1)
                        forum_pic = "<td align=right><img src='" & forum_pic & "' border=0></td>"
                    Else
                        forum_pic = ""
                    End If

                    Response.Write vbcrlf & "<tr" & format_table(3,1) & "><td width='10%' rowspan=2 align=center><img src='images/small/label_" & forum_type & ".gif' border=0></td>" & _
                    vbcrlf & "<td width='24%' align=center height=20 " & forum_table2 & "><a href='forum_list.asp?forum_id=" & forumid & "'>�� " & forumname & " ��</a></td>" & _
                    vbcrlf & "<td width='38%'" & forum_table2 & ">" & _
                    vbcrlf & "  <table border=0 width='100%'><tr align=center>" & _
                    vbcrlf & "  <td width='45%'>��̳������&nbsp;&nbsp;<font class=red_3>" & rsforum("forum_topic_num") & "</font></td>" & _
                    vbcrlf & "  <td width='45%'>��̳������&nbsp;&nbsp;<font class=red_3>" & rsforum("forum_data_num") & "</font></td>" & _
                    vbcrlf & "  <td width='10%'></td><td width='16%'><a href='forum_write.asp?forum_id=" & forumid & "'><img src='images/small/mini_write.gif' align=absmiddle title='��������' border=0></a></td></tr></table>" & _
                    vbcrlf & "</td>" & _
                    vbcrlf & "<td width='30%'" & forum_table2 & ">������" & forum_power(rsforum("forum_power"),ptnums) & "</td></tr>" & _
                    vbcrlf & "<tr" & format_table(3,1) & "><td colspan=2 align=center><table border=0 width='99%'><tr><td class=htd>" & code_html(rsforum("forum_remark"),2,0) & "</td>" & forum_pic & "</tr></table></td>" & _
                    vbcrlf & "<td align=left valign=top class=htd>������" & new_info_dim(3) & "<br>���ߣ�" & new_info_dim(0) & "<br>ʱ�䣺" & new_info_dim(1) & "</td></tr>"
                    Erase new_info_dim

                    rsforum.movenext
                Loop

                rsclass.movenext
                Response.Write "</table>" & kong
            Loop

            Set rsclass = Nothing:Set rsforum = Nothing
        End Function

        '------------------------------------forum_list-------------------------------------

        Sub forum_view()

            Select Case Int(istop)
                Case 1
                    folder_type = "istop"
                Case 2
                    folder_type = "istops"
                Case Else

                    If Int(isgood) = 1 Then
                        folder_type = "isgood"
                    Else

                        If Int(islock) = 1 Then
                            folder_type = "islock"
                        ElseIf Int(re_counter) >= 10 Then
                            folder_type = "ishot"
                        End If

                    End If

            End Select

            view_url       = "forum_view.asp?forum_id=" & forumnid & "&view_id=" & id

            If Int(re_counter) > 0 Then
                topic_head = "<img loaded=no src='images/small/fk_plus.gif' border=0 id=followImg" & id & " style=""cursor:hand;"" onclick=""load_tree(" & forumnid & "," & id & ")"" title='չ�������б�'>"
            Else
                topic_head = "<img src='images/small/fk_minus.gif' border=0 id=followImg" & id & ">"
            End If

            Response.Write vbcrlf & "<tr align=center" & format_table(3,1) & ">" & _
            vbcrlf & "<td bgcolor=" & web_var(web_color,5) & "><img src='images/small/" & folder_type & ".gif' border=0></td>" & _
            vbcrlf & "<td bgcolor=" & web_var(web_color,6) & ">"

            If action = "manage" Then
                Response.Write "<input type=checkbox name=del_id value='" & id & "' class=bg_3>"
            Else
                Response.Write "<img src='images/icon/" & icon & ".gif' border=0>"
            End If

            Response.Write "</td>" & _
            vbcrlf & "<td align=left>" & topic_head & "<a href='" & view_url & "' title='���⣺" & code_html(topic,1,0) & "<br>����ʱ�䣺" & tim & "<br>���ظ���" & re_username & "<br>�ظ�ʱ�䣺" & re_tim & "'>" & code_html(topic,0,22) & "</a>&nbsp;" & index_pagecute(view_url,re_counter + 1,web_var(web_num,3),"#cc3300") & "</td>" & _
            vbcrlf & "<td bgcolor=" & web_var(web_color,6) & ">" & format_user_view(username,1,"") & "</td>" & _
            vbcrlf & "<td><a href='" & view_url & "' target=_blank><img src='images/small/new_win.gif' alt='���´����������' border=0 width=13 height=11></a></td>" & _
            vbcrlf & "<td bgcolor=" & web_var(web_color,6) & ">" & re_counter & "<font class=gray>/</font>" & counter & "</td>" & _
            vbcrlf & "<td align=left><font class=timtd>" & time_type(re_tim,6) & "</font><font class=red>��</font>" & format_user_view(re_username,1,"") & "</td>" & _
            vbcrlf & "</tr>" & _
            vbcrlf & "<tr" & format_table(3,1) & " style=""display:none"" id=follow" & id & " height=30><td colspan=2>&nbsp;</td><td colspan=6 id=followTd" & id & " style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:" & web_var(web_color,5) & ";color:" & web_var(web_color,7) & ";padding:2px"" onclick=""load_tree(" & forumnid & "," & id & ")"">���ڶ�ȡ���ڱ�����ĸ��������Ժ��</div></td></tr>"
            del_temp = del_temp + 1
        End Sub

        '------------------------------------forum_view-------------------------------------

        Function view_type()
            Dim up:up = Int(popedom_format(u_popedom,42))
            table_bg = format_table(3,1)
            If ii Mod 2 = 0 Then table_bg = forum_table3

            If var_null(u_whe) <> "" Then
                u_whe = "���ԣ�" & u_whe & "<br>"
            End If

            If var_null(u_nname) <> "" Then
                u_nname = "ͷ�Σ�" & u_nname & "<br>"
            End If

            view_type = vbcrlf & "<tr align=center valign=top" & table_bg & "><td width='20%' bgcolor='" & web_var(web_color,6) & "'>" & _
            vbcrlf & "<table border=0 width='94%'><tr><td align=center height=30><table border=0><tr><td><font class=blue><b>" & u_username & "</b></font></td><td>&nbsp;" & user_view_sex(u_sex) & "</td></tr></table></td></tr>" & _
            vbcrlf & "<tr><td align=center height=96><img src='images/face/" & rs("u_face") & ".gif' border=0></td></tr>" & _
            vbcrlf & "<tr><td height=15><img src='images/star/star_" & user_star(u_integral,u_power,1) & ".gif' border=0></td></tr>" & _
            vbcrlf & "<tr><td>�ȼ���" & user_view_power(u_power,0) & user_star(u_integral,u_power,2) & "<br>" & u_nname & "������" & u_bbs_counter & "<br>���֣�" & u_integral & "<br>" & u_whe & "ע�᣺" & FormatDateTime(rs("u_tim"),2) & "</td></tr>" & _
            vbcrlf & "</table></td><td width='80%' height='100%'>" & _
            vbcrlf & "<table border=0 width='99%' cellspacing=2 cellpadding=0 height='100%'><tr height=25><td width='85%'>" & _
            vbcrlf & "<a target=_blank href='user_view.asp?username=" & Server.urlencode(u_username) & "'><img src='images/small/forum_profile.gif' title='�鿴 " & u_username & " ����ϸ��Ϣ' border=0></a>&nbsp;" & _
            vbcrlf & "<a target=_blank href='user_friend.asp?action=add&add_username=" & Server.urlencode(u_username) & "'><img src='images/small/forum_friend.gif' title='�� " & u_username & " ��Ϊ�ҵĺ���' border=0></a>&nbsp;" & _
            vbcrlf & "<a target=_blank href='user_message.asp?action=write&accept_uaername=" & Server.urlencode(u_username) & "'><img src='images/small/forum_message.gif' title='�� " & u_username & " ������' border=0></a>&nbsp;" & _
            vbcrlf & "<a href='forum_edit.asp?forum_id=" & forumid & "&edit_id=" & qid & "'><img src='images/small/forum_edit.gif' title='�༭�������' border=0></a>&nbsp;"

            If Int(fir_islock) <> 1 Then
                view_type = view_type & vbcrlf & "<a href='forum_reply.asp?forum_id=" & forumid & "&quote=yes&view_id=" & qid & "'><img src='images/small/forum_quote.gif' title='���ò��ظ��������' border=0></a>&nbsp;" & _
                vbcrlf & "<a href='forum_reply.asp?forum_id=" & forumid & "&view_id=" & qid & "'><img src='images/small/forum_reply.gif' title='�ظ��������' border=0></a>"
            Else
                view_type = view_type & vbcrlf & "<img src='images/small/forum_reply.gif' title='��������ѱ�����' border=0>"
            End If

            view_type     = view_type & vbcrlf & "</td><td width='15%' align=center>�� <font class=red_3><b>" & ii + (viewpage - 1)*nummer & "</b></font> ¥</td></tr>" & _
            vbcrlf & "<tr><td colspan=2 height=1 bgcolor=" & web_var(web_color,3) & "></td></tr>"

            If up = 0 Then
                view_type = view_type & vbcrlf & "<tr><td colspan=2 valign=top align=center>" & _
                vbcrlf & "<table border=0 width='98%' class=tf><tr><td height=30>" & _
                vbcrlf & "<img src='images/icon/" & rs("icon") & ".gif' align=absMiddle border=0>&nbsp;<font class=red_3><b>" & code_html(rs("topic"),1,0) & "</b></font></td></tr>" & _
                vbcrlf & "<tr><td class=bw><font class=htd>" & code_jk(rs("word")) & "</font></td></tr>" & _
                vbcrlf & "</table></td></tr>" & _
                vbcrlf & "<tr><td colspan=2 height=20 align=right>" & img_small("signature") & "</td></tr>" & _
                vbcrlf & "<tr><td colspan=2 height=30 align=center valign=top><table border=0 width='96%' class=tf><tr><td class=bw><font class=htd>" & u_remark & "</font></a></td></tr></table></td></tr>"
            Else
                view_type = view_type & vbcrlf & "<tr><td colspan=2 valign=top><br><table border=0><tr><td class=htd><font class=red_2>========================<br>&nbsp;&nbsp;���û�����̳��������ʱ���������Σ�<br>========================</font></td></tr></table></td></tr>"
            End If

            view_type     = view_type & vbcrlf & "<tr><td height=25 colspan=2><table border=0 width='100%'><tr><td>" & img_small("forum_tim") & "<font class=gray>��������ʱ�䣺" & rs("tim") & "</font></td><td align=right>" & ip_types(rs("ip"),u_username,1) & "��<img src='images/small/sys.gif' align=absMiddle title='" & view_sys(rs("sys")) & "' border=0>��<a href=""javascript:" & del_type & "('" & forumid & "','" & iid & "');""><img src='images/small/forum_del.gif' align=absMiddle border=0></a></td></tr></table></td></tr>" & _
            vbcrlf & "</table>"
        End Function

        Function user_view_sex(us)

            If us = False Then
                user_view_sex = "<img src='images/small/forum_girl.gif' align=absmiddle title='�ഺŮ��' border=0>":Exit Function
            Else
                user_view_sex = "<img src='images/small/forum_boy.gif' align=absmiddle title='�����к�' border=0>":Exit Function
            End If

        End Function

        Function user_view_power(uvp,ut)
            user_view_power = img_small("icon_" & uvp)
            If ut = 1 Then user_view_power = user_view_power & "<font class=red_3>" & format_power(uvp,1) & "</font>"
        End Function %>