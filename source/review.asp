<!-- #include file="include/config.asp" -->
<!-- #include file="include/skin.asp" -->
<!-- #include file="include/config_review.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim rsort
Dim rurl
Dim re_id
Dim rerr:rerr = "":sql = ""
tit = "��������"
Call web_head(0,2,0,0,0)

Select Case action
    Case "delete"
        Call review_delete()
    Case "del"
        Call review_del()
    Case Else
        Call review_main()
End Select

Call close_conn()

Sub review_delete()
    Call review_d()
    On Error Resume Next
    conn.execute(sql)

    If Err Then
        Err.Clear
        Call review_err("����Ĵ�������վ����ϵ��\nhttp://beyondest.com/\n")

        Exit Sub
        End If

        Response.Write vbcrlf & "<script lanuage=javascript><!--" & _
        vbcrlf & "alert(""�ѳɹ�ɾ�������⣨n_sort��" & rsort & "�� id��" & re_id & "�����������ۣ�\n\n�������..."");"

        If Len(rurl) < 5 Then
            Response.Write vbcrlf & "location.href='main.asp';"
        Else
            Response.Write vbcrlf & "location.href='" & rurl & "';"
        End If

        Response.Write vbcrlf & "--></script>"
    End Sub

    Sub review_del()
        Call review_d()
        Dim rid:rid = Trim(Request.querystring("rid"))

        If Not(IsNumeric(rid)) Then
            rerr = rerr & "ɾ�����۵� RID ����\n"
        End If

        If rerr <> "" Then Call review_err(rerr):Exit Sub
            sql = sql & " and rid=" & rid
            On Error Resume Next
            conn.execute(sql)

            If Err Then
                Err.Clear
                Call review_err("����Ĵ�������վ����ϵ��\nhttp://beyondest.com/\n")

                Exit Sub
                End If

                Response.Write vbcrlf & "<script lanuage=javascript><!--" & _
                vbcrlf & "alert(""�ѳɹ�ɾ����һ�����⣨n_sort��" & rsort & "�� id��" & re_id & "�����ۣ�rid��" & rid & "����\n\n�������..."");"

                If Len(rurl) < 5 Then
                    Response.Write vbcrlf & "location.href='main.asp';"
                Else
                    Response.Write vbcrlf & "location.href='" & rurl & "';"
                End If

                Response.Write vbcrlf & "--></script>"
            End Sub

            Sub review_d()

                If login_mode <> format_power2(1,1) Then
                    Call close_conn()
                    Call review_err("��û��ɾ�����۵�Ȩ�ޣ�����\n")
                    Response.End
                End If

                rsort    = Trim(Request.querystring("rsort"))
                re_id    = Trim(Request.querystring("re_id"))
                rurl     = Trim(Request.querystring("rurl"))

                If review_rsort(rsort) <> "yes" Then
                    rerr = rerr & "ɾ�����۵����ͳ���\n"
                End If

                If Not(IsNumeric(re_id)) Then
                    rerr = rerr & "ɾ�����۵� ID ����\n"
                End If

                If rerr <> "" Then
                    Call close_conn()
                    Call review_err(rerr)
                    Response.End
                End If

                sql = "delete from review where rsort='" & rsort & "' and re_id=" & re_id
            End Sub

            Sub review_main()
                Dim rusername
                Dim remail
                Dim rword
                rusername = code_form(Trim(Request.form("rusername")))
                remail    = code_form(Trim(Request.form("remail")))
                rword     = code_form(Trim(Request.form("rword")))
                rsort     = Trim(Request.form("rsort"))
                re_id     = Trim(Request.form("re_id"))
                rurl      = Trim(Request.form("rurl"))

                If review_rsort(rsort) <> "yes" Then
                    rerr  = rerr & "�������۵����ͳ�������\n"
                End If

                If Not(IsNumeric(re_id)) Then
                    rerr = rerr & "�������۵� ID ��������\n"
                End If

                If symbol_name(rusername) <> "yes" Then
                    rerr = rerr & "�������������ƣ������ú��зǷ��ַ���\n"
                End If

                If Len(remail) > 0 Then

                    If email_ok(remail) <> "yes" Or Len(remail) > 50 Then
                        rerr = rerr & "���� E-mail ���ú��зǷ��ַ���\n"
                    End If

                End If

                If Len(rword) < 1 Then
                    rerr = rerr & "��û��������������ݣ�\n"
                ElseIf Len(rword) > 250 Then
                    rerr = rerr & "���������������̫����(<=250�ֽ�)\n"
                End If

                If rerr <> "" Then Call review_err(rerr):Exit Sub

                    On Error Resume Next
                    sql = "insert into review(rsort,re_id,rusername,remail,rword,rtim,rtype) values('" & rsort & "'," & re_id & ",'" & rusername & "','" & remail & "','" & rword & "','" & now_time & "',"

                    If rusername = login_username Then
                        sql = sql & "1"
                    Else
                        sql = sql & "0"
                    End If

                    sql = sql & ")"
                    conn.execute(sql)

                    If Err Then
                        Err.Clear
                        Call review_err("����Ĵ�������վ����ϵ��\nhttp://beyondest.com/\n")

                        Exit Sub
                        End If

                        Response.Write vbcrlf & "<script lanuage=javascript><!--" & _
                        vbcrlf & "alert(""���ɹ��ķ������й��������ۣ�\n\nлл���Ĳ��룡�������..."");"

                        If Len(rurl) < 5 Then
                            Response.Write vbcrlf & "location.href='main.asp';"
                        Else
                            Response.Write vbcrlf & "location.href='" & rurl & "';"
                        End If

                        Response.Write vbcrlf & "--></script>"
                    End Sub

                    Sub review_err(revar)
                        Response.Write vbcrlf & "<script lanuage=javascript><!--" & _
                        vbcrlf & "alert(""���ڷ�������ʱ�������´���\n\n" & revar & "\n�������..."");" & _
                        vbcrlf & "history.back(-1);" & _
                        vbcrlf & "--></script>"
                    End Sub %>