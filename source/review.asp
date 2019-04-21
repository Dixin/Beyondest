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
tit = "发表评论"
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
        Call review_err("意外的错误！请与站长联系。\nhttp://beyondest.com/\n")

        Exit Sub
        End If

        Response.Write vbcrlf & "<script lanuage=javascript><!--" & _
        vbcrlf & "alert(""已成功删除了主题（n_sort：" & rsort & "， id：" & re_id & "）的所有评论！\n\n点击返回..."");"

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
            rerr = rerr & "删除评论的 RID 出错！\n"
        End If

        If rerr <> "" Then Call review_err(rerr):Exit Sub
            sql = sql & " and rid=" & rid
            On Error Resume Next
            conn.execute(sql)

            If Err Then
                Err.Clear
                Call review_err("意外的错误！请与站长联系。\nhttp://beyondest.com/\n")

                Exit Sub
                End If

                Response.Write vbcrlf & "<script lanuage=javascript><!--" & _
                vbcrlf & "alert(""已成功删除了一条主题（n_sort：" & rsort & "， id：" & re_id & "）评论（rid：" & rid & "）！\n\n点击返回..."");"

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
                    Call review_err("您没有删除评论的权限！！！\n")
                    Response.End
                End If

                rsort    = Trim(Request.querystring("rsort"))
                re_id    = Trim(Request.querystring("re_id"))
                rurl     = Trim(Request.querystring("rurl"))

                If review_rsort(rsort) <> "yes" Then
                    rerr = rerr & "删除评论的类型出错！\n"
                End If

                If Not(IsNumeric(re_id)) Then
                    rerr = rerr & "删除评论的 ID 出错！\n"
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
                    rerr  = rerr & "发表评论的类型出错！！！\n"
                End If

                If Not(IsNumeric(re_id)) Then
                    rerr = rerr & "发表评论的 ID 出错！！！\n"
                End If

                If symbol_name(rusername) <> "yes" Then
                    rerr = rerr & "请输入您的名称！（不得含有非法字符）\n"
                End If

                If Len(remail) > 0 Then

                    If email_ok(remail) <> "yes" Or Len(remail) > 50 Then
                        rerr = rerr & "您的 E-mail 不得含有非法字符！\n"
                    End If

                End If

                If Len(rword) < 1 Then
                    rerr = rerr & "您没有输入的评论内容！\n"
                ElseIf Len(rword) > 250 Then
                    rerr = rerr & "您输入的评论内容太长！(<=250字节)\n"
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
                        Call review_err("意外的错误！请与站长联系。\nhttp://beyondest.com/\n")

                        Exit Sub
                        End If

                        Response.Write vbcrlf & "<script lanuage=javascript><!--" & _
                        vbcrlf & "alert(""您成功的发表了有关您的评论！\n\n谢谢您的参与！点击返回..."");"

                        If Len(rurl) < 5 Then
                            Response.Write vbcrlf & "location.href='main.asp';"
                        Else
                            Response.Write vbcrlf & "location.href='" & rurl & "';"
                        End If

                        Response.Write vbcrlf & "--></script>"
                    End Sub

                    Sub review_err(revar)
                        Response.Write vbcrlf & "<script lanuage=javascript><!--" & _
                        vbcrlf & "alert(""您在发表评论时出现如下错误：\n\n" & revar & "\n点击返回..."");" & _
                        vbcrlf & "history.back(-1);" & _
                        vbcrlf & "--></script>"
                    End Sub %>