<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
Sub configs_load()
    conn.execute("insert into configs(id,num_topic,num_data,num_reg,new_username,num_news,num_article,num_down,num_product,counter,max_online,max_tim,start_tim) values(1,0,0,0,'',0,0,0,0,1,1,'" & now_time & "','" & time_type(now_time,33) & "')")
End Sub

Function counter_type(cmain,ct)
    Dim rs
    Dim sql
    Dim ft
    Dim counters
    Dim max_online
    Dim max_tim
    Dim start_tim
    Dim types
    Dim counts
    types  = 1:i = 0:counter_type = ""
    sql    = "select counter,max_online,max_tim,start_tim from configs where id=1"
    Set rs = conn.execute(sql)

    If rs.eof And rs.bof Then
        rs.Close
        Call configs_load()
        Set rs = conn.execute(sql)
    End If

    counters   = rs("counter")
    max_online = rs("max_online")
    max_tim    = rs("max_tim")
    start_tim  = rs("start_tim")
    rs.Close:Set rs = Nothing
    counters   = Int(counters)
    max_online = Int(max_online)
    ft         = Mid(web_setup,7,1)
    If Not(IsNumeric(ft)) Then ft = 1
    ft         = Int(ft)

    If ft = 0 Then

        If Trim(Request.cookies(web_cookies)("counters")) <> "yes" Then
            counters                                  = counters + types
            Response.cookies(web_cookies)("counters") = "yes"
        End If

    Else
        Response.cookies(web_cookies)("counters") = ""
        counters                                  = counters + types
    End If

    If online_num > max_online Then max_online = online_num:max_tim = now_time
    sql = "update configs set counter=" & counters & ",max_online=" & max_online & ",max_tim='" & max_tim & "' where id=1"
    conn.execute(sql)

    If cmain = "view" Then
        counts = "本站总访问量:&nbsp;<font class=red_3 title=从&nbsp;" & start_tim & "&nbsp;至今>" & counters & "</font>&nbsp;人次" & _
        "&nbsp;┋&nbsp;最高峰&nbsp;<font class=red_3 title=最高峰发生在：" & max_tim & ">" & max_online & "</font>&nbsp;人在线" & _
        "&nbsp;┋&nbsp;当前有&nbsp;<font class=red_3>" & online_num & "</font>&nbsp;人在线"
        counter_type = counts
    End If

End Function %>