<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
Function symbol_name(sn_var)
    Dim safety_char:symbol_name = "yes":sn_var = Trim(sn_var)
    If sn_var = "" Or IsNull(sn_var) Or Len(sn_var) > 20 Or InStr(sn_var,"|") > 0 Or InStr(sn_var,":") > 0 Or InStr(sn_var,Chr(9)) > 0 Or InStr(sn_var,Chr(32)) > 0 Or InStr(sn_var,"'") > 0 Or InStr(sn_var,"""") > 0 Then symbol_name = "no":Exit Function
    safety_char = web_var(web_safety,1)

    For i = 1 To Len(sn_var)
        If InStr(1,safety_char,Mid(sn_var,i,1)) > 0 Then symbol_name = "no":Exit Function
    Next

End Function

Function symbol_ok(symbol_var)
    Dim safety_char:symbol_ok = "yes":symbol_var = Trim(symbol_var)
    If symbol_var = "" Or IsNull(symbol_var) Or Len(symbol_var) > 20 Then symbol_ok = "no":Exit Function
    safety_char = web_var(web_safety,2)

    For i = 1 To Len(symbol_var)
        If InStr(1,safety_char,Mid(symbol_var,i,1)) = 0 Then symbol_ok = "no":Exit Function
    Next

End Function

Sub time_load(tt,t1,t2)

    If tt = 0 Then
        Response.cookies(web_cookies)("time_load") = now_time

        Call cookies_yes():Exit Sub
        End If

        Dim tims
        Dim vars
        Dim tts
        tims    = Int(web_var(web_num,16))
        vars    = time_type(Request.cookies(web_cookies)("time_load"),9)

        If vars = "" Or Not(IsDate(vars)) Then Exit Sub
            tts = Int(DateDiff("s",vars,now_time))

            If tts <= tims Then
                If t1 = 1 Then Set rs = Nothing
                If t2 = 1 Then Call close_conn()
                Call cookies_type("time_load")
                Response.End
            End If

        End Sub

        Function health_name(hnn)
            health_name = "yes":Dim ti
            Dim tnum
            Dim tdim
            Dim hn
            hn = hnn:tdim = Split(web_var(web_safety,3),":"):tnum = UBound(tdim)

            For ti = 0 To tnum
                If InStr(hn,tdim(ti)) > 0 Then health_name = "no":Erase tdim:Exit Function
            Next

            Erase tdim
            tdim = Split(web_var(web_safety,4),":"):tnum = UBound(tdim)

            For ti = 0 To tnum
                If InStr(hn,tdim(ti)) > 0 Then health_name = "no":Erase tdim:Exit Function
            Next

            Erase tdim
        End Function

        Function health_var(hnn,vt)
            Dim ti
            Dim tj
            Dim tdim
            Dim ht
            Dim hn:hn = hnn

            If Int(Mid(web_setup,4,1)) = 1 And vt = 1 Then
                tdim       = Split(web_var(web_safety,4),":")

                For ti = 0 To UBound(tdim)
                    ht     = ""

                    For tj = 1 To Len(tdim(ti))
                        ht = ht & "*"
                    Next

                    hn     = Replace(hn,tdim(ti),ht)
                Next

                Erase tdim
            End If

            health_var = hn
        End Function

        Function found_error(error_type,error_len)
            If error_len > 600 Or error_len < 200 Then error_len = 300
            found_error = VbCrLf & "<table border=0 width=" & error_len & " clas=fr><tr><td align=center height=50><font class=red>系统发现你输入的数据有以下错误：</font></td></tr><tr><td class=htd>" & error_type & "</td></tr><tr><td align=center height=50>" & go_back & "</td></tr></table>"
        End Function

        Function url_true(puu,pus)
            Dim puuu
            Dim pu:puuu = puu:pu = pus

            If InStr(1,pu,"://") <> 0 Then
                url_true = pu
            Else
                If Right(puuu,1) <> "/" Then puuu = puuu & "/"
                url_true = puuu & pu
            End If

        End Function

        Function time_type(tvar,tt)
            Dim ttt:ttt = tvar
            If Not(IsDate(ttt)) Then time_type = "":Exit Function

            Select Case tt
                Case 1	'10-10
                    time_type = Month(ttt) & "-" & Day(ttt)
                Case 11	'月-日
                    time_type = Month(ttt) & "月" & Day(ttt) & "日"
                Case 2	'年(2)-月-日 00-10-10
                    time_type = Right(Year(ttt),2) & "-" & Month(ttt) & "-" & Day(ttt)
                Case 3	'2000-10-10
                    time_type = Year(ttt) & "-" & Month(ttt) & "-" & Day(ttt)
                Case 33	'年(4)-月-日
                    time_type = Year(ttt) & "年" & Month(ttt) & "月" & Day(ttt) & "日"
                Case 4	'23:45
                    time_type = Hour(ttt) & ":" & Minute(ttt)
                Case 44	'时:分
                    time_type = Hour(ttt) & "时" & Minute(ttt) & "分"
                Case 5	'23:45:36
                    time_type = Hour(ttt) & ":" & Minute(ttt) & ":" & Second(ttt)
                Case 55	'时:分:秒
                    time_type = Hour(ttt) & "时" & Minute(ttt) & "分" & Second(ttt) & "秒"
                Case 6	'10-10 23:45
                    time_type = Month(ttt) & "-" & Day(ttt) & " " & Hour(ttt) & ":" & Minute(ttt)
                Case 66	'月-日 时:分
                    time_type = Month(ttt) & "月" & Day(ttt) & "日 " & Hour(ttt) & "时" & Minute(ttt) & "分"
                Case 7	'年(2)-月-日 时:分  00-10-10 23:45
                    time_type = Right(Year(ttt),2) & "-" & Month(ttt) & "-" & Day(ttt) & " " & Hour(ttt) & ":" & Minute(ttt)
                Case 8	'2000-10-10 23:45
                    time_type = Year(ttt) & "-" & Month(ttt) & "-" & Day(ttt) & " " & Hour(ttt) & ":" & Minute(ttt)
                Case 88	'年(4)-月-日 时:分
                    time_type = Year(ttt) & "年" & Month(ttt) & "月" & Day(ttt) & "日 " & Hour(ttt) & "时" & Minute(ttt) & "分"
                Case 9	'2000-10-10 23:45:45
                    time_type = Year(ttt) & "-" & Month(ttt) & "-" & Day(ttt) & " " & Hour(ttt) & ":" & Minute(ttt) & ":" & Second(ttt)
                    time_type = FormatDateTime(time_type)
                Case 99	'年(4)-月-日 时:分:秒
                    time_type = Year(ttt) & "年" & Month(ttt) & "月" & Day(ttt) & "日 " & Hour(ttt) & "时" & Minute(ttt) & "分" & Second(ttt) & "秒"
                Case Else
                    time_type = ttt
            End Select

        End Function

        Function upload_time(tt)
            Dim ttt:ttt = tt
            ttt         = Replace(ttt,":",""):ttt = Replace(ttt,"-","")
            ttt         = Replace(ttt," ",""):ttt = Replace(ttt,"/","")
            ttt         = Replace(ttt,"PM",""):ttt = Replace(ttt,"AM","")
            ttt         = Replace(ttt,"上午",""):ttt = Replace(ttt,"下午","")
            upload_time = ttt
        End Function

        Function code_form(strers)
            Dim strer:strer = Trim(strers)
            If IsNull(strer) Or strer = "" Then code_form = "":Exit Function
            strer     = Replace(strer,"'","""")
            code_form = strer
        End Function

        Function code_word(strers)
            Dim strer:strer = Trim(strers)
            If IsNull(strer) Or strer = "" Then code_word = "":Exit Function
            strer     = Replace(strer,"'","&#39;")
            code_word = strer
        End Function

        Function code_html(strers,chtype,cutenum)
            Dim strer:strer = strers
            If IsNull(strer) Or strer = "" Then code_html = "":Exit Function
            strer = health_var(strer,1)
            If cutenum > 0 Then strer = cuted(strer,cutenum)
            strer = Replace(strer,"<","&lt;")
            strer = Replace(strer,">","&gt;")
            strer = Replace(strer,Chr(39),"&#39;")
            strer = Replace(strer,Chr(34),"&quot;")
            strer = Replace(strer,Chr(32),"&nbsp;")

            Select Case chtype
                Case 1
                    strer = Replace(strer,Chr(9),"&nbsp;")
                    strer = Replace(strer,Chr(10),"")
                    strer = Replace(strer,Chr(13),"")
                Case 2
                    strer = Replace(strer,Chr(9),"&nbsp;　&nbsp;")
                    strer = Replace(strer,Chr(10),"<br>")
                    strer = Replace(strer,Chr(13),"<br>")
            End Select

            code_html     = strer
        End Function

        Function cuted(types,num)
            Dim ctypes
            Dim cnum
            Dim ci
            Dim tt
            Dim tc
            Dim cc
            Dim cmod
            cmod   = 3
            ctypes = types:cnum = Int(num):cuted = "":tc = 0:cc = 0

            For ci = 1 To Len(ctypes)
                If cnum < 0 Then cuted = cuted & "...":Exit For
                tt        = Mid(ctypes,ci,1)

                If Int(Asc(tt)) >= 0 Then
                    cuted = cuted & tt:tc = tc + 1:cc = cc + 1
                    If tc = 2 Then tc = 0:cnum = cnum - 1
                    If cc > cmod Then cnum = cnum - 1:cc = 0
                Else
                    cnum  = cnum - 1
                    If cnum <= 0 Then cuted = cuted & "...":Exit For
                    cuted = cuted & tt
                End If

            Next

        End Function

        Function post_chk()
            Dim server_v1
            Dim server_v2
            post_chk  = "no"
            server_v1 = Request.ServerVariables("HTTP_REFERER")
            server_v2 = Request.ServerVariables("SERVER_NAME")
            If Mid(server_v1,8,Len(server_v2)) = server_v2 Then post_chk = "yes":Exit Function
        End Function

        Function email_ok(email)
            Dim names
            Dim name
            Dim i
            Dim c
            email_ok = "yes":names = Split(email, "@")
            If UBound(names) <> 1 Then email_ok = "no":Exit Function

            For Each name in names
                If Len(name) <= 0 Then email_ok = "no":Exit Function

                For i = 1 To Len(name)
                    c = LCase(Mid(name, i, 1))
                    If InStr("abcdefghijklmnopqrstuvwxyz-_.", c) <= 0 And Not IsNumeric(c) Then email_ok = "no":Exit Function
                Next

                If Left(name, 1) = "." Or Right(name, 1) = "." Then email_ok = "no":Exit Function
            Next

            If InStr(names(1), ".") <= 0 Then email_ok = "no":Exit Function
            i = Len(names(1)) - InStrRev(names(1), ".")
            If i <> 2 And i <> 3 Then email_ok = "no":Exit Function
            If InStr(email, "..") > 0 Then email_ok = "no"
        End Function

        Function ip_sys(isu,iun)	'0,*=ip_sys  1,0=ip  1,1=ip:port  2,*=port  3,*=sys
            Dim userip
            Dim userip2

            Select Case isu
                Case 2
                    ip_sys  = Request.ServerVariables("REMOTE_PORT")
                Case 3
                    ip_sys  = Request.Servervariables("HTTP_USER_AGENT")
                Case Else
                    userip  = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
                    userip2 = Request.ServerVariables("REMOTE_ADDR")
                    If InStr(userip,",") > 0 Then userip = Left(userip,InStr(userip,",") - 1)
                    If InStr(userip2,",") > 0 Then userip2 = Left(userip2,InStr(userip2,",") - 1)

                    If userip = "" Then
                        ip_sys = userip2
                    Else
                        ip_sys = userip
                    End If

                    If isu = 0 Then ip_sys = "您的真实ＩＰ是：" & ip_sys & ":" & Request.ServerVariables("REMOTE_PORT") & "，" & view_sys(Request.Servervariables("HTTP_USER_AGENT")):Exit Function
                    If iun = 1 Then ip_sys = ip_sys & ":" & Request.ServerVariables("REMOTE_PORT")
            End Select

        End Function

        Function view_sys(vss)
            Dim vs
            Dim ver
            Dim strUserAgentArr
            Dim strTempUserInfo1
            Dim strTempUserInfo2
            Dim Mozilla
            Dim Agent
            Dim BcType
            Dim Browser
            Dim sSystem
            Dim strSystem
            Dim strBrowser
            On Error Resume Next
            vs               = Trim(vss):strUserAgentArr = Split(vs, "; "):strTempUserInfo1 = strUserAgentArr(1)
            If InStr(strTempUserInfo1, "MSIE") > 0 Then strTempUserInfo1 = Replace(strTempUserInfo1, "MSIE", "Internet Explorer")
            strTempUserInfo2 = Trim(Left(strUserAgentArr(2), Len(strUserAgentArr(2)) - 1))
            If InStr(vs, "Mozilla/4.0 (compatible;") > 0 And strTempUserInfo2 = "Windows NT 5.0" Then strTempUserInfo2 = "Windows 2000"
            Mozilla          = vs:Agent = Split(Mozilla,"; "):BcType = 0
            If InStr(Agent(1),"U") Or InStr(Agent(1),"I") Then BcType = 1
            If InStr(Agent(1),"MSIE") Then BcType = 2

            Select Case BcType
                Case 0
                    Browser = "其它":sSystem = "其它"
                Case 1
                    Ver     = Mid(Agent(0),InStr(Agent(0), "/") + 1)
                    Ver     = Mid(Ver,1,InStr(Ver, " ") - 1)
                    Browser = "Netscape" & Ver
                    sSystem = Mid(Agent(0), InStr(Agent(0), "(") + 1)
                    sSystem = Replace(sSystem, "Windows", "Win")
                Case 2
                    Browser = Agent(1):sSystem = Replace(Agent(2), ")", ""):sSystem = Replace(sSystem, "Windows", "Win")
            End Select

            strSystem       = Replace(sSystem, "Win", "Windows")
            If InStr(strSystem,"98") > 0 And InStr(Mozilla,"Win 9x") > 0 Then strSystem = Replace(strSystem, "98", "Me")
            strSystem       = Replace(strSystem, "NT 5.0", "2000")
            strSystem       = Replace(strSystem, "NT5.0", "2000")
            strSystem       = Replace(strSystem, "NT 5.1", "XP")
            strSystem       = Replace(strSystem, "NT5.1", "XP")
            strSystem       = Replace(strSystem, "NT 5.2", "2003")
            strSystem       = Replace(strSystem, "NT5.2", "2003")
            strBrowser      = Replace(Browser, "MSIE", "Internet Explorer")
            Set Browser     = Nothing:Set sSystem = Nothing
            view_sys        = "操作系统：" & Trim(strSystem) & " ，浏览器：" & Trim(strBrowser)
            If Err Then Err.Clear:view_sys = "未知的操作系统和浏览器"
        End Function

        Function ip_true(tips)
            Dim tip
            Dim iptemp
            Dim iptemp1
            Dim iptemp2:tip = tips:ip_true = "no":tip = Trim(tip)
            iptemp  = tip:iptemp = Replace(Replace(iptemp,".",""),":","")
            iptemp1 = tip:iptemp1 = Len(tip) - Len(Replace(iptemp1,".",""))
            iptemp2 = tip:iptemp2 = Len(tip) - Len(Replace(iptemp2,":",""))
            If IsNumeric(iptemp) And CInt(iptemp1) = 3 And (CInt(iptemp2) = 1 Or CInt(iptemp2) = 0) Then ip_true = "yes"
        End Function

        Function ip_ip(tips)
            Dim ipn
            Dim tip:tip = tips:tip = Trim(tip)
            If ip_true(tip) = "no" Then ip_ip = "no":Exit Function
            ipn   = InStr(tip,":")
            If ipn > 0 Then ip_ip = Left(tip,ipn - 1):Exit Function
            ip_ip = tip
        End Function

        Function ip_types(tips,tu,tt)
            Dim ipn
            Dim tip2
            Dim wip
            Dim ip_type:tip2 = tips:tip2 = Trim(tip2):ip_types = "error"
            If ip_true(tip2) = "no" Then ip_types = "no":Exit Function
            wip = Int(Mid(web_setup,5,1))
            If login_mode = format_power2(1,1) Then wip = 2

            Select Case wip
                Case 0
                    ip_types = "*.*.*.*":ip_types = tu & " 的IP是：" & ip_types
                    If tt <> 0 Then ip_types = "<img src='images/small/ip.gif' align=absMiddle title='" & ip_types & "' border=0>"
                Case 1
                    ipn          = InStr(tip2,":")

                    If ipn > 0 Then
                        ip_types = Left(tip2,ipn - 1)
                    Else
                        ip_types = tip2
                    End If

                    ip_type      = Split(ip_types,"."):ip_types = ip_type(0) & "." & ip_type(1) & ".*.*"
                    Erase ip_type:ip_types = tu & " 的IP是：" & ip_types
                    If tt <> 0 Then ip_types = "<img src='images/small/ip.gif' align=absMiddle title='" & ip_types & "' border=0>"
                Case Else
                    ip_types = tu & " 的IP是：" & tip2
                    If tt <> 0 Then ip_types = "<a href='ip_address.asp?ip=" & tip2 & "'><img src='images/small/ip.gif' align=absMiddle title='" & ip_types & "' border=0></a>"
            End Select

        End Function

        Function ip_port(pips)
            Dim ipnn
            Dim iptemp
            Dim pip:pip = pips
            pip     = Trim(pip)
            If ip_true(pip) = "no" Then ip_port = "no":Exit Function
            ipnn    = InStr(pip,":")
            If ipnn > 0 Then ip_port = Right(pip,Len(pip) - ipnn):Exit Function
            ip_port = "yes"
        End Function

        Function ip_address(sips)
            Dim str1
            Dim str2
            Dim str3
            Dim str4
            Dim num
            Dim country
            Dim city
            Dim irs
            Dim sip:sip = sips
            If Not(IsNumeric(Left(sip,2))) Then ip_address = "未知":Exit Function
            If sip = "127.0.0.1" Then sip = "192.168.0.1"
            str1        = Left(sip,InStr(sip,".") - 1):sip = Mid(sip,InStr(sip,".") + 1)
            str2        = Left(sip,InStr(sip,".") - 1):sip = Mid(sip,InStr(sip,".") + 1)
            str3        = Left(sip,InStr(sip,".") - 1):str4 = Mid(sip,InStr(sip,".") + 1)

            If Not(IsNumeric(str1) = 0 Or IsNumeric(str2) = 0 Or IsNumeric(str3) = 0 Or IsNumeric(str4) = 0) Then
                num     = CInt(str1)*256*256*256 + CInt(str2)*256*256 + CInt(str3)*256 + CInt(str4) - 1
                sql     = "select Top 1 country,city from ip_address where ip1 <=" & num & " and ip2 >=" & num & ""
                Set irs = Server.CreateObject("adodb.recordset")
                irs.open sql,conn,1,1

                If irs.eof And irs.bof Then
                    country = "亚洲":city = ""
                Else
                    country = irs(0):city = irs(1)
                End If

                irs.Close:Set irs = Nothing
            End If

            ip_address = country & city
        End Function %>