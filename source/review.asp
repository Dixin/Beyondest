<!-- #include file="include/config.asp" -->
<!-- #include file="include/skin.asp" -->
<!-- #include file="INCLUDE/config_review.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim rsort,rurl,re_id,rerr:rerr="":sql=""
tit="��������"
call web_head(0,2,0,0,0)

select case action
case "delete"
  call review_delete()
case "del"
  call review_del()
case else
  call review_main()
end select

call close_conn()

sub review_delete()
  call review_d()
  on error resume next
  conn.execute(sql)
  if err then
    err.clear
    call review_err("����Ĵ�������վ����ϵ��\nhttp://beyondest.com/\n")
    exit sub
  end if
  response.write vbcrlf&"<script lanuage=javascript><!--" & _
		 vbcrlf&"alert(""�ѳɹ�ɾ�������⣨n_sort��"&rsort&"�� id��"&re_id&"�����������ۣ�\n\n�������..."");"
  if len(rurl)<5 then
    response.write vbcrlf&"location.href='main.asp';"
  else
    response.write vbcrlf&"location.href='"&rurl&"';"
  end if
  response.write vbcrlf&"--></script>"
end sub

sub review_del()
  call review_d()
  dim rid:rid=trim(request.querystring("rid"))
  if not(isnumeric(rid)) then
    rerr=rerr&"ɾ�����۵� RID ����\n"
  end if
  if rerr<>"" then call review_err(rerr):exit sub
  sql=sql&" and rid="&rid
  on error resume next
  conn.execute(sql)
  if err then
    err.clear
    call review_err("����Ĵ�������վ����ϵ��\nhttp://beyondest.com/\n")
    exit sub
  end if
  response.write vbcrlf&"<script lanuage=javascript><!--" & _
		 vbcrlf&"alert(""�ѳɹ�ɾ����һ�����⣨n_sort��"&rsort&"�� id��"&re_id&"�����ۣ�rid��"&rid&"����\n\n�������..."");"
  if len(rurl)<5 then
    response.write vbcrlf&"location.href='main.asp';"
  else
    response.write vbcrlf&"location.href='"&rurl&"';"
  end if
  response.write vbcrlf&"--></script>"
end sub

sub review_d()
  if login_mode<>format_power2(1,1) then
    call close_conn()
    call review_err("��û��ɾ�����۵�Ȩ�ޣ�����\n")
    response.end
  end if
  rsort=trim(request.querystring("rsort"))
  re_id=trim(request.querystring("re_id"))
  rurl=trim(request.querystring("rurl"))
  if review_rsort(rsort)<>"yes" then
    rerr=rerr&"ɾ�����۵����ͳ���\n"
  end if
  if not(isnumeric(re_id)) then
    rerr=rerr&"ɾ�����۵� ID ����\n"
  end if
  if rerr<>"" then
    call close_conn()
    call review_err(rerr)
    response.end
  end if
  sql="delete from review where rsort='"&rsort&"' and re_id="&re_id
end sub

sub review_main()
  dim rusername,remail,rword
  rusername=code_form(trim(request.form("rusername")))
  remail=code_form(trim(request.form("remail")))
  rword=code_form(trim(request.form("rword")))
  rsort=trim(request.form("rsort"))
  re_id=trim(request.form("re_id"))
  rurl=trim(request.form("rurl"))
  if review_rsort(rsort)<>"yes" then
    rerr=rerr&"�������۵����ͳ�������\n"
  end if
  if not(isnumeric(re_id)) then
    rerr=rerr&"�������۵� ID ��������\n"
  end if
  if symbol_name(rusername)<>"yes" then
    rerr=rerr&"�������������ƣ������ú��зǷ��ַ���\n"
  end if
  if len(remail)>0 then
    if email_ok(remail)<>"yes" or len(remail)>50 then
      rerr=rerr&"���� E-mail ���ú��зǷ��ַ���\n"
    end if
  end if
  if len(rword)<1 then
    rerr=rerr&"��û��������������ݣ�\n"
  elseif len(rword)>250 then
    rerr=rerr&"���������������̫����(<=250�ֽ�)\n"
  end if
  if rerr<>"" then call review_err(rerr):exit sub

  on error resume next
  sql="insert into review(rsort,re_id,rusername,remail,rword,rtim,rtype) values('"&rsort&"',"&re_id&",'"&rusername&"','"&remail&"','"&rword&"','"&now_time&"',"
  if rusername=login_username then
    sql=sql&"1"
  else
    sql=sql&"0"
  end if
  sql=sql&")"
  conn.execute(sql)
  if err then
    err.clear
    call review_err("����Ĵ�������վ����ϵ��\nhttp://beyondest.com/\n")
    exit sub
  end if

  response.write vbcrlf&"<script lanuage=javascript><!--" & _
		 vbcrlf&"alert(""���ɹ��ķ������й��������ۣ�\n\nлл���Ĳ��룡�������..."");"
  if len(rurl)<5 then
    response.write vbcrlf&"location.href='main.asp';"
  else
    response.write vbcrlf&"location.href='"&rurl&"';"
  end if
  response.write vbcrlf&"--></script>"
end sub

sub review_err(revar)
  response.write vbcrlf&"<script lanuage=javascript><!--" & _
		 vbcrlf&"alert(""���ڷ�������ʱ�������´���\n\n"&revar&"\n�������..."");" & _
		 vbcrlf&"history.back(-1);" & _
		 vbcrlf&"--></script>"
end sub
%>