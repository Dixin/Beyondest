<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

'---------------------���ݿ����ͼ�·������---------------------
Dim conn,connstr
connstr  = "DBQ=" + Server.mappath("data/ip_address.mdb") + ";DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn = Server.CreateObject("ADODB.CONNECTION")
conn.open connstr

Sub close_conn()
    conn.Close
    Set conn = Nothing
End Sub %>