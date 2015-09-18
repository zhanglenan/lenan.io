<%
Const Db_Type = "ACC"
Const Db_DateTime = "Now()"

Dim Conn,ConnStr,Rs

Dim SysSiteDbPath
SysSiteDbPath = "Db/Ok3w#47#ASPACCFREE.ASP"

ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbdns & SysSiteDbPath)

Call OpenConn()

Set Rs = Server.CreateObject("Adodb.RecordSet")

If Application("Ok3w_SiteTitle") = "" Then
	Application.Lock()
	Sql = "select * from Ok3w_SiteConfig"
	Rs.Open Sql,Conn,1,1
	For i = 0 To Rs.Fields.Count - 1
		Application("Ok3w_" & Rs.Fields(i).Name) = Rs(i).Value
	Next
	Rs.Close
	Application.UnLock()
End If

Private Sub CloseConn()
	Conn.Close
	Set Conn = Nothing
End Sub

Private Sub OpenConn()
	On Error Resume Next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open ConnStr
	If Err.Number<>0 Then
		Response.Write("数据库连接错误")
		Err.Clear
		Response.End()
		Else
			On Error Goto 0
	End If
End Sub
%>
