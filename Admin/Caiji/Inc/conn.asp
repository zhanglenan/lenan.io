<%
Const CMSDataBase=0
Const DataBaseType=0 
Const SysVer=1
Dim SqlNowString,SiteSN,Conn,DBPath,DataServer,DataUser,DataBaseName,DataBasePsw,ConnStr,CollcetConnStr
Dim sql_databasename,sql_password,sql_username,sql_localname
IF CMSDataBase=0 Then
	If DataBaseType=0 Then
		'ACCESS数据库设置
		DBPath       = "../../Db/Ok3w#47#ASPACCFREE.ASP"
		SqlNowString = "Now()"
	End If
End if
dim db,ErrMsg,ErrMsg_lx,Site
Site="ok3w.glzy8.com"
dim connstrItem
dim dbItem
dim connItem
dbItem="Data/Ok3w_Caiji_Data.asp"
Set connItem = Server.CreateObject("ADODB.Connection")
connstrItem="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbItem)
connItem.Open connstrItem
If Err Then
   err.Clear
   Set ConnItem = Nothing
   Response.Write "采集数据库连接出错，请检查连接字串。"
   Response.End
End If
Sub CloseConnItem()
   On Error Resume Next
   ConnItem.close
   Set ConnItem=nothing
End sub
'---------------------数据库CMS------------------------------------------------------------------
Sub OpenConn()
	On Error Resume Next
	If DataBaseType=0 Then
		ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBPath)
	Else
		ConnStr = "Provider = Sqloledb; User ID = " & sql_username & "; Password = " & sql_password & "; Initial Catalog = " & sql_databasename & "; Data Source = " & sql_localname & ";"
	End If
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.open ConnStr
	If Err.Number<>0 Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "数据库连接出错，请检查Admin/Caiji/Inc/Conn.asp文件中的数据库参数设置。"
		Response.End
		Else
			On Error Goto 0
	End If
End Sub

Sub CloseConn()
	Conn.close
	Set Conn=nothing
End sub
'-------------------------------------------------------------------------------------------------------------
%>