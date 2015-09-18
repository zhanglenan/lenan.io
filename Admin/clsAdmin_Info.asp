<%
Class Admin_Info
	Dim AdminId
	Dim AdminName
	Dim AdminPwd
	Dim GroupId
	
	Private Sub Class_Initialize
		AdminId = Request.Cookies("Ok3w")("AdminId")
		AdminName = Request.Cookies("Ok3w")("AdminName")
		AdminPwd = Request.Cookies("Ok3w")("AdminPwd")
		GroupId = Request.Cookies("Ok3w")("GroupId")
	End Sub
	
	Public Function AdminLogin(sAdminName,sAdminPwd,sType)
		AdminName = sAdminName
		
		Sql = "select * from Ok3w_Admin where AdminName=? and AdminPwd=?"
		Set AdminCmd = Server.CreateObject("Adodb.Command")
		AdminCmd.ActiveConnection = Conn
		AdminCmd.CommandType = 1
		AdminCmd.CommandText = Sql
		AdminCmd.Parameters.Append(AdminCmd.CreateParameter("@AdminName",200,1,50,sAdminName))
		AdminCmd.Parameters.Append(AdminCmd.CreateParameter("@AdminPwd",200,1,50,sAdminPwd))
		Set AdminRs = Server.CreateObject("Adodb.RecordSet")
		Set AdminRs = AdminCmd.Execute
		Set AdminCmd = Nothing
		If AdminRs.Eof And AdminRs.Bof Then
			AdminLogin = 1'�û������������
			Else
				If AdminRs("AdminLock") Then
					AdminLogin = 2'�û�������
					Else
						Response.Cookies("Ok3w")("AdminId") = AdminRs("AdminId")
						Response.Cookies("Ok3w")("AdminName") = AdminRs("AdminName")
						Response.Cookies("Ok3w")("AdminPwd") = AdminRs("AdminPwd")
						Response.Cookies("Ok3w")("GroupId") = AdminRs("GroupId")

						If sType="IsLogin" Then Call AdminActionLog("�ɹ���½")
						AdminLogin = -1'�ɹ���½
				End If
		End If
		AdminRs.Close
		Set AdminRs = Nothing
	End Function
	
	'����Ա������־
	Public Sub AdminActionLog(logInfo)
		Set logRs = Server.CreateObject("Adodb.RecordSet")
		Sql = "select * from Ok3w_Adlog where 1=2"
		logRs.Open Sql,Conn,1,3
		logRs.AddNew
		logRs("logUser") = AdminName
		logRs("logIp") = Request.ServerVariables("REMOTE_ADDR")
		logRs("LogTime") = Now()
		logRs("LogInfo") = logInfo
		logRs("LogType") = 0
		logRs.Update
		logRs.Close
		Set logRs = Nothing
	End Sub
	
	Public Function AdminIsLogin()
		If Trim(AdminName) = "" Then
			AdminIsLogin = 0'û�е�½
			Else
				If AdminLogin(AdminName,AdminPwd,"IsCheck")<>-1 Then
					AdminIsLogin = 0'Cookies����
					Else
						AdminIsLogin = -1'�Ѿ���½
				End If
		End If
	End Function
	
	Public Function AdminIsFlag(Flag)
		If InStr(GroupId,"," & Flag & ",") <= 0 Then
			AdminIsFlag = 0'û�и�Ȩ��
			Else
				AdminIsFlag = -1'�д�Ȩ��
		End If
	End Function
	
	Public Function AdminEditPassWord(oldpwd,newpwd)
		If AdminPwd <> MD5(oldpwd) Then
			AdminEditPassWord = 0'ԭ���벻��ȷ
			Else
				Sql = "select * from Ok3w_Admin where AdminId=" & AdminId
				Set AdminRs = Server.CreateObject("Adodb.RecordSet")
				AdminRs.Open Sql,Conn,1,3
				AdminRs("AdminPwd") = MD5(newpwd)
				AdminRs.Update
				AdminRs.Close
				Set AdminRs = Nothing
				
				Response.Cookies("Ok3w")("AdminPwd") = MD5(newpwd)
				
				AdminEditPassWord = -1'�޸ĳɹ�
		End If
	End Function
End Class
%>
	
