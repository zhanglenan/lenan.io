<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Dim action
Dim Info
Dim dns
dns = Request.QueryString("dns")
action = Request.QueryString("action")
Select Case action
	Case "chkuserlogin"
		If Request.Cookies("Ok3w")("User_Name")<>"" Then
          	Info = "�װ���<strong>" & Request.Cookies("Ok3w")("User_Name") & "</strong>��ӭ������[<a href=""" & dns & "User_Index.asp"">�����������</a>] [<a href=""" & dns & "save.asp?action=LoginOut"">��ȫ�˳�</a>]"
			Response.Write("UserLoginInfo.innerHTML='" & Info & "';")
		End If
	Case "siteisclose"
		If Application("Ok3w_SiteIsClose") = "1" Then
			Info = "document.location.href='" & dns & "c/siteclose.asp';"
			Response.Write(Info)
		End If
End Select
%>
