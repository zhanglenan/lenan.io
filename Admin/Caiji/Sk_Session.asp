<%
If Session("Ok3w_Caiji")="" Then
	Response.Write("�㻹û�е�½�����ǵ�½�Ѿ���ʱ��")
	Response.End()
End If

response.Cookies(Site)("IsAdmin")=1
response.Cookies(Site)("Admin_name")="Admin"
response.Cookies(Site)("Admin_type")="��������Ա"
%>