<%
If Session("Ok3w_Caiji")="" Then
	Response.Write("你还没有登陆，或是登陆已经超时。")
	Response.End()
End If

response.Cookies(Site)("IsAdmin")=1
response.Cookies(Site)("Admin_name")="Admin"
response.Cookies(Site)("Admin_type")="超级管理员"
%>