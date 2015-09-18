<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Const dbdns="../"%>
<!--#include file="../AppCode/Conn.asp"-->
<%
ID = Cdbl(Request.QueryString("ID"))
Select Case Trim(Request.QueryString("type"))
	Case "soft"
		Sql="select ID,Hits from Ok3w_Soft where ID=" & ID
		Rs.Open Sql,Conn,0,1
		Hits = Rs("Hits")
		Rs.Close
	Case Else
		Sql="select ID,Hits from Ok3w_Article where ID=" & ID
		Rs.Open Sql,Conn,1,3
		Rs("Hits") = Rs("Hits") + 1
		Rs.Update
		Hits = Rs("Hits")
		Rs.Close
End Select

Set Rs = Nothing
Call CloseConn()
%>

<script language="javascript">
parent.Article_Hits.innerHTML = "<%=Hits%>";
</script>