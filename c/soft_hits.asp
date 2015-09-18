<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Const dbdns="../"%>
<!--#include file="../AppCode/Conn.asp"-->
<%
Dim id,Hits,Ding_Hits,Cai_Hits,Sp1,Sp2
id = Cdbl(Request.QueryString("id"))
Select Case Request.QueryString("type")
	Case "display"
		action = Request.QueryString("action")
		If action = "" Then
		ElseIf action="1" Then
			Call ChkVote(id)
			Sql="update Ok3w_Soft set Ding_Hits=Ding_Hits+1 where Id=" & id
			Conn.Execute Sql
			Call IsVote(id)
		ElseIf action="2" Then
			Call ChkVote(id)
			Sql="update Ok3w_Soft set Cai_Hits=Cai_Hits+1 where Id=" & id
			Conn.Execute Sql
			Call IsVote(id)
		End If		
		Sql="select Hits,Ding_Hits,Cai_Hits from Ok3w_Soft where Id=" & id
		Rs.Open Sql,Conn,0,1
		Hits = Rs(0)
		Ding_Hits = Rs(1)
		Cai_Hits = Rs(2)
		If Ding_Hits=0 And Cai_Hits=0 Then
			Sp1 = 0
			Sp2 = 0
			Else
				Sp2 = Int(Cai_Hits/(Ding_Hits+Cai_Hits) * 100)
				Sp1 = 100 - Sp2
		End If
		Call Js_DisplayHits()
End Select

Set Rs = Nothing
Call CloseConn()

Private Sub IsVote(ID)
	Response.Cookies("Ok3w")("VoteStrID") = Request.Cookies("Ok3w")("VoteStrID") & "," & id & ","
End Sub

Private Sub ChkVote(ID)
	If InStr(Request.Cookies("Ok3w")("VoteStrID"),"," & id & ",")>=1 Then
%>
		<script language="javascript">
		alert("谢谢，你已经投过票了！");
		</script>
<%
		Set Rs = Nothing
		Call CloseConn()
		Response.End()
		Else
%>
		<script language="javascript">
		alert("谢谢，投票成功！");
		</script>
<%
	End If
End Sub		

Private Sub Js_DisplayHits()
%>
<script language="javascript">
parent.downCount.innerHTML = "<%=Hits%>";
parent.s1.innerHTML = "<%=Ding_Hits%>";
parent.s2.innerHTML = "<%=Cai_Hits%>";
parent.sp1.innerHTML = "<%=Sp1%>%";
parent.sp2.innerHTML = "<%=Sp2%>%";
parent.eimg1.style.width = "<%=Int(Sp1/100*55)%>px";
parent.eimg2.style.width = "<%=Int(Sp2/100*55)%>px";
</script>
<%End Sub%>