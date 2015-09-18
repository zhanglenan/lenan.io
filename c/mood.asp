<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Const dbdns="../"%>
<!--#include file="../AppCode/Conn.asp"-->
<%
ID = Cdbl(Request.QueryString("ID"))
action = Request.QueryString("action")

If action="mood" And InStr("," & Request.Cookies("Ok3w")("pMoodArticleID"),"," & ID & ",")>=1 Then
%>
<script language="javascript">
alert("囧，-_-|||，你不是表过态了嘛？");
</script>
<%
Else

pMoodStr = "0,0,0,0,0,0,0,0"
Sql="select pMoodStr,Hits from Ok3w_Article where ID=" & ID
Rs.Open Sql,Conn,0,1
If Not IsNull(Rs("pMoodStr")) Then
	pMoodStr = Rs("pMoodStr")
End If
Hits = Rs("Hits")
mTmp = Split(pMoodStr,",")

If action="mood" Then
	mII = Cint(Request.Form("mood"))
	mTmp(mII-1) = mTmp(mII-1) + 1
	pMoodStr = mTmp(0) & "," & mTmp(1) & "," & mTmp(2) & "," & mTmp(3) & "," & mTmp(4) & "," & mTmp(5) & "," & mTmp(6) & "," & mTmp(7)
	
	Sql="update  Ok3w_Article set pMoodStr='" & pMoodStr & "' where ID=" & ID
	Conn.Execute Sql
	
	Response.Cookies("Ok3w")("pMoodArticleID") = Request.Cookies("Ok3w")("pMoodArticleID") & ID & ","
End If

moodCount = Cint(mTmp(0)) + Cint(mTmp(1)) + Cint(mTmp(2)) + Cint(mTmp(3)) + Cint(mTmp(4)) + Cint(mTmp(5)) + Cint(mTmp(6)) + Cint(mTmp(7))
mHH = 60
If moodCount = 0 Then
	moodImg = "1,1,1,1,1,1,1,1"
Else
	moodImg = Int(mTmp(0)/moodCount*mHH) & "," &  Int(mTmp(1)/moodCount*mHH) & "," & Int(mTmp(2)/moodCount*mHH) & "," & Int(mTmp(3)/moodCount*mHH) & "," & Int(mTmp(4)/moodCount*mHH) & "," & Int(mTmp(5)/moodCount*mHH) & "," & Int(mTmp(6)/moodCount*mHH) & "," & Int(mTmp(7)/moodCount*mHH)
End If
mITmp = Split(moodImg,",")
%>
<script language="javascript">
<%For i=1 To 8
If mITmp(i-1)=0 Then mITmp(i-1)=1
%>
parent.document.getElementById("moodimg<%=i%>").style.height = "<%=mITmp(i-1)%>px";
parent.document.getElementById("moodnum<%=i%>").innerHTML = "<%=mTmp(i-1)%>";
<%Next%>
parent.document.getElementById("News_Hits").innerHTML = "<%=Hits%>";
</script>
<%
End If

Set Rs = Nothing
Call CloseConn()
%>