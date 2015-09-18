<%
'--------------------------------------------------------------------
'名称：ASP分页类 v2009
'
'作者：zhengbi（QQ：124895502 Email:zhengbi888@yahoo.com.cn）
'
'仅供参考，你可以任意修改、传播；如果你对它进行了改良，希望你能给我一个
'副本，谢谢。
'--------------------------------------------------------------------
Class TurnPage
	Dim sPageNo,sPageSize,sPageCount,sRecordCount,sAbsoluteRecord
	
	Private Sub Class_Initialize()
		sPageNo=Trim(Request.QueryString("PageNo"))
		If sPageNo<>"" Then
			sPageNo = Cdbl(sPageNo)
			Else
				sPageNo = 1
		End If
		sAbsoluteRecord = 1
	End Sub

	Public Sub GetRs(ByRef Conn,ByRef Rs,ByVal Sql,ByVal PageSize)
		Rs.Open Sql,Conn,1,1
		
		Rs.PageSize		= PageSize
		sPageSize		= Rs.PageSize
		sPageCount		= Rs.PageCount
		sRecordCount	= Rs.RecordCount
		
		If Not Rs.Eof Then Rs.AbsolutePage = sPageNo
	End Sub
	
	Public Function Eof()
		If sAbsoluteRecord<=sPageSize Then
			'sAbsoluteRecord = sAbsoluteRecord + 1
			Eof = False
			Else
				Eof = True
		End If
	End Function
	
	Public Sub MoveNext()
		sAbsoluteRecord = sAbsoluteRecord + 1
	End Sub
	
	
	Public Sub GetPageList()
		Dim sURL,sTmp,sQUERY_STRING,p,n,i,a,b
		
		sURL = Request.ServerVariables("URL")
		sQUERY_STRING = Request.ServerVariables("QUERY_STRING")
		sTmp = Split(sURL,"/")
		sURL = sTmp(Ubound(sTmp))
		If sQUERY_STRING <> "" Then	sQUERY_STRING=Replace(sQUERY_STRING,"PageNo=" & sPageNo,"")
		If sQUERY_STRING = "" Then
			sURL = sURL & "?"
			Else
				sURL = sURL & "?" & sQUERY_STRING & "&"
		End If
		sURL = Replace(sURL,"&&","&")
		
		p = sPageNo-1
		n = sPageNo+1
		If p<1 Then p = 1
		If n>sPageCount Then n = sPageCount
		
		a = sPageNo-5
		b = sPageNo+5
		If a<1 Then a = 1
		If b>sPageCount Then b = sPageCount
				
		Response.Write("<div class=""page_nav"">")
		
		Response.Write("<a href=""" & sURL & "PageNo=1"">首页</a> <a href=""" & sURL &"PageNo=" & p & """>上页</a>")
		
		For i=a To b
			Response.Write(" <a href=""" & sURL & "PageNo=" & i & """")
			If i = sPageNo Then	Response.Write(" style=""font-weight:bold; color:#FF0000;""")
			Response.Write(">" & i & "</a>")
		Next
		Response.Write(" <a href=""" & sURL & "PageNo=" & n & """>下页</a> <a href=""" & sURL & "PageNo=" & sPageCount & """>末页</a>")

		Response.Write("</div>")
	End Sub
End Class
%>