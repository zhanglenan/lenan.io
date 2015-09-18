<%
Private Sub Ok3w_Article_List(ClassID,Rows,Cels,LeftN,IsCommend,DisClass,DisTime,TimeFormat,DisHits,OrderType)
Sql="select top " & Rows * Cels & " Id,Title,TitleColor,TitleURL,AddTime,Hits from Ok3w_Article where ChannelID=1 and IsPass=1 and IsDelete=0"
If ClassID<>"" Then Sql=Sql & " and SortPath like '%," & ClassID & ",%'"
If IsCommend Then Sql=Sql & " and IsCommend=1"
Select Case OrderType
	Case "hot"
		Sql = Sql & " order by Hits desc,AddTime desc,ID desc"
	Case "rnd"
		Randomize
		Sql = Sql & " order by Rnd(-(ID+"&Rnd()&")),ID desc"
	Case "new"
		Sql = Sql & " order by AddTime desc,ID desc"
	Case Else
		Sql = Sql & " order by IsTop desc,IsCommend desc,AddTime desc,ID desc"
End Select
Rs.Open Sql,Conn,0,1
StrTmp = "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
Do While Not Rs.Eof
StrTmp = StrTmp & "<tr>"
For II=1 To Cels 
If Rs.Eof Then
	StrTmp = StrTmp & "<td>&nbsp;</td>"
Else
	StrTmp = StrTmp & "<td class=""newstitle""><div>"
	If DisClass Then
		StrTmp = StrTmp & "[<a href=""" & Htmldns & """list.asp?id=" & Rs("ClassID") & """>" & GetClassName(Rs("ClassID")) & "</a>]"
	Else
		StrTmp = StrTmp & "·"
	End If
	StrTmp = StrTmp & Format_TitleURL("show.asp",Rs("ID"),Rs("Title"),Rs("TitleColor"),Rs("TitleURL"),LeftN)
	StrTmp = StrTmp & "</div><span>"
	If DisHits Then
		StrTmp = StrTmp & " 查看:" & Rs("Hits")
	End If
	If DisTime Then
		StrTmp = StrTmp & " " & Format_Time(Rs("AddTime"),TimeFormat)
	End If
	StrTmp = StrTmp & "</span>"
	StrTmp = StrTmp & "</td>"
	Rs.MoveNext
End If
If Cels<>1 And II<>Cels Then
	StrTmp = StrTmp & "<td>&nbsp;</td>"
End If
Next
StrTmp = StrTmp & "</tr>"  
Loop
Rs.Close 
StrTmp = StrTmp & "</table>"
Response.Write(StrTmp)
End Sub


Private Sub Ok3w_Article_PreNext(ClassID,ID)
	Sql="select ID,Title,TitleColor,TitleURL from Ok3w_Article where SortPath like '%," & ClassID & ",%' and IsPass=1 and IsDelete=0 and ID>" & ID & " order by ID asc"
	Set oRs = Conn.Execute(Sql)
	Response.Write("<div style=""width:50%; float:left;"">")
	If oRs.Eof And oRs.bof Then
		Response.Write("上一篇：没有了")
		Else
			Response.Write("上一篇：" & Format_TitleURL("show.asp",oRs("ID"),oRs("Title"),oRs("TitleColor"),oRs("TitleURL"),50))
	End If
	Response.Write("</div>")
	Sql="select ID,Title,TitleColor,TitleURL from Ok3w_Article where SortPath like '%," & ClassID & ",%' and IsPass=1 and IsDelete=0 and ID<" & ID & " order by ID desc"
	Set oRs = Conn.Execute(Sql)
	Response.Write("<div style=""width:50%; float:right;"">")
	If oRs.Eof And oRs.bof Then
		Response.Write(" 下一篇：没有了")
		Else
			Response.Write(" 下一篇：" & Format_TitleURL("show.asp",oRs("ID"),oRs("Title"),oRs("TitleColor"),oRs("TitleURL"),50))
	End If
	Response.Write("</div>")
End Sub


Private Sub Ok3w_Article_Gundong(ClassID,TopN,Width,Height,Speed)
Randomize
RndID = Int(Rnd()*10000)
RndFun = "__Ok3w_Article_Gundong__00" & RndID
Sql="select top " & TopN & " Id,Title,TitleColor,TitleURL from Ok3w_Article where ChannelID=1 and IsPass=1 and IsDelete=0 and IsMove=1"
If ClassID<>"" Then Sql=Sql & " and SortPath like '%," & ClassID & ",%'"
Sql=Sql & " order by IsTop desc,IsCommend desc,Id desc"
Rs.Open Sql,Conn,0,1
StrTmp = ""
Do While Not Rs.Eof
	StrTmp = StrTmp & "·" & Format_TitleURL("show.asp",Rs("ID"),Rs("Title"),Rs("TitleColor"),Rs("TitleURL"),50) & "&nbsp;&nbsp;"
	Rs.MoveNext
Loop
Rs.Close
StrTmp = Replace(StrTmp,"'","\'")

Response.Write("<script language=""javascript"">")
Response.Write("function " & RndFun & "()")
Response.Write("{")
Response.Write("	var RndID = 'pro" & RndID & "';")
Response.Write("	var StrGD = '" & StrTmp & "';")
Response.Write("	Ok3w_Marquee(RndID,StrGD," & Width & "," & Height & "," & Speed & ");")
Response.Write("}")
Response.Write(RndFun & "();")
Response.Write("</script>")
End Sub


Private Sub Ok3w_Article_ImgGD(ClassID,Rows,Cels,Width,Height,ImgW,ImgH,IsCommend,OrderType,Speed)
Randomize
RndID = Int(Rnd()*10000)
RndFun = "__Ok3w_Article_ImgGD__00" & RndID
Sql="select top " & Rows * Cels & " Id,Title,TitleColor,TitleURL,PicFile from Ok3w_Article where ChannelID=1 and IsPass=1 and IsDelete=0 and IsPic=1 and IsPlay=0"
If ClassID<>"" Then Sql=Sql & " and SortPath like '%," & ClassID & ",%'"
If IsCommend Then Sql=Sql & " and IsCommend=1"
Select Case OrderType
	Case "hot"
		Sql = Sql & " order by Hits desc,AddTime desc,ID desc"
	Case "rnd"
		Randomize
		Sql = Sql & " order by Rnd(-(ID+"&Rnd()&")),ID desc"
	Case "new"
		Sql = Sql & " order by AddTime desc,ID desc"
	Case Else
		Sql = Sql & " order by IsTop desc,IsCommend desc,AddTime desc,ID desc"
End Select
StrTmp = ""
Rs.Open Sql,Conn,0,1
StrTmp = "<table border=""0"" cellspacing=""5"" cellpadding=""0"">"
Do While Not Rs.Eof
StrTmp = StrTmp & "  <tr>"
For img=1 To Cels
StrTmp = StrTmp & "    <td align=""center"">"
If Rs.Eof Then
	StrTmp = StrTmp & "&nbsp;"
Else
	StrTmp = StrTmp & "<div style=""width:" & ImgW & "px; height:" & ImgH & "px; border:1px solid #CCCCCC; margin:0px 0px 5px 0px; padding:1px;"">"
	If Rs("TitleURL")="" Then
		uTmp = Htmldns & "show.asp?id=" & Rs("ID") & ""
	Else
		uTmp = Rs("TitleURL")
	End If
	StrTmp = StrTmp & "<a href=""" & uTmp & """ target=""_blank"">"
	StrTmp = StrTmp & "<img src=""" & Rs("PicFile") & """ width=""" & ImgW & """ height=""" & ImgH & """ alt=""" & Rs("Title") & """ border=""0"" />"
	StrTmp = StrTmp & "</a>"
	StrTmp = StrTmp & "</div>"
	StrTmp = StrTmp & Format_TitleURL("show.asp",Rs("ID"),Rs("Title"),Rs("TitleColor"),Rs("TitleURL"),ImgW\16)
	Rs.MoveNext
End If
StrTmp = StrTmp & "</td>"
Next
StrTmp = StrTmp & "  </tr>"
Loop
Rs.Close 
StrTmp = StrTmp & "</table>"

StrTmp = Replace(StrTmp,"'","\'")

Response.Write("<script language=""javascript"">")
Response.Write("function " & RndFun & "()")
Response.Write("{")
Response.Write "	var RndID = 'pro" & RndID & "';"
Response.Write "	var StrGD = '" & StrTmp & "';"
Response.Write "	Ok3w_Marquee(RndID,StrGD," & Width & "," & Height & "," & Speed & ");"
Response.Write "}"
Response.Write RndFun & "();"
Response.Write "</script>"
End Sub


Private Sub Ok3w_Article_ImgFlash(ClassID,Width,Height)
Randomize
RndID = Int(Rnd()*10000)
RndFun = "__Ok3w_Article_ImgFlash__00" & RndID
Response.Write "<script language=""javascript"">"
Response.Write "function " & RndFun & "()"
Response.Write "{"
	Sql="select top 5 Id,Title,PicFile,TitleColor,TitleURL from Ok3w_Article where ChannelID=1 and IsPass=1 and IsPlay=1 and IsDelete=0"
	If ClassID<>"" Then Sql=Sql & " and SortPath like '%," & ClassID & ",%'"
	Sql=Sql & " order by IsTop desc,IsCommend desc,Id desc"
	Rs.Open Sql,Conn,0,1
	pics = ""
	links = ""
	texts = ""
	Do While Not Rs.Eof
		ID = Rs("ID")
		Title = Rs("Title")
		PicFile = Rs("PicFile")
		TitleURL = Rs("TitleURL")
		If TitleURL = "" Then TitleURL = "show.asp?id=" & ID 
		
		pics = pics & PicFile & "|"
		links = links & TitleURL & "|"
		texts = texts & Title & "|"
		Rs.MoveNext
	Loop
	Rs.Close
	If texts<>"" Then
		pics = Left(pics,Len(pics)-1)
		links = Left(links,Len(links)-1)
		texts = Left(texts,Len(texts)-1)
	End If
	
Response.Write "	var pics=""" & pics & """;"
Response.Write "	var links=""" & links & """;"
Response.Write "	var texts=""" & texts & """;"
	
Response.Write "	var focus_width=" & Width & ";"
Response.Write "	var focus_height=" & Height - 18 & ";"
Response.Write "	var text_height=18;"
Response.Write "	var swf_height = focus_height + text_height;"
	 
Response.Write "	Ok3w_insertFlash(""" & Htmldns & """, focus_width, focus_height, swf_height, text_height, pics, links, texts);"
Response.Write "}"
Response.Write RndFun & "();"
Response.Write "</script>"
End Sub


Private Sub Ok3w_Article_aList(ClassID,ImgCels,ImgW,ImgH)
	Sql="select * from Ok3w_Class where ParentID=" & ClassID & " and gotoURL='' order by OrderID"
	Set oRs = Conn.Execute(Sql)
	If oRs.Eof And oRs.Bof Then
		oRs.Close
		Set oRs = Nothing
		
		Sql="select * from Ok3w_Class where ID=" & ClassID
		Set oRs=Conn.Execute(Sql)
		IsPic = oRs("IsPic")
		PageSize = oRs("PageSize")
		If IsPic=1 Then
			Call Ok3w_Article_pList(ClassID,PageSize,ImgCels,ImgW,ImgH)
			Else
				Call Ok3w_Article_bList(ClassID,PageSize)
		End If
	Else
		Call Ok3w_Article_sList(oRs)
		oRs.Close
		Set oRs = Nothing
	End If
End Sub
%>

<%
Private Sub Ok3w_Article_sList(ByRef oRs)
Do While Not oRs.Eof
Response.Write("<div class=""xtitle_box"">")
Response.Write("<div class=""xtitle""><h3>" & oRs("SortName") & "<font color=""#999999"">|</font></h3><span><a href=""" & "list.asp?id=" & oRs("ID") & """>更多&gt;&gt;</a></span></div>")
Response.Write("<div style=""padding:5px;"">")
Call Ok3w_Article_List(oRs("ID"),10,1,50,False,False,True,1,False,"new")
Response.Write("</div>")
Response.Write("</div>")
oRs.MoveNext
Loop
End Sub
%>

<%
Private Sub Ok3w_Article_bList(ClassID,TopN)
Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
Sql="select Id,Title,Content,TitleColor,TitleURL,AddTime from Ok3w_Article where IsPass=1 and IsDelete=0 and SortPath like '%," & ClassID & ",%' order by AddTime desc,ID desc"
Call Page.GetRs(Conn,Rs,Sql,TopN)
i_Count = 0
Do While Not Rs.Eof And Not Page.Eof
i_Count = i_Count + 1
Response.Write("  <tr>")
Response.Write("    <td class=""clist"">·" & Format_TitleURL("show.asp",Rs("ID"),Rs("Title"),Rs("TitleColor"),Rs("TitleURL"),50) & "</td>")
Response.Write("    <td align=""right"">" & Format_Time(Rs("AddTime"),1) & "</td>")
Response.Write("  </tr>")
Rs.MoveNext
Page.MoveNext
If i_Count Mod 5 = 0 Or Rs.Eof Then
Response.Write("  <tr>")
Response.Write("    <td colspan=""2""><div style=""border-bottom:1px dotted #CCCCCC; margin:0px 0px 15px 0px; width:100%;""></div></td>")
Response.Write("  </tr>")
End If
Loop
Rs.Close
Response.Write("</table>")
Call Page.GetPageList()
End Sub


Private Sub Ok3w_Article_pList(ClassID,TopN,Cels,ImgW,ImgH)
Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""8"" cellspacing=""0"">")
Sql="select Id,Title,TitleColor,TitleURL,PicFile from Ok3w_Article where IsPass=1 and IsDelete=0 and SortPath like '%," & ClassID & ",%' order by AddTime desc,ID desc"
Call Page.GetRs(Conn,Rs,Sql,TopN)
Do While Not Rs.Eof And Not Page.Eof
Response.Write("  <tr>")
For iCount = 1 To Cels  
Response.Write("    <td width=""" & 100\Cels & "%"" align=""center"">")
If Rs.Eof Or Page.Eof Then
	Response.Write("&nbsp;")
Else
	Response.Write("<div style=""width:" & ImgW & "px; height:" & ImgH & "px; border:1px solid #CCCCCC; margin:0px 0px 5px 0px; padding:1px;"">")
	If Rs("TitleURL")="" Then
		uTmp = "show.asp?id=" & Rs("ID")
	Else
		uTmp = Rs("TitleURL")
	End If
    Response.Write("<a href=""" & uTmp & """ target=""_blank"">")
	Response.Write("<img src=""" & Rs("PicFile") & """ alt=""" & Rs("Title") & """ width=""" & ImgW & """ height=""" & ImgH & """ border=""0"" /></a></div>")
    Response.Write(Format_TitleURL("show.asp",Rs("ID"),Rs("Title"),Rs("TitleColor"),Rs("TitleURL"),ImgW\16))
	Rs.MoveNext
	Page.MoveNext
End If
Response.Write("</td>")
Next
Response.Write("  </tr>")
Loop
Rs.Close
Response.Write("</table>")
Response.Write("<div style=""border-bottom:1px dotted #CCCCCC; margin:8px 0px;""></div>")
Call Page.GetPageList()
End Sub


Private Sub Ok3w_Search_List(ClassID,sType,Keyword,TopN)
Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
Sql="select Id,Title,TitleColor,TitleURL,AddTime from Ok3w_Article where IsPass=1 and IsDelete=0 and ChannelID=1"
Select Case sType
	Case "Title"
		Sql = Sql & " and Title like '%" & Keyword & "%'"
	Case "Content"
		Sql = Sql & " and Content like '%" & Keyword & "%'"
	Case Else
		Sql = Sql & " and (Title like '%" & Keyword & "%' or Content like '%" & Keyword & "%')"
End Select
Sql = Sql & " order by AddTime desc,ID desc"
Call Page.GetRs(Conn,Rs,Sql,TopN)
If Not(Rs.Eof And Rs.Bof) Then
i_Count = 0
Do While Not Rs.Eof And Not Page.Eof
i_Count = i_Count + 1
Response.Write("  <tr>")
Response.Write("    <td class=""clist"">·" & Format_TitleURL("show.asp",Rs("ID"),Rs("Title"),Rs("TitleColor"),Rs("TitleURL"),50) & "</td>")
Response.Write("    <td align=""right"">" & Format_Time(Rs("AddTime"),1) & "</td>")
Response.Write("  </tr>")
Rs.MoveNext
Page.MoveNext
If i_Count Mod 5 = 0 Or Rs.Eof Then
Response.Write("  <tr>")
Response.Write("    <td colspan=""2""><div style=""border-bottom:1px dotted #CCCCCC; margin:0px 0px 15px 0px; width:100%;""></div></td>")
Response.Write("  </tr>")
End If
Loop
Else
Response.Write("			<tr>")
Response.Write("              <td><br /><br /><h1>Sorry!没有找到任何相关内容。</h1><br /><br /></td>")
Response.Write("            </tr>")
End If
Rs.Close
Response.Write("</table>")
Call Page.GetPageList()
End Sub


Private Sub Ok3w_Article_SmallClass(ClassID)	
	Sql="select ID,SortName,gotoURL from Ok3w_Class where ParentID=" & ClassID & " and IsNav=1 order by OrderID"
	Rs.Open Sql,Conn,0,1
	If Rs.Eof And Rs.Bof Then
		Sql = "select ParentID from Ok3w_Class where ID=" & ClassId
		a_ParentID = ExecSqlReturnOneValue(Sql)
		If a_ParentID<>0 Then
			Rs.Close
			Sql="select ID,SortName,gotoURL from Ok3w_Class where ParentID=" & a_ParentID & " and IsNav=1 order by OrderID"
			Rs.Open Sql,Conn,0,1
		End If
	End If
	If Not Rs.Eof Then
	Response.Write("<div class=""classlist"">")
	Response.Write("<strong style=""font-size:12px;"">" & GetParentClassName(ClassID) & "<span>&gt;&gt;</span></strong>")
	Do While Not Rs.Eof
		If Rs("gotoURL")="" Then
			Response.Write("<a href=""" & "list.asp?id=" & Rs("ID") & """>" & Rs("SortName") & "</a>")
		Else
			Response.Write("<a href=""" & Rs("gotoURL") & """ target=""_blank"">" & Rs("SortName") & "</a>")
		End If
		Rs.MoveNext
		If Not Rs.Eof Then Response.Write("<span>|</span>")
	Loop
	Response.Write("</div>")
	End If
	Rs.Close
End Sub


Private Sub Ok3w_Article_Class_Nav(SortPath)
cTmp = Left(SortPath,Len(SortPath)-1)
Sql="select ID,SortName from Ok3w_Class where ID in(" & cTmp & ") order by ID"
Set oRs = Conn.Execute(Sql)
Do While Not oRs.Eof
	Response.Write("<a href=""" & "list.asp?id=" & oRs("ID") & """>" & oRs("SortName") & "</a>")
	oRs.MoveNext
	If Not oRs.Eof Then Response.Write(" &gt;&gt; ")
Loop
oRs.Close
Set oRs = Nothing
End Sub


Private Sub Ok3w_Article_Class_PageTitle(SortPath)
cTmp = Left(SortPath,Len(SortPath)-1)
Sql="select ID,SortName from Ok3w_Class where ID in(" & cTmp & ") order by ID desc"
Set oRs= Conn.Execute(Sql)
Do While Not oRs.Eof
	Response.Write(oRs("SortName"))
	oRs.MoveNext
	If Not oRs.Eof Then Response.Write("_")
Loop
oRs.Close
Set oRs = Nothing
End Sub

Private Sub Ok3w_Article_IndexClassImg(ClassID,ww,hh,TopN,LeftN)
	Sql="select top " & TopN & " Id,Title,TitleColor,TitleURL,Description,PicFile from Ok3w_Article where ChannelID=1 and IsPass=1 and IsDelete=0 and IsIndexImg=1"
	If ClassID<>"" Then Sql=Sql & " and SortPath like '%," & ClassID & ",%'"
	Sql=Sql & " order by IsTop desc,IsCommend desc,AddTime desc,Id desc"
	Rs.Open Sql,Conn,0,1
	If Rs.Eof And Rs.Bof Then
	Else
		Do While Not Rs.Eof
			Response.Write("<div class=""indeclassimg"" style=""height:" & hh & "px;"">")
			If Rs("TitleURL")="" Then
				TitleURL = "./show.asp?id=" & Rs("ID") & ""
			Else
				TitleURL = Rs("TitleURL")
			End If
			Response.Write("<a href=""" & TitleURL & """ title=""" & Rs("Title") & """ target=""_blank"">")
			Response.Write("<img src=""" & Rs("PicFile") & """ width=""" & ww & """ height=""" & hh & """ border=""0"" />")
			Response.Write("<strong>" & LeftX(Rs("Title"),LeftN*2) & "</strong><br />" & OutStr(Rs("Description")))
			Response.Write("</a>")
			Response.Write("</div>")
			Rs.MoveNext
		Loop
	End If
	Rs.Close
End Sub

Private Sub Ok3w_Soft_List(ClassID,Rows,Cels,LeftN,IsCommend,DisClass,DisTime,TimeFormat,DisHits,OrderType)
Sql="select top " & Rows * Cels & " ID,SoftName,TitleColor,TitleURL,ClassID,Updatetime,Hits from Ok3w_Soft where ChannelID=3 and IsPass=1 and IsDelete=0"
If ClassID<>"" Then Sql=Sql & " and SortPath like '%," & ClassID & ",%'"
If IsCommend Then Sql=Sql & " and IsCommend=1"
Select Case OrderType
	Case "hot"
		Sql = Sql & " order by Hits desc,Updatetime desc,ID desc"
	Case "rnd"
		Randomize
		Sql = Sql & " order by Rnd(-(ID+"&Rnd()&")),ID desc"
	Case "new"
		Sql = Sql & " order by Updatetime desc,ID desc"
	Case Else
		Sql = Sql & " order by IsTop desc,IsCommend desc,Updatetime desc,ID desc"
End Select
Rs.Open Sql,Conn,0,1
Response.Write("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
Do While Not Rs.Eof
Response.Write("  <tr>")
For II=1 To  Cels
If Rs.Eof Then
Response.Write("	<td>&nbsp;</td>")
Else
Response.Write("    <td class=""newstitle""><div>")
If DisClass Then
	Response.Write("[<a href=""soft_list.asp?id=" & Rs("ClassID") & """>" & GetClassName(Rs("ClassID")) & "</a>]")
Else
	Response.Write("·")
End If
Response.Write(Format_TitleURL("soft_show.asp",Rs("ID"),Rs("SoftName"),Rs("TitleColor"),Rs("TitleURL"),LeftN))
Response.Write("</div><span>")
If DisHits Then
	Response.Write("下载:" & Rs("Hits") & " ")
End If
If DisTime Then
	Response.Write(Format_Time(Rs("Updatetime"),TimeFormat))
End If
Response.Write("</span></td>")
Rs.MoveNext
End If
If Cels<>1 And II<>Cels Then
	Response.Write("<td>&nbsp;</td>")
End If
Next
Response.Write("  </tr>")
Loop
Rs.Close 
Response.Write("</table>")
End Sub


Private Sub Ok3w_Soft_ImgFlash(ClassID,Width,Height)
Randomize
RndFun = "__Ok3w_Soft_ImgFlash__00" & Int(Rnd()*10000)
Response.Write("<script language=""javascript"">")
Response.Write("function " & RndFun & "()")
Response.Write("{")
	Sql="select top 5 ID,SoftName,Softimageurl,TitleColor,TitleURL from Ok3w_Soft where ChannelID=3 and Softimageurl<>'' and IsPlay=1 and IsDelete=0"
	If ClassID<>"" Then Sql=Sql & " and SortPath like '%," & ClassID & ",%'"
	Sql=Sql & " order by IsTop desc,IsCommend desc,Id desc"
	Rs.Open Sql,Conn,0,1
	pics = ""
	links = ""
	texts = ""
	Do While Not Rs.Eof
		ID = Rs("ID")
		Title = Rs("SoftName")
		PicFile = Htmldns & Rs("Softimageurl")
		TitleURL = Rs("TitleURL")
		If TitleURL = "" Then TitleURL = "soft_show.asp?id=" & ID & ""
		
		pics = pics & PicFile & "|"
		links = links & TitleURL & "|"
		texts = texts & Title & "|"
		Rs.MoveNext
	Loop
	Rs.Close
	If texts<>"" Then
		pics = Left(pics,Len(pics)-1)
		links = Left(links,Len(links)-1)
		texts = Left(texts,Len(texts)-1)
	End If
	
Response.Write "	var pics=""" & pics & """;"
Response.Write "	var links=""" & links & """;"
Response.Write "	var texts=""" & texts & """;"
	
Response.Write "	var focus_width=" & Width & ";"
Response.Write "	var focus_height=" & Height - 18 & ";"
Response.Write "	var text_height=18;"
Response.Write "	var swf_height = focus_height + text_height;"
	 
Response.Write "	Ok3w_insertFlash(""" & Htmldns & """, focus_width, focus_height, swf_height, text_height, pics, links, texts);"
Response.Write "}"
Response.Write RndFun & "();"
Response.Write "</script>" 
End Sub
%>

<%
Private Sub Ok3w_Soft_aList(ClassID,ImgCels,ImgW,ImgH)
	Sql="select * from Ok3w_Class where ID=" & ClassID
	Set oRs=Conn.Execute(Sql)
	IsPic = oRs("IsPic")
	PageSize = oRs("PageSize")
	If IsPic=1 Then
		Call Ok3w_Soft_pList(ClassID,PageSize,ImgCels,ImgW,ImgH)
	Else
		Call Ok3w_Soft_sList(ClassID,PageSize)
	End If
End Sub

Private Sub Ok3w_Soft_sList(ClassID,TopN)
Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
Sql="select * from Ok3w_Soft where IsPass=1 and IsDelete=0 and SortPath like '%," & ClassID & ",%' order by Updatetime desc,ID desc"
Call Page.GetRs(Conn,Rs,Sql,TopN)
i_Count = 0
Do While Not Rs.Eof And Not Page.Eof
i_Count = i_Count + 1
Response.Write("  <tr>")
Response.Write("    <td class=""slist"">" & Format_TitleURL("soft_show.asp",Rs("ID"),Rs("SoftName"),Rs("TitleColor"),Rs("TitleURL"),50) & "</td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td class=""xlist"">日期：" & Rs("Updatetime") & " 分类：" & GetClassName(Rs("ClassID")) & " 大小：" & Rs("Softsize") & Rs("Softsizeunit") & " 语言：" & Rs("Softlanguage") & " 授权：" & Rs("Softlicence") & "</td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td class=""slist""><div>" & Left(Replace(filterhtml(Rs("Softintro"))," ",""),100) & "...</div></td>")
Response.Write("  </tr>")
	Rs.MoveNext
	Page.MoveNext
Loop
Rs.Close
Response.Write("</table>")
Call Page.GetPageList()
End Sub


Private Sub Ok3w_Soft_pList(ClassID,TopN,Cels,ImgW,ImgH)
Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""8"" cellspacing=""0"">")
Sql="select ID,SoftName,Softimageurl,TitleColor,TitleURL from Ok3w_Soft where IsPass=1 and IsDelete=0 and SortPath like '%," & ClassID & ",%' order by Updatetime desc,ID desc"
Call Page.GetRs(Conn,Rs,Sql,TopN)
Do While Not Rs.Eof And Not Page.Eof
Response.Write("  <tr>")
For iCount = 1 To Cels  
Response.Write("    <td width=""" & 100\Cels & "%"" align=""center"">")
If Rs.Eof Or Page.Eof Then
	Response.Write("&nbsp;")
Else
	If Rs("TitleURL")="" Then
		aTmp = "soft_show.asp?id=" & Rs("ID") & ""
	Else
		aTmp = Rs("TitleURL")
	End If
	Response.Write("<div style=""width:" & ImgW & "px; height:" & ImgH & "px; border:1px solid #CCCCCC; margin:0px 0px 5px 0px; padding:1px;"">")
	Response.Write("      <a href=""" & aTmp & """ target=""_blank""><img src=""" & Rs("Softimageurl") & """ alt=""" & Rs("SoftName") & """ width=""" & ImgW & """ height=""" & ImgH & """ border=""0"" /></a></div>")
    Response.Write(Format_TitleURL("soft_show.asp",Rs("ID"),Rs("SoftName"),Rs("TitleColor"),Rs("TitleURL"),ImgW\16))
	Rs.MoveNext
	Page.MoveNext
End If
Response.Write("</td>")
Next	
Response.Write("  </tr>")
Loop
Rs.Close
Response.Write("</table>")
Response.Write("<div style=""border-bottom:1px dotted #CCCCCC; margin:8px 0px;""></div>")
Call Page.GetPageList()
End Sub


Private Sub Ok3w_Soft_Search(ClassID,sType,Keyword,TopN)
Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
Sql="select * from Ok3w_Soft where IsPass=1 and IsDelete=0 and ChannelID=3"
Select Case sType
	Case "Title"
		Sql = Sql & " and SoftName like '%" & Keyword & "%'"
	Case "Content"
		Sql = Sql & " and Softintro like '%" & Keyword & "%'"
	Case Else
		Sql = Sql & " and (SoftName like '%" & Keyword & "%' or Softintro like '%" & Keyword & "%')"
End Select
Sql = Sql & " order by Updatetime desc,ID desc"
Call Page.GetRs(Conn,Rs,Sql,TopN)
If Not(Rs.Eof And Rs.Bof) Then
Do While Not Rs.Eof And Not Page.Eof
Response.Write("  <tr>")
Response.Write("    <td class=""slist"">" & Format_TitleURL("soft_show.asp",Rs("ID"),Rs("SoftName"),Rs("TitleColor"),Rs("TitleURL"),50) & "</td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td class=""xlist"">日期：" & Rs("Updatetime") & " 分类：" & GetClassName(Rs("ClassID")) & " 大小：" & Rs("Softsize") & Rs("Softsizeunit") & " 语言：" & Rs("Softlanguage") & " 授权：" & Rs("Softlicence") & "</td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td class=""slist""><div>" & Left(Replace(filterhtml(Rs("Softintro"))," ",""),100) & "...</div></td>")
Response.Write("  </tr>")
	Rs.MoveNext
	Page.MoveNext
Loop
Else
Response.Write("			<tr>")
Response.Write("              <td><br /><br /><h1>Sorry!没有找到任何相关内容。</h1><br /><br /></td>")
Response.Write("            </tr>")
End If
Rs.Close
Response.Write("</table>")
Call Page.GetPageList()
End Sub


Private Sub Ok3w_Soft_Class_Nav(SortPath)
cTmp = Left(SortPath,Len(SortPath)-1)
Sql="select ID,SortName from Ok3w_Class where ID in(" & cTmp & ") order by ID"
Set oRs = Conn.Execute(Sql)
Do While Not oRs.Eof
	Response.Write("<a href=""" & "soft_list.asp?id=" & oRs("ID") & """>" & oRs("SortName") & "</a>")
	oRs.MoveNext
	If Not oRs.Eof Then Response.Write(" &gt;&gt; ")
Loop
oRs.Close
Set oRs = Nothing
End Sub


Private Sub Ok3w_Soft_Class_PageTitle(SortPath)
cTmp = Left(SortPath,Len(SortPath)-1)
Sql="select ID,SortName from Ok3w_Class where ID in(" & cTmp & ") order by ID desc"
Set oRs= Conn.Execute(Sql)
Do While Not oRs.Eof
	Response.Write(oRs("SortName"))
	oRs.MoveNext
	If Not oRs.Eof Then Response.Write("_")
Loop
oRs.Close
Set oRs = Nothing
End Sub


Private Sub Ok3w_Site_Link(TopN,LeftN,oT,cT)
StrTmp = ""
StrTmp = "<table width=""100%"" border=""0"" align=""center"" cellpadding=""5"" cellspacing=""0"">"
If oT = 1 Then
Sql = "select top " & TopN & " * from Ok3w_Ljnk where Ltype=1 and Ctype=" & cT & " order by Lorder,Lid"
Rs.Open Sql,Conn,0,1
Do While Not Rs.Eof
StrTmp = StrTmp & "  <tr>"
For link=1 To LeftN
StrTmp = StrTmp & "    <td align=""center"" width=""" & 100\LeftN & "%"">"
If Not Rs.Eof Then
StrTmp = StrTmp & "	<a href=""" & Rs("Lurl") & """ target=""_blank""><img src=""" & Rs("Lpic") & """ alt=""" & Rs("Lname") & """ width=""88"" height=""31"" border=""0"" style=""border:1px solid #CCCCCC; padding:1px;"" /></a>"
	Rs.MoveNext
Else
StrTmp = StrTmp & "	&nbsp;"
End If
StrTmp = StrTmp & "    </td>"
Next	
StrTmp = StrTmp & "  </tr>"
Loop
Rs.Close
End If
If oT=0 Then
Sql = "select top " & TopN & " * from Ok3w_Ljnk where Ltype=0 and Ctype=" & cT & "  order by Lorder,Lid"
Rs.Open Sql,Conn,0,1
Do While Not Rs.Eof
StrTmp = StrTmp & "  <tr>"
For link=1 To LeftN
StrTmp = StrTmp & "    <td align=""center"" width=""" & 100\LeftN & "%"">"
If Not Rs.Eof Then
StrTmp = StrTmp & "	<a href=""" & Rs("Lurl") & """ target=""_blank"">" & Rs("Lname") & "</a>"
Rs.MoveNext
Else
StrTmp = StrTmp & "	&nbsp;"
End If
StrTmp = StrTmp & "    </td>"
Next
StrTmp = StrTmp & "  </tr>"
Loop
Rs.Close
End If
StrTmp = StrTmp & "</table>"
Response.Write(StrTmp)
End Sub


Function Ok3w_Site_Tongji()
	Response.Write(Application("Ok3w_SiteTongji"))
End Function
%>