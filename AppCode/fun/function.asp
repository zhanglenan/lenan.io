<%
Function IsSelfRefer()
	server_v1=Lcase(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Lcase(Request.ServerVariables("SERVER_NAME"))
	If InStr(server_v1,server_v2)<1 Then
		IsSelfRefer = False'����
		Else 
			IsSelfRefer = True'��
	End If
End Function

Private Sub Page_Err(Msg)
	Response.Write(Msg)
	Response.End()
End Sub

Function myCdbl(str)
	If str = "" Or Not IsNumeric(str) Then
		Call Page_Err("��������Ҫ��Ϊ�����͡�")
		Else
			myCdbl = Cdbl(str)
	End If
End Function

Private Sub MessageBox(Msg,gotoUrl)
	Response.Write("<script language=""javascript"">")
	Response.Write("alert(""" & Msg & """);")
	If gotoUrl="" Then
		Response.Write("history.back();")
		Else
			Response.Write("document.URL='" & gotoUrl & "';")
	End If
	Response.Write("</script>")
	Response.End()
End Sub

Function LeftX(Str,N)
	Dim i,j,ch,StrTmp
	j = 0
	StrTmp = ""
	For i = 1 To Len(Str)
		ch = Mid(Str,i,1)
		StrTmp = StrTmp & ch
		If Asc(ch)<0 Then
			j = j + 2
			Else
				j = j + 1
		End If
		If j >= N Then Exit For
	Next
	LeftX = StrTmp
End Function

Function OutStr(Str)
  strer=Str
  if strer="" or isnull(strer) then
    OutStr="":exit function
  end if
  strer=replace(strer,"<","&lt;")
  strer=replace(strer,">","&gt;")
  strer=replace(strer,CHR(13) & Chr(10),"<br>")    '����
  strer=replace(strer,CHR(32),"&nbsp;")    '�ո�
  strer=replace(strer,CHR(9),"&nbsp;")    'table
  strer=replace(strer,CHR(39),"&#39;")    '������
  strer=replace(strer,CHR(34),"&quot;")    '˫����
  OutStr = strer
End Function

Function UBBCode(strer)
	If strer="" Then Exit Function
	dim re
	set re=new RegExp
	re.IgnoreCase =true
	re.Global=true
	  
	re.Pattern="(javascript)"
	strer=re.Replace(strer,"&#106avascript")
	re.Pattern="(jscript:)"
	strer=re.Replace(strer,"&#106script:")
	re.Pattern="(js:)"
	strer=re.Replace(strer,"&#106s:")
	re.Pattern="(value)"
	strer=re.Replace(strer,"&#118alue")
	re.Pattern="(about:)"
	strer=re.Replace(strer,"about&#58")
	re.Pattern="(file:)"
	strer=re.Replace(strer,"file&#58")
	re.Pattern="(document.cookie)"
	strer=re.Replace(strer,"documents&#46cookie")
	re.Pattern="(vbscript:)"
	strer=re.Replace(strer,"&#118bscript:")
	re.Pattern="(vbs:)"
	strer=re.Replace(strer,"&#118bs:")
	re.Pattern="(on(mouse|exit|error|click|key))"
	strer=re.Replace(strer,"&#111n$2")
	  
	're.Pattern="\[pic\](http|https|ftp):\/\/(.[^\[]*)\[\/pic\]"
	'strer=re.Replace(strer,"<IMG SRC='http://$2' border=0 onload=""javascript:if(this.width>screen.width-430)this.width=screen.width-430"">")
	re.Pattern="\[IMG\](http|https|ftp):\/\/(.[^\[]*)\[\/IMG\]"
	strer=re.Replace(strer,"<a href='http://$2' target=_blank><IMG SRC='http://$2' border=0 alt='�������´������ͼƬ' onload=""javascript:if(this.width>350)this.width=350""></a>")
	're.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(.[^\[]*)(\[\/FLASH\])"
	'strer= re.Replace(strer,"<a href='$4' TARGET=_blank>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE='$4'><PARAM NAME=quality VALUE=high><embed src='$4' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")
	're.Pattern="(\[FLASH\])(http://.[^\[]*(.swf))(\[\/FLASH\])"
	'strer= re.Replace(strer,"<a href=""$2"" TARGET=_blank>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")
	re.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
	strer= re.Replace(strer,"<A HREF='$2' class=blue TARGET=_blank>$2</A>")
	re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
	strer= re.Replace(strer,"<A HREF='$2' class=blue TARGET=_blank>$3</A>")
	re.Pattern="(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
	strer= re.Replace(strer,"<A HREF=""mailto:$2"">$2</A>")
	re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
	strer= re.Replace(strer,"<A HREF=""mailto:$2"" TARGET=_blank>$3</A>")
	re.Pattern = "^(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strer = re.Replace(strer,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strer = re.Replace(strer,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "([^>='])(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strer = re.Replace(strer,"$1<a target=_blank href=$2>$2</a>")
	re.Pattern = "^(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strer = re.Replace(strer,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strer = re.Replace(strer,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "[^>='](ftp://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
	re.Pattern="\[color=(.[^\[]*)\](.[^\[]*)\[\/color\]"
	strer=re.Replace(strer,"<font color=$1>$2</font>")
	re.Pattern="\[face=(.[^\[]*)\](.[^\[]*)\[\/face\]"
	strer=re.Replace(strer,"<font face=$1>$2</font>")
	re.Pattern="\[align=(.[^\[]*)\](.[^\[]*)\[\/align\]"
	strer=re.Replace(strer,"<div align=$1>$2</div>")
	re.Pattern="\[align=(.[^\[]*)\](.*)\[\/align\]"
	strer=re.Replace(strer,"<div align=$1>$2</div>")
	re.Pattern="\[center\](.[^\[]*)\[\/center\]"
	strer=re.Replace(strer,"<div align=center>$1</div>")
	re.Pattern="\[i\](.[^\[]*)\[\/i\]"
	strer=re.Replace(strer,"<i>$1</i>")
	re.Pattern="\[u\](.[^\[]*)(\[\/u\])"
	strer=re.Replace(strer,"<u>$1</u>")
	re.Pattern="\[b\](.[^\[]*)(\[\/b\])"
	strer=re.Replace(strer,"<b>$1</b>")
	re.Pattern="\[size=([1-6])\](.[^\[]*)\[\/size\]"
	strer=re.Replace(strer,"<font size=$1>$2</font>")
	  
	set re=Nothing
	UBBCode = strer
End Function

Function filterhtml(fstring)
    if isnull(fstring) or trim(fstring)="" then
        filterhtml=""
        exit function
    end if
    set  re = new  regexp
    re.ignorecase=true
    re.global=true
    re.pattern="<(.+?)>"
    fstring = re.replace(fstring, "")
    set   re=nothing
	filterhtml = fstring
End Function

Function ExecSqlReturnOneValue(Sql)
	Set opRs = Server.CreateObject("Adodb.RecordSet")
	opRs.Open Sql,Conn,0,1
	If  opRs.Eof And opRs.Bof Then 
		ExecSqlReturnOneValue = ""
		Else
			ExecSqlReturnOneValue = opRs(0)
	End If
	opRs.Close
	Set opRs = Nothing
End Function

Function Format_Time(s_Time,n_Flag)
	Dim y, m, d, h, mi, s
	Format_Time = ""
	If IsDate(s_Time) = False Then Exit Function
	y = cstr(year(s_Time))
	m = cstr(month(s_Time))
	If len(m) = 1 Then m = "0" & m
	d = cstr(day(s_Time))
	If len(d) = 1 Then d = "0" & d
	h = cstr(hour(s_Time))
	If len(h) = 1 Then h = "0" & h
	mi = cstr(minute(s_Time))
	If len(mi) = 1 Then mi = "0" & mi
	s = cstr(second(s_Time))
	If len(s) = 1 Then s = "0" & s
	Select Case n_Flag
	Case 0
		Format_Time = m & "-" & d
	Case 1
		' yyyy-mm-dd hh:mm:ss
		Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	Case 2
		' yyyy-mm-dd
		Format_Time = y & "-" & m & "-" & d
	Case 3
		' hh:mm:ss
		Format_Time = h & ":" & mi & ":" & s
	Case 4
		' yyyy��mm��dd��
		Format_Time = y & "��" & m & "��" & d & "��"
	Case 5
		' yyyymmdd
		Format_Time = y & m & d
	End Select
End Function

Function CmdSafeLikeSqlStr(Str)
	Str = Replace(Str,"'","''")
	Str = Replace(Str,"[","[[]")
	Str = Replace(Str,"%","[%]")
	Str = Replace(Str,"_","[_]")
	CmdSafeLikeSqlStr = Str
End Function

Function ReplaceTest(str,patrn, replStr)
  Dim regEx, str1 
  str1 = str
  Set regEx = New RegExp
  regEx.Pattern = patrn
  regEx.IgnoreCase = True
  regEx.global=true
  ReplaceTest = regEx.Replace(str1, replStr)
End Function

Function Format_TitleURL(pType,ID,Title,TitleColor,TitleURL,LeftN)
	oTitle = LeftX(Title,LeftN * 2)
	If TitleColor<>"" Then oTitle = "<font color=""" & TitleColor & """>" & oTitle & "</font>"
	If TitleURL="" Then
		Target = " " & Base_Target
		TitleURL = Htmldns & pType & "?id=" & ID
		Else
			Target = " target=""_blank"""
	End If
	Format_TitleURL = "<a href=""" & TitleURL & """" & Target & ">" & oTitle & "</a>"
End Function

Function GetPageUrlPath()
	SERVER_NAME = Request.ServerVariables("SERVER_NAME")
	SERVER_PORT = Request.ServerVariables("SERVER_PORT")
	PATH_INFO = Request.ServerVariables("PATH_INFO")
	PATH_TMP = Split(PATH_INFO,"/")
	PATH_INFO = Replace(PATH_INFO,PATH_TMP(Ubound(PATH_TMP)),"")
	URL = "http://" & SERVER_NAME
	If SERVER_PORT<>80 Then URL = URL & ":" & SERVER_PORT
	URL = URL & PATH_INFO
	GetPageUrlPath = URL
End Function

Function GetAdSense(SN)
	Response.Write("<script language=""javascript"" src=""" & Htmldns & "js/ok3w_" & SN & ".js""></script>")
End Function

'Ƶ������
Function GetChannelName(ChannelId)
	If ChannelId="" Then
		GetChannelName = "--------"
		Exit Function
	End If
	Sql = "select ChannelName from Ok3w_Channel where ChannelId=" & ChannelId
	GetChannelName = ExecSqlReturnOneValue(Sql)
End Function

'��Ŀ����
Function GetClassName(ClassId)
	If ClassId="0" Then
		GetClassName = "--------"
		Exit Function
	End If
	Sql = "select SortName from Ok3w_Class where ID=" & ClassId
	GetClassName = ExecSqlReturnOneValue(Sql)
End Function

Function GetSortPath(ClassId)
	If ClassId = "" Then
		GetSortPath = ""
	Else
		GetSortPath = Conn.Execute("select SortPath from Ok3w_Class where ID=" & ClassID)(0)
	End If
End Function 

'����Ŀ����
Function GetParentClassName(ClassId)
	Sql="select ParentID,SortName from Ok3w_Class where ID=" & ClassId
	Set oRs = Conn.Execute(Sql)
	a_ParentID = oRs(0)
	a_SortName = oRs(1)
	oRs.Close
	If a_ParentID <> 0 Then
		Sql="select ID,SortName from Ok3w_Class where ID=" & a_ParentID
		Set oRs = Conn.Execute(Sql)
		a_ParentID = oRs(0)
		a_SortName = oRs(1)
		oRs.Close
	End If
	Set oRs = Nothing
	GetParentClassName = a_SortName
End Function

Function IsHaveNextClass(ClassID)
	Sql="select ID from Ok3w_Class where ParentID=" & ClassID & " and IsNav=1"
	Set oRs = Conn.Execute(Sql)
	If oRs.Eof And oRs.Bof Then
		Sql = "select ParentID from Ok3w_Class where ID=" & ClassId
		a_ParentID = ExecSqlReturnOneValue(Sql)
		If a_ParentID<>0 Then
			oRs.Close
			Sql="select ID from Ok3w_Class where ParentID=" & a_ParentID & " and IsNav=1"
			Set oRs = Conn.Execute(Sql)
			If oRs.Eof And oRs.Bof Then
				IsHaveNextClass = False
			Else
				IsHaveNextClass = True
			End If
		Else
			IsHaveNextClass = False
		End If
	Else
		IsHaveNextClass = True
	End If
	oRs.Close
	Set oRs = Nothing
End Function

Function GetCommentsCount(TypeID,TableID)
	Sql="select count(ID) from Ok3w_Guest where TypeID=" & TypeID & " and TableID=" & TableID
	If Application("Ok3w_SiteIsGuest")="1" Then Sql = Sql & " and IsPass=1"
	GetCommentsCount = ExecSqlReturnOneValue(Sql)
End Function

'����ID�����б�ѡ��˵�
Private Sub InitClassSelectOption(ChannelId,ParentID,ChkID)
	Dim opRs,cTmp,cLen,cCount
	Set opRs = Server.CreateObject("Adodb.RecordSet")
	Sql = "select ID,SortName,SortPath from Ok3w_Class where ChannelId=" & ChannelId & " and ParentID=" & ParentID & " and gotoURL='' order by OrderID"
	opRs.Open Sql,Conn,0,1
	Do While Not opRs.Eof
		Response.Write("<option value=""" & opRs("ID") & """")
		If ChkID = opRs("ID") Then Response.Write(" selected=""selected""")
		Response.Write(">")
		cTmp = Split(opRs("SortPath"),",")
		cLen = Ubound(cTmp) - 2
		For cCount=1 To cLen
			Response.Write("��&nbsp;")
		Next
		Response.Write("��" & opRs("SortName") & "</option>")
		
		Call InitClassSelectOption(ChannelId,opRs("ID"),ChkID)
		opRs.MoveNext
	Loop
	opRs.Close
	Set opRs = Nothing
End Sub

Function ReplaceUpFilePath(uPath,aStr,bStr)
	If Left(uPath,7)="http://" Then
		ReplaceUpFilePath = uPath
	Else
		ReplaceUpFilePath = Replace(uPath,aStr,bStr)
	End If
End Function

Private Sub OutThisPageContent(aID,Content,PageType)
	thisPage = Request.QueryString("thisPage")
	If thisPage<>"" Then thisPage = myCdbl(thisPage)
	If thisPage="" Then thisPage=1
	thisPage = Cint(thisPage)
	
	If InStr(Content,"[Ok3w_NextPage]")>0 Then
		Content_Tmp = Split(Content,"[Ok3w_NextPage]")
		Page_Count = Ubound(Content_Tmp)+1
		If thisPage> Page_Count Then thisPage = Page_Count
		OutContent = Content_Tmp(thisPage-1)
	Else
		OutContent = Content
	End If
	
	OutContent = ReplaceTest(OutContent,"<img ","<img style=""cursor:pointer;"" onclick=""ImageOpen(this)"" onload=""ImageZoom(this,560,700)"" ")
	OutContent = ReplaceTest(OutContent,"""upfiles/","""" & Htmldns & "upfiles/")
	OutContent = ReplaceTest(OutContent,"""editor/","""" & Htmldns & "editor/")
	
	If Trim(Application("Ok3w_SitePublicKeyWords"))<>"" And Application("Ok3w_SitePublicKeyWords")<>"0" Then
		kTmp = Split(Application("Ok3w_SitePublicKeyWords"),vbCrLf)
		For kk=0 To Ubound(kTmp)
			uTmp = Split(kTmp(kk),"|")
			If Ubound(uTmp)=1 Then
				OutContent = ReplaceTest(OutContent,uTmp(0),"<span class=keyword><a href=" & uTmp(1) & " target=_blank>" & uTmp(0) & "<a></span>")
			End If
		Next
	End If
	
	Response.Write(OutContent)
	
	If Page_Count>1 Then
		Response.Write("<span></span><div class=""thisPageNav"">")
		For iPage=1 To Page_Count
			If iPage = 1 Then
				URL = "?id=" & aID
			Else
				URL = "?id=" & aID & "&thisPage=" & iPage
			End If
			If iPage = thisPage Then
				Response.Write("<strong style=""color:#FF0000;"">��" & iPage & "ҳ</strong> ")
				Else
					Response.Write("<a href=""" & URL & """>��" & iPage & "ҳ</a> ")
			End If
		Next
		Response.Write("</div>")
	End If
End Sub

Function GetUserDengjiValue(id,dj)
	SiteUserDengji = Application("Ok3w_SiteUserDengji")
	If SiteUserDengji<>"" Then
		DJ1 = Split(SiteUserDengji,"|")
		DJ2 = Split(DJ1(id),",")
		GetUserDengjiValue = DJ2(dj-1)
	End If
End Function

Function GetUserDengjiID(Jifen)
	SiteUserDengji = Application("Ok3w_SiteUserDengji")
	DJ1 = Split(SiteUserDengji,"|")
	DJ2 = Split(DJ1(1),",")
	Max_ADJ = Ubound(DJ2)
	For dj_item = Max_ADJ To 0 Step -1
		If DJ2(dj_item)<>"" Then
			If Jifen>=Cdbl(DJ2(dj_item)) Then
				GetUserDengjiID = dj_item + 1
				Exit Function
			End If
		End If
	Next
End Function

Private Sub InitUserDengjiSelectOption(dj)
	SiteUserDengji = Application("Ok3w_SiteUserDengji")
	DJ_Tmp = Split(SiteUserDengji,"|")
	dj_Name = Split(DJ_Tmp(0),",")
	For dj_item = 0 To Ubound(dj_Name)
		If dj_Name(dj_item)="" Then Exit Sub
		Response.Write("<option value=""" & dj_item + 1 & """")
		If Cint(dj) = dj_item + 1 Then Response.Write(" selected=""selected""")
		Response.Write(">" & dj_Name(dj_item) & "</option>")
	Next
End Sub

Private Sub InitSoftBaseChkItem(ChkStr,objName,objType)
	ItemSTr = Application("Ok3w_Site" & objName)
	ItmpTmp = Split(ItemSTr,"|")
	Select Case objType
		Case "checkbox"
			For II=0 To Ubound(ItmpTmp)
				Response.Write("<input name=""" & objName & """ type=""" & objType & """ value=""" & ItmpTmp(II) & """")
				If InStr("," & ChkStr & "," , "," & Trim(ItmpTmp(II)) & ",")>0 Then	Response.Write(" checked")
				Response.Write(">" & ItmpTmp(II) & " ")
			Next
		Case "select"
			For II=0 To Ubound(ItmpTmp)
				Response.Write("<option value=""" & ItmpTmp(II) & """")
				If Trim(ItmpTmp(II)) = ChkStr Then	Response.Write(" selected")
				Response.Write(">" & ItmpTmp(II) & "</option>")
			Next
	End Select
End Sub

Private Sub SoftstarImg(star)
	For II=1 To star
		'Response.Write("<img src=""" & Htmldns & "images/star.gif"" />")
		Response.Write("��")
	Next
End Sub
%>
