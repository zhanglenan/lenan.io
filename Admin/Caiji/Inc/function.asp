<%
'*************************************************************************************
	'������:SK_Class_Channel_add
	'��  ��:Ƶ������--����Ƶ��
'*************************************************************************************
	Function SK_Class_Channel_add()
		Dim cTmp,Rs1,Rs2
		Call OpenConn()
		Set Rs1 = Server.CreateObject("adodb.recordset")
		Set Rs2 = Server.CreateObject("adodb.recordset")
		ConnItem.Execute("delete from SK_Class")
		Rs1.Open "select * from Ok3w_Class where ChannelID=1 and gotoURL=''",Conn,1,1
		Rs2.Open "select * from SK_Class",ConnItem,1,3
		Do While Not Rs1.Eof
			Rs2.AddNew()
			cTmp					= Split(Rs1("SortPath"),",")
			Rs2("ChannelID")		= 1
			Rs2("ClassID")			= Rs1("ID")
			Rs2("className")		= Rs1("SortName")
			Rs2("ParentID")			= Rs1("ParentID")
			Rs2("ParentPath")		= Rs1("SortPath")
			Rs2("depth")			= Ubound(cTmp)-2
			Rs2("OrderID")			= Rs1("OrderID")
			Rs2.Update
			Rs1.MoveNext
		Loop
		Rs1.Close
		Rs2.Close
		
		Call CloseConn()
		
		Response.Redirect("sk_class.asp?Colleclx="& Colleclx)
	End Function

'*************************************************************************************
	'������:DisClassTable
	'��  ��:�ݹ���ʾ���з���
'*************************************************************************************
Private Sub DisClassTable(ChannelId,ParentID)
	Dim opRs,cTmp,cLen,cCount
	Set opRs = Server.CreateObject("Adodb.RecordSet")
	Sql = "select ClassID,ClassName,ParentPath from SK_Class where ChannelId=" & ChannelId & " and ParentID=" & ParentID & " order by OrderID"
	opRs.Open Sql,ConnItem,0,1
	Do While Not opRs.Eof
		Response.Write("<TR class=""tdbg"">")
        Response.Write("<TD bgcolor=""#F2F2F2"" class=tdbg>" & opRs("ClassID") & "</TD>")
        Response.Write("<TD bgcolor=""#F2F2F2"" class=tdbg>")
		cTmp = Split(opRs("ParentPath"),",")
		cLen = Ubound(cTmp) - 2
		For cCount=1 To cLen
			Response.Write("��&nbsp;")
		Next
		Response.Write("��" & opRs("ClassName"))
		Response.Write("</TD>")
        Response.Write("</TR>")		
		Call DisClassTable(ChannelId,opRs("ClassID"))
		opRs.MoveNext
	Loop
	opRs.Close
	Set opRs = Nothing
End Sub	

'*************************************************************************************
	'������:InitClassSelectOption
	'��  ��:����ID�����б�ѡ��˵�
'*************************************************************************************
Private Sub InitClassSelectOption(ChannelId,ParentID,ChkID)
	Dim opRs,cTmp,cLen,cCount
	Set opRs = Server.CreateObject("Adodb.RecordSet")
	Sql = "select ClassID,ClassName,ParentPath from SK_Class where ChannelId=" & ChannelId & " and ParentID=" & ParentID & " order by OrderID"
	opRs.Open Sql,ConnItem,0,1
	Do While Not opRs.Eof
		Response.Write("<option value=""" & opRs("ClassID") & """")
		If Cint(ChkID) = opRs("ClassID") Then Response.Write(" selected=""selected""")
		Response.Write(">")
		cTmp = Split(opRs("ParentPath"),",")
		cLen = Ubound(cTmp) - 2
		For cCount=1 To cLen
			Response.Write("��&nbsp;")
		Next
		Response.Write("��" & opRs("ClassName") & "</option>")
		
		Call InitClassSelectOption(ChannelId,opRs("ClassID"),ChkID)
		opRs.MoveNext
	Loop
	opRs.Close
	Set opRs = Nothing
End Sub

'*************************************************************************************
	'������:SK_Class_Channel_edit
	'��  ��:Ƶ������--�޸�Ƶ��
'*************************************************************************************
	Function SK_Class_Channel_edit(className,ParentPath,classID)
		dim sql
		if className="" or ParentPath="" then
			ErrMsg="<font color=red>����д������</font>"
		Else
			sql = "select top 1 * from SK_class where classID=" & classID 
			Set Rs = Server.CreateObject("adodb.recordset")
			Rs.Open SQL, ConnItem, 1,3
			Rs("className")=className
			rs("ParentPath")=ParentPath
			rs.update
			set rs=nothing
			response.redirect("sk_class.asp?Colleclx="& Colleclx)
			response.end
		end if	
		SK_Class_Channel_edit=ErrMsg
	End Function
'*************************************************************************************
'������:SK_Class_small_add
'��  ��:Ƶ������--������Ŀ
'*************************************************************************************
	Function SK_Class_small_add(ClassName,ParentPath,ParentID)
		dim sql,Ordermax
		if ClassName="" or ParentPath="" then
		'Response.Write  ClassName & ParentPath & ("asdfsadf") & ParentID
				ErrMsg="<font color=red>����д������</font>"
		Else
			set rs=ConnItem.execute("select top 1 * from SK_class where classID=" & ParentID)
			Ordermax=rs("OrderID")+1
			depth=rs("depth")+1			
			
			Set Rs=ConnItem.execute("select * from SK_class where OrderID >=" & Ordermax)
			while not rs.eof
			Response.Write rs("classid")
			ConnItem.execute("update SK_class set OrderID="& rs("OrderID")+1 &" where classid="& rs("classid") )
			rs.movenext
			wend
			
			sql = "select top 1 * from SK_class order by ClassID desc"
			Set Rs = Server.CreateObject("adodb.recordset")
			Rs.Open SQL, ConnItem, 1,3
			rs.addnew
			if Colleclx<>0 then
				rs("ChannelID")=Colleclx
			Else
				rs("ChannelID")=0
			end if
			Rs("className")=className
			rs("ParentPath")=ParentPath
			rs("ParentID")=ParentID
			rs("OrderID")=Ordermax
			rs("depth")=depth
			rs.update
			set rs=nothing
			response.redirect("sk_class.asp?Colleclx="& Colleclx)
			response.end
				
				if ConnItem.execute("select count(ClassID) from sk_Class where ClassDir='"& ClassDir &"'")(0)>0  then
				ErrMsg="<font color=red>Ŀ¼�Դ��ڣ����ñ��Ŀ¼</font>"
					Else
					sql = "insert into sk_Class (ChannelID,ClassName,Readme,ClassDir,ParentPath) values ('" & ChannelID & "','" & ClassName &"','" & Readme &"','" & ClassDir &"','" & ClassDir &"')"
					
					ConnItem.execute(sql)
					response.redirect("sk_smallclass.asp?ChannelID=" & ChannelID)
					response.end
	  			end if
		end if		
		SK_Class_small_add=ErrMsg
	End Function
'==================================================
'��������SK_Showclass
'��  �ã���ʾƵ����Ŀ����
'==================================================	
sub SK_Showclass(FolderID,ChannelID) 
if ChannelID=0 then
Response.Write"<option selected value='555555555'>5555555555555555</option>"

Else
			Dim RS,FolderName,TreeStr,ID		
			Set  RS=Server.CreateObject("ADODB.Recordset")
			FolderID = Trim(FolderID)
			If Not IsNumeric(ChannelID) Then Return
			RS.Open ("select ID,FolderName from KS_Class Where ChannelID=" & ChannelID & " AND tj=1 Order BY FolderOrder ASC"),conn, 1, 1
			Do While Not RS.EOF
			 ID = Trim(RS(0))
			 FolderName = Trim(RS(1))
				 If FolderID = ID Then
				   TreeStr = TreeStr & "<option selected value='" & ID & "'>" & FolderName & "</option>"
				 Else
				  TreeStr = TreeStr & "<option value='" & ID & "'>" & FolderName & " </option>"
				 End If
			  TreeStr = TreeStr & ReturnSubList(ID, FolderID)
			RS.MoveNext
		   Loop
		   RS.Close:Set RS = Nothing
		   if TreeStr="" then
		   Response.Write "<option value='0'> ��û����� </option>"
		   Else
		   Response.Write TreeStr
		   end if
		   
		   'KSCache.add ReturnTree,dateadd("n",1000000,now)
end if
end sub
'**************************************************
'��������ReturnSubClass
'��  �ã����Ҳ����������ɼ���ĿID��
'��  ����ParentID ----���ڵ�ID,   K ----ѡ����ID
'����ֵ������
'************************************************** 
Public Function ReturnSubClass(ParentID,k)
	Dim  SubRS, SpaceStr, Num,FolderName,ID,RS
	'-----------------------------------------
	IF k=1 then
		Set RS = Server.CreateObject("ADODB.RECORDSET")
		RS.Open ("Select * From Item Where  ClassID ='"& ParentID &"' order by ItemID DESC" ), connItem, 1, 1
		Do while Not Rs.Eof
			SpaceStr= SpaceStr & RS("ItemID") & ","
			Rs.Movenext
		Loop
		Rs.Close : Set Rs=Nothing
		k=0
	End If
	'-----------------------------------------
	Set SubRS = Server.CreateObject("ADODB.RECORDSET")
	SubRS.Open ("Select ClassID,ClassName,ParentID from SK_Class Where  ParentID =" & ParentID & " "), connItem, 1, 1
	Do While Not SubRS.EOF
		ID = Trim(SubRS(0))
		'-------------------------------------
		Set RS = Server.CreateObject("ADODB.RECORDSET")
		RS.Open ("Select * From Item Where  ClassID ='"& ID &"' order by ItemID DESC" ), connItem, 1, 1
		Do while Not Rs.Eof
			SpaceStr= SpaceStr & RS("ItemID") & ","
			Rs.Movenext
		Loop
		Rs.Close : Set Rs=Nothing
		'-------------------------------------
		SpaceStr= SpaceStr & ReturnSubClass(ID,k)
		SubRS.MoveNext
	Loop 
	SubRS.Close : Set SubRS=Nothing
	ReturnSubClass=SpaceStr
End Function 
'**************************************************
'��������ReturnSubList
'��  �ã����Ҳ����������ݡ�
'��  ����ParentID ----���ڵ�ID,   FolderID ----ѡ����ID
'����ֵ������
'**************************************************
Public Function ReturnSubList(ParentID, FolderID)
	  Dim SubTypeList, SubRS, SpaceStr, k, Total, Num,FolderName, ID,TJ
	  Set SubRS = Server.CreateObject("ADODB.RECORDSET")
	  SubRS.Open ("Select count(ID) AS total from KS_Class Where  TN='" & ParentID & "'"), conn, 1, 1
	  Total = SubRS("Total")
	  SubRS.Close
	  SubRS.Open ("Select ID,FolderName,TJ from KS_Class Where  TN='" & ParentID & "' Order BY FolderOrder ASC"), conn, 1, 1
	  Num = 0
	  Do While Not SubRS.EOF
	   Num = Num + 1
	   SpaceStr = ""
		TJ = CInt(SubRS(2))
		For k = 1 To TJ - 1
		  If k = 1 And k <> TJ - 1 Then
		  SpaceStr = SpaceStr & "&nbsp;&nbsp;��"
		  ElseIf k = TJ - 1 Then
			If Num = Total Then
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;�� "
			Else
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;�� "
			End If
		  Else
		   SpaceStr = SpaceStr & "&nbsp;&nbsp;��"
		  End If
		Next
	  ID = Trim(SubRS(0))
	  FolderName = Trim(SubRS(1))
	  If FolderID = ID Then
	   SubTypeList = SubTypeList & "<option selected value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  Else
	   SubTypeList = SubTypeList & "<option value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  End If
	   SubTypeList = SubTypeList & ReturnSubList(ID, FolderID)
	  SubRS.MoveNext
	 Loop
	  SubRS.Close:Set SubRS = Nothing
	  ReturnSubList = SubTypeList
End Function
'*************************************************************************************
	'������:GetFileID
	'��  ��:�����ļ�ID��,6λ���+�ļ�
	'lx=�ɼ�����
	'��  ��:��
'*************************************************************************************
	Function GetFileID(dir,filename,lx)
		  Dim RSC,TempUrlArray
          Set RSC=Server.CreateObject("ADODB.RECORDSET")
			 '6λ���+�ļ�
				Do While True
				 GetFileID = dir + MakeRandom(6) + filename
				 select case lx
				 case 1
				 'RSC.Open "Select all from sk_photo Where PicUrls='" & GetFileID & "'", ConnItem, 1, 1
				 case 2
				 'RSC.Open "Select PicUrls from sk_photo Where PicUrls='" & GetFileID & "'", ConnItem, 1, 1
				 case 3
				 RSC.Open "Select PicUrls from sk_photo Where PicUrls LIKE '%" & GetFileID & "%'", ConnItem, 1, 1
				 case 6
				 RSC.Open "Select FileUrls from sk_all Where PicUrls LIKE '%" & GetFileID & "%'", ConnItem, 1, 1
				 end select
				 
				  If RSC.EOF And RSC.BOF Then
				  Exit Do
				  End If
				Loop
		RSC.Close
		Set RSC = Nothing
	End Function
'*************************************************************************************
	'������:GetClassID
	'��  ��:������Ŀ¼��Ƶ����ID��
	'��  ��:��
'*************************************************************************************
	Function GetClassID()
		  Dim RSC
          Set RSC=Server.CreateObject("ADODB.RECORDSET")
			 '����Ŀ¼ID ��+10λ���
				Do While True
				 GetClassID = Year(Now()) & MakeRandom(10)
				 RSC.Open "Select ID from KS_Class Where ID='" & GetClassID & "'", ConnItem, 1, 1
				  If RSC.EOF And RSC.BOF Then
				  Exit Do
				  End If
				Loop
		RSC.Close
		Set RSC = Nothing
	End Function
	'**************************************************
	'��������GetFileName
	'��  �ã������ļ�����
	'��  ����ArticleFsoType  ----��������
	'        addDate   -----����ʱ��,GetFileNameType--��չ��
	'**************************************************
	Public Function GetFileName(ArticleFsoType, AddDate, GetFileNameType)
		Dim N
		Randomize
	 Select Case ArticleFsoType
	  Case 1
	  GetFileName = Year(AddDate) & "/" & Month(AddDate) & "-" & Day(AddDate) & "/" & GetFileNameType  '��/��-��/�����+��չ��
	  Case 2
	  GetFileName = Year(AddDate) & "/" & Month(AddDate) & "/" & Day(AddDate) & "/" & GetFileNameType '��/��/��/�����+��չ��
	  Case 3
	  GetFileName = Year(AddDate) & "-" & Month(AddDate) & "-" & Day(AddDate) & "/" & GetFileNameType '��-��-��/�����+��չ��
	  Case 4
	  GetFileName = Year(AddDate) & "/" & Month(AddDate) & "/" & GetFileNameType '��/��/�����+��չ��
	  Case 5
	  GetFileName = Year(AddDate) & "-" & Month(AddDate) & "/"  & GetFileNameType '��-��/�����+��չ��
	  Case 6
	  GetFileName = Year(AddDate) & Month(AddDate) & Day(AddDate) & "/"  & GetFileNameType '������/�����+��չ��
	  Case 7
	  GetFileName = Year(AddDate) & "/"  & GetFileNameType '��/�����+��չ��
	  Case 8
	  GetFileName = Year(AddDate) & Month(AddDate) & Day(AddDate)  & GetFileNameType '��+��+��+�����+��չ��
	  Case 9
	  GetFileName =  GetFileNameType
	  Case 10
	  GetFileName =  GetFileNameType '����ַ�
	  Case Else
	   GetFileName = Year(AddDate) & Month(AddDate) & Day(AddDate) & GetFileNameType '12λ�����+��չ��
	End Select
	End Function
	'*************************************************************************************
	'������:GetInfoID_CMS
	'��  ��:��������,ͼƬ�����صȵ�ΨһID
	'��  ��:ChannelID--Ƶ��ID
	'*************************************************************************************
	Public Function GetInfoID_CMS(ChannelID)
	   On Error Resume Next
	   Dim RSC, TableNameStr
       Set RSC=Server.CreateObject("ADODB.RECORDSET")
	   Select Case ChannelID
		 Case 1
		   TableNameStr = "Select NewsID From KS_Article Where NewsID='"
		 Case 2
		   TableNameStr = "Select PicID From KS_Photo Where PicID='"
		 Case 3
		   TableNameStr = "Select DownID From KS_DownLoad Where DownID='"
		 Case 4
		   TableNameStr = "Select FlashID From KS_Flash Where FlashID='"
	   End Select
	   
	   Do While True
		 GetInfoID_CMS = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Now(), "-", ""), " ", ""), ":", ""), "PM", ""), "AM", ""), "����", ""), "����", "") & MakeRandom(3)
			RSC.Open TableNameStr & GetInfoID_CMS & "'", conn, 1, 1
				  If RSC.EOF And RSC.BOF Then Exit Do
	   Loop
		RSC.Close:Set RSC = Nothing
	End Function
	'*************************************************************************************
	'������:GetInfoID
	'��  ��:��������,ͼƬ�����صȵ�ΨһID
	'��  ��:ChannelID--Ƶ��ID
	'*************************************************************************************
	Function GetInfoID(ChannelID)
	   On Error Resume Next
	   Dim RSC, TableNameStr
       Set RSC=Server.CreateObject("ADODB.RECORDSET")
	   Select Case ChannelID
		 Case 1
		    TableNameStr = "Select ArticleID From SK_Article Where ArticleID='"
		 Case 2
		    TableNameStr = "Select FlashID From sk_Flash Where FlashID='"
		 Case 3
			TableNameStr = "Select PicID From sk_Photo Where PicID='"
	   End Select
	   Do While True
		 GetInfoID = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Now(), "-", ""), " ", ""), ":", ""), "PM", ""), "AM", ""), "����", ""), "����", "") & MakeRandom(3)
			RSC.Open TableNameStr & GetInfoID & "'", ConnItem, 1, 1
				If RSC.EOF And RSC.BOF Then
					Exit Do
				End If
	   Loop
		RSC.Close
		Set RSC = Nothing
	End Function
'=============================================================================
'��  ���� Tvar  ----����   sType---ѡ����ʾ����
'����ֵ���ɹ����������
'��  �ã�ʱ���ʽ����, 
'==============================================================================
	Public Function Format_Time(Tvar,sType)
        dim Tt,sYear,sMonth,sDay,sHour,sMinute,sSecond
        If Not IsDate(Tvar) Then Format_Time = "" : Exit Function
        Tt			= Tvar
        sYear		= Year(Tt)
        sMonth		= Right("0" & Month(Tt),2)
		sDay		= Right("0" & Day(Tt),2)
		sHour		= Right("0" & Hour(Tt),2)
		sMinute		= Right("0" & Minute(Tt),2)
		sSecond		= Right("0" & Second(Tt),2)
        Select Case sType
        Case 1	'2005-10-01 23:45:45
			Format_Time = sYear & "-" & sMonth & "-" & sDay & " " & sHour & ":" & sMinute & ":" & sSecond
        Case 2	'��-��-�� ʱ:��:��
			Format_Time = sYear & "��" & sMonth & "��" & sDay & "�� " & sHour & "ʱ" & sMinute & "��" & sSecond & "��"
        Case 3	'10-01 23:45
			Format_Time = sMonth & "-" & sDay & " " & sHour & ":" & sMinute
        Case 4	'2005-10-01
			Format_Time = sYear & "-" & sMonth & "-" & sDay
        Case 5	'2005��10��01��
			Format_Time = sYear & "��" & sMonth & "��" & sDay & "��"
        Case 6	'10-01
			Format_Time = sMonth & "-" & sDay
        Case 7	'20051001234545
			Format_Time = sYear & sMonth & sDay & sHour & sMinute & sSecond		
		Case 8	'20051001234545
			Format_Time = sYear & sMonth & sDay & sHour & sMinute
        Case Else
			Format_Time = Tt
        End Select
    End Function
'===================================================
'С��ͨ�òɼ�ϵͳ
'��������MakeRandom
'��  �ã�����ָ��λ���������
'��  ���� maxLen  ----����λ��
'����ֵ���ɹ����������
'===================================================
	Function MakeRandom(ByVal maxLen)
	
	  Dim strNewPass
	  Dim whatsNext, upper, lower, intCounter
	  Randomize
	
	 For intCounter = 1 To maxLen
	   upper = 57
	   lower = 48
	   strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	 Next
	   MakeRandom = strNewPass
	End Function
'===================================================
'С��ͨ�òɼ�ϵͳ
'��������km_template()
'��  �ã���ʾ��Ŀ
'===================================================
sub km_template()
	 set rs= ConnItem.execute("Select * From ks_template")
	 if rs.eof then
	 	response.write "���޷���"
	 Else 
	
	 	Response.Write "<table width=100% border=0 cellspacing=0 cellpadding=0>"
		Response.Write "<tr><td width='40%'>ID��</td><td width='60%'>����</td></tr>"
		do while not rs.eof
		 if rs("ChannelID")<>0 then		
			Response.Write "<tr height='22'><td><input name='km_asdfsf' type='text'value="&rs("TemplateID")&" size='15' /></td><td>&nbsp;"&rs("TemplateName")&"--ģ��</td></tr>"
		  end if
		rs.movenext
		loop
		Response.Write "</table>"
	 end if
	 rs.close
	 set rs=nothing
end sub	
'==================================================
'С��ͨ�òɼ�ϵͳ
'��������km_class()
'��  �ã���ʾ���Ŀ����Ŀ
'==================================================
sub km_class()
	 set rs= ConnItem.execute("Select * From ks_class")
	 if rs.eof then
	 	response.write "���޷���"
	 Else 
	
	 	Response.Write "<table width=100% border=0 cellspacing=0 cellpadding=0>"
		Response.Write "<tr><td width='40%'>ID��</td><td width='60%'>����</td></tr>"
		do while not rs.eof
		 if rs("tn")=0 then
		Response.Write "<tr height='22'><td><font color='#FF0000'>"&rs("id")&"</font></td><td><font color='#FF0000'>&nbsp;"&rs("FolderName")&"--Ƶ��</font></td></tr>"
			Else
			Response.Write "<tr height='22'><td><input name='km_asdfsf' type='text'value="&rs("id")&" size='15' /></td><td>&nbsp;"&rs("FolderName")&"--��Ŀ</td></tr>"
		  end if
		rs.movenext
		loop
		Response.Write "</table>"
	 end if
	 rs.close
	 set rs=nothing
end sub	 

'==================================================
'С��ͨ�òɼ�ϵͳ
'��������sk_dir_get()
'��  �ã���ȡĿ¼
'==================================================
Function sk_dir_get(ClassID,lx)
Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,i,PathTemp,SaveTf
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
'strtemp = strtemp & strInstallDir 

set rs = ConnItem.execute("select top 1 ArticleDir,flashdir,photoDir from SK_Config")
if lx =1 then strtemp = strtemp & rs("ArticleDir")
if lx =2 then strtemp = strtemp  & rs("flashdir")
if lx =3 then strtemp = strtemp  & rs("photoDir") 
rs.close
set rs = ConnItem.execute("select top 1 * from SK_Class where ClassID=" & ClassID)
strtemp = strtemp & rs("ClassDir") &"/"
sk_dir_get = strtemp
rs.close
set rs=nothing
end function
'==================================================
'С��ͨ�òɼ�ϵͳ
'��������sk_dir()
'PicUrls=Զ���ļ���ַ
'��  �ã���������Ŀ¼
'==================================================
Function sk_dir(ClassID,lx,FileUrl)
Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,i,PathTemp,SaveTf
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
strtemp = strtemp & strInstallDir 

set rs = ConnItem.execute("select top 1 ArticleDir,flashdir,photoDir from SK_Config")
if lx =1 then strtemp = strtemp & rs("ArticleDir")
if lx =2 then strtemp = strtemp  & rs("flashdir")
if lx =3 then strtemp = strtemp  & rs("photoDir") 
rs.close
set rs = ConnItem.execute("select top 1 * from SK_Class where ClassID=" & ClassID)
strtemp = strtemp & rs("ClassDir") &"/"
Call SaveRemoteFile(strtemp,FileUrl)'����Զ���ļ�
sk_dir = strtemp
rs.close
set rs=nothing
end function
'==================================================
'�ɼ�ϵͳ
'��������Sk_GetSaveDir()
'lx=����1=���� 2=flash 3=ͼƬ 4=���� 5=���� 6=�Զ�
'��  �ã���ȡ�ļ�����Ŀ¼����
'==================================================
Function Sk_GetSaveDir(lx)
	Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,Tempi,PathTemp,SaveTf,TempUrlArray,Ranfilestr
	strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
	strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
	set rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=" & lx)
	strtemp = strtemp  & rs("Dir")
	Sk_GetSaveDir = strtemp & SaveFileUrl
	rs.close
	set rs=nothing
end function
'==================================================
'С��ͨ�òɼ�ϵͳ
'��������Sk_SaveFile()
'Lx=Ƶ��
'FileUrl=Զ���ļ���ַ
'��  �ã�������Ƶ������Զ���ļ��滻��ַ
'==================================================
Function Sk_SaveFile(lx,FileUrl)
Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,Tempi,PathTemp,SaveTf,TempUrlArray,Ranfilestr,Ranfilestr1
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
'strtemp = strtemp & strInstallDir 
FileUrl=replace(replace(FileUrl,"""","")," ","")
Select Case lx
Case 1
	Set rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=1")
Case 2

Case 3
	Set rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=2")
Case 5
	Set rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=3")
Case 6
End Select
strtemp = strtemp & rs("Dir")
strtemp = strtemp & SaveFileUrl
      Arr_Path=Split(strtemp,"/")
      PathTemp=""
      For Tempi=0 To Ubound(Arr_Path)
         If Tempi=0 Then
            PathTemp=Arr_Path(0) & "/"
         ElseIf Tempi=Ubound(Arr_Path) Then
            Exit For
         Else
            PathTemp=PathTemp & Arr_Path(Tempi) & "/"
         End If
         If CheckDir(PathTemp)=False Then
            If MakeNewsDir(PathTemp)=False Then
               SaveTf=False
               Exit For
            End If
         End If
      Next
TempUrlArray=Split(FileUrl,"/")  
Ranfilestr=GetFileID(strtemp,TempUrlArray(Ubound(TempUrlArray)),3)'�����ļ���
'Call SaveRemoteFile(Ranfilestr,FileUrl)'����Զ���ļ�
If SaveRemoteFile(Ranfilestr,FileUrl)<>False then'����Զ���ļ�
Ranfilestr1=Ranfilestr
if Thumb_WaterMark=1 then call SKThumb.AddWaterMark(Ranfilestr)'ˮӡ
Sk_SaveFile = Ranfilestr1
Else
Sk_SaveFile = False
End if
Rs.close
Set Rs=nothing
End function
'==================================================
'С��ͨ�òɼ�ϵͳ
'��������SaveFile()
'��  �ã�Զ�̱���-
'==================================================
Sub SaveFile()
dim Savelx,ChannelID,ClassID,id,FoundErr,ErrMsg,sql,rs,strChannelDir,pici,picii,SaveErr,TempArray,i,lx,n
Savelx=Request("Savelx")
	ChannelID=Request("ChannelID")
	ClassID=Request("ClassID")
	lx=0
		If ChannelID="" or ChannelID=0 Then
		   FoundErr=True
		   ErrMsg=ErrMsg & "<br><li>δָ��Ƶ��</li>"
		Else
		   ChannelID=Clng(ChannelID)
		End If
		If ClassID="" or ClassID = 0  Then
		   FoundErr=True
		   ErrMsg=ErrMsg & "<br><li>δָ����Ŀ</li>"
		Else
		   ClassID=CLng(ClassID)
		  set rs=ConnItem.execute("select * From SK_Class Where ClassID=" & ClassID)
		   If rs.bof and rs.eof then
				 FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ</li>"
			End If
			strChannelDir=rs("ClassDir")
			rs.close
			Set rs=Nothing


		End if
		if FoundErr=True then'������Ϣ
			Response.Write ErrMsg
		call Main()	
		Else
		if ConnItem.execute("select count(ArticleID) from PE_Article where ClassID="& ClassID & " and Passed=true and SaveFile<>True")(0) >0 then lx=1
		if ConnItem.execute("select count(ID) from SK_Flash where TID="& ClassID &" and Verific=1 and SaveFile=0")(0) >0 then lx=2
		if ConnItem.execute("select count(ID) from sk_photo where TID="& ClassID &" and Verific=1 and SaveFile=0" )(0) >0 then lx=3
		if lx=0 then
				Response.Write "û�ҵ���Ҫ������.����Ƿ���� �� ���Ծ��������"
		Else
		
			Response.Redirect "savefile.asp?lx="& lx &"&ClassID="& ClassID 
		end if
			call Main()	
		end if
end sub
'--------------------------SQL������------------------------
'===================================================
'С��ͨ�òɼ�ϵͳ
'��  �ã�SQL�����¼������
'===================================================
'==================================================
'��������Admin_ShowChannel_Name
'��  �ã���ʾƵ������
'��  ����ChannelID ------Ƶ��ID
'==================================================
Sub Admin_ShowChannel_Name(ChannelID)   
   Dim Sqlc,Rsc,TempStr
   ChannelID=Clng(ChannelID)
   Sqlc ="select top 1 ChannelName from sk_Channel Where ChannelID=" & ChannelID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,ConnItem,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="��ָ��Ƶ��"   
   Else   
      TempStr=Rsc("ChannelName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   response.write TempStr   
End Sub  

'==================================================
'��������Admin_ShowChannel_Option
'��  �ã���ʾƵ��ѡ��
'��  ����ChannelID ------Ƶ��ID
'==================================================
'--ipq��
Sub Admin_ShowChannel_Option(ChannelID)   
   Dim Sqlc,Rsc,ChannelName,TempStr
   ChannelID=Clng(ChannelID)
   Sqlc ="select ChannelID,ChannelName from sk_Channel where ModuleType=1 order by ChannelID asc"   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.Open Sqlc,ConnItem,1,1  
   TempStr="<option value=""0"">��ѡ��Ƶ��</option>"    
   If Rsc.Eof and Rsc.Bof Then
      TempStr=TempStr & "<option value=""0"">������Ƶ��</option>"   
   Else
      Do while not Rsc.Eof   
         TempStr=TempStr & "<option value=" & """" & Rsc("ChannelID") & """" & "" 
         If ChannelID=Rsc("ChannelID") Then
            TempStr=TempStr & "selected"
			End If
		 TempStr=TempStr & ">" & Rsc("ChannelName")
         TempStr=TempStr & "</option>"  
      Rsc.Movenext   
      Loop   
   End if
   Rsc.Close   
   Set Rsc=Nothing   
   Response.Write TempStr   
End sub 
'--ipq��
Sub Admin_ShowChannel_Opin(ChannelID)   
   Dim Sqlc,Rsc,ChannelName,TempStr
   ChannelID=Clng(ChannelID)
   Sqlc ="select ChannelID,ChannelName from sk_Channel where ModuleType=1 order by ChannelID asc"   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.Open Sqlc,ConnItem,1,1  
   TempStr="<option value=""0"" selected>��ѡ��Ƶ��</option>"    
   If Rsc.Eof and Rsc.Bof Then
      TempStr=TempStr & "<option value=""0"">������Ƶ��</option>"   
   Else
      Do while not Rsc.Eof   
         TempStr=TempStr & "<option value=" & """" & Rsc("ChannelID") & """" & "" 
         If ChannelID=Rsc("ChannelID") Then
            TempStr=TempStr & ""
         End If
         TempStr=TempStr & ">" & Rsc("ChannelName")
         TempStr=TempStr & "</option>"  
      Rsc.Movenext   
      Loop   
   End if
   Rsc.Close   
   Set Rsc=Nothing   
   Response.Write TempStr   
End sub 

'==================================================
'��������Admin_ShowClass_Name
'��  �ã���ʾ��Ŀ����
'��  ����ChannelID ------Ƶ��ID
'��  ����ClassID ------��ĿID
'==================================================
Sub Admin_ShowClass_Name(ChannelID,ClassID)   
   Dim SqlC,RsC,TempStr
   Sqlc ="Select top 1 ClassName from SK_Class Where  ClassID=" & ClassID 
   'Sqlc ="Select top 1 FolderName from ks_Class Where ID='"& ClassID &"'"   
   Set RsC=server.CreateObject("adodb.recordset")   
   RsC.Open SqlC,connItem,1,1   
   If RsC.Eof And RsC.Bof Then   
      TempStr="��ָ����Ŀ"   
   Else   
      TempStr=RsC("ClassName")
   End if   
   RsC.Close   
   Set RsC=Nothing
   Response.Write TempStr   
End Sub  

'==================================================
'��������Admin_ShowSpecial_Name
'��  �ã���ʾר������
'��  ����ChannelID ------Ƶ��ID
'��  ����SpecialID ------ר��ID
'==================================================
Sub Admin_ShowSpecial_Name(ChannelID,SpecialID)   
   Dim Sqlc,Rsc,TempStr
   ChannelID=Clng(ChannelID)
   SpecialID=Clng(SpecialID)
   Sqlc ="select top 1 SpecialName from SK_Special Where ChannelID=" & ChannelID & " and SpecialID=" & SpecialID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,ConnItem,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="��ָ��ר��"   
   Else   
      TempStr=Rsc("SpecialName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   Response.Write TempStr   
End Sub  

'==================================================
'��������Admin_ShowClass_Option
'��  �ã���ʾ��Ŀѡ��
'��  ����ChannelID ------Ƶ��ID
'��  ����ClassID ------��ĿID
'==================================================
sub Admin_ShowClass_Option(ChannelID,ClassID)
	dim rsClass,sqlClass,strTempC,tmpDepth,i
	dim arrShowLine(20)
	ChannelID=Clng(ChannelID)
	ClassID=Clng(ClassID)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
        StrTempC=""
	sqlClass="Select * From SK_Class where ChannelID="& ChannelID & " order by RootID,OrderID"
	set rsClass=ConnItem.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		strTempC="<option value='0'>����������Ŀ</option>"
	Else
                do while not rsClass.eof
                        tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			Else
				arrShowLine(tmpDepth)=False
			end if
	        strTempC=StrTempC & "<option value='" & rsClass("ClassID") & "'"
		        if rsClass("ClassID")=ClassID then
			        strTempC=strTempC & "selected"
		        end if
			strTempC=strTempC & ">" &rsClass("ClassName")
			strTempC=strTempC & "</option>"
		rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
	response.write strTempc
end sub

'==================================================
'��������Admin_ShowSpecial_Option
'��  �ã���ʾר��ѡ��
'��  ����ChannelID ------Ƶ��ID
'��  ����SpecialID ------ר��ID
'==================================================
sub Admin_ShowSpecial_Option(ChannelID,SpecialID)
    ChannelID=Clng(ChannelID)
    SpecialID=Clng(SpecialID)
    Dim TempStr
	TempStr="<select name='SpecialID' id='SpecialID'><option value='0'"
	if SpecialID=0 then
		TempStr=TempStr & " selected"
	end if
	TempStr=TempStr & ">�������κ�ר��</option>"
	                
	dim sqlSpecial,rsSpecial
        sqlSpecial = "select * from SK_Special where ChannelID=" & ChannelID
	set rsSpecial=server.CreateObject("adodb.recordset")
	rsSpecial.open sqlSpecial,ConnItem,1,1
	do while not rsSpecial.eof
		if rsSpecial("SpecialID")=SpecialID then
			TempStr=TempStr & "<option value='" & rsSpecial("SpecialID") & "' selected>" & rsSpecial("SpecialName") & "</option>"
		Else
			TempStr=TempStr & "<option value='" & rsSpecial("SpecialID") & "'>" & rsSpecial("SpecialName") & "</option>"
		end if
	rsSpecial.movenext
	loop
	rsSpecial.close
        set rsSpecial = nothing
        Response.write TempStr
end sub



'==================================================
'��������Admin_ShowTemplate_Option
'��  �ã���ʾ���ģ��ѡ��
'��  ����TemplateID ------���ģ��ID
'��  ����ChannelID-----
'==================================================
sub Admin_ShowTemplate_Option(ChannelID,TemplateType,TemplateID)
	dim sqlTemplate,rsTemplate,TempStr
    ChannelID=Clng(ChannelID)
    TempLateType=Clng(TempLateType)
    TempLateID=Clng(TempLateID)
	TempStr="<select name='TemplateID' id='TemplateID'><option value='0'>ϵͳĬ������ҳģ��</option>"
	sqlTemplate="select * from PE_Template where TemplateType=" & TemplateType & " And ChannelID=" & ChannelID
	set rsTemplate=server.CreateObject("adodb.recordset")
	rsTemplate.open sqlTemplate,ConnItem,1,1
	if rsTemplate.bof and rsTemplate.eof then
	  	TempStr= TempStr & "<option value='0'>����������ģ��</option>"
	Else
	  	do while not rsTemplate.eof
	  		if rsTemplate("TemplateID")=TemplateID then
				TempStr= TempStr & "<option value='" & rsTemplate("TemplateID") & "' selected>" & rsTemplate("TemplateName") & "</option>"
			Else		
				TempStr= TempStr & "<option value='" & rsTemplate("TemplateID") & "'>" & rsTemplate("TemplateName") & "</option>"
	  		end if		
			rsTemplate.movenext
	  	loop
	end if
	rsTemplate.close
	set rsTemplate=nothing
    TempStr= TempStr & "</select>"
    Response.Write TempStr
end sub

'==================================================
'��������Admin_ShowItem_Name
'��  �ã���ʾ��Ŀ����
'��  ����ItemID ------��ĿID
'==================================================
Sub Admin_ShowItem_Name(ItemID)   
   Dim Sqlc,Rsc,TempStr
   ItemID=Clng(ItemID)
   Sqlc ="select top 1 ItemName from Item Where ItemID=" & ItemID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,ConnItem,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="��ָ����Ŀ"   
   Else   
      TempStr=Rsc("ItemName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   Response.Write TempStr   
End Sub  


'==================================================
'��������Admin_ShowItem_Option
'��  �ã���ʾ��Ŀѡ��
'��  ����ItemID ------��ĿID
'==================================================
Sub Admin_ShowItem_Option(ItemID,lx)   
   Dim SqlI,RsI,TempStr
   ItemID=Clng(ItemID)
   if lx="" or lx=0 then
   		SqlI ="select ItemID,ItemName from Item order by ItemID desc"
   Else
   		SqlI ="select ItemID,ItemName from Item where Colleclx="& lx &" order by ItemID desc"   
   end if
   Set RsI=server.CreateObject("adodb.recordset")   
   RsI.Open SqlI,ConnItem,1,1
   TempStr="<select Name=""ItemID"" ID=""ItemID"">"   
   If RsI.Eof and RsI.Bof Then
      TempStr=TempStr & "<option value=""0"">��������Ŀ</option>"   
   Else   
      TempStr=TempStr & "<option value=""0"">��ѡ����Ŀ</option>"
      Do while not RsI.Eof   
         TempStr=TempStr & "<option value=" & """" & RsI("ItemID") & """" & "" 
         If ItemID=RsI("ItemID") Then
            TempStr=TempStr & " Selected"
         End If
         TempStr=TempStr & ">" & RsI("ItemName")
         TempStr=TempStr & "</option>"  
      RsI.Movenext   
      Loop   
   End if
   RsI.Close   
   Set RsI=Nothing   
   TempStr=TempStr & "</select>"
   Response.Write TempStr   
End sub   
'==================================================
'��������GetHttpPage
'��  �ã���ȡ��ҳԴ��
'��  ����HttpUrl ------��ҳ��ַ,Cset ����
'==================================================
Function GetHttpPage(ByVal URL, ByVal Cset)
On Error Resume Next
Dim Http
 If IsNull(URL)=True Or Len(URL)<18 Or URL="$False$" Then
      GetHttpPage="$False$"
      Exit Function
   End If
   Set Http=server.createobject("MSXML2.XMLHTTP")
   Http.open "GET",URL,False
   Http.Send()
   If Http.Readystate<>4 then
      Set Http=Nothing 
      GetHttpPage="$False$"
      Exit function
   End if
   GetHTTPPage=bytesToBSTR(Http.responseBody,Cset)
   Set Http=Nothing
   
   If Err.number<>0 then
   	  If IsNull(URL)=True Or Len(URL)<18 Or URL="$False$" Then
      GetHttpPage="$False$"
      Exit Function
   End If
   Set Http=server.createobject("MSXML2.XMLHTTP")
   'On Error Resume Next
   Http.open "GET",URL,False
   Http.Send()
   If Http.Readystate<>4 or Http.Status > 300 then
      Set Http=Nothing 
      GetHttpPage="$False$"
      Exit function
   End if
   GetHTTPPage=bytesToBSTR(Http.responseBody,Cset)
   Set Http=Nothing
   	  Err.Clear
   End If
   
End Function
'==================================================
'��������BytesToBstr
'��  �ã�����ȡ��Դ��ת��Ϊ����
'��  ����Body ------Ҫת���ı���
'��  ����Cset ------Ҫת��������
'==================================================
Function BytesToBstr(Body,Cset)
   Dim Objstream
   Set Objstream = Server.CreateObject("adodb.stream")
   objstream.Type = 1
   objstream.Mode =3
   objstream.Open
   objstream.Write body
   objstream.Position = 0
   objstream.Type = 2
   objstream.Charset = Cset
   BytesToBstr = objstream.ReadText 
   objstream.Close
   set objstream = nothing
End Function

'==================================================
'��������PostHttpPage
'��  �ã���¼
'==================================================
Function PostHttpPage(RefererUrl,PostUrl,PostData) 
    Dim xmlHttp 
    Dim RetStr      
    Set xmlHttp = CreateObject("Msxml2.XMLHTTP")  
    xmlHttp.Open "POST", PostUrl, False
    XmlHTTP.setRequestHeader "Content-Length",Len(PostData) 
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlHttp.setRequestHeader "Referer", RefererUrl
    xmlHttp.Send PostData 
    If Err.Number <> 0 Then 
        Set xmlHttp=Nothing
        PostHttpPage = "$False$"
        Exit Function
    End If
    PostHttpPage=bytesToBSTR(xmlHttp.responseBody,"GB2312")
    Set xmlHttp = nothing
End Function 

'==================================================
'��������UrlEncoding
'��  �ã�ת������
'==================================================
Function UrlEncoding(DataStr)
    Dim StrReturn,Si,ThisChr,InnerCode,Hight8,Low8
    StrReturn = ""
    For Si = 1 To Len(DataStr)
        ThisChr = Mid(DataStr,Si,1)
        If Abs(Asc(ThisChr)) < &HFF Then
            StrReturn = StrReturn & ThisChr
        Else
            InnerCode = Asc(ThisChr)
            If InnerCode < 0 Then
               InnerCode = InnerCode + &H10000
            End If
            Hight8 = (InnerCode  And &HFF00)\ &HFF
            Low8 = InnerCode And &HFF
            StrReturn = StrReturn & "%" & Hex(Hight8) &  "%" & Hex(Low8)
        End If
    Next
    UrlEncoding = StrReturn
End Function

'==================================================
'��������GetBody
'��  �ã���ȡ�ַ���
'��  ����ConStr ------��Ҫ��ȡ���ַ���
'��  ����StartStr ------��ʼ�ַ���
'��  ����OverStr ------�����ַ���
'��  ����IncluL ------�Ƿ����StartStr
'��  ����IncluR ------�Ƿ����OverStr
'==================================================
Function GetBody(ConStr,StartStr,OverStr,IncluL,IncluR)
   If ConStr="$False$" or ConStr="" or IsNull(ConStr)=True Or StartStr="" or IsNull(StartStr)=True Or OverStr="" or IsNull(OverStr)=True Then
      GetBody="$False$"
      Exit Function
   End If
   Dim ConStrTemp
   Dim Start,Over
   ConStrTemp=Lcase(ConStr)
   StartStr=Lcase(StartStr)
   OverStr=Lcase(OverStr)
   Start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
   If Start<=0 then
      GetBody="$False$"
      Exit Function
   Else
      If IncluL=False Then
         Start=Start+LenB(StartStr)
      End If
   End If
   Over=InStrB(Start,ConStrTemp,OverStr,vbBinaryCompare)
   If Over<=0 Or Over<=Start then
      GetBody="$False$"
      Exit Function
   Else
      If IncluR=True Then
         Over=Over+LenB(OverStr)
      End If
   End If
   GetBody=MidB(ConStr,Start,Over-Start)
End Function


'==================================================
'��������GetArray
'��  �ã���ȡ���ӵ�ַ����$Array$�ָ�
'��  ����ConStr ------��ȡ��ַ��ԭ�ַ�
'��  ����StartStr ------��ʼ�ַ���
'��  ����OverStr ------�����ַ���
'��  ����IncluL ------�Ƿ����StartStr
'��  ����IncluR ------�Ƿ����OverStr
'==================================================
Function GetArray(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
On Error Resume Next
   If ConStr="$False$" or ConStr="" Or IsNull(ConStr)=True or StartStr="" Or OverStr="" or  IsNull(StartStr)=True Or IsNull(OverStr)=True Then
      GetArray="$False$"
      Exit Function
   End If
   Dim TempStr,TempStr2,objRegExp,Matches,Match,Templisturl
   TempStr=""
   Set objRegExp = New Regexp 
   objRegExp.IgnoreCase = True 
   objRegExp.Global = True
   objRegExp.Pattern = "("&StartStr&").+?("&OverStr&")"
   Set Matches =objRegExp.Execute(ConStr) 
   For Each Match in Matches
      if Templisturl =Match.Value then
	  Else
	  TempStr=TempStr & "$Array$" & Match.Value
	  Templisturl=Match.Value
	  end if
   Next 
   Set Matches=nothing

   If TempStr="" Then
      GetArray="$False$"
      Exit Function
   End If
   TempStr=Right(TempStr,Len(TempStr)-7)
   If IncluL=False then
      objRegExp.Pattern =StartStr
      TempStr=objRegExp.Replace(TempStr,"")
   End if
   If IncluR=False then
      objRegExp.Pattern =OverStr
      TempStr=objRegExp.Replace(TempStr,"")
   End if
   Set objRegExp=nothing
   Set Matches=nothing
   
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")
   'TempStr=Replace(TempStr,"(","")
   'TempStr=Replace(TempStr,")","")

   If TempStr="" then
      GetArray="$False$"
   Else
      GetArray=TempStr
   End if
End Function


'==================================================
'��������DefiniteUrl
'��  �ã�����Ե�ַת��Ϊ���Ե�ַ
'��  ����PrimitiveUrl ------Ҫת������Ե�ַ
'��  ����ConsultUrl ------��ǰ��ҳ��ַ
'==================================================
Function DefiniteUrl(Byval PrimitiveUrl,Byval ConsultUrl)
   Dim ConTemp,PriTemp,Pi,Ci,PriArray,ConArray
   If PrimitiveUrl="" or ConsultUrl="" or PrimitiveUrl="$False$" or ConsultUrl="$False$" Then
      DefiniteUrl="$False$"
      Exit Function
   End If
   If Left(Lcase(ConsultUrl),7)<>"http://" Then
      ConsultUrl= "http://" & ConsultUrl
   End If
   ConsultUrl=Replace(ConsultUrl,"\","/")
   ConsultUrl=Replace(ConsultUrl,"://",":\\")
   PrimitiveUrl=Replace(PrimitiveUrl,"\","/")

   If Right(ConsultUrl,1)<>"/" Then
      If Instr(ConsultUrl,"/")>0 Then
         If Instr(Right(ConsultUrl,Len(ConsultUrl)-InstrRev(ConsultUrl,"/")),".")>0 then   
         Else
            ConsultUrl=ConsultUrl & "/"
         End If
      Else
         ConsultUrl=ConsultUrl & "/"
      End If
   End If
   ConArray=Split(ConsultUrl,"/")

   If Left(LCase(PrimitiveUrl),7) = "http://" then
      DefiniteUrl=Replace(PrimitiveUrl,"://",":\\")
   ElseIf Left(PrimitiveUrl,1) = "/" Then
      DefiniteUrl=ConArray(0) & PrimitiveUrl
   ElseIf Left(PrimitiveUrl,2)="./" Then
      PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-2)
      If Right(ConsultUrl,1)="/" Then   
         DefiniteUrl=ConsultUrl & PrimitiveUrl
      Else
         DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
      End If
   ElseIf Left(PrimitiveUrl,3)="../" then
      Do While Left(PrimitiveUrl,3)="../"
		 PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-3)
         Pi=Pi+1
      Loop            
      For Ci=0 to (Ubound(ConArray)-1-Pi)
         If DefiniteUrl<>"" Then
            DefiniteUrl=DefiniteUrl & "/" & ConArray(Ci)
         Else
            DefiniteUrl=ConArray(Ci)
         End If
      Next
      DefiniteUrl=DefiniteUrl & "/" & PrimitiveUrl
   Else
      If Instr(PrimitiveUrl,"/")>0 Then
         PriArray=Split(PrimitiveUrl,"/")
         If Instr(PriArray(0),".")>0 Then
            If Right(PrimitiveUrl,1)="/" Then
               DefiniteUrl="http:\\" & PrimitiveUrl
            Else
               If Instr(PriArray(Ubound(PriArray)-1),".")>0 Then 
                  DefiniteUrl="http:\\" & PrimitiveUrl
               Else
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               End If
            End If      
         Else
            If Right(ConsultUrl,1)="/" Then   
               DefiniteUrl=ConsultUrl & PrimitiveUrl
            Else
               DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
            End If
         End If
      Else
         If Instr(PrimitiveUrl,".")>0 Then
            If Right(ConsultUrl,1)="/" Then
               If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               Else
                  DefiniteUrl=ConsultUrl & PrimitiveUrl
               End If
            Else
               If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               Else
                  DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl
               End If
            End If
         Else
            If Right(ConsultUrl,1)="/" Then
               DefiniteUrl=ConsultUrl & PrimitiveUrl & "/"
            Else
               DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl & "/"
            End If         
         End If
      End If
   End If
   If Left(DefiniteUrl,1)="/" then
     DefiniteUrl=Right(DefiniteUrl,Len(DefiniteUrl)-1)
   End if
   If DefiniteUrl<>"" Then
      DefiniteUrl=Replace(DefiniteUrl,"//","/")
      DefiniteUrl=Replace(DefiniteUrl,":\\","://")
   Else
      DefiniteUrl="$False$"
   End If
End Function
'==================================================
'��������ReplaceSaveRemoteFile
'��  �ã��滻������Զ��ͼƬ
'��  ����ConStr ------ Ҫ�滻���ַ���
'��  ����SaveTf ------ �Ƿ񱣴��ļ���False�����棬True����
'��  ��: TistUrl------ ��ǰ��ҳ��ַ
'==================================================
Function ReplaceSaveRemoteFile(ConStr,strInstallDir,strChannelDir,SaveTf,TistUrl)
   If ConStr="$False$" or ConStr="" or strChannelDir="" Then
      ReplaceSaveRemoteFile=ConStr
      Exit Function
   End If
   Dim TempStr,TempStr2,TempStr3,Re,Matches,Match,Tempi,TempArray,TempArray2

   Set Re = New Regexp 
   Re.IgnoreCase = True 
   Re.Global = True
   Re.Pattern ="<img.+?[^\>]>"
   Set Matches =Re.Execute(ConStr) 
   For Each Match in Matches
      If TempStr<>"" then 
         TempStr=TempStr & "$Array$" & Match.Value
      Else
         TempStr=Match.Value
      End if
   Next
   If TempStr<>"" Then
      TempArray=Split(TempStr,"$Array$")
      TempStr=""
      For Tempi=0 To Ubound(TempArray)
         Re.Pattern ="src\s*=\s*.+?\.(gif|jpg|bmp|jpeg|psd|png|svg|dxf|wmf|tiff)"
         Set Matches =Re.Execute(TempArray(Tempi)) 
         For Each Match in Matches
            If TempStr<>"" then 
               TempStr=TempStr & "$Array$" & Match.Value
            Else
               TempStr=Match.Value
            End if
         Next
      Next
   End if
   If TempStr<>"" Then
   	  IncludePic=1'ͼƬ����
      Re.Pattern ="src\s*=\s*"
      TempStr=Re.Replace(TempStr,"")
   End If
   Set Matches=nothing
   Set Re=nothing
   If TempStr="" or IsNull(TempStr)=True Then
      ReplaceSaveRemoteFile=ConStr
      Exit function
   End if
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")
   
   Dim RemoteFileurl,SavePath,PathTemp,DtNow,strFileName,strFileType,ArrSaveFileName,RanNum,Arr_Path
   DtNow=Now()
   If SaveTf=True then
 '***********************************
      SavePath= strChannelDir & year(DtNow) & right("0" & month(DtNow),2) & "/"
	  response.write "����·����" & savepath & "<br>"
      Arr_Path=Split(SavePath,"/")
      PathTemp=""
      For Tempi=0 To Ubound(Arr_Path)
         If Tempi=0 Then
            PathTemp=Arr_Path(0) & "/"
         ElseIf Tempi=Ubound(Arr_Path) Then
            Exit For
         Else
            PathTemp=PathTemp & Arr_Path(Tempi) & "/"
         End If
         If CheckDir(PathTemp)=False Then
            If MakeNewsDir(PathTemp)=False Then
               SaveTf=False
               Exit For
            End If
         End If
      Next
   End If

   'ȥ���ظ�ͼƬ��ʼ
   TempArray=Split(TempStr,"$Array$")
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      If Instr(Lcase(TempStr),Lcase(TempArray(Tempi)))<1 Then
         TempStr=TempStr & "$Array$" & TempArray(Tempi)
      End If
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempArray=Split(TempStr,"$Array$")
   'ȥ���ظ�ͼƬ����

   'ת�����ͼƬ��ַ��ʼ
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      TempStr=TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi),TistUrl)
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempStr=Replace(TempStr,Chr(0),"")
   TempArray2=Split(TempStr,"$Array$")
   TempStr=""
   'ת�����ͼƬ��ַ����
	'ͼƬ�滻/����
   Set Re = New Regexp
   Re.IgnoreCase = True 
   Re.Global = True
   For Tempi=0 To Ubound(TempArray2)
      RemoteFileUrl=TempArray2(Tempi)
      If RemoteFileUrl<>"$False$" And SaveTf=True Then'����ͼƬ
         ArrSaveFileName = Split(RemoteFileurl,".")
	 strFileType=Lcase(ArrSaveFileName(Ubound(ArrSaveFileName)))'�ļ�����
         If strFileType="asp" or strFileType="asa" or strFileType="aspx" or strFileType="cer" or strFileType="cdx" or strFileType="exe" or strFileType="rar" or strFileType="zip" then
            UploadFiles=""
            ReplaceSaveRemoteFile=ConStr
            Exit Function
         End If

         Randomize
         RanNum=Int(900*Rnd)+100
	 strFileName = year(DtNow) & right("0" & month(DtNow),2) & right("0" & day(DtNow),2) & right("0" & hour(DtNow),2) & right("0" & minute(DtNow),2) & right("0" & second(DtNow),2) & ranNum & "." & strFileType
         Re.Pattern =TempArray(Tempi)
		 
	 If SaveRemoteFile(SavePath & strFileName,RemoteFileUrl)=True Then
'********************************
            PathTemp=SavePath & strFileName 
            ConStr=Re.Replace(ConStr,PathTemp)
            Re.Pattern=strInstallDir & strChannelDir 
            UploadFiles=UploadFiles & "|" & Re.Replace(SavePath &strFileName,"")
			Response.Flush()
			response.write " &nbsp;&nbsp;&nbsp;ͼƬ�����ַ��" & PathTemp & "<br>"
			if Thumb_WaterMark=1 then call SKThumb.AddWaterMark(PathTemp)'ˮӡ
         Else
            PathTemp=RemoteFileUrl
            ConStr=Re.Replace(ConStr,PathTemp)
            'UploadFiles=UploadFiles & "|" & RemoteFileUrl
         End If
      ElseIf RemoteFileurl<>"$False$" and SaveTf=False Then'������ͼƬ
         Re.Pattern =TempArray(Tempi)
         ConStr=Re.Replace(ConStr,RemoteFileUrl)
         UploadFiles=UploadFiles & "|" & RemoteFileUrl
      End If
   Next   
   Set Re=nothing
   If UploadFiles<>"" Then
      UploadFiles=Right(UploadFiles,Len(UploadFiles)-1)
   End If
   ReplaceSaveRemoteFile=ConStr
End function
	'================================================
	'��������ReSaveRemoteFile
	'��  �ã������ļ������滻
	'��  ����Str   ----ԭ�ַ���
	'��  ����url   ----��Ȼ��վURL
	'��  ����Dir -----����Ŀ¼
	'��  ����InSave ------�Ƿ񱣴�,True,False
	'����ֵ����ʽ��ȡ����ַ���
	'================================================
	Public Function ReSaveRemoteFile(ByVal str, ByVal URL, ByVal Dir,InSave)
		Dim s_Content
		Dim re
		Dim ContentFile, ContentFileUrl
		Dim strTempUrl,strFileUrl,DirTemp,PathTemp,FileTemp,Tempi,TempUrlArray,Arr_Path
		Dim sAllowExtName
		sAllowExtName="rm|swf"
		
		s_Content = str
		On Error Resume Next
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((src=|href=)((\S)+[.]{1}(" & sAllowExtName & ")))"
		Set ContentFile = re.Execute(s_Content)
		Dim sContentUrl(), n, i, bRepeat
		n = 0
		For Each ContentFileUrl In ContentFile
			strFileUrl = Replace(Replace(Replace(Replace(ContentFileUrl.Value, "src=", "", 1, -1, 1), "href=", "", 1, -1, 1), "'", ""), Chr(34), "")
			If n = 0 Then
				n = n + 1
				ReDim sContentUrl(n)
				sContentUrl(n) = strFileUrl
			Else
				bRepeat = False
				For i = 1 To UBound(sContentUrl)
					If UCase(strFileUrl) = UCase(sContentUrl(i)) Then
						bRepeat = True
						Exit For
					End If
				Next
				If bRepeat = False Then
					n = n + 1
					ReDim Preserve sContentUrl(n)
					sContentUrl(n) = strFileUrl
				End If
			End If
		Next
		If n = 0 Then
			ReSaveRemoteFile = s_Content
			Exit Function
		End If
		For i = 1 To n 
			strTempUrl = sContentUrl(i) : strTempUrl = FormatRemoteUrl(strTempUrl,URL)'�õ��ļ���ַ
			Response.Write(strTempUrl)
			IF InSave=True then
				Arr_Path=Split(Dir,"/")
				'----------��Ŀ¼-----------------------
				  For Tempi=0 To Ubound(Arr_Path)
					 If Tempi=0 Then
						PathTemp=Arr_Path(0) & "/"
					 ElseIf Tempi=Ubound(Arr_Path) Then
						Exit For
					 Else
						PathTemp=PathTemp & Arr_Path(Tempi) & "/"
					 End If
					 If CheckDir(PathTemp)=False Then
						If MakeNewsDir(PathTemp)=False Then
						   SaveTf=False
						   Exit For
						End If
					 End If
				  Next
				 '------------------------------------------------------
				TempUrlArray=Split(strTempUrl,"/")
				'----------����ļ��Ƿ����.������ڻ��ļ���------------------
				Do while True 
					FileTemp=Dir &  MakeRandom(5) & TempUrlArray(Ubound(TempUrlArray))'��������ļ���
					If CheckFile(FileTemp)=False then
						Exit Do
					end if
				loop 
				'-------------------------------------------------------------------
				Response.Write(FileTemp)
				If SaveRemoteFile(FileTemp,strTempUrl)=True then
					Response.Write("����ɹ�")&"<Br>"
					s_Content = Replace(s_Content,sContentUrl(i),FileTemp, 1, -1, 1)'�滻��ַ	
				Else
					Response.Write("����ʧ��")&"<Br>"
				End if
			Else
				s_Content = Replace(s_Content,sContentUrl(i),strTempUrl, 1, -1, 1)'�滻��ַ		
			End If	
		Next
		Set re = Nothing
		PictureExist = True
		ReSaveRemoteFile = s_Content
		Exit Function
	End Function
'================================================
'��������FormatRemoteUrl
'��  �ã���ʽ���ɵ�ǰ��վ������URL-����Ե�ַת��Ϊ���Ե�ַ
'��  ���� url ----Url�ַ���
'��  ���� CurrentUrl ----��Ȼ��վURL
'����ֵ����ʽ��ȡ���Url
'================================================
	Public Function FormatRemoteUrl(ByVal URL,ByVal CurrentUrl)
		Dim strUrl
		If Len(URL) < 2 Or Len(URL) > 255 Or Len(CurrentUrl) < 2 Then
			FormatRemoteUrl = vbNullString
			Exit Function
		End If
		CurrentUrl = Trim(Replace(Replace(Replace(Replace(Replace(CurrentUrl, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))
		URL = Trim(Replace(Replace(Replace(Replace(Replace(URL, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))	
		If InStr(9, CurrentUrl, "/") = 0 Then
			strUrl = CurrentUrl
		Else
			strUrl = Left(CurrentUrl, InStr(9, CurrentUrl, "/") - 1)
		End If

		If strUrl = vbNullString Then strUrl = CurrentUrl
		Select Case Left(LCase(URL), 6)
			Case "http:/", "https:", "ftp://", "rtsp:/", "mms://"
				FormatRemoteUrl = URL
				Exit Function
		End Select

		If Left(URL, 1) = "/" Then
			FormatRemoteUrl = strUrl & URL
			Exit Function
		End If
		
		If Left(URL, 3) = "../" Then
			Dim ArrayUrl
			Dim ArrayCurrentUrl
			Dim ArrayTemp()
			Dim strTemp
			Dim i, n
			Dim c, l
			n = 0
			ArrayCurrentUrl = Split(CurrentUrl, "/")
			ArrayUrl = Split(URL, "../")
			c = UBound(ArrayCurrentUrl)
			l = UBound(ArrayUrl) + 1
			
			If c > l + 2 Then
				For i = 0 To c - l
					ReDim Preserve ArrayTemp(n)
					ArrayTemp(n) = ArrayCurrentUrl(i)
					n = n + 1
				Next
				strTemp = Join(ArrayTemp, "/")
			Else
				strTemp = strUrl
			End If
			URL = Replace(URL, "../", vbNullString)
			FormatRemoteUrl = strTemp & "/" & URL
			Exit Function
		End If
		strUrl = Left(CurrentUrl, InStrRev(CurrentUrl, "/"))
		FormatRemoteUrl = strUrl & Replace(URL, "./", vbNullString)
		Exit Function
	End Function	
'================================================
'��������ReplaceTrim
'��  �ã����˵��ַ������е�tab�ͻس��ͻ���
'================================================
	Public Function ReplaceTrim(ByVal strContent)
		On Error Resume Next
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(" & Chr(8) & "|" & Chr(9) & "|" & Chr(10) & "|" & Chr(13) & ")"
		strContent = re.Replace(strContent, vbNullString)
		Set re = Nothing
		ReplaceTrim = strContent
		Exit Function
	End Function
'==================================================
'������: CheckFile
'��  �ã����ĳһ�ļ��Ƿ����
'��  ����FileName ------ �ļ���ַ ��:/swf/1.swf
'����ֵ��False  ----  True
'==================================================
	Public Function CheckFile(FileName)
		 On Error Resume Next
		 Dim FsoObj
		 Set FsoObj = Server.CreateObject("Scripting.FileSystemObject")
		  If Not FsoObj.FileExists(Server.MapPath(FileName)) Then
			  CheckFile = False
			  Exit Function
		  End If
		 CheckFile = True:Set FsoObj = Nothing
	End Function
'==================================================
'��������SaveRemoteFile
'��  �ã�����Զ�̵��ļ�������
'��  ����LocalFileName ------ �����ļ���
'��  ����RemoteFileUrl ------ Զ���ļ�URL
'==================================================
Function SaveRemoteFile(LocalFileName,RemoteFileUrl)
    SaveRemoteFile=True
	dim Ads,Retrieval,GetRemoteData	
	On Error Resume Next
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", RemoteFileUrl, False, "", ""
		.Send
        If .Readystate<>4 or .Status > 300 then
            SaveRemoteFile=False
            Exit Function
        End If
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	'If LenB(GetRemoteData) < 100 Then Exit Function
	'If MaxFileSize > 0 Then
			'If LenB(GetRemoteData) > 5000 Then Exit Function
			Response.Write(Round(LenB(GetRemoteData)/1024)) & "KB"
	'End If
	Set Ads = Server.CreateObject("Adodb.Stream")
	With Ads
		.Type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile server.MapPath(LocalFileName),2
		.Cancel()
		.Close()
	End With
	If Err.number<>0 then
	  SaveRemoteFile=False
      Exit Function
   	  Err.Clear
   	End If
	Set Ads=nothing
	
end Function
'==================================================
'��������FpHtmlEnCode
'��  �ã��������
'��  ����fString ------�ַ���
'==================================================
Function FpHtmlEnCode(fString)
   If IsNull(fString)=False or fString<>"" or fString<>"$False$" Then
       fString=nohtml(fString)
       fString=FilterJS(fString)
       fString = Replace(fString,"&nbsp;"," ")
       fString = Replace(fString,"&quot;","")
       fString = Replace(fString,"&#39;","")
       fString = replace(fString, ">", "")
       fString = replace(fString, "<", "")
       fString = Replace(fString, CHR(9), " ")'&nbsp;
       fString = Replace(fString, CHR(10), "")
       fString = Replace(fString, CHR(13), "")
       fString = Replace(fString, CHR(34), "")
       fString = Replace(fString, CHR(32), " ")'space
       fString = Replace(fString, CHR(39), "")
       fString = Replace(fString, CHR(10) & CHR(10),"")
       fString = Replace(fString, CHR(10)&CHR(13), "")
       fString=Trim(fString)
       FpHtmlEnCode=fString
   Else
       FpHtmlEnCode="$False$"
   End If
End Function

'==================================================
'��������GetPaing
'��  �ã���ȡ��ҳ
'==================================================
Function GetPaing(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
If ConStr="$False$" or ConStr="" Or StartStr="" Or OverStr="" or IsNull(ConStr)=True or IsNull(StartStr)=True Or IsNull(OverStr)=True Then
   GetPaing="$False$"
   Exit Function
End If

Dim Start,Over,ConTemp,TempStr
TempStr=LCase(ConStr)
StartStr=LCase(StartStr)
OverStr=LCase(OverStr)
Over=Instr(1,TempStr,OverStr)
If Over<=0 Then
   GetPaing="$False$"
   Exit Function
Else
   If IncluR=True Then
      Over=Over+Len(OverStr)
   End If
End If
TempStr=Mid(TempStr,1,Over)
Start=InstrRev(TempStr,StartStr)
If IncluL=False Then
   Start=Start+Len(StartStr)
End If

If Start<=0 Or Start>=Over Then
   GetPaing="$False$"
   Exit Function
End If
ConTemp=Mid(ConStr,Start,Over-Start)

ConTemp=Trim(ConTemp)
ConTemp=Replace(ConTemp," ","")
ConTemp=Replace(ConTemp,",","")
ConTemp=Replace(ConTemp,"'","")
ConTemp=Replace(ConTemp,"""","")
ConTemp=Replace(ConTemp,">","")
ConTemp=Replace(ConTemp,"<","")
ConTemp=Replace(ConTemp,"&nbsp;","")
GetPaing=ConTemp
End Function

'==================================================
'��������ScriptHtml
'��  �ã�����html���
'��  ����ConStr ------ Ҫ���˵��ַ���
'==================================================
Function ScriptHtml(Byval ConStr,TagName,FType)
    Dim Re
    Set Re=new RegExp
    Re.IgnoreCase =true
    Re.Global=True
    Select Case FType
    Case 1
       Re.Pattern="<" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    Case 2
       Re.Pattern="<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"") 
	Case 3
       Re.Pattern="<" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
       Re.Pattern="</" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    End Select
    ScriptHtml=ConStr
    Set Re=Nothing
End Function

'-------------------------
'--���Ŀ¼�Ƿ����
'----------------------
Function CheckDir(byval FolderPath)
	dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(Server.MapPath(folderpath)) then
	'����
		CheckDir = True
	Else
	'������
		CheckDir = False
	End if
	Set fso = nothing
End Function

'-------------------------
'--����Ŀ¼
'----------------------
Function MakeNewsDir(byval foldername)
	dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder(Server.MapPath(foldername))
        If fso.FolderExists(Server.MapPath(foldername)) Then
           MakeNewsDir = True
        Else
           MakeNewsDir = False
        End If
	Set fso = nothing
End Function

'**************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'**************************************************
Function IsObjInstalled(strClassString)
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'**************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'**************************************************
sub WriteErrMsg(ErrMsg)
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='css/Admin_STYLE.CSS' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & ErrMsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
	response.end
end sub

'**************************************************
'��������WriteSucced
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'**************************************************
sub WriteSucced(ErrMsg)
	dim strErr
	strErr=strErr & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='../admin/Admin_STYLE.CSS' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>��ϲ�㣡</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top' align='center'>" & ErrMsg &"</td></tr>" & vbcrlf
	'strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'��������ShowPage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sFileName  ----���ӵ�ַ
'       TotalNumber ----������
'       MaxPerPage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'����ֵ������һҳ ��һҳ������Ϣ��HTML����
'**************************************************
function ShowPage(sFileName,TotalNumber,MaxPerPage,ShowTotal,ShowAllPages,strUnit)
	dim TotalPage,strTemp,strUrl,i

	if TotalNumber=0 or MaxPerPage=0 or isNull(MaxPerPage) then
		ShowPage=""
		exit function
	end if
	if totalnumber mod maxperpage=0 then
    	TotalPage= totalnumber \ maxperpage
  	Else
    	TotalPage= totalnumber \ maxperpage+1
  	end if
	if CurrentPage>TotalPage then CurrentPage=TotalPage
		
  	strTemp= "<table align='center'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    	strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	Else
    	strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    	strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if CurrentPage>=TotalPage then
    	strTemp=strTemp & "��һҳ βҳ"
  	Else
    	strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    	strTemp=strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong>ҳ "
        strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;&nbsp;ת����<input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "page=" & "'+this.value;""'>ҳ"
		 'strTemp  = strTemp &"&nbsp;<Input type=""button""  onClick=""window.location.href='" & strUrl & "page='+document.all.page.value;""  name=button1  value=GO >"
	end if
	strTemp=strTemp & "</td></tr></table>"
	ShowPage=strTemp
end function

'**************************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'**************************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			Else
				JoinChar=strUrl
			end if
		Else
			JoinChar=strUrl & "?"
		end if
	Else
		JoinChar=strUrl
	end if
end function

'**************************************************
'��������CreateKeyWord
'��  �ã��ɸ������ַ������ɹؼ���
'��  ����Constr---Ҫ���ɹؼ��ֵ�ԭ�ַ���
'����ֵ�����ɵĹؼ���
'**************************************************
Function CreateKeyWord(byval Constr,Num)
   If Constr="" or IsNull(Constr)=True or Constr="$False$" Then
      CreateKeyWord="$False$"
      Exit Function
   End If
   If Num="" or IsNumeric(Num)=False Then
      Num=2
   End If
   Constr=Replace(Constr,CHR(32),"")
   Constr=Replace(Constr,CHR(9),"")
   Constr=Replace(Constr,"&nbsp;","")
   Constr=Replace(Constr," ","")
   Constr=Replace(Constr,"(","")
   Constr=Replace(Constr,")","")
   Constr=Replace(Constr,"<","")
   Constr=Replace(Constr,">","")
   Constr=Replace(Constr,"""","")
   Constr=Replace(Constr,"?","")
   Constr=Replace(Constr,"*","")
   Constr=Replace(Constr,"|","")
   Constr=Replace(Constr,",","")
   Constr=Replace(Constr,".","")
   Constr=Replace(Constr,"/","")
   Constr=Replace(Constr,"\","")
   Constr=Replace(Constr,"-","")
   Constr=Replace(Constr,"@","")
   Constr=Replace(Constr,"#","")
   Constr=Replace(Constr,"$","")
   Constr=Replace(Constr,"%","")
   Constr=Replace(Constr,"&","")
   Constr=Replace(Constr,"+","")
   Constr=Replace(Constr,":","")
   Constr=Replace(Constr,"��","")   
   Constr=Replace(Constr,"��","")
   Constr=Replace(Constr,"��","")
   Constr=Replace(Constr,"��","")
   Constr=Replace(Constr,"&","")  
   Constr=Replace(Constr,"gt;","")      
   Dim i,ConstrTemp
   For i=1 To Len(Constr)
      ConstrTemp=ConstrTemp & "|" & Mid(Constr,i,Num)
   Next
   If Len(ConstrTemp)<254 Then
      ConstrTemp=ConstrTemp & "|"
   Else
      ConstrTemp=Left(ConstrTemp,254) & "|"
   End If
   ConstrTemp=left(ConstrTemp,len(ConstrTemp)-1)
   ConstrTemp= Right(ConstrTemp,len(ConstrTemp)-1)
   CreateKeyWord=ConstrTemp
End Function

Function CheckUrl(strUrl)
   Dim Re
   Set Re=new RegExp
   Re.IgnoreCase =true
   Re.Global=True
   Re.Pattern="http://([\w-]+\.)+[\w-]+(/[\w-./?%&=]*)?"
   If Re.test(strUrl)=True Then
      CheckUrl=strUrl
   Else
      CheckUrl="$False$"
   End If
   Set Rs=Nothing
End Function
Function GetItemId(ItemID,Itemlx) 'ѡ��ɼ�
	Dim ItemIdstr,ItemID_id : ItemID_id=Split(ItemID,",")  
	for i=1 to Ubound(ItemID_id)
			If i=1 then 
				ItemIdstr=ItemID_id(i)
			Else
				ItemIdstr=ItemIdstr & "," & ItemID_id(i)
			End if
	NExt
	if Itemlx=0 then GetItemId=ItemID_id(0)	Else GetItemId=ItemIdstr end if
End Function 
%>