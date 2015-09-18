<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>文件上传</title>
<style type="text/css">
body{margin:0px;}
body,td{font-size:12px; color:#000000;}
a{color:#0000FF; text-decoration:underline;}
a:hover{ color:#FF0000;}
.myinput{border:1px solid #CCCCCC; background-color:#FFFFFF; margin-top:5px; padding:2px;}
form{margin:0px; padding:0px;}
</style>
</head>

<body>
<%
If Session("Ok3w_eWebEditor")="" Then
	Response.Write("你还没有登陆或是登陆已经超时。")
	Response.End()
End If

Dim Data_5xsoft
Dim formID,objID
Dim ImgW,ImgH
Dim FilePath,upload,FileExe,FileExeTmp,fso
Const AllowExe = "jpg,gif,bmp,png"
Const AllowSize = 1024
formID = Request.QueryString("formID")
objID = Request.QueryString("objID")
FilePath = ""

If Trim(Request.QueryString("action")) = "save" Then
	Set upload = New upload_5xsoft
	Set file=upload.file("file")
	ImgW = Cint(upload.Form("ImgW"))
	ImgH = Cint(upload.Form("ImgH"))
	If file.FileSize>0 then
		FileExeTmp = Split(file.FileName,".")
		FileExe = Lcase(FileExeTmp(Ubound(FileExeTmp)))
		
		If file.FileSize>AllowSize*1024 Then
			Response.Write("文件大小不能超过" & AllowSize & "KB，请<a href=""upload_files.asp?formID=" & formID & "&objID=" & objID & """>返回</a>。")
			Response.End()
		End If
		If InStr("," & AllowExe & ",","," & FileExe & ",") <=0 Then
			Response.Write("不允许上传该类型文件，请<a href=""upload_files.asp?formID=" & formID & "&objID=" & objID & """>返回</a>。")
			Response.End()
		End If
		
		FilePath = "upfiles/" & Year(Date()) & Right("0"&Month(Date()),2) & "/"
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Not fso.FolderExists(Server.MapPath("../" & FilePath)) Then
			fso.CreateFolder(Server.MapPath("../" & FilePath))
		End If
  		Set fso = Nothing
		
		Randomize
		sRnd = Int(900 * Rnd) + 100
		FilePath = FilePath & year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & sRnd & "." & FileExe

		file.SaveAs Server.mappath("../" & FilePath)
		
		'================aspJpeg 开始====================
		BG_File = "../" & FilePath
		
		If Application("Ok3w_Sitesy_Type")<>0 Then
			
			Set BG = Server.CreateObject("Persits.Jpeg")
			BG.Open Server.MapPath(BG_File)
			BG_W = BG.Width
			BG_H = BG.Height
			
			If Application("Ok3w_Sitesy_Type") = 1 Then
				BG.Canvas.Font.Color = "&H" & Application("Ok3w_Sitesy_Color")
				BG.Canvas.Font.ShadowColor = &HFFFFFF
				BG.Canvas.Font.Family = Application("Ok3w_Sitesy_Family")
				BG.Canvas.Font.Size = Application("Ok3w_Sitesy_Size")
				BG.Canvas.Font.Bold = False
				BG.Canvas.Font.Quality = 3
				Select Case Application("Ok3w_Sitesy_Location")
					Case 0
						x = 20 : y = 20
					Case 1
						x = BG_W - Len(Application("Ok3w_Sitesy_Text")) * 20 : y = 20
					Case 2
						x = 20 : y = BG_H - 20
					Case 3
						x = BG_W - Len(Application("Ok3w_Sitesy_Text")) * 20 : y = BG_H - 20*2
					Case 4
						x = BG_W\2 - Len(Application("Ok3w_Sitesy_Text")) * 20\2 : y = BG_H\2 - 20*2
				End Select
				BG.Canvas.PrintText x, y, Application("Ok3w_Sitesy_Text")
			End If
			
			If Application("Ok3w_Sitesy_Type") = 2 Then
				Set logo = Server.CreateObject("Persits.Jpeg")
				logo.Open Server.MapPath("../" & Application("Ok3w_Sitesy_Logo"))
				logo_w = logo.Width
				logo_h = logo.Height
				Select Case Application("Ok3w_Sitesy_Location")
					Case 0
						x = 20 : y = 20
					Case 1
						x = BG_W - logo_w - 20 : y = 20
					Case 2
						x = 20 : y = logo_h - 20
					Case 3
						x = BG_W - logo_w - 20 : y = BG_H - logo_h - 20
					Case 4
						x = BG_W\2 - logo_w\2 : y = BG_H\2 - logo_h - 20
				End Select
				BG.DrawImage x, y, logo, 0.8, &HFFFFFF 
				Set logo = Nothing
			End If
			
			BG.Quality = 85
			
			BG.Save Server.MapPath(BG_File)
			Set BG = Nothing
		End If
		
		If ImgW<>0 Or ImgH<>0 Then
			Set BG = Server.CreateObject("Persits.Jpeg")
			BG.Open Server.MapPath(BG_File)
			
			If ImgW = 0 Then
				ImgW = Int(BG.OriginalWidth * ImgH / BG.OriginalHeight)
			End If
			If ImgH = 0 Then
				ImgH = Int(BG.OriginalHeight * ImgW / BG.OriginalWidth)
			End If
			
			BG.Width = ImgW
			BG.Height = ImgH
			
			BG.Save Server.MapPath(BG_File)
			Set BG = Nothing 
		End If
		
		'================aspJpeg 结束====================
	End If
	Set file = Nothing
	Set upload = Nothing
End If
%>
<div style="border:1px solid #999999; padding:5px; margin-top:5px; background-color:#EBEBEB; width:420px;">
<script language="javascript">
var IsAspJpeg = "<%=Lcase(IsObjInstalled("Persits.Jpeg"))%>";
function chkdata(frm)
{
	if(frm.file.value=="")
	{
		alert("请选择需要上传的本地文件");
		return false;
	}
	if(IsAspJpeg=="false" && (frm.ImgW.value!=0 || frm.ImgH.value!=0))
	{
		alert("你没有按装AspJpeg，无法自动生成缩略图");
		return false;
	}
	
	divProcessing.style.display="";
	
	return true;
}
</script>
<form id="form1" name="form1" enctype="multipart/form-data" method="post" action="?action=save&formID=<%=formID%>&objID=<%=objID%>" onsubmit="return chkdata(this);">
<%If FilePath="" Then%>
  文件：<input name="file" type="file" style="width:255px; border:1px solid #999999; background-color:#FFFFFF; padding:2px; " onKeyPress="return false;" />
  <br />
  缩放：宽
  <input name="ImgW" type="text" class="myinput" id="ImgW" value="0" size="4" />
  px
  ×高
  <input name="ImgH" type="text" class="myinput" id="ImgH" value="0" size="4" />
  px
  “0”表示自动适应<br />
  <input name="Submit" type="submit" class="myinput" value="开始上传" />
<div id="divProcessing" style="width:200px;height:30px;position:absolute;left:50px;top:30px;display:none;">
<table border=0 cellpadding=0 cellspacing=1 bgcolor="#000000" width="100%" height="100%"><tr><td bgcolor=#3A6EA5><marquee align="middle" behavior="alternate" scrollamount="5"><font color=#FFFFFF>...文件上传中...请等待...</font></marquee></td></tr></table>
</div>
  <%Else%>  
  <br /><br /><a href="upload_files.asp?formID=<%=formID%>&objID=<%=objID%>">上传成功，返回点这里...</a><br /><br /><br />
  <script language="javascript">
  parent.<%=formID%>.<%=objID%>.value = "<%=FilePath%>";
  </script>  
  <%End If%>  
</form>
</div>
</body>
</html>

<%
Function IsObjInstalled(ObjName)
	On Error ReSume Next
	Set Obj = Server.CreateObject(ObjName)
	If Err.Number<>0 Then
		Err.Clear
		IsObjInstalled = False
		Else
			IsObjInstalled = True
	End If
	On Error Goto 0
End Function

Class upload_5xsoft
  
dim objForm,objFile,Version

Public function Form(strForm)
   strForm=lcase(strForm)
   if not objForm.exists(strForm) then
     Form=""
   else
     Form=objForm(strForm)
   end if
 end function

Public function File(strFile)
   strFile=lcase(strFile)
   if not objFile.exists(strFile) then
     set File=new FileInfo
   else
     set File=objFile(strFile)
   end if
 end function


Private Sub Class_Initialize 
  dim RequestData,sStart,vbCrlf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,theFile
  dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
  dim iFindStart,iFindEnd
  dim iFormStart,iFormEnd,sFormName
  Version="化境HTTP上传程序 Version 2.1"
  set objForm=Server.CreateObject("Scripting.Dictionary")
  set objFile=Server.CreateObject("Scripting.Dictionary")
  if Request.TotalBytes<1 then Exit Sub
  set tStream = Server.CreateObject("adodb.stream")
  set Data_5xsoft = Server.CreateObject("adodb.stream")
  Data_5xsoft.Type = 1
  Data_5xsoft.Mode =3
  Data_5xsoft.Open
  Data_5xsoft.Write  Request.BinaryRead(Request.TotalBytes)
  Data_5xsoft.Position=0
  RequestData =Data_5xsoft.Read 

  iFormStart = 1
  iFormEnd = LenB(RequestData)
  vbCrlf = chrB(13) & chrB(10)
  sStart = MidB(RequestData,1, InStrB(iFormStart,RequestData,vbCrlf)-1)
  iStart = LenB (sStart)
  iFormStart=iFormStart+iStart+1
  while (iFormStart + 10) < iFormEnd 
	iInfoEnd = InStrB(iFormStart,RequestData,vbCrlf & vbCrlf)+3
	tStream.Type = 1
	tStream.Mode =3
	tStream.Open
	Data_5xsoft.Position = iFormStart
	Data_5xsoft.CopyTo tStream,iInfoEnd-iFormStart
	tStream.Position = 0
	tStream.Type = 2
	tStream.Charset ="gb2312"
	sInfo = tStream.ReadText
	tStream.Close
	'取得表单项目名称
	iFormStart = InStrB(iInfoEnd,RequestData,sStart)
	iFindStart = InStr(22,sInfo,"name=""",1)+6
	iFindEnd = InStr(iFindStart,sInfo,"""",1)
	sFormName = lcase(Mid (sinfo,iFindStart,iFindEnd-iFindStart))
	'如果是文件
	if InStr (45,sInfo,"filename=""",1) > 0 then
		set theFile=new FileInfo
		'取得文件名
		iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
		iFindEnd = InStr(iFindStart,sInfo,"""",1)
		sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
		theFile.FileName=getFileName(sFileName)
		theFile.FilePath=getFilePath(sFileName)
		theFile.FileExt=GetFileExt(sFileName)
		'取得文件类型
		iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14
		iFindEnd = InStr(iFindStart,sInfo,vbCr)
		theFile.FileType =Mid (sinfo,iFindStart,iFindEnd-iFindStart)
		theFile.FileStart =iInfoEnd
		theFile.FileSize = iFormStart -iInfoEnd -3
		theFile.FormName=sFormName
		if not objFile.Exists(sFormName) then
		  objFile.add sFormName,theFile
		end if
	else
	'如果是表单项目
		tStream.Type =1
		tStream.Mode =3
		tStream.Open
		Data_5xsoft.Position = iInfoEnd 
		Data_5xsoft.CopyTo tStream,iFormStart-iInfoEnd-3
		tStream.Position = 0
		tStream.Type = 2
		tStream.Charset ="gb2312"
	        sFormValue = tStream.ReadText 
	        tStream.Close
		if objForm.Exists(sFormName) then
		  objForm(sFormName)=objForm(sFormName)&", "&sFormValue		  
		else
		  objForm.Add sFormName,sFormValue
		end if
	end if
	iFormStart=iFormStart+iStart+1
	wend
  RequestData=""
  set tStream =nothing
End Sub

Private Sub Class_Terminate  
 if Request.TotalBytes>0 then
	objForm.RemoveAll
	objFile.RemoveAll
	set objForm=nothing
	set objFile=nothing
	Data_5xsoft.Close
	set Data_5xsoft =nothing
 end if
End Sub
   
 
 Private function GetFilePath(FullPath)
  If FullPath <> "" Then
   GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
  Else
   GetFilePath = ""
  End If
 End  function

 Private function GetFileExt(FullPath)
  If FullPath <> "" Then
   GetFileExt = mid(FullPath,InStrRev(FullPath, ".")+1)
  Else
   GetFileExt = ""
  End If
 End  function
 
 Private function GetFileName(FullPath)
  If FullPath <> "" Then
   GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
  Else
   GetFileName = ""
  End If
 End  function
End Class

Class FileInfo
  dim FormName,FileName,FilePath,FileSize,FileExt,FileType,FileStart
  Private Sub Class_Initialize 
    FileName = ""
    FilePath = ""
    FileSize = 0
    FileStart= 0
    FormName = ""
    FileType = ""
    FileExt  = ""
  End Sub
  
 Public function SaveAs(FullPath)
    dim dr,ErrorChar,i
    SaveAs=true
    if trim(fullpath)="" or FileStart=0 or FileName="" or right(fullpath,1)="/" then exit function
    set dr=CreateObject("Adodb.Stream")
    dr.Mode=3
    dr.Type=1
    dr.Open
    Data_5xsoft.position=FileStart
    Data_5xsoft.copyto dr,FileSize
    dr.SaveToFile FullPath,2
    dr.Close
    set dr=nothing 
    SaveAs=false
  end function
  End Class
%>