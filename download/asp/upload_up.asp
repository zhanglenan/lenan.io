<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Server.ScriptTimeOut = 1800%>
<!--#include file="upload_class.asp"-->
<%
If Session("Ok3w_eWebEditor")="" Then
	Response.Write("�㻹û�е�½���ǵ�½�Ѿ���ʱ��")
	Response.End()
End If

Const AllowExe = "zip,rar"
Const AllowSize = 10000

action = Trim(Request.QueryString("action"))
If action="save" Then
	ErrMsg = ""
	Set upload = New upload_5xsoft
	Set file=upload.file("file")
	If file.FileSize>0 then
		FileExeTmp = Split(file.FileName,".")
		FileExe = Lcase(FileExeTmp(Ubound(FileExeTmp)))
		
		If file.FileSize>AllowSize*1024 Then
			ErrMsg = "�ļ���С���ܳ���" & AllowSize & "KB��������ѡ��"
		End If
		If ErrMsg="" Then
			If InStr("," & AllowExe & ",","," & FileExe & ",") <=0 Then
				ErrMsg = "ֻ�����ϴ���" & AllowExe & "�����͵��ļ���������ѡ��"
			End If
		End If
		
		If ErrMsg = "" Then
			FilePath = "soft/" & Year(Date()) & Right("0" & Month(Date()),2) & "/"
			Set fso = CreateObject("Scripting.FileSystemObject")
			If Not fso.FolderExists(Server.MapPath("../" & FilePath)) Then
				fso.CreateFolder(Server.MapPath("../" & FilePath))
			End If
			FilePath = "soft/" & Year(Date()) & Right("0" & Month(Date()),2) & "/" & Right("0" & Day(Date()),2) & "/"
			Set fso = CreateObject("Scripting.FileSystemObject")
			If Not fso.FolderExists(Server.MapPath("../" & FilePath)) Then
				fso.CreateFolder(Server.MapPath("../" & FilePath))
			End If
			Set fso = Nothing
			FilePath = FilePath & file.FileName'Year(Date()) & Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Right("0" & Hour(Time()),2) & Right("0" & Minute(Time()),2) & Right("0" & Second(Time()),2) & "." & FileExe
			file.SaveAs Server.mappath("../" & FilePath)
			fsize = file.FileSize \ 1024
		End If
		Else
			ErrMsg = "��ѡ����Ҫ�ϴ����ļ�"
	End If
	Set file = Nothing
	Set upload = Nothing
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ļ��ϴ�</title>
<style type="text/css">
body,td{
	margin:0px;
	font-size:12px;
}
</style>
<%If ErrMsg<>"" Then%>
<script language="javascript">
alert("<%=ErrMsg%>");
</script>
<%End If%>
</head>

<body bgcolor="menu">
<%If FilePath<>"" And ErrMsg="" Then%>
<script language="javascript">
window.returnValue = "<%=FilePath%>|<%=fsize%>"; 
window.close(); 
</script>
<%Else%>
<br />
<script language="javascript">
function chkup(frm)
{
	if(frm.file.value=="")
	{
		alert("��ѡ���ļ�");
		frm.file.focus();
		return false;
	}
	
	divProcessing.style.display="";
	
	frm.bntSubmit.disabled=true;
	frm.bntSubmit.value=" �� �� ";
	frm.submit();
}
</script>
<table border="0" align="center" cellpadding="10" cellspacing="0">
  <form id="myform" name="myform" enctype="multipart/form-data" method="post" action="?action=save">
    <tr>
      <td>��ѡ���ļ�</td>
      <td><input type="file" name="file" onKeyPress="return false;" /></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input name="bntSubmit" type="button" id="bntSubmit" onClick="chkup(this.form)" value=" �� �� " />
      <input type="button" name="Submit" value=" ȡ �� " onClick="window.close()"></td>
    </tr>
  </form>
</table>
<div id="divProcessing" style="width:200px;height:30px;position:absolute;left:50px;top:50px;display:none">
<table border=0 cellpadding=0 cellspacing=1 bgcolor="#000000" width="100%" height="100%"><tr><td bgcolor=#3A6EA5><marquee align="middle" behavior="alternate" scrollamount="5"><font color=#FFFFFF>...�ļ��ϴ���...��ȴ�...</font></marquee></td></tr></table>
</div>
<%End If%>
</body>
</html>