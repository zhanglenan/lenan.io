
<!--#include file="inc/setup.asp"-->

<%
Dim Rs,Sql,SqlItem,RsItem,Action,FoundErr
Dim HistrolyID,ItemID,ChannelID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result
Dim  Arr_Histroly,Arr_ArticleID,i_Arr,Del,Flag
Dim MaxPerPage,CurrentPage,AllPage,HistrolyNum,i_His
MaxPerPage=20
FoundErr=False
Del=Trim(Request("Del"))
Action=Trim(Request("Action"))
If Del="Del" Then
   Call DelHistroly()
End If
If FoundErr<>True Then
   Call Main()
else
  Call WriteErrMsg(ErrMsg)
End If
'�ر����ݿ�����
Call CloseConnItem()
%>

<%Sub Main%>
<html>
<head>
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/Admin_Style.css">
<style type="text/css">
.ButtonList {
	BORDER-RIGHT: #000000 2px solid; BORDER-TOP: #ffffff 2px solid; BORDER-LEFT: #ffffff 2px solid; CURSOR: default; BORDER-BOTTOM: #999999 2px solid; BACKGROUND-COLOR: #e6e6e6
}
</style>
<SCRIPT language=javascript>
function unselectall(thisform)
{
    if(thisform.chkAll.checked)
	{
		thisform.chkAll.checked = thisform.chkAll.checked&0;
    } 	
}

function CheckAll(thisform)
{
	for (var i=0;i<thisform.elements.length;i++)
    {
	var e = thisform.elements[i];
	if (e.Name != "chkAll"&&e.disabled!=true)
		e.checked = thisform.chkAll.checked;
    }
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr class="topbg">
	<td height="22" colspan="2" align="center"><strong>�� �� ϵ ͳ �� ʷ �� ¼ �� ��</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr class="tdbg"> 
    <td height="30" width="65"><strong>����������</strong></td>  
    <td height="30"><a href="Sk_ItemHistroly.asp">������ҳ</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Sk_ItemHistroly.asp?Action=Succeed">�ɹ���¼</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Sk_ItemHistroly.asp?Action=Failure">ʧ�ܼ�¼</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Sk_ItemHistroly.asp?Action=LoseEf">ʧЧ��¼</a></td>     
  </tr>         
</table>     
<%                             
Set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from Histroly"
If Action="Succeed"  Then
   SqlItem=SqlItem  &  " Where  Result=True"
   Flag="�� �� �� ¼"
ElseIf Action="Failure"  Then
   SqlItem=SqlItem  &  " Where  Result=False"
   Flag="ʧ �� �� ¼"
ElseIf  Action="LoseEf"  Then
   Flag="ʧ Ч �� ¼"
   Set Rs=server.createobject("adodb.recordset")
   Sql="Select ArticleID From SK_Article Where Deleted=False"
   Rs.open  Sql,ConnItem,1,1
   If (Not Rs.Eof) And (Not Rs.Bof) Then
      Do While not rs.eof
         Arr_ArticleID=Arr_ArticleID & "," & CStr(rs("ArticleID"))
      Rs.MoveNext
      Loop
   End  If
   Rs.Close
   Set Rs=Nothing
   If Arr_ArticleID<>"" Then
      Arr_ArticleID=0 & Arr_ArticleID
   Else
      Arr_ArticleID=0
   End If
   SqlItem=SqlItem  &  " Where ArticleID Not In("  &  Arr_ArticleID & ")"
Else
   Flag="�� �� �� ¼"
End  If
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>��  ʷ  ��  ¼ ����<%=Flag%></strong></div></td>
    </tr>
</table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">         
   <form name="form1" method="POST" action="?">         
     <tr class="tdbg" style="padding: 0px 2px;">         
      <td width="56" height="22" align="center" bgcolor="#EFEFEF" >ѡ��</td>                 
      <td width="214" align="center" bgcolor="#EFEFEF" >��Ŀ����</td>         
      <td width="435" align="center" bgcolor="#EFEFEF">����</td>
      <td width="123" height="22" align="center" bgcolor="#EFEFEF" >Ƶ��</td> 
      <td width="120" height="22" align="center" bgcolor="#EFEFEF" >��Ŀ</td> 
      <td width="126" align="center" bgcolor="#EFEFEF" >��Դ</td>        
      <td width="87" align="center" bgcolor="#EFEFEF" >���</td>
      <td width="93" height="22" align="center" bgcolor="#EFEFEF" >����</td>         
     </tr>         
<%
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if 
SqlItem=SqlItem  &  " order by HistrolyID DESC"
RsItem.open SqlItem,ConnItem,1,1
If (Not RsItem.Eof) and (Not RsItem.Bof) then
   RsItem.PageSize=MaxPerPage
   Allpage=RsItem.PageCount
   If Currentpage>Allpage Then Currentpage=1
   HistrolyNum=RsItem.RecordCount
   RsItem.MoveFirst
   RsItem.AbsolutePage=CurrentPage
   i_His=0
   Do While not RsItem.Eof
%>
    <tr class="tdbg"  onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">          
      <td width="56" align="center" style="border-bottom:#999999 1xp solid">          
        <input type="checkbox" value="<%=RsItem("HistrolyID")%>" name="HistrolyID" onClick="unselectall(this.form)" style="border: 0px;background-color: #E1F4EE;">         
      </td>               
      <td width="214" align="center" style="border-bottom:#999999 1xp solid">          
        <%Call Admin_ShowItem_Name(RsItem("ItemID"))%>         
      </td>         
      <td width="435" align="left" style="border-bottom:#999999 1xp solid"><%=RsItem("Title")%>
      </td>  
      <td width="123" align="center" style="border-bottom:#999999 1xp solid"><%Call Admin_ShowChannel_Name(RsItem("ChannelID"))%></td>
      <td width="120" align="center" style="border-bottom:#999999 1xp solid"><%Call Admin_ShowClass_Name(RsItem("ChannelID"),RsItem("ClassID"))%></td>
      <td width="126" align="center" style="border-bottom:#999999 1xp solid"><a href="<%=RsItem("NewsUrl")%>" target=_blank title=<%=RsItem("NewsUrl")%>>�������</a></td>     
      <td width="87" align="center" style="border-bottom:#999999 1xp solid">
      <%If RsItem("Result")=True Then
           Response.write "�ɹ�"
        ElseIf RsItem("Result")=False Then
           Response.Write "<font color=red>ʧ��</font>"
        Else
           Response.Write "<font color=red>�쳣</font>"
        End If
      %>
      </td>
      <td width="93" align="center" style="border-bottom:#999999 1xp solid">                           
      <a href="Sk_ItemHistroly.asp?Action=<%=Action%>&Del=Del&HistrolyID=<%=RsItem("HistrolyID")%>" onclick='return confirm("ȷ��Ҫɾ���˼�¼��");'>ɾ��</a>                 
      </td>         
    </tr>         
<%         
           i_His=i_His+1
           If i_His > MaxPerPage Then
              Exit Do
           End If
        RsItem.Movenext         
   Loop         
%>         
    <tr class="tdbg">          
      <td colspan=8 height="30">       
        <input name="Del" type="hidden" id="Del" value="Del">   
        <input name="Action" type="hidden" id="Action" value="<%=Action%>"> 
        <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #E1F4EE;">ȫѡ
      </td>         
    </tr> 
    <tr class="tdbg">          
      <td colspan=8 height="30" align=center>
        <input type="submit" value="���ѡ�м�¼" name="DelFlag" onclick='return confirm("ȷ��Ҫ�����ѡ��¼��");' style="cursor: hand;background-color: #cccccc;">&nbsp;&nbsp;       
        <input type="submit" value="���ʧ�ܼ�¼" name="DelFlag" onclick='return confirm("ȷ��Ҫ�������ʧ�ܼ�¼��");' style="cursor: hand;background-color: #cccccc;">&nbsp;&nbsp;
        <input type="submit" value="���ʧЧ��¼" name="DelFlag"  onclick='return confirm("ȷ��Ҫ�������ʧЧ��¼��");' style="cursor: hand;background-color: #cccccc;">&nbsp;&nbsp;
        <input type="submit" value="������м�¼" name="DelFlag" onclick='return confirm("ȷ��Ҫ������м�¼��");' style="cursor: hand;background-color: #cccccc;">
      </td>         
    </tr> 
    <tr class="tdbg">          
      <td colspan=8 height="30">          
        
      </td>         
    </tr>       
<%Else%>
<tr class="tdbg">
        <td colspan='9' class="tdbg" align="center"><br>ϵͳ��������ʷ��¼��</td>
</tr> 
<%End  If%>       
<%         
RsItem.Close         
Set RsItem=nothing           
%>         
</form>         
</table>  
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
    <tr> 
      <td height="22" colspan="2" class="tdbg">
<%
Response.Write ShowPage("Sk_ItemHistroly.asp?Action="& Action,HistrolyNum,MaxPerPage,True,True," ����¼")
%>

      </td>
    </tr>
</table>
</body>         
</html>
<%End Sub%>
<%Sub DelHistroly

Dim DelFlag
DelFlag=Trim(Request("DelFlag"))
HistrolyID=Trim(Request("HistrolyID"))
If HistrolyID<>"" Then
   HistrolyID=Replace(HistrolyID," ","")
End If
If DelFlag="���ѡ�м�¼" Then
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫɾ���ļ�¼</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [Histroly] Where HistrolyID in(" & HistrolyID & ")"
   End If
ElseIf DelFlag="���ʧ�ܼ�¼" Then
   SqlItem="Delete From [Histroly] Where Result=False"
ElseIf DelFlag="���ʧЧ��¼" Then
   Set Rs=server.createobject("adodb.recordset")
   Sql="Select ArticleID From SK_Article Where Deleted=False"
   Rs.open Sql,ConnItem,1,1
   If (Not Rs.Eof) And (Not Rs.Bof) Then
      Do While not rs.eof
         Arr_ArticleID=Arr_ArticleID & "," & CStr(rs("ArticleID"))
      Rs.MoveNext
      Loop
   End  If
   Rs.Close
   Set Rs=Nothing
   If Arr_ArticleID<>"" Then
      Arr_ArticleID=0 & Arr_ArticleID
      SqlItem="Delete From [Histroly] Where ArticleID Not In(" & Arr_ArticleID & ")"
   Else
      SqlItem="Delete From [Histroly]"
   End If
ElseIf DelFlag="������м�¼" Then
   SqlItem="Delete From [Histroly]"
Else
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫɾ���ļ�¼</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [Histroly] Where HistrolyID In(" & HistrolyID & ")"
   End If
End if

If FoundErr<>True Then
   ConnItem.Execute(SqlItem)
End If
End Sub
%>