

<%
option explicit
response.buffer=true
%>
<!--#include file="inc/setup.asp"-->
<!--#include file="inc/cj_cls.asp"-->

<%
Dim SqlItem,RsItem,Action,ItemID,FilterName,FilterID,FilterObject,FilterType,FilterContent,FisString,FioString,FilterRep,Flag,PublicTf
Dim FoundErr,MaxPerPage,CurrentPage,AllPage,HistrolyNum,i_His,lx,DelFlag
Action=Trim(Request("Action"))
lx=Request("radiobutton")
FilterID=Trim(Request("FilterID"))
DelFlag=Trim(Request.Form("DelFlag"))
MaxPerPage=20
If lx="" or lx=0 then lx=1
Call top()
Select Case Action
Case "add"
	 If Trim(Request.Form("Saveok"))="ok" then
	 	Call Save_Data
	 Else 
	 	Call add_edit(1)
	 End if
Case "edit"
	 If Trim(Request.Form("Saveok"))="ok" then
	 	Call Save_Data
	 Else 
	 	Call add_edit(2)
	 End if
Case "flag"
	 flag=Trim(Request("flag"))
	 ConnItem.execute("update Filters set flag="& flag &"  where FilterID="&FilterID )
	 response.redirect request.ServerVariables("HTTP_REFERER")''������Դҳ����ˢ��
Case "del"
	 Select Case DelFlag
	 Case "ɾ����ѡ��¼"
	 Response.Write("sadfsadfasdfasdfasfasfd")
	 	if FilterID<>"" then 
			 ConnItem.execute("Delete From Filters Where FilterID in(" & FilterID & ")")
		end if
	 Case "ɾ��ȫ����¼"
	 	if lx=1 then ConnItem.execute("Delete From Filters where colleclx=1")
		if lx=3 then ConnItem.execute("Delete From Filters where colleclx=3")
		if lx=5 then ConnItem.execute("Delete From Filters where colleclx=5")
	 Case Else
	 	if FilterID<>"" then 
			 ConnItem.execute("Delete From Filters Where FilterID in(" & FilterID & ")")
			 response.redirect request.ServerVariables("HTTP_REFERER")''������Դҳ����ˢ��
		end if	
	 End Select
	 	Response.Redirect "sk_itemfilters.asp?radiobutton="& lx 
Case Else
	Select Case lx
	case 1
	   Call Main(1)'����
	   Call Show_Page
	case 3
	   Call Main(3)'ͼƬ
	   Call Show_Page
	case 5
	   Call Main(5)'����
	   Call Show_Page
	case 6
	   Call Main(6)'�Զ�
		Call Show_Page
	case else
	   Call Main(1)
	   Call Show_Page
	end select
End select
'�ر����ݿ�����
Call CloseConnItem()
%>
<%sub top()%>

<html>
<head>
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/Admin_Style.css">
<SCRIPT language=javascript>
function showset(thisform)
{
		if(thisform.FilterType.selectedIndex==1)
		{
        	FilterType1.style.display = "none";
			FilterType2.style.display = "";
		}
		else
		{
        	FilterType1.style.display = "";
			FilterType2.style.display = "none";
	    }

}
</script>
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
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder">
  <tr class='topbg'> 
    <td height="22" colspan="2" align="center" bgcolor="#F2F2F2" ><strong>�� �� ϵ ͳ �� �� �� ��</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder">
  <tr bgcolor="#F2F2F2" class="tdbg"> 
    <td width="65" height="30"><strong>����������</strong></td>
    <td height="30"><a href=sk_ItemFilters.asp>������ҳ</a> | <a href="?Action=add&radiobutton=<%=lx%>">�����¹���</a></td>
  </tr>
</table>
<table width="100%" align="center"  border="0" cellspacing="0" cellpadding="0" class="tableBorder">
  <tr>
    <td height=30 align="center" bgcolor="#F2F2F2"> ѡ�����ͣ�
      <input name="radiobutton" type="radio" value="1" <%if lx =1 then Response.Write "checked" %>  onClick="location.href='?radiobutton=1';" >
      ���Ųɼ�
      <input type="radio" name="radiobutton" value="3"  <%if lx =3  then Response.Write "checked" %> onClick="location.href='?radiobutton=3';">
      ͼƬ�ɼ�
      <input type="radio" name="radiobutton" value="5"  <%if lx =5  then Response.Write "checked" %> onClick="location.href='?radiobutton=5';">
    �����ɼ� </td>
  </tr>
</table>

<%
end sub 
'---------�����б�------------------
Sub Main(lx)
%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>�� �� �� ¼</strong></div></td>
    </tr>
</table>
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <form name="form1" method="POST" action="?action=del">
    <tr class="tdbg" style="padding: 0px 2px;"> 
      <td width="57" height="22" align="center" bgcolor="#EFEFEF" class=ButtonList>ѡ��</td>
      <td width="176" align="center" bgcolor="#EFEFEF" class=ButtonList>�����ɼ���Ŀ</td>
      <td width="246" align="center" bgcolor="#EFEFEF" class=ButtonList>��������</td>
      <td width="114" height="22" align="center" bgcolor="#EFEFEF" class=ButtonList>��������</td>
      <td width="134" align="center" bgcolor="#EFEFEF" class=ButtonList>��������</td>
	  <td width="30" align="center" bgcolor="#EFEFEF" class=ButtonList>״̬</td>
      <td width="154" height="22" align="center" bgcolor="#EFEFEF" class=ButtonList>����</td>
    </tr>
    <%                          
Set RsItem=server.createobject("adodb.recordset")
SqlItem="select * from Filters where colleclx=" & lx 
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if 

SqlItem=SqlItem  &  " order by FilterID DESC"
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
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
      <td width="57" align="center"> <input type="checkbox" value="<%=RsItem("FilterID")%>" name="FilterID" onClick="unselectall(this.form)" style="border: 0px;background-color: #E1F4EE;">      </td>
      <td width="176" align="center"> 
	  <%dim rs : set RS=ConnItem.execute("select * from Item where ItemID=" & RsItem("ItemID"))
	  if rs.eof then
	  	Response.Write("û����Ŀ!")
	  else
	  	Response.Write RS("ItemName") 
	  end if
	   rs.close : set rs=nothing%> </td>
      <td width="246" align="left"><%=RsItem("FilterName")%></td>
      <td width="114" align="center"><%if RsItem("FilterObject")=1 then Response.Write "�������" else Response.Write "���Ĺ���" end if%></td>
      <td align="center"> <%if RsItem("FilterType")=1 then Response.Write "���滻" else Response.Write "�߼��滻" end if%></td>
      <td width="30" align="center">
	  <%if RsItem("Flag")=1 then Response.Write "��" else Response.Write "<font color=""#FF0000"">��</font>" end if%>
	  </td>
	  <td width="154" align="center">
	  <%if RsItem("Flag")=0 then Response.Write "<a href=""?Action=flag&flag=1&FilterID="& RsItem("FilterID") &"&radiobutton="& lx &""">����</a>" else Response.Write "<a href=""?Action=flag&flag=0&FilterID="& RsItem("FilterID") &"&radiobutton="& lx &""">����</a>" end if%>
	   &nbsp;<a href="?Action=edit&FilterID=<%=RsItem("FilterID")%>&radiobutton=<%=lx%>" >�༭</a> &nbsp;
      <a href="?Action=del&FilterID=<%=RsItem("FilterID")%>&radiobutton=<%=lx%>" onclick='return confirm("ȷ��Ҫɾ���˼�¼��");'>ɾ��</a>      </td>
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
        <input name="Action" type="hidden" id="Action" value="ok"> 
		<input name="radiobutton" type="hidden" id="Action" value="<%=lx%>"> 
        <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #E1F4EE;">
        ȫѡ </td>
    </tr> 
    <tr class="tdbg"> 
      <td colspan=8 height="30" align=center>&nbsp;&nbsp;&nbsp;&nbsp; <input name="DelFlag" type="submit" class="lostfocus" style="cursor: hand;background-color: #cccccc;"  onclick='return confirm("ȷ��Ҫɾ��ѡ�еļ�¼��");' value="ɾ����ѡ��¼"> 
        &nbsp;&nbsp;&nbsp;&nbsp; <input name="DelFlag" type="submit" class="lostfocus" style="cursor: hand;background-color: #cccccc;" onclick='return confirm("ȷ��Ҫɾ�����еļ�¼��");' value="ɾ��ȫ����¼"> 
      &nbsp;&nbsp;
	  </td></tr>
    <tr class="tdbg"> 
      <td colspan=8 height="30"> </td>
    </tr>
    <%Else%>
    <tr class="tdbg"> 
      <td colspan='9' class="tdbg" align="center"><br>
        ϵͳ�����޼�¼��</td>
    </tr>
    <%End  If%>
    <%         
RsItem.Close         
Set RsItem=nothing           
%>
  </form>
</table>  
<%End Sub%>

<% '----------��ʾ����һҳ ��һҳ������Ϣ-----------
sub Show_Page() %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
    <tr> 
      <td height="22" colspan="2" class="tdbg">
<%
Response.Write ShowPage("sk_checkDatabase.asp?Action="& Action,HistrolyNum,MaxPerPage,True,True," ����¼")
%>

      </td>
    </tr>
</table>
<%end sub

'--------------�������Ӻͱ༭--------------
sub add_edit(Type_1)
If Type_1=2 then Call Test
%>
<form method="post" action="?Action=<% if Type_1 =1 then Response.Write("add") else Response.Write("edit") end if %>" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" >                                
    <tr>                                 
      <td height="22" colspan="2" class="title"> <div align="center"><strong><% if Type_1=1 then Response.Write "�� �� �� ��" else  Response.Write "��  �� �� ��" end if %></strong></div></td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> �������ƣ�</strong></td>                                
      <td class="tdbg"><input name="FilterName" type="text" id="FilterName" value="<%=FilterName%>" size="25" maxlength="30">                                 
        &nbsp;</td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> ������Ŀ��</strong></td>                                
      <td class="tdbg"><%Call Admin_ShowItem_Option(ItemID,lx)%>      </td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> ���˶���<strong></td>                               
      <td class="tdbg">                               
         <select name="FilterObject" id="FilterObject">                              
            <option value="1" <%if FilterObject=1 Then Response.Write "selected"%>>�������</option>                              
            <option value="2" <%if FilterObject=2 or FilterObject="" or FilterObject=0  Then Response.Write "selected"%>>���Ĺ���</option>             
         </select>      </td>                               
    </tr>                               
    <tr class="tdbg">                                
      <td width="150" class="tdbg"><strong> �������ͣ�</strong></td>                               
      <td class="tdbg">                               
         <select name="FilterType" id="FilterType" onchange=showset(this.form)>                              
            <option value="1" <%if FilterType=1 Then Response.Write "selected"%> >���滻</option>                              
            <option value="2" <%if FilterType=2 Then Response.Write "selected"%>>�߼�����</option>                              
         </select>      </td>                                
    </tr>     
    <tr class="tdbg">                                
      <td width="150" class="tdbg"><strong> ʹ��״̬��</strong></td>                               
      <td class="tdbg">                            
         <select name="Flag" id="Flag">    
		 	<option value="0" <%If Flag=0 Then Response.Write "selected"%>>����</option>                           
            <option value="1" <%If Flag=1 or Flag="" Then Response.Write "selected"%>>����</option>                            
         </select>      </td>                                
    </tr>                                                                         
</table>                                
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" id="FilterType1" style="display:<%if FilterType<>1 and FilterType<>"" Then Response.Write "none"%>">                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> ���ݣ�</strong></td>                               
      <td class="tdbg"><textarea name="FilterContent" cols="70" rows="5" class="lostfocus"><%=FilterContent%></textarea></td>                               
    </tr>                            
</table>                 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" id="FilterType2" style="display:<%if FilterType<>2 Then Response.Write "none"%>">                                   
    <tr class="tdbg">                               
      <td width="150" class="tdbg"><strong> ��ʼ��ǣ�</strong></td>                              
      <td class="tdbg"><textarea name="FisString" cols="70" rows="5" class="lostfocus"><%=FisString%></textarea>                               
        &nbsp;</td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> ������ǣ�</strong></td>                               
      <td class="tdbg"><textarea name="FioString" cols="70" rows="5" class="lostfocus"><%=FioString%></textarea>                                
        &nbsp;</td>                                
    </tr>                         
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" >
    <tr class="tdbg"> 
      <td width="150" class="tdbg"><strong> �滻��</strong></td>
      <td class="tdbg"><textarea name="FilterRep" cols="70" rows="5" class="lostfocus" id="FilterRep"><%=FilterRep%></textarea></td>
    </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" >
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg">
        <input name="Saveok" type="hidden" id="Saveok" value="ok">
		<input name="radiobutton" type="hidden" id="Action" value="<%=lx%>"> 
        <input name="FilterID" type="hidden" id="FilterID" value="<%=FilterID%>">
        <input name="Cancel" type="button" id="Cancel" value="��&nbsp;&nbsp;��" onClick="window.location.href='sk_ItemFilters.asp'" style="cursor: hand;background-color: #cccccc;">
        &nbsp;
        <input  type="submit" name="Submit" value="ȷ&nbsp;&nbsp;��" style="cursor: hand;background-color: #cccccc;"> 
      </td>
    </tr>
</table>
</form>
         
</body>         
</html>
<%End Sub%>

<%
Sub Save_Data
FilterID=Trim(Request("FilterID"))
FilterName=Trim(Request.Form("FilterName"))
ItemID=Trim(Request.Form("ItemID"))
FilterObject=Trim(Request.Form("FilterObject"))
FilterType=Trim(Request.Form("FilterType"))
FilterContent=Request.Form("FilterContent")
FisString=Request.Form("FisString")
FioString=Request.Form("FioString")
FilterRep=Request.Form("FilterRep")
Flag=Trim(Request.Form("Flag"))
'PublicTf=Trim(Request.Form("PublicTf"))
If Action<>"edit" and  Action<>"add" then exit sub
If Action="edit" then 
	If FilterID="" Then                                
	   FoundErr=True                                
	   ErrMsg=ErrMsg & "<br><li>�������������Ч���ӽ���</li>"
	Else
	   FilterID=Clng(FilterID)                                 
	End If                                
End If
If FilterName="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>�������Ʋ���Ϊ��</li>"                                 
End If                                
If ItemID="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>��ѡ�����������Ŀ</li>"                                 
Else                                
   ItemID=Clng(ItemID)  
   If ItemID=0 Then  
      FoundErr=True                                
      ErrMsg=ErrMsg & "<br><li>��ѡ�����������Ŀ</li>"   
   End If                                
End If       
If FilterObject="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>��ѡ����˶���</li>"                                 
Else                                
   FilterObject=Clng(FilterObject)   
End If                                
                            
If FilterType="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>��ѡ���������</li>"                                 
Else                                
   FilterType=Clng(FilterType)                                
   If FilterType=1 Then     
      If FilterContent="" Then                              
         FoundErr=True                                
         ErrMsg=ErrMsg & "<br><li>���˵����ݲ���Ϊ��</li>"   
      End If   
   ElseIf FilterType=2 Then   
      If FisString="" or FioString="" Then   
         FoundErr=True                                
         ErrMsg=ErrMsg & "<br><li>��ʼ/������ǲ���Ϊ��</li>"   
      End If   
   Else   
      FoundErr=True                                
      ErrMsg=ErrMsg & "<br><li>�������������Ч���ӽ���</li>"   
   End If   
End If                     
If FoundErr<>True Then     
   if Action="add" then                            
   		SqlItem ="select top 1 *  from Filters"
   else 
   		SqlItem ="select top 1 *  from Filters Where FilterID=" & FilterID                                
   end if
   Set RsItem=server.CreateObject("adodb.recordset")                                
   RsItem.open SqlItem,ConnItem,2,3
   if Action="add" then RsItem.addnew                                           
   RsItem("FilterName")=FilterName                                
   RsItem("ItemID")=ItemID   
   RsItem("FilterObject")=FilterObject                                
   RsItem("FilterType")=FilterType                                
   If FilterType=1 Then                                
      RsItem("FilterContent")=FilterContent                                
   ElseIf FilterType=2 Then                                
      RsItem("FisString")=FisString                                
      RsItem("FioString")=FioString                                
   End If                                                
   RsItem("FilterRep")=FilterRep        
   RsItem("Flag")=Flag
   RsItem("colleclx")=lx  
   'RsItem("PublicTf")=PublicTf                        
   RsItem.Update                                
   RsItem.Close                                
   Set RsItem=Nothing 
   	Response.Redirect "SK_ItemFilters.asp?radiobutton=" &lx                         
Else                                
   Call WriteErrMsg(ErrMsg)                                
End If                  
End Sub

Sub Test
Set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from Filters Where FilterID=" & FilterID         
RsItem.open SqlItem,ConnItem,1,1      
If Not RsItem.Eof Then
   ItemID=RsItem("ItemID")
   FilterName=RsItem("FilterName")
   FilterObject=RsItem("FilterObject")
   FilterType=RsItem("FilterType")
   FilterContent=RsItem("FilterContent")
   FisString=RsItem("FisString")
   FioString=RsItem("FioString")
   FilterRep=RsItem("FilterRep")
   Flag=RsItem("Flag")
   'PublicTf=RsItem("PublicTf")
Else
   FoundErr=True
   ErrMsg=ErrMsg & "<br></li>���������Ҳ�������Ŀ</li>"
End If
RsItem.Close
Set RsItem=Nothing
End Sub
%>