<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Const dbdns="../"%>
<!--#include file="chk.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����ϵͳ</title>
<link rel="stylesheet" type="text/css" href="images/Style.css">

<style type="text/css">
<!--
.ttl {	CURSOR: hand; COLOR: #000000; PADDING-TOP: 4px
}
BODY {
	MARGIN-TOP: 5px; MARGIN-LEFT: 2px; BACKGROUND-COLOR: #fda700; TEXT-ALIGN: center
}
td{
	line-height:170%;
}
-->
</style>

<script language="javascript">
function showHide(obj)
{
	obj.style.display = obj.style.display == "none" ? "block" : "none";
}
</script>
</head>

<body>
<table cellspacing="1" cellpadding="2" width="150" align="center" bgcolor="#999999" 
border="0">
  <tbody>
    <tr>
      <td class="ttl" onClick="showHide(m0)" valign="top" align="left" background="images/top-bj3.jpg"><table cellspacing="0" cellpadding="0" width="127" border="0">
        <tbody>
          <tr>
            <td width="8">&nbsp;</td>
            <td align="left" width="117"><strong class="mtitle">�������</strong></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr id="m0" style="display: block">
      <td valign="top" align="middle" bgcolor="#f3f5f1"><table width="100%" cellpadding="2">
        <tbody>
          
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="ad_right.asp" target="right">������ҳ</a></td>
          </tr>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="pass.asp" target="right">�޸�����</a></td>
          </tr>
          
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="quit.asp" target="_parent" onClick="if(!confirm('���Ҫ�˳���?')) return false;">��ȫ�˳�</a></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br>
<%If InStr(Admin.GroupId,",1,")>=1 Then%>
<table cellspacing="1" cellpadding="2" width="150" align="center" bgcolor="#999999" 
border="0">
  <tbody>
    <tr>
      <td class="ttl" onClick="showHide(menu6)" valign="top" align="left" background="images/top-bj3.jpg"><table cellspacing="0" cellpadding="0" width="127" border="0">
        <tbody>
          <tr>
            <td width="8">&nbsp;</td>
            <td align="left" width="117"><strong class="mtitle">ϵͳ����</strong></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr id="menu6" style="display: block">
      <td valign="top" align="middle" bgcolor="#f3f5f1"><table width="100%" cellpadding="2">
        <tbody>
          
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Sys_Config.asp" target="right">վ����Ϣ����</a></td>
          </tr>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Sys_admin.asp" target="right">����Ա����</a></td>
          </tr>
		  <%If Db_Type="ACC" Then%>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Sys_db.asp" target="right">���ݿ����</a></td>
          </tr>
		  <%End If%>
		   <tr>
		     <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="ad_weblog.asp" target="right">��̨��־</a></td>
	        </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br />
<%End If%>
<%If InStr(Admin.GroupId,",3,")>=1 Then%>
<table cellspacing="1" cellpadding="2" width="150" align="center" bgcolor="#999999" 
border="0">
  <tbody>
    <tr>
      <td class="ttl" onClick="showHide(m2)" valign="top" align="left" background="images/top-bj3.jpg"><table cellspacing="0" cellpadding="0" width="127" border="0">
        <tbody>
          <tr>
            <td width="8">&nbsp;</td>
            <td align="left" width="117"><span class="mtitle">���Ź���</span></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr id="m2" style="display: block">
      <td valign="top" align="middle" bgcolor="#f3f5f1"><table width="100%" cellpadding="2">
        <tbody>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Class_Manage.asp?ChannelId=1&amp;ParentID=0&amp;Depth=0" target="right">���ŷ���</a></td>
          </tr>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Article_Edit.asp?ChannelId=1" target="right">��������</a> <a href="Caiji/" target="_blank" style="color:#0000FF;">�ɼ�</a></td>
          </tr>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Article_List.asp?ChannelId=1" target="right">��������</a> <a href="Article_List.asp?ChannelId=1&IsPass=0" target="right" style="color:#0000FF;">����</a> <a  href="Article_Move.asp?ChannelId=1" target="right">ת��</a></td>
          </tr>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Article_Code.asp" target="right">�ⲿ���ô���</a></td>
          </tr>
		  <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Article_List.asp?ChannelId=1&amp;IsDelete=1" target="right">����վ</a></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br />
<%End If%>
<%If InStr(Admin.GroupId,",2,")>=1 Then%>
<table cellspacing="1" cellpadding="2" width="150" align="center" bgcolor="#999999" 
border="0">
  <tbody>
    <tr>
      <td class="ttl" onClick="showHide(m1)" valign="top" align="left" background="images/top-bj3.jpg"><table cellspacing="0" cellpadding="0" width="127" border="0">
        <tbody>
          <tr>
            <td width="8">&nbsp;</td>
            <td align="left" width="117"><span class="mtitle">���˵��</span></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr id="m1" style="display: block">
      <td valign="top" align="middle" bgcolor="#f3f5f1"><table width="100%" cellpadding="2">
        <tbody>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Class_Manage.asp?ChannelId=2&ParentID=0" target="right">��Ŀ����</a></td>
          </tr>
<%
Sql="select * from Ok3w_Class where ChannelId=2 and gotoURL='' order by OrderID"
Set oRs = Conn.Execute(Sql)
Do While Not oRs.Eof
Sql="select * from Ok3w_Article where ClassID=" & oRs("ID")
Rs.Open Sql,Conn,1,3
If Rs.Eof And Rs.Bof Then
	Rs.AddNew
	Set maxRs = Conn.Execute("select max(ID) from Ok3w_Article")
	If IsNull(maxRs(0)) Then
		maxID = 1
	Else
		maxID = maxRs(0)
	End If
	maxRs.Close
	Set maxRs = Nothing
	Rs("ID") =  maxID + 1
    Rs("ChannelID") = 2
    Rs("ClassID") = oRs("ID")
	Rs("SortPath") = "0," & oRs("ID") & ","
    Rs("Title") = oRs("SortName")
    Rs("Content") = oRs("SortName")
    Rs("Author") = "Systemp"
   	Rs("ComeFrom") = "Systemp"
    Rs("AddTime") = Now()
   	Rs("Inputer") = ""
    Rs("IsPass") = 1
    Rs("IsPic") = 0
    Rs("PicFile") = ""
   	Rs("IsTop") = 0
    Rs("IsCommend") = 0
	Rs("IsDelete") = 0
	Rs("IsMove") = 0
	Rs("IsPlay") = 0
	Rs("IsUserAdd") = 0
	Rs("GiveJifen") = 0
	Rs("vUserGroupID") = 0
	Rs("vUserMore") = 1
	Rs("vUserJifen") = 0
    Rs("Hits") = 0
	Rs.Update
End If
ID = Rs("ID")
Rs.Close
%>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Article_Edit.asp?ChannelId=2&action=edit&ID=<%=ID%>" target="right"><%=oRs("SortName")%></a></td>
          </tr>
<%
	oRs.MoveNext
Loop
oRs.Close
Set oRs = Nothing
%>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br />
<%End If%>
<table cellspacing="1" cellpadding="2" width="150" align="center" bgcolor="#999999" 
border="0">
  <tbody>
    
    <tr id="m3" style="display: block">
      <td valign="top" align="middle" bgcolor="#f3f5f1"><table width="100%">
        <tbody>
          <tr>
            <td align="left">��ҳ��lenan zhang<br />
              ������zhengbi<br />
              Q &nbsp;Q��1111111</td>
          </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br />
</body>
</html>
<%
Set Rs = Nothing
Call CloseConn()
Set Admin = Nothing
%>