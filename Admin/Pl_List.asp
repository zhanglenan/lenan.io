<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Const dbdns="../"%>
<!--#include file="chk.asp"-->
<!--#include file="../AppCode/Pager.asp"-->
<!--#include file="../AppCode/fun/function.asp"-->
<!--#include file="../AppCode/fun/CreateHtml.asp"-->
<!--#include file="../AppCode/Class/Ok3w_Guest.asp"-->
<%
TypeID = Request.QueryString("TypeID")
Select Case TypeID
	Case 2
		Call CheckAdminFlag(3)
	Case 3
		Call CheckAdminFlag(6)
	Case Else
		Response.End()
End Select

PageNo = Request.QueryString("PageNo")
If PageNo="" Then PageNo = "1"
IsPass = Request.QueryString("IsPass")
IdList = Request.Form("IdList")

If IdList<>"" Then
	Set Guest = New Ok3w_Guest
	
	Select Case Trim(Request.Form("Cmd"))
		Case "删除"
			Call Guest.Del(IdList)
			Set Guest = Nothing
			Tmp = Split(IdList,",")
			Call SaveAdminLog("删除留言，ID=" & IdList)
			Call ActionOk("Pl_List.asp?IsPass=" & IsPass & "&PageNo=" & PageNo & "&TypeID=" & TypeID)
		Case "开通"
				Call Guest.Pass(1,IdList)
				Set Guest = Nothing
				Tmp = Split(IdList,",")
				Response.Write(IdList)
				Call SaveAdminLog("回复/开通/关闭留言，ID=" & IdList)
				Call ActionOk("Pl_List.asp?IsPass=" & IsPass & "&PageNo=" & PageNo & "&TypeID=" & TypeID)
		Case "关闭"
				Call Guest.Pass(0,IdList)
				Set Guest = Nothing
				Tmp = Split(IdList,",")
				Call SaveAdminLog("回复/开通/关闭留言，ID=" & IdList)
				Call ActionOk("Pl_List.asp?IsPass=" & IsPass & "&PageNo=" & PageNo & "&TypeID=" & TypeID)
	End Select
End If
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理系统</title>
<link rel="stylesheet" type="text/css" href="images/Style.css">
<script language="javascript">
function ChkAll()
{
	var obj = document.form2.IdList;
	for(var i=0;i<obj.length;i++)
		obj[i].checked = !obj[i].checked;
}
</script>
</head>

<body>
<!--#include file="top.asp"-->
<br />
<table cellspacing="0" cellpadding="0" width="98%" align="center" border="0">
  <tbody>
    <tr>
      <td style="PADDING-LEFT: 2px; HEIGHT: 22px" 
    background="images/tab_top_bg.gif"><table cellspacing="0" cellpadding="0" width="477" border="0">
        <tbody>
          <tr>
            <td width="147"><table height="22" cellspacing="0" cellpadding="0" border="0">
              <tbody>
                <tr>
                  <td width="3"><img id="tabImgLeft__0" height="22" 
                  src="images/tab_active_left.gif" width="3" /></td>
                  <td class="mtitle" 
                background="images/tab_active_bg.gif">留言/评论管理</td>
                  <td width="3"><img id="tabImgRight__0" height="22" 
                  src="images/tab_active_right.gif" 
            width="3" /></td>
                </tr>
              </tbody>
            </table></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr>
      <td bgcolor="#ffffff"><table cellspacing="0" cellpadding="0" width="100%" border="0">
        <tbody>
          <tr>
            <td width="1" background="images/tab_bg.gif"><img height="1" 
            src="images/tab_bg.gif" width="1" /></td>
            <td 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          valign="top"><div id="tabContent__0" style="DISPLAY: block; VISIBILITY: visible">
              <table cellspacing="1" cellpadding="1" width="100%" align="center" 
            bgcolor="#8ccebd" border="0">
                <tbody>
                  <tr>
                    <td 
                style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
                valign="top" bgcolor="#fffcf7"><%
stype = Trim(Request.QueryString("stype"))
keyword = Trim(Request.QueryString("keyword"))
Sql = "select * from Ok3w_Guest where 1=1 and TypeID=" & TypeID
If IsPass<>"" Then Sql=Sql & " and IsPass=" & IsPass
If keyword<>"" Then Sql = Sql & " and " & stype & " like '%" & keyword & "%'"
Sql = Sql & " order by Id desc"
Set Page = New TurnPage
Call Page.GetRs(Conn,Rs,Sql,10)
%>
                      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
                        <form id="form1" name="form1" method="get" action="">
                          <tr>
                            <td height="30" colspan="6" align="left" bgcolor="#EBEBEB">&nbsp;
                              <select name="stype" id="stype">
                                <option value="UserName">按用户</option>
                              <option value="Title">按标题</option>
                              <option value="Content">按内容</option>
                              </select>
                                <input name="keyword" type="text" id="keyword" /><input name="TypeID" type="hidden" value="<%=TypeID%>">
                                <input type="submit" name="Submit2" value="搜索" /></td>
                          </tr>
                        </form>
                        <tr>
                          <td height="25" align="center" bgcolor="#EBEBEB">发布者信息</td>
                          <td align="center" bgcolor="#EBEBEB">内容</td>
                          <td align="center" bgcolor="#EBEBEB">状态</td>
                          <td align="center" bgcolor="#EBEBEB">回复（点击回复）</td>
                          <td align="center" bgcolor="#EBEBEB">选择<br>
                            <input type="checkbox" onClick="ChkAll()"></td>
                          </tr>
                        <form id="form2" name="form2" method="post" action="?IsPass=<%=IsPass%>&PageNo=<%=PageNo%>&TypeID=<%=TypeID%>">
                          <%
	Do While Not Rs.Eof And Not Page.Eof
	%>
                          <tr>
                            <td height="25" align="center" bgcolor="#FFFFFF">
                              <textarea name="textarea" cols="40" rows="6">姓名：<%= Rs("UserName")%>
标题：<%= Rs("Title")%>
邮箱：<%= Rs("Mail")%>
主页：<%= Rs("Homepage")%>
OICQ：<%= Rs("QQ")%>
时间：<%= Rs("AddTime")%>
IP：<%= Rs("Ip")%></textarea></td>
                            <td align="center" bgcolor="#FFFFFF"><span style="color:#0000FF">评论：《<%If Rs("TypeID")=2 Then%><%=ExecSqlReturnOneValue("select Title from Ok3w_Article where ID=" & Rs("TableID"))%><%Else%><%=ExecSqlReturnOneValue("select SoftName from Ok3w_Soft where ID=" & Rs("TableID"))%><%End If%>》</span><br>
                              <textarea name="Content" cols="50" rows="6" id="Content"><%= Rs("Content")%></textarea></td>
                            <td align="center" bgcolor="#FFFFFF"><%If Rs("IsPass") Then%>开通<%Else%><font color="#FF0000">关闭</font><%End If%></td>
                            <td align="center" bgcolor="#FFFFFF"><a href="#" onClick="g_show(<%=Rs("ID")%>)"><%If Rs("Ad_Ask")="" Then%><font color="#FF0000">未回复</font><%Else%>已回复<%End If%></a></td>
                            <td align="center" bgcolor="#FFFFFF"><input name="IdList" type="checkbox" id="IdList" value="<%=Rs("ID")%>"></td>
                            </tr> 
                          <%
		Rs.MoveNext
		Page.MoveNext
	Loop
	Rs.Close
	%>
                         
                          <tr>
                            <td height="25" colspan="6" align="right" bgcolor="#FFFFFF"><input name="Cmd" type="submit" id="Cmd" value="开通">
                                <input name="Cmd" type="submit" id="Cmd" value="关闭">
                                <input name="Cmd" type="submit" id="Cmd" onClick="if(!confirm('真的要删除吗？')) return false;" value="删除"></td>
                          </tr> </form>
                          <tr>
                            <td height="25" colspan="6" bgcolor="#FFFFFF"><%Call Page.GetPageList()%></td>
                          </tr>
                      </table></td>
                  </tr>
                </tbody>
              </table>
            </div></td>
            <td width="1" background="images/tab_bg.gif"><img height="1" 
            src="images/tab_bg.gif" width="1" /></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr>
      <td background="images/tab_bg.gif" bgcolor="#ffffff"><img height="1" 
      src="images/tab_bg.gif" width="1" /></td>
    </tr>
  </tbody>
</table>
<div id="g_edit" style="z-index:9999; position:absolute; top:85px; left:20px; display:none;">
<iframe scrolling="auto" id="g_url" name="g_url" width="500" height="400" frameborder="1"></iframe>
<script language="javascript">
function g_show(id)
{
	document.getElementById("g_edit").style.display = "";
	g_url.location.href = "Guest_Edit.asp?ID="+id +"&TypeID=2";
}
function g_hidde()
{
	document.getElementById("g_edit").style.display = "none";
}
</script>
</div>
</body>
</html>
<%
Set Guest = Nothing
Set Rs = Nothing
Call CloseConn()
Set Admin = Nothing
%>

