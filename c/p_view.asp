<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "../"
Const DisSmallClass = False
Const Htmldns = "../"
%>
<!--#include file="../AppCode/Conn.asp"-->
<!--#include file="../AppCode/fun/function.asp"-->
<!--#include file="../AppCode/Class/Ok3w_Article.asp"-->
<!--#include file="../AppCode/Class/Ok3w_User.asp"-->
<!--#include file="../vbs.asp"-->
<%
id=myCdbl(Request.QueryString("id"))
Set Article = New Ok3w_Article
Call Article.HitsAdd(Id)
Call Article.Load(Id)
If Article.IsPass=0 Then Call Page_Err("文章已经关闭")
If Article.IsDelete=1 Then Call Page_Err("文章已经删除")

Set User = New Ok3w_User
User_Name = Replace(User.IsLogin(),"'","")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=Article.Title%></title>
<script language="JavaScript" src="../images/js.js" type="text/javascript"></script>
<link href="../css/css2.css" rel="stylesheet" type="text/css" />
</head>

<body style="background-color:#FFFFFF; background-image:url();">
<div style="font-size:14px; line-height:14px; margin-top:1px;">
<%If User_Name="" Then%>
<div style="color:#0000FF;"><br />提示：只有本站会员<%If Article.vUserJifen<>0 Then%>，且积分大于<strong><%=Article.vUserJifen%></strong><%End If%><%If Article.vUserGroupID<>0 Then%>，且等级属于<strong><%=GetUserDengjiValue(0,Article.vUserGroupID)%></strong><%If Article.vUserMore=1 Then%>及以上<%End If%><%End If%>才能查看该文章。</div>
<br /><br />
<strong>请登陆...</strong>
<br /><br />
<form id="frmuserlogin" name="frmuserlogin" method="post" target="_parent" action="../save.asp?ArticleID=<%=Article.ID%>" onSubmit="return ChkUserLogin(this);">
  帐号：
    <input name="User_Name" type="text" class="ainput" id="User_Name" size="13" maxlength="20" />
  密码：
  <input name="User_Password" type="password" class="ainput" id="User_Password" size="13" maxlength="20" />
  <input name="action" type="hidden" id="action" value="Login" />
  <input type="submit" name="Submit" value=" 登录 " class="abnt" />
  <input type="button" name="Submit2" value=" 注册 "  class="abnt" onClick="parent.location.href='../Reg.asp';" />
</form>
<br /><br />
<%Else
Jifen = Conn.Execute("select Jifen from Ok3w_User where User_Name='" & User_Name & "'")(0)
GroupID = GetUserDengjiID(Jifen)
If (Jifen>=Article.vUserJifen And Article.vUserJifen<>0) Or (GroupID=Article.vUserGroupID Or (GroupID>Article.vUserGroupID And Article.vUserMore=1) And Article.vUserGroupID<>0) Then
%>
<%Call OutThisPageContent(Article.ID,Article.Content,"asp")%>
<%Else%>
<br />提示：只有本站会员<%If Article.vUserJifen<>0 Then%>，且积分大于<strong><%=Article.vUserJifen%></strong><%End If%><%If Article.vUserGroupID<>0 Then%>，且等级属于<strong><%=GetUserDengjiValue(0,Article.vUserGroupID)%></strong><%If Article.vUserMore=1 Then%>及以上<%End If%><%End If%>才能查看该文章。<br /><br /><br />
你目前的积分是：<strong><%=Jifen%></strong> 等级为：<strong><%=GetUserDengjiValue(0,GroupID)%></strong><br /><br />
<%End If%>
<%End If%>
</div>
</body>
</html>
<%
Set Rs = Nothing
Call CloseConn()
%>