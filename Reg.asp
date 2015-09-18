<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "./"
Const Htmldns = "./"
Const Base_Target = ""
ClassID = ""
%>
<!--#include file="AppCode/Conn.asp"-->
<!--#include file="AppCode/fun/function.asp"-->
<!--#include file="AppCode/Pager.asp"-->
<!--#include file="vbs.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>会员注册_<%=Application("Ok3w_SiteName")%></title>
<script language="javascript" src="js/js.js"></script>
<link href="css/css2.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="head.asp"-->
<table width="970" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF"><table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="left" valign="top">
	<div class="pathnav">您当前位置：<a href="./">网站首页</a> &gt;&gt; 会员注册</div>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="cbox">
        
        <tr>
          <td valign="top"  style="padding:5px;">
		  <div style="padding:8px;border:1px solid #EBEBEB; background-color:#f2f6fb; font-size:14px; font-weight:bold;">欢迎加入我们^_^</div>
		  <div style="border:1px solid #EBEBEB; margin-top:4px;">
		    
		      <table border="0" cellspacing="0" cellpadding="5"><form id="form1" name="form1" method="post" action="save.asp">
<script language="javascript" src="images/DatePicker/WdatePicker.js"></script>			  
<script language="javascript">
function chkuser(str)
{
	var patrn=/^[a-zA-Z0-9_\u4E00-\u9FA5]{4,20}$/;
	if(!patrn.test(str))
	{
		alert("用户名不正确：\n\n它应该是由4-20位的中文、英文、数字及下划线组成的字符");
		return false;
	}
	else
		return true;
}

function chkpass(str)
{
	var patrn=/^[a-zA-Z0-9_]{6,20}$/;
	if(!patrn.test(str))
	{
		alert("密码不正确：\n\n它应该是由6-20位的英文、数字及下划线组成的字符");
		return false;
	}
	else
		return true;
}

function chkemail(str)
{
	var patrn= /^[_a-zA-Z0-9\-]+(\.[_a-zA-Z0-9\-]*)*@[a-zA-Z0-9\-]+([\.][a-zA-Z0-9\-]+)+$/;
	if(!patrn.test(str))
	{
		alert("邮箱地址格式错误，请重新输入");
		return false;
	}
	else
		return true;
}

function ChkReg(frm)
{
	if(!chkuser(frm.User_Name.value))
	{
		frm.User_Name.focus();
		return false;
	}
	if(!chkpass(frm.User_Password.value))
	{
		frm.User_Password.focus();
		return false;
	}
	if(frm.User_Password.value!=frm.User_Password2.value)
	{
		alert("两次输入的密码不一致，请重新输入");
		frm.User_Password2.focus();
		return false;
	}
	if(!chkemail(frm.Mail.value))
	{
		frm.Mail.focus();
		return false;
	}
	
	frm.bntSubmit.disabled=true;
	
	frm.submit();
}
</script>			  
                <tr>
                  <td colspan="2" class="red12">登陆信息（必填）</td>
                </tr>
                <tr>
                  <td align="right">用户名：</td>
                  <td><input name="User_Name" type="text" class="ginput" id="User_Name" size="20" maxlength="20" />
                    <span class="red12real">*</span>4-20位字符，可以是中文</td>
                </tr>
                <tr>
                  <td align="right">密码：</td>
                  <td><input name="User_Password" type="password" class="ginput" id="User_Password" size="20" maxlength="20" />
                    <span class="red12real">*</span>6-20位字符</td>
                </tr>
                <tr>
                  <td align="right">密认密码：</td>
                  <td><input name="User_Password2" type="password" class="ginput" id="User_Password2" size="20" maxlength="20" />
                    <span class="red12real">*</span></td>
                </tr>
                <tr>
                  <td align="right">邮箱：</td>
                  <td><input name="Mail" type="text" class="ginput" id="Mail" size="35" maxlength="50" />
                    <span class="red12real">*</span>请填最常用邮箱，方便联系</td>
                </tr>
                <tr>
                  <td colspan="2" class="red12">个人信息（选填）</td>
                </tr>
                <tr>
                  <td align="right">姓名：</td>
                  <td><input name="Name" type="text" class="ginput" id="Name" size="10" maxlength="8" /></td>
                </tr>
                <tr>
                  <td align="right">性别：</td>
                  <td><input type="radio" name="Sex" value="男" />
                    男
                      <input type="radio" name="Sex" value="女" />
                      女 
                      <input name="Sex" type="radio" value="保密" checked="checked" />
                      保密</td>
                </tr>
                <tr>
                  <td align="right">出生年月日：</td>
                  <td><input name="Birthday" type="text" class="ginput" id="Birthday" size="10" maxlength="8" onClick="WdatePicker()" readonly="readonly" /></td>
                </tr>
                <tr>
                  <td align="right">电话：</td>
                  <td><input name="Tel" type="text" class="ginput" id="Tel" size="15" maxlength="15" /></td>
                </tr>
                <tr>
                  <td align="right">QQ：</td>
                  <td><input name="QQ" type="text" class="ginput" id="QQ" size="15" maxlength="15" /></td>
                </tr>
                <tr>
                  <td align="right">联系地址：</td>
                  <td><input name="Address" type="text" class="ginput" id="Address" size="35" maxlength="50" /></td>
                </tr>
                <tr>
                  <td align="right">邮编：</td>
                  <td><input name="Zip" type="text" class="ginput" id="Zip" size="10" maxlength="6" /></td>
                </tr>
                <tr>
                  <td align="right">自我介绍：</td>
                  <td><textarea name="Content" cols="60" rows="8" class="ginput" id="Content"></textarea></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><textarea name="textarea" cols="60" rows="8" class="ginput"><%=Application("Ok3w_SiteUserRegTK")%></textarea></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><input name="bntSubmit" type="button" class="gbnt" id="bntSubmit" onClick="ChkReg(this.form);" value="同意以上《注册条款》并提交注册" />
                    <input name="action" type="hidden" id="action" value="Reg" /></td>
                </tr>
                
</form>
              </table>
             
		    </div>
		  
		  </td>
        </tr>
      </table>
    </td>
    <td width="353" align="right" valign="top">
	<!--#include file="right.asp"-->
    </td>
  </tr>
</table></td>
  </tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
