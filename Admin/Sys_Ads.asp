<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%dbdns="../"%>
<!--#include file="chk.asp"-->
<!--#include file="../AppCode/Class/Ok3w_AdSense.asp"-->
<!--#include file="../AppCode/fun/html-js.asp"-->
<!--#include file="../AppCode/fun/inc_file.asp"-->
<%
Call CheckAdminFlag(1)

Set Ads = New Ok3w_AdSense

Select Case Trim(Request.Form("action"))
	Case "add"
		Call Ads.Add()
		Call SaveAdminLog("添加广告：" & Ads.Title)
		Call CloseConn()
		Call ActionOk("Sys_Ads.asp")
	Case "edit"
		Call Ads.Edit()
		Call SaveAdminLog("修改广告(ID=" & Ads.ID & ")为：" & Ads.Title)
		Call CloseConn()
		Call ActionOk("Sys_Ads.asp")
	Case "del"
		Call Ads.Del()
		Call SaveAdminLog("删除广告(ID=" & Ads.ID & ")")
		Call CloseConn()
		Call ActionOk("Sys_Ads.asp")
End Select
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理系统</title>
<link rel="stylesheet" type="text/css" href="images/Style.css">
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
                background="images/tab_active_bg.gif">广告管理</td>
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
                valign="top" bgcolor="#fffcf7"><table border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                      <form action="" method="post">
                        <tr>
                          <td align="center" bgcolor="#EBEBEB">广告编号</td>
                          <td align="center" bgcolor="#EBEBEB">广告描述</td>
                          <td align="center" bgcolor="#EBEBEB">广告代码</td>
                          <td align="center" bgcolor="#EBEBEB">操作</td>
                        </tr>
                        <tr>
                          <td align="center" bgcolor="#FFFFFF"><input name="SN" type="text" id="SN" value="<%=Ads.GetMaxSN()+1%>" size="4" readonly="readonly"></td>
                          <td align="center" bgcolor="#FFFFFF"><input name="Title" type="text" id="Title" value="" size="25"></td>
                          <td align="center" bgcolor="#FFFFFF"><textarea name="Code" cols="40" rows="5" id="Code"></textarea></td>
                          <td align="center" bgcolor="#FFFFFF"><input type="button" name="Submit" value="添加" onClick="javascript:formsubmit(this.form,'add');" />
                              <input name="B_Time" type="hidden" id="B_Time" value="<%=Now()%>" />
                              <input name="E_Time" type="hidden" id="E_Time" value="<%=Now()%>" />
                              <input name="Hits" type="hidden" id="Hits" value="0" />
                              <input name="Key" type="hidden" id="Key" value="1" />
<input name="action" type="hidden" id="action" /></td>
                        </tr>
                      </form>
                    </table>
                      <br />
                      <table border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                        <tr>
                          <td align="center" bgcolor="#EBEBEB">广告编号</td>
                          <td align="center" bgcolor="#EBEBEB">广告描述</td>
                          <td width="310" align="center" bgcolor="#EBEBEB">广告代码</td>
                          <td align="center" bgcolor="#EBEBEB">JS调用代码</td>
                          <td align="center" bgcolor="#EBEBEB">操作</td>
                        </tr>
                        <%
sql = "select * from Ok3w_AdSense order by SN,ID"
Rs.Open sql,Conn,1,1
n = 0
Do While Not Rs.Eof
n = n + 1
%>
                        <form id="form1" name="form1" method="post" action="">
                          <tr>
                            <td align="center" bgcolor="#FFFFFF"><input name="SN" type="text" id="SN" value="<%=Rs("SN")%>" size="4" readonly="readonly"></td>
                            <td align="center" bgcolor="#FFFFFF"><input name="Title" type="text" id="Title" value="<%=Rs("Title")%>" size="25"></td>
                            <td align="center" bgcolor="#FFFFFF"><a href="###" onClick="ads<%=n%>.style.display=ads<%=n%>.style.display=='none'?'block':'none'">显示/隐藏/编辑</a>
							<div style="display:none;" id="ads<%=n%>">	
							<textarea name="Code" cols="40" rows="5" id="Code"><%=Rs("Code")%></textarea>
							</div>
							</td>
                            <td align="center" bgcolor="#FFFFFF"><input name="textfield" type="text" value="<%="<%" & "=GetAdSense(" & Rs("SN") &")%" & ">" %>" size="20"></td>
                            <td align="center" bgcolor="#FFFFFF"><input type="button" name="Submit4" value="修改" onClick="javascript:formsubmit(this.form,'edit');"  /> 
                              <input type="button" name="Submit5" value="删除" onClick="javascript:if(!confirm('真的要删除吗？')) return false;formsubmit(this.form,'del');"  />
                                <input name="ID" type="hidden" id="ID" value="<%=Rs("ID")%>" />
                                <input name="action" type="hidden" id="action" />
                                <input name="B_Time" type="hidden" id="B_Time" value="<%=Now()%>" />
                                <input name="E_Time" type="hidden" id="E_Time" value="<%=Now()%>" />
                                <input name="Hits" type="hidden" id="Hits" value="0" />
                                <input name="Key" type="hidden" id="Key" value="1" /></td>
                          </tr>
                        </form>
                        <%
	Rs.MoveNext
Loop
Rs.Close
%>
                      </table>
<script language="JavaScript" type="text/javascript">
function formsubmit(frm,action)
{
	if(frm.SN.value.trim()=="")
	{
		ShowErrMsg("广告编号不能为空，请输入");
		frm.SN.focus();
		return false;
	}
	if(frm.Title.value.trim()=="")
	{
		ShowErrMsg("广告描述不能为空，请输入");
		frm.Title.focus();
		return false;
	}
	if(frm.Code.value.trim()=="")
	{
		ShowErrMsg("广告代码不能为空，请输入");
		frm.Code.focus();
		return false;
	}
	
	frm.action.value = action;
	frm.submit();
}
                      </script>
					  
					  
					  </td>
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
</body>
</html>
<%
Call CloseConn()
Set Rs = Nothing
Set Admin = Nothing
%>

