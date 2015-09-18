<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Const dbdns="../"%>
<!--#include file="chk.asp"-->
<!--#include file="../AppCode/Class/Ok3w_Article.asp"-->
<!--#include file="../AppCode/fun/function.asp"-->
<!--#include file="../AppCode/fun/CreateHtml.asp"-->
<%
Set Article = New Ok3w_Article
action = Request.QueryString("action")
Id = Request.QueryString("Id")
ChannelID = Request.QueryString("ChannelID")
ClassID = Request.QueryString("ClassID")
Select Case ChannelID
	Case 1
		Call CheckAdminFlag(3)
	Case 2
		Call CheckAdminFlag(2)
	Case Else
		Response.End()
End Select

action_ok = Request.Form("action_ok")
If action = "" Then action = "add"

Select Case action_ok
	Case "add"
		Call Article.Add()
		Call SaveAdminLog("添加文章：" & Article.Title)
		Call ActionOk("Article_Edit.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&action=" & action & "&ID=" & ID)
	Case "edit"
		Call Article.Edit()
		Call SaveAdminLog("修改文章：Id=" & Article.Id)
		Call ActionOk("Article_Edit.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&action=" & action & "&ID=" & ID)
End Select
If ClassID<>"" Then Article.ClassID = Cdbl(ClassID)
If action="edit" Or action="copy" Then Call Article.Load(Id)
If action="" Or action="copy" Then
	action = "add"
	Id = ""
End If
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
                  <td 
                background="images/tab_active_bg.gif" class="tab"><strong class="mtitle">文章编辑</strong></td>
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
                valign="top" bgcolor="#fffcf7"><table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EBEBEB">
                      <form id="Form" name="Form" method="post" action="?action=<%=action%>&Id=<%=Id%>&ChannelID=<%=ChannelID%>">
                        
                        <tr bgcolor="#FFFFFF">
                          <td width="70" align="right">文章标题<span class="red">*</span></td>
                          <td><input name="Title" type="text" id="Title" value="<%=Article.Title%>" size="50" maxlength="255" />
                          标题颜色：
                            <input name="TitleColor" type="text" value="<%=Article.TitleColor%>" size="8" maxlength="7">
                            <select name="select" onChange="this.form.TitleColor.value=this.value;">
                              <option value="">选择</option>
                              <option value="#FF0000">红色</option>
                              <option value="#0000FF">蓝色</option>
                              <option value="#008800">绿色</option>
                            </select>                            </td>
                          </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="right">跳转连接</td>
                          <td><input name="TitleURL" type="text" id="TitleURL" value="<%=Article.TitleURL%>" size="50" maxlength="255" />
                            点击标题直接打开的链接</td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="right">文章分类<span class="red">*</span></td>
                          <td><select name="ClassID">
                            <%Call InitClassSelectOption(ChannelId,0,Article.ClassID)%>
                          </select>                         </td>
                        </tr>
                        
                        <tr bgcolor="#FFFFFF">
                          <td align="right">文章内容<span class="red">*</span></td>
                          <td>
						  <input name="Content" type="hidden" id="Content" value="<%=Server.HTMLEncode(Replace(Replace(Article.Content,"""upfiles/","""../upfiles/"),"""editor/","""../editor/"))%>" />
						  <IFRAME ID="eWebEditor1" SRC="../editor/ewebeditor.htm?id=Content&style=Ok3w&savepathfilename=UpFiles" FRAMEBORDER="0" SCROLLING="no" WIDTH="600" HEIGHT="350" style="border:1px solid #CCCCCC;"></IFRAME>
						  <input name="UpFiles" type="hidden" id="UpFiles">
						  <input name="eWebEditorUpFile" type="hidden" id="eWebEditorUpFile" value="1"></td>
                          </tr>
                        
                        <tr bgcolor="#FFFFFF">
                          <td align="right">关键词</td>
                          <td><input name="Keywords" type="text" id="Keywords" value="<%=Article.Keywords%>" size="50" maxlength="255">
                            <a href="###" onClick="Get_Keywords()">自动获取</a></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="right">内容提要</td>
                          <td><textarea name="Description" cols="49" rows="5" id="Description"><%=Article.Description%></textarea>
                            <a href="###" onClick="Get_Description()">自动获取</a></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="right">访问权限</td>
                          <td><select name="vUserGroupID" id="vUserGroupID">
                            <option value="0">==游客==</option>
							<%Call InitUserDengjiSelectOption(Article.vUserGroupID)%>
                          </select>
                            <input name="vUserMore" type="checkbox" id="vUserMore" value="1" <%If Article.vUserMore=1 Then%>checked<%End If%>>
                            及以上级别，或是指定最低积分：
                            <input name="vUserJifen" type="text" id="vUserJifen" size="4" value="<%=Article.vUserJifen%>">                             </td>
                        </tr>						
                        <tr bgcolor="#FFFFFF">
                          <td align="right">文章属性</td>
                          <td><input name="IsPass" type="checkbox" id="IsPass" value="1" <%If Article.IsPass Then%>checked="checked"<%End If%> />
通过审核

                            <input name="IsPic" type="checkbox" id="IsPic" onClick="javascript:ChkIsPic();" value="1" <%If Article.IsPic Then%>checked="checked"<%End If%> />图片新闻
                            <input name="IsPlay" type="checkbox" id="IsPlay" onClick="javascript:ChkIsPic();" value="1" <%If Article.IsPlay Then%>checked="checked"<%End If%> />图片轮播
<input name="IsIndexImg" type="checkbox" id="IsIndexImg" onClick="javascript:ChkIsPic();" value="1" <%If Article.IsIndexImg Then%>checked="checked"<%End If%>>首页分类略图
<input name="IsMove" type="checkbox" id="IsMove" value="1" <%If Article.IsMove Then%>checked="checked"<%End If%> />滚动新闻
<input name="IsCommend" type="checkbox" id="IsCommend" value="1" <%If Article.IsCommend Then%>checked="checked"<%End If%> />推荐闻闻
<input name="IsTop" type="checkbox" id="IsTop" value="1" <%If Article.IsTop Then%>checked="checked"<%End If%> />置顶新闻                            
                            <input name="IsDelete" type="hidden" id="IsDelete" value="<%=Article.IsDelete%>"></td>
                          </tr>
                        <tr style="display:none;" id="DisIsPic" bgcolor="#FFFFFF">
                          <td align="right">&nbsp;图片地址<span class="red">*</span></td>
                          <td><input name="PicFile" type="text" id="PicFile" size="50" value="<%=Article.PicFile%>" />
                            <a href="###" onClick="Get_eWebEditor_Img();">从编辑器中选择</a>
                            <style type="text/css">
							#eImg{border:1px solid #CCCCCC; padding:5px; margin-top:5px; display:none;}
							#eImg img{border:1px solid #666666; cursor:pointer; width:100px; height:80px;}
							</style>
							<div id="eImg"></div>
                            <iframe scrolling="no" frameborder="0" width="100%" height="100" src="../editor/upload_files.asp?formID=Form&objID=PicFile"></iframe></td>
                        </tr>
						<tr bgcolor="#FFFFFF">
                          <td align="right">文章来源</td>
                          <td><input name="ComeFrom" type="text" id="ComeFrom" value="<%=Article.ComeFrom%>" size="20" />
                            作者
                            <input name="Author" type="text" id="Author" value="<%=Article.Author%>" size="20" />
                            时间
                            <input name="AddTime" type="text" id="AddTime" value="<%=Article.AddTime%>" size="18" />
                            查看
                            <input name="Hits" type="text" id="Hits" value="<%=Article.Hits%>" size="4" />
                            投票
                            <input name="pMoodStr" type="text" id="pMoodStr" value="<%=Article.pMoodStr%>" size="15"></td>
                          </tr>                    
                        <tr bgcolor="#FFFFFF">
                          <td align="right">&nbsp;</td>
                          <td><input name="action_ok" type="hidden" id="action_ok" value="<%=action%>" />
                            <input name="IsUserAdd" type="hidden" id="IsUserAdd" value="<%=Article.IsUserAdd%>">
                            <input name="GiveJifen" type="hidden" id="GiveJifen" value="<%=Article.GiveJifen%>">
                            <input name="Inputer" type="hidden" id="Inputer" value="<%=Article.Inputer%>">
                            <input name="bntSubmit" type="button" class="bnt14" id="bntSubmit" onClick="submitform(forms[0]);" value=" 立即保存 "/>
                                <input name="Submit" type="button" class="bnt14" onClick="javascript:document.URL='Article_List.asp?ChannelID=<%=ChannelID%>';" value=" 取消 " /></td>
                          </tr>
                      </form>
                    </table>
<script language="JavaScript" type="text/javascript">
function submitform(frm)
{
	if(frm.Title.value.trim()=="")
	{
		ShowErrMsg("标题不能为空，请输入");
		frm.Title.focus();
		return false;
	}
	if(frm.ClassID.value=="")
	{
		ShowErrMsg("栏目不能为空，请选择");
		frm.ClassID.focus();
		return false;
	}
	if(eWebEditor1.eWebEditor.document.body.innerHTML.trim()=="")
	{
		ShowErrMsg("内容不能为空，请输入");
		eWebEditor1.eWebEditor.focus();
		return false;
	}
	if(frm.IsPic.checked && frm.PicFile.value.trim()=="")
	{
		ShowErrMsg("文章属性选择了“图片”，但还没有上传图片,请上传");
		frm.PicFile.focus();
		return false;
	}
	
	frm.action = frm.action + "&ClassID=" + frm.ClassID.value;
	
	frm.bntSubmit.disabled = true;
	frm.bntSubmit.value = "请稍候...";
	
	frm.submit();
}

function ChkIsPic()
{
	obj = document.getElementById("DisIsPic");
	var _IsPic = document.getElementById("IsPic").checked;
	var _IsPlay = document.getElementById("IsPlay").checked;
	var IsIndexImg  = document.getElementById("IsIndexImg").checked;
	if(_IsPic || _IsPlay || IsIndexImg)
	{
		obj.style.display = "";
	}
	else
	{
		obj.style.display = "none";
	}
}

function Get_eWebEditor_Img()
{	
	var imgs = eWebEditor1.eWebEditor.document.getElementsByTagName('img');
	var imgstr = "";
	for(var img=0;img<imgs.length;img++)
	{
		imgstr = imgstr + '<img onclick="Set_Img(this.src)" src="'+imgs[img].src+'" /> ';
	}
	if(imgstr!="")
	{
		eImg.innerHTML = imgstr;
		eImg.style.display = "block";
	}
	else
	{
		alert("编辑器没有图片");
	}
}
function Set_Img(src)
{
	var sPath = document.location.host + document.location.pathname;
	sPath = sPath.substr(0, sPath.length-16);
	var tmp = sPath.split("/");
	var url = "http://";
	for(var i=0;i<tmp.length-2;i++)
		url = url + tmp[i] + "/";
	Form.PicFile.value = src.replace(url,"");
	eImg.style.display = "none";
}

function Get_Keywords()
{
	var tmp = Form.Title.value;
	if(tmp=="")
	{
		alert("请先输入标题");
		return false;
	}
	var kk="";
	var tt="";
	var i=0;
	var j=0;
	for(j=0;j<tmp.length-1;j++)
	{
		tt = tmp.substring(i,i+2)
		if(kk.indexOf(tt)==-1)
			kk = kk + "|" + tt;
		i = i+1;
	}
	Form.Keywords.value = kk.substring(1,kk.length);
}

function Get_Description()
{
	var tmp = eWebEditor1.eWebEditor.document.body.innerText;
	if(tmp=="")
	{
		alert("请先输入内容");
		return false;
	}
	Form.Description.value = tmp.substring(0,150);
}

ChkIsPic();
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
Set Article = Nothing
Set Admin = Nothing
%>

