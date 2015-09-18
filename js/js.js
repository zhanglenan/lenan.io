window.onerror = function(){return true;}
String.prototype.trim = function(){ return this.replace(/(^\s*)|(\s*$)/g, "");}

function ImageZoom(Img,width,height)
{ 
	var image=new Image(); 
	image.src=Img.src;
	if(image.width>width||image.height>height)
	{
		w=image.width/width; 
		h=image.height/height; 
		if(w>h)
		{
			Img.width=width; 
			Img.height=image.height/w; 
		}
		else
		{
			Img.height=height; 
			Img.width=image.width/h; 
		} 
	}
}

function ImageOpen(Img)
{
	window.open(Img.src);
}

function ChkUserLogin(frm)
{
	if(frm.User_Name.value.trim()=='')
	{
		alert("请输入您的用户名。");
		frm.User_Name.focus();
		return false;
	}
	if(frm.User_Password.value.trim()=='')
	{
		alert("请输入您的密码。");
		frm.User_Password.focus();
		return false;
	}
	return true;
}

function ChkUserSearch(frm)
{
	if(frm.keyword.value.trim()=="")
	{
		alert("请输入你要查询的关键词。");
		frm.keyword.focus()
		return false;
	}
	return true;
}

function Ok3w_G_Submit(frm)
{
	if(frm.UserName.value.trim()=="" || frm.UserName.value=="请输入您的姓名")
	{
		alert("请输入姓名");
		frm.UserName.focus();
		return false;
	}
	if(frm.Content.value.trim()=="" || frm.Content.value=="请输入您的评论")
	{
		alert("请输入内容");
		frm.Content.focus();
		return false;
	}
	
	frm.bntSubmit.disabled=true;
	
	frm.submit();
}

function SetCwinHeight(obj)
{
  var cwin=obj;
  if (document.getElementById)
  {
    if (cwin && !window.opera)
    {
      if (cwin.contentDocument && cwin.contentDocument.body.offsetHeight)
        cwin.height = cwin.contentDocument.body.offsetHeight; 
      else if(cwin.Document && cwin.Document.body.scrollHeight)
        cwin.height = cwin.Document.body.scrollHeight;
    }
  }
}

function oCopy(a,type,str){
	if(str=="" || str=="http://")
	{
		alert("该网友没有填写相关内容");
		return false;
	}
	a.target="_blank";
	if(type==1)
		a.href = str;
	if(type==2)
		a.href = "mailto:" + str;
	if(type==3)
		a.href = "tencent://message/?uin="+ str + "&Site=im.qq.com&Menu=yes";
}

function Vote(id,action)
{
	mydown.location.href = "./c/soft_hits.asp?id=" + id + "&type=display&action=" + action;
}

function Ok3w_insertFlash(base_url, focus_width, focus_height, swf_height, text_height, pics, links, texts)
{
	document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
	document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="'+base_url+'images/focus.swf"><param name="quality" value="high"><param name="bgcolor" value="#f5fafe">');
	document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
	document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
	document.write('</object>');
}

function Ok3w_Marquee(RndID,StrGD,Width,Height,Speed)
{
	document.write('<div id="'+RndID+'" style="overflow:hidden;height:'+Height+'px;width:'+Width+'px;">');
	document.write('<table border="0" cellspacing="0" cellpadding="0">');
	document.write('  <tr>');
	document.write('    <td id="'+RndID+'1" height="'+Height+'" nowrap="nowrap">'+StrGD+'</td>');
	document.write('    <td id="'+RndID+'2" height="'+Height+'" nowrap="nowrap"></td>');
	document.write('  </tr>');
	document.write('</table>');
	document.write('</div>');
	
	if(Speed==0)
		return;
	var speed = Speed;
	var pro = document.getElementById(RndID);
	var pro1 = document.getElementById(RndID+"1");
	var pro2 = document.getElementById(RndID+"2");
	pro2.innerHTML=pro1.innerHTML;
	var MyMar=setInterval(Marquee,speed) 
	pro.onmouseover=function() {clearInterval(MyMar)} 
	pro.onmouseout=function() {MyMar=setInterval(Marquee,speed)} 
	function Marquee()
	{ 
		var mm_mo = pro.offsetWidth - pro1.offsetWidth;
		if(mm_mo<0) mm_mo=0;
		if(pro2.offsetWidth-pro.scrollLeft<=mm_mo) 
			pro.scrollLeft-=pro1.offsetWidth;
			else
				pro.scrollLeft+=5;
	} 
}

function settab(bobj,id,num)
{
	for(var i=1;i<=num;i++)
	{
		var obj = document.getElementById(bobj + i);
		obj.className = "hdout";
		var obj = document.getElementById(bobj + "_con" + i);
		obj.style.display = "none";
	}
	var obj = document.getElementById(bobj + id);
	obj.className = "hdover";
	var obj = document.getElementById(bobj + "_con" + id);
	obj.style.display = "";
}