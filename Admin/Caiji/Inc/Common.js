//---------------����ѡ����ɫ����-------------------------
function OpenThenSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;status:0;help:0;scroll:0;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}
//--------------------------------------------------
function unselectall(thisform){
        if(thisform.chkAll.checked){
            thisform.chkAll.checked = thisform.chkAll.checked&0;
        }   
    }
    function CheckAll(thisform){
        for (var i=0;i<thisform.elements.length;i++){
            var e = thisform.elements[i];
            if (e.Name != "chkAll"&&e.disabled!=true)
                e.checked = thisform.chkAll.checked;
        }
    }

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function AddReplace(){
	var strthis='�滻ǰ���ַ���|�滻����ַ���'; 
	var str=prompt('�������滻ǰ���ַ������滻����ַ������м��á�|��������',strthis);
	if(str!=null&&str!=''){document.myform.strReplace.options[document.myform.strReplace.length]=new Option(str,str);}
}
function ModifyReplace(){
	if(document.myform.strReplace.length==0) return false;
	var strthis=document.myform.strReplace.value; 
	if (strthis=='') {alert('����ѡ��һ���ַ������ٵ��޸İ�ť��');return false;}
	var str=prompt('�������滻ǰ���ַ������滻����ַ������м��á�|��������',strthis);
	if(str!=strthis&&str!=null&&str!=''){document.myform.strReplace.options[document.myform.strReplace.selectedIndex]=new Option(str,str);}
}
function DelReplace(){
	if(document.myform.strReplace.length==0) return false;
	var strthis=document.myform.strReplace.value; 
	if (strthis=='') {alert('����ѡ��һ���ַ������ٵ�ɾ����ť��');return false;}
	document.myform.strReplace.options[document.myform.strReplace.selectedIndex]=null;
}

function CheckForm(){
	if (document.myform.ItemName.value==''){
		alert('��Ŀ���Ʋ���Ϊ�գ�');
		return false;
	}
	if (document.myform.ListStr.value==''){
		alert('������Զ���б�URL��');
		return false;
	}

		for(var n=0;n<document.myform.strReplace.length;n++){
			if (document.myform.ReplaceList.value=='') document.myform.ReplaceList.value=document.myform.strReplace.options[n].value;
			else document.myform.ReplaceList.value+='$$$'+document.myform.strReplace.options[n].value;
		}
}
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->