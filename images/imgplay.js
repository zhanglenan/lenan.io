var showNum = document.getElementById("num");
function Mea(value){
	n=value;
	setBg(value);
	plays(value);
	cons(value);
	}
function setBg(value){
	for(var i=0;i<rowsNews+1;i++)
   if(value==i){
		showNum.getElementsByTagName("td")[i].className='bigon';
		} 
	else{	
		showNum.getElementsByTagName("td")[i].className='bigoff';
		}  
	} 
function plays(value){
	try
	{
		with (fc)
		{
			filters[0].Apply();
			for(i=0;i<rowsNews+1;i++)i==value?children[i].style.display="block":children[i].style.display="none"; 
			filters[0].play(); 
		}
	}
	catch(e)
	{
		var divlist = document.getElementById("fc").getElementsByTagName("div");
		for(i=0;i<rowsNews+1;i++)
		{
			i==value?divlist[i].style.display="block":divlist[i].style.display="none";
		}
	}

	
}
function cons(value){
	try
	{
		with (con)
		{
				for(i=0;i<rowsNews+1;i++)i==value?children[i].style.display="block":children[i].style.display="none"; 		
		}
	}
	catch(e)
	{
		var divlist = document.getElementById("con").getElementsByTagName("div");
		for(i=0;i<rowsNews+1;i++)
		{
			i==value?divlist[i].style.display="block":divlist[i].style.display="none";
		}		
	}
}

var FilterArr= new Array("RevealTrans(duration=2,transition=23)","BlendTrans(duration=2)","progid:DXImageTransform.Microsoft.Fade(duration=2,overlap=0)","progid:DXImageTransform.Microsoft.Wipe(duration=3,gradientsize=0.25,motion=reverse)");
function clearAuto(){clearInterval(autoStart)}
function setAuto(){autoStart=setInterval("auto(n)", 3000)}
function auto(){
	var jj = parseInt(Math.random() * FilterArr.length);
	document.getElementById("fc").style.filter = FilterArr[jj];
	n++;
	if(n>rowsNews)n=0;
	Mea(n);
} 
function sub(){
	n--;
	if(n<0)n=rowsNews;
	Mea(n);
} 
setAuto(); 