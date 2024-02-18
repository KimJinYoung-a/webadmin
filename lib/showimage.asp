<%
dim img
img=request("img")
%>
<HTML>
<HEAD>
<TITLE>이미지를 클릭하시면 창이 닫힙니다.</TITLE>
<script language="javascript">
<!--
function jsResize(){
	var strImgWidth = document.all.imgMain.width+10;
	var strImgHeight = document.all.imgMain.height+100;
	var CheckScrollbar = 0;
	
	if (strImgHeight > screen.availHeight) 
	{
		strImgHeight = (screen.availHeight - 39);
		strImgWidth = strImgWidth + 30;
		CheckScrollbar = 1;
	}
	
	if (strImgWidth > screen.availWidth) 
	{
		strImgWidth = screen.availWidth;
		strImgHeight = strImgHeight + 30;
		CheckScrollbar = 1;	
	}
	
	if(CheckScrollbar == 0)
	{		
		document.body.style.overflow='hidden';
	}
	
	window.resizeTo(strImgWidth,strImgHeight);	
	}
//-->
</script>
</HEAD>
<BODY style="margin:0" onload="jsResize();">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="center" valign="middle">
		<img src="<%= img %>" border="0" onClick="window.close()" style="cursor:hand" id="imgMain">
	</td>
</tr>
</table>
</BODY>
</HTML>
