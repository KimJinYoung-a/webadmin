<%@ language=vbscript %>
<% option Explicit %>
<%
	Dim sImgUrl 
	sImgUrl = Request("simgUrl")
%>
<html>
<head>
<script language="javascript">
<!--
function jsResize(){
	var strImgWidth = document.all.imgMain.width;
	var strImgHeight = document.all.imgMain.height;
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
	
	window.resizeTo(strImgWidth+30,strImgHeight+80);
	}
//-->
</script>
</head>
<body  leftmargin="0" topmargin="10" onload="jsResize();">
<div align="center" valign="middle">
<img id="imgMain" src="<%=sImgUrl%>" border="0">
</div>
</body>
</html>