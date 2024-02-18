<%@ language=vbscript %>
<% option Explicit %>
<%
'####################################################
' Description :  전체이미지 보여주기
' History : 2011.04.06 한용민 생성
'####################################################
%>
<%
Dim sImgUrl 
	sImgUrl = Request("sUrl")
%>
<html>
<head>
<script language="javascript">

function jsResize(){
	var strImgWidth = document.all.imgMain.width+10;
	var strImgHeight = document.all.imgMain.height+60;
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

//	window.print();
	
</script>

</head>
<body leftmargin="0" topmargin="0" onload="jsResize();">
<img id="imgMain" src="<%=sImgUrl%>" border="0">
</body>
</html>