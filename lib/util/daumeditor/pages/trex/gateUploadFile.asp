<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 에디터 이미지 등록
' History : 2015.01.29 정윤정  생성
' /admin/incSessionAdmin.asp => /common/incSessionBctId.asp '2016/07/11 eastone
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
	Dim sFile,sFilePath
	sFile = ReplaceRequestSpecialChar(request("sFN")) 
	sFilePath= webImgUrl&"/"&ReplaceRequestSpecialChar(request("sFP"))  
%>
<script src="../../js/popup.js" type="text/javascript" charset="utf-8"></script>  
<link rel="stylesheet" href="../../css/popup.css" type="text/css"  charset="utf-8"/>
<script type="text/javascript"  charset="utf-8">  
		var _opener = PopupUtil.getOpener();
		if(!_opener) {
			alert('잘못된 경로로 접근하셨습니다.'); 
			closeWindow();
		}
 
	    var _attacher = getAttacher('image', _opener);
	    registerAction(_attacher);
	    
	    if (typeof(execAttach) == 'undefined') { //Virtual Function
	        alert("데이터 처리에 문제가 발생했습니다.");
	    } 
		
		var _mockdata = {
			'imageurl': '<%=sFilePath%>/<%=sFile%>',
			'filename': '<%=sFile%>',
			'filesize': 640,
			'imagealign': 'C',
			'originalurl': '<%=sFilePath%>/<%=sFile%>',
			'thumburl': '<%=sFilePath%>/<%=sFile%>'
		};
		execAttach(_mockdata);
		closeWindow();
</script>
 
 