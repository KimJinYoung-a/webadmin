<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이미지 등록
' History : 2015.01.29 정윤정 생성
' /admin/incSessionAdmin.asp => /common/incSessionBctId.asp '2016/07/11 eastone
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->   
<script src="../../js/popup.js" type="text/javascript" charset="utf-8"></script>
<link rel="stylesheet" href="../../css/popup.css" type="text/css"  charset="utf-8"/>
<script type="text/javascript">
// <![CDATA[
	
	function done() { 
	 
		if(!document.frmImg.sfile.value){
			alert("이미지를 선택해주세요");
			return;
		}
		document.frmImg.submit();
//		if (typeof(execAttach) == 'undefined') { //Virtual Function
//	        return;
//	    }
		
//		var _mockdata = {
//			'imageurl': 'http://cfile284.uf.daum.net/image/116E89154AA4F4E2838948',
//			'filename': 'editor_bi.gif',
//			'filesize': 640,
//			'imagealign': 'C',
//			'originalurl': 'http://cfile284.uf.daum.net/original/116E89154AA4F4E2838948',
//			'thumburl': 'http://cfile284.uf.daum.net/P150x100/116E89154AA4F4E2838948'
//		};
//		execAttach(_mockdata);
//		closeWindow();
	}

	function initUploader(){
	    var _opener = PopupUtil.getOpener();
	    if (!_opener) {
	        alert('잘못된 경로로 접근하셨습니다.');
	        return;
	    }
	    
	    var _attacher = getAttacher('image', _opener);
	    registerAction(_attacher);
	}
// ]]>
</script>
</head>
<body onload="initUploader();">
<div class="wrapper">
	<div class="header">
		<h1>사진 첨부</h1>
	</div>	
	<div class="body">
		<dl> 
		    <dd style="text-align:center;margin:10px;">
		    	<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/d_editor/uploadImage.asp" enctype="MULTIPART/FORM-DATA">
		    	<input type="hidden"  name="iML" value="1">
		    	 <input type="file" name="sfile">
		    	</form>
			</dd>
		</dl>
	</div>
	<div class="footer">
		<p><a href="#" onclick="closeWindow();" title="닫기" class="close">닫기</a></p>
		<ul>
			<li class="submit"><a href="#" onclick="done();" title="등록" class="btnlink">등록</a> </li>
			<li class="cancel"><a href="#" onclick="closeWindow();" title="취소" class="btnlink">취소</a></li>
		</ul>
	</div>
</div>
</body>
</html>