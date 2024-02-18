<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 답글쓰기"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/apps/academy/notice/lecturerNoticecls.asp"-->
<%
Dim Doc_Id, Ans_Idx, MakerID, i, mode, olectview
Doc_Id = RequestCheckVar(request("Doc_Id"),10)
Ans_Idx = RequestCheckVar(request("Ans_Idx"),10)
mode = RequestCheckVar(request("mode"),32)
MakerID = requestCheckVar(request.cookies("partner")("userid"),32)

Set olectview = New ClecturerList
olectview.FrectAns_Idx = Ans_Idx
olectview.FRectMakerID = MakerID
olectview.fnGetolectView()

%>

<script>
<!--
function fnAppCallWinConfirm(){
	if($("#commContents").val()==""){
		alert("답변 내용을 입력해 주세요.");
		return false;
	}else{
		if(confirm("작성된 내용을 등록하시겠습니까?")){
			document.rfrm.target="FrameCKP";
			document.rfrm.submit();
		}
	}
}

function fnFreeBoardRelyEnd(){
	fnAPPParentsWinReLoad();
	setTimeout(function(){
		fnAPPclosePopup();
	}, 300);
}

//-->
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form method="post" name="rfrm" action="/apps/academy/notice/dofreeboardWrite.asp">
		<input type="hidden" name="didx" value="<%=Doc_Id%>">
		<input type="hidden" name="aidx" value="<%=Ans_Idx%>">
		<input type="hidden" name="mode" value="<%=mode%>">
		<div class="content bgGry">
			<h1 class="hidden">답글쓰기</h1>
			<div class="spcNote">
				<div class="linkInsert">
					<textarea rows="5" name="ans_content" placeholder="내용을 입력해주세요"><%=olectview.foneitem.FAns_Content%></textarea>
				</div>
			</div>
		</div>
		</form>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
	fnAPPShowRightConfirmBtns();
	$('.linkInsert').on( 'keyup', 'textarea', function (e){
		$(this).css('height', 'auto' );
		$(this).height( this.scrollHeight );
	});
	$('.linkInsert').find( 'textarea' ).keyup();
});
//-->
</script>
<% Set olectview = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->