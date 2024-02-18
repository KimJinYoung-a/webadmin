<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% 
	response.Charset="UTF-8"
	Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 응원톡 쓰기"
'####################################################
' Description : 응원톡 수정/답글
' History : 2017-01-11 이종화 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/apps/academy/etc/talk_cls.asp" -->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<%
Dim lecturer_id, mode, gubun, paramid, ridx, idx, mycomment, gubunholderStr
Dim oMycomm
mode	= requestCheckVar(Request("mode"),10)
gubun	= requestCheckVar(Request("gubun"),1)
paramid	= requestCheckVar(Request("paramid"),32)
ridx	= getNumeric(requestCheckVar(Request("ridx"),10))
idx		= getNumeric(requestCheckVar(Request("idx"),10))

lecturer_id = requestCheckVar(request.cookies("partner")("userid"),32)

If mode = "" Then
	Response.Write "<script>alert('잘못된 경로입니다.');</script>"
	dbACADEMYget.close()
	Response.End
End If

If mode = "reply" and ridx = "" Then
	Response.Write "<script>alert('잘못된 경로입니다.');</script>"
	dbACADEMYget.close()
	Response.End
End If

If gubun = "D" Then
	gubunholderStr = "내용을 입력해주세요"
Else
	gubunholderStr = "내용을 입력해주세요"
End If

If mode = "edit" Then
	Set oMycomm = new cCorner
		oMycomm.FRectIdx		= idx
		oMycomm.FRectUserid		= lecturer_id
		oMycomm.FRectParamid	= paramid
		oMycomm.getMycommRead
		mycomment = oMycomm.FOneItem.FComment
	Set oMycomm = nothing
End if

%>
<script>
function fnAppCallWinConfirm(){
	var frm = document.commform;
	if(frm.commContents.value.length < 1){
		alert("내용을 입력해주세요.");
		frm.commContents.focus();
		return;
	}
	if(confirm("저장 하시겠습니까?")){
		frm.target="FrameCKP";
		frm.submit();
	}
}

function fnQnARelyEnd(msg,mode,qnacount){
	alert(msg);
	fnAPPopenerJsCallClose("fnTalkListRelold(\"\")");
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<div class="content bgGry">
			<h1 class="hidden">답글쓰기</h1>
			<div class="spcNote">
				<form method="post" action="/apps/academy/etc/Talk_Proc.asp" name="commform">
				<input type="hidden" name="mode" value="<%= mode %>">
				<input type="hidden" name="gubun" value="<%= gubun %>">
				<input type="hidden" name="paramid" value="<%= paramid %>">
				<input type="hidden" name="ridx" value="<%= ridx %>">
				<input type="hidden" name="idx" value="<%= idx %>">
				<div class="linkInsert">
					<textarea rows="5" name="commContents" placeholder="<%= gubunholderStr %>"><%= mycomment %></textarea>
				</div>
				</form>
			</div>
		</div>
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
</body>
</html>
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
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->