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
<!-- #include virtual="/apps/academy/etc/LecDiyqnaCls.asp"-->
<%
Dim ridx, idx, MakerID, i, mode, Comment
ridx = RequestCheckVar(request("ridx"),10)
idx = RequestCheckVar(request("rowidx"),32)
mode = RequestCheckVar(request("mode"),32)
MakerID = requestCheckVar(request.cookies("partner")("userid"),32)

Dim oMyqna, masterQitemid, masterQlec_idx, masterQmakerid, masterQRegID, masterGubun, masterQSmsOK, masterQTitle, masterQRegName, masterQEmail, oMyReply

SET oMyqna = new CQna
	oMyqna.FRectIdx = idx
	oMyqna.FRectGroupIdx = ridx
	'oMyqna.FRectsearchDiv = qnagubun
	oMyqna.getOnemyqna

	masterQitemid		= oMyqna.FOneItem.Fitemid
	masterQlec_idx		= oMyqna.FOneItem.Flec_idx
	masterQmakerid		= oMyqna.FOneItem.Fmakerid
	masterQRegID		= oMyqna.FOneItem.FUserid
	masterGubun		= oMyqna.FOneItem.Fpagegubun
	masterQSmsOK		= oMyqna.FOneItem.FSmsok
	masterQTitle		= oMyqna.FOneItem.FTitle

	Call getMyinfo(masterQRegID, masterQRegName, masterQEmail)

If mode = "edit" Then
SET oMyReply = new CQna
oMyReply.FRectIdx = idx
oMyReply.getOneMyReply
Comment = oMyReply.FOneItem.FComment
Else
Comment=""
End If
%>

<script>
<!--
function fnAppCallWinConfirm(){
	if($("#commContents").val()==""){
		alert("답변 내용을 입력해 주세요.");
		return false;
	}else{
		<% If mode="edit" Then %>
		document.rfrm.mode.value="edit";
		<% Else %>
		document.rfrm.mode.value="addreply";
		<% End If %>
		document.rfrm.action="/apps/academy/etc/doqnareply.asp";
		document.rfrm.target="FrameCKP";
		document.rfrm.submit();
	}
}

function fnQnARelyEnd(msg,mode,qnacount){
	setTimeout(function(){
			fnAPPParentsWinReLoad();
		}, 300);
	alert(msg);
	setTimeout(function(){
		fnAPPopenerJsCallClose("fnQnaViewRelold(\"<%=ridx%>\")");
	}, 600);
	if(mode=="addreply"){
		setTimeout(function(){
			fnAPPChangeBadgeCount("qnacount",qnacount);
		}, 900);
	}
}

//-->
</script>
</head>
<body>
<%
Dim lastqna, qstContents, lastRegdate, lastSMSok, lastSmsNum
Dim QnaColor
SET oMyqna = new CQna
	oMyqna.FCurrPage = 1
	oMyqna.FPageSize = 500
	oMyqna.FRectGroupIdx = ridx
	oMyqna.getqnaDetailList
For i = 0 to oMyqna.FResultCount - 1
	lastqna			= oMyqna.FItemList(i).FQna 
	If lastqna = "Q" Then
		qstContents		= oMyqna.FItemList(i).Fcomment
		lastRegdate		= oMyqna.FItemList(i).FRegdate
		lastSMSok		= oMyqna.FItemList(i).FSmsok
		lastSmsNum		= oMyqna.FItemList(i).FSmsnum
	End If
Next
%>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">답글쓰기</h1>
			<div class="spcNote">
				<div class="linkInsert">
				<form name="rfrm" id="rfrm" method="post" style="margin:0px;">
				<input type="hidden" name="mode" value="<%=mode%>">
				<input type="hidden" name="ridx" value="<%=ridx %>">
				<input type="hidden" name="idx" value="<%= idx %>">
				<input type="hidden" name="gubunVal" value="<%= masterGubun %>" >
				<input type="hidden" name="lastSmsNum" value="<%= lastSmsNum %>" >
				<input type="hidden" name="makerid" value="<%= masterQmakerid %>" >
				<input type="hidden" name="pagegubun" value="<%= masterGubun %>" >
				<input type="hidden" name="diyitemid" value="<%= masterQitemid %>" >
				<input type="hidden" name="lec_idx" value="<%= masterQlec_idx %>" >
				<!-- 메일에 필요한 내용 hidden 처리 -->
				<input type="hidden" name="usermail" value="<%= masterQEmail %>" >
				<input type="hidden" name="qstContents" value="<%= qstContents %>" >
				<input type="hidden" name="lastRegdate" value="<%= lastRegdate %>" >
				<input type="hidden" name="masterQTitle" value="<%= masterQTitle %>" >
				<!-- ################################-->
				<!-- SMS전송에 필요한 내용 hidden 처리 -->
				<input type="hidden" name="lastSMSok" value="<%= lastSMSok %>" >
				<input type="hidden" name="lastSmsNum" value="<%= lastSmsNum %>" >
				<!-- ################################-->
					<textarea rows="5" name="ansContents" placeholder="내용을 입력해주세요"><%=Comment%></textarea>
				</form>
				</div>
			</div>
		</div>
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
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->