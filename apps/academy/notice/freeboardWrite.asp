<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 자유게시판 글쓰기"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/notice/lecturerNoticecls.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<%
Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status, sDoc_Type, sDoc_Import, sDoc_part_sn
dim sDoc_Diffi, sDoc_Subj, sDoc_Content ,sDoc_UseYN, sDoc_Regdate
dim sDoc_WorkerView , i , tContents , olect, lectFile , arrFileList
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	
	If iDoc_Idx = "" Then
		sDoc_Id 		= requestCheckVar(request.cookies("partner")("userid"),32)
		sDoc_Name		= session("ssBctCname")
		sDoc_Regdate	= Left(now(),10)
		sDoc_Status = "K001"
	Else
		
		Set olect = New ClecturerList
		olect.FrectDoc_Idx = iDoc_Idx
		olect.fnGetlecturerView
	
		sDoc_Id 		= olect.FOneItem.FDoc_Id
		sDoc_Name		= olect.FOneItem.FDoc_Name
		sDoc_Status		= olect.FOneItem.FDoc_Status
		if sDoc_Status = "" then sDoc_Status = "K001"			
		sDoc_Type		= olect.FOneItem.FDoc_Type
		sDoc_Import		= olect.FOneItem.FDoc_Import
		sDoc_Diffi		= olect.FOneItem.FDoc_Diffi
		sDoc_Subj		= olect.FOneItem.FDoc_Subj
		tContents	= olect.FOneItem.FDoc_Content	
		sDoc_UseYN		= olect.FOneItem.FDoc_UseYN
		sDoc_Regdate	= olect.FOneItem.FDoc_Regdate
		sDoc_part_sn	= olect.FOneItem.fpart_sn

		set lectFile = new ClecturerList
	 	lectFile.FrectDoc_Idx = iDoc_Idx
		arrFileList = lectFile.fnGetFileList	
	End If

if sDoc_Type = "" then sDoc_Type = "G020"
if sDoc_Import = "" then sDoc_Import = "L002"
%>
<script>
<!--
$(function() {
	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});
});

function fnAppCallWinConfirm(){
	if(document.frm.G000.value == ""){
		alert("구분을 선택해주세요.");
		return;
	}else if(document.frm.doc_subject.value == ""){
		alert("제목을 입력해주세요.");
		return;
	}else if(document.frm.brd_content.value == ""){
		alert("내용을 입력해주세요.");
		return;
	}else{
		if(confirm("작성된 내용을 등록하시겠습니까?")){
		document.frm.target="FrameCKP";
		document.frm.submit();
		}
	}
}

function fnFreeBoardAskEnd(){
	fnAPPParentsWinReLoad();
	setTimeout(function(){
		fnAPPclosePopup();
		}, 300);
}

function fnFreeBoardAskEditEnd(){
	fnAPPParentsWinReLoad();
	setTimeout(function(){
		fnAPPopenerJsCallClose("fnAskEditEnd(\"\")");
		}, 300);
}

function fnDocImportant(objval){
	document.frm.L000.value=objval;
}
//-->
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form name="frm" action="lecturer_proc.asp" method="post" style="margin:0px;" onsubmit="fnAppCallWinConfirm(); return false;">
		<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
		<input type="hidden" name="gubun" value="write">
		<input type="hidden" name="mode" value="edit">
		<input type="hidden" name="K000" value="<%=sDoc_Status%>">
		<input type="hidden" name="doc_difficult" value="2">
		<input type="hidden" name="L000" value="<%=sDoc_Import%>">
		<div class="content bgGry">
			<h1 class="hidden">글쓰기</h1>
			<div class="askWrite">
				<ul class="artList">
					<li class="selectBtn">
						<%=CommonCode("w","G000",sDoc_Type)%>
					</li>
					<li class="list">
						<dfn><b>중요도</b></dfn>
						<div class="btnGroup selectBtn">
							<button type="button" class="btn btnGry <% If sDoc_Import="L001" Then Response.write "selected"%>" onClick="fnDocImportant('L001');">상</button>
							<button type="button" class="btn btnGry <% If sDoc_Import="L002" Then Response.write "selected"%>" onClick="fnDocImportant('L002');">중</button>
							<button type="button" class="btn btnGry <% If sDoc_Import="L003" Then Response.write "selected"%>" onClick="fnDocImportant('L003');">하</button>
						</div>
					</li>
					<li>
						<input type="text" name="doc_subject" value="<%=sDoc_Subj%>" placeholder="제목을 입력해주세요" />
					</li>
					<li class="linkInsert">
						<textarea rows="5" name="brd_content" placeholder="내용을 입력해주세요"><%=stripHTML(tContents)%></textarea>
					</li>
				</ul>
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
	$('.artList').on( 'keyup', 'textarea', function (e){
		$(this).css('height', 'auto' );
		$(this).height( this.scrollHeight );
	});
	$('.artList').find( 'textarea' ).keyup();
});
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->