<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 구매후기"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/apps/academy/etc/reviewcls.asp" -->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<%
Dim searchdiv, searchtxt, MakerID
searchdiv = RequestCheckVar(request("searchdiv"),1)
searchtxt = RequestCheckVar(request("searchtxt"),32)
MakerID = requestCheckVar(request.cookies("partner")("userid"),32)
If searchdiv="" Then searchdiv=0

If (MakerID="") Then
	Response.Write "<script>alert('리뷰 정보가 없습니다.');fnAPPclosePopup();</script>"
	Response.End
End If

Dim oEval, q, oEvalPage
If oEvalPage = "" Then oEvalPage = 1
Set oEval = new cCorner
	oEval.FCurrPage = oEvalPage
	oEval.FPageSize = 8
	oEval.FRectLecturer_id = MakerID
	oEval.FRectSearchDIV=searchdiv
	oEval.FRectSearchTXT=searchtxt
	oEval.getDiyvaluationList
%>
<script>
$(function() {
	// search button control
	$(".searchInput input").keyup(function () {
		$(this).parent().find('button').fadeIn();
	});

	// search box hidden scroll top auto change
	var schH = $(".artSearchTop").outerHeight();
	$("body").css("min-height",schH+$(window).height());
	setTimeout(function(){
		$('html, body').animate({scrollTop:schH}, 'fast');
	}, 300);
});

function jsGoPage(iP){
	document.searchForm.page.value = iP;
	document.searchForm.submit();
}

function fnSearchList(){
	if(document.searchForm.searchdiv.value==0){
		alert("구분을 선택해주세요.");
		document.searchForm.searchdiv.focus();
	}else if(document.searchForm.searchtxt.value==""){
		alert("검색어를 입력해주세요.");
		document.searchForm.searchtxt.focus();
	}else{
		document.searchForm.submit();
	}
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">구매후기</h1>
			<% If searchtxt<>"" Or oEval.FresultCount > 0 Then %>
			<form name="searchForm" id="searchForm" method="get" style="margin:0px;" onSubmit="fnSearchList();return false;" action="">
			<div class="artSearchTop">
				<div class="searchInput hasOpt">
					<span class="schSlt">
						<select name="searchdiv">
							<option value="0"<% If searchdiv=0 Then Response.write " selected"%>>구분선택</option>
							<option value="1"<% If searchdiv=1 Then Response.write " selected"%>>작품코드</option>
							<option value="2"<% If searchdiv=2 Then Response.write " selected"%>>작품명</option>
							<option value="3"<% If searchdiv=3 Then Response.write " selected"%>>작성자</option>
							<option value="4"<% If searchdiv=4 Then Response.write " selected"%>>글 내용</option>
						</select>
					</span>
					<input type="Search" name="searchtxt" placeholder="작품코드, 작품명, 작성자, 글 내용 검색" value="<%=searchtxt%>" onKeyPress="if (event.keyCode == 13){ fnSearchList(); return false;}" maxlength="32" />
					<button type="button" class="btnSearch" onClick="fnSearchList();return false;">검색</button>
				</div>
			</div>
			</form>
			<% End If %>
			<% If oEval.FResultCount > 0 Then %>
			<div class="boardList reviewListWrap">
				<ul class="artList">
					<% For q = 0 to oEval.FResultCount - 1 %>
					<li>
						<p class="title">[<%= UCase(oEval.FItemList(q).FBrandName) %>] <%= oEval.FItemList(q).FItemname %></p>
						<div class="reviewCont">
							<div class="star score0<%= oEval.FItemList(q).FTotalPoint %>"><span></span></div>
							<p class="txt"><%= nl2br(oEval.FItemList(q).FContents) %></p>
							<div class="reviewImg">
								<% If oEval.FItemList(q).Flinkimg1 <> "" Then %>
									<img src="<%= oEval.FItemList(q).getLinkImage1 %>"/>
								<% End if %>
								<% If oEval.FItemList(q).Flinkimg2 <>"" Then %>
									<img src="<%= oEval.FItemList(q).getLinkImage2 %>"/>
								<% End If %>
							</div>
							<p class="txtInfo"><%= printUserId(oEval.FItemList(q).FUserID,2,"*") %><span>l</span><%= FormatDate(oEval.FItemList(q).FRegdate, "0000.00.00") %></p>
						</div>
					</li>
					<% Next %>
				</ul>
				<% if oEval.FTotalCount>oEval.FPageSize then %>
				<div class="paging">
					<%=fnDisplayPaging_New(page,oEval.FTotalCount,oEval.FPageSize,"jsGoPage")%>
				</div>
				<% end if %>
			</div>
			<% Else %>
			<div class="artNo" style="display:">
				<div class="linkNotice">
					<p class="fs1-5r">수신된 구매후기 내역이 없습니다.</p>
				</div>
			</div>
			<% End If %>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<%
SEt oEval = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->