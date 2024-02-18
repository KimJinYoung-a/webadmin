<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - Q&A"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/apps/academy/etc/qnaCls.asp" -->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<%
Dim ridx, page, MakerID, oDiyItemQnAList, searchdiv, searchtxt, isFilter, statediv, i

searchdiv = RequestCheckVar(request("searchdiv"),1)
searchtxt = RequestCheckVar(request("searchtxt"),32)
ridx	= requestCheckVar(request("ridx"),10)	''qna그룹 idx
page	= getNumeric(requestCheckVar(request("page"),3))
MakerID	= requestCheckVar(request.cookies("partner")("userid"),32)
isFilter = RequestCheckVar(request("isFilter"),1)
statediv = RequestCheckVar(request("statediv"),1)

if (statediv="") then statediv="0"
If searchdiv="" Then searchdiv=0
If page="" Then page=1

If (MakerID="") Then
	Response.Write "<script>alert('잘못된 접속 입니다.');fnAPPclosePopup();</script>"
	Response.End
End If

dim strParam
strParam = "statediv=" & statediv & "&searchdiv=" & searchdiv & "&searchtxt=" & searchtxt

'상품 QnA 리스트
set oDiyItemQnAList = new DiyItemCls
oDiyItemQnAList.FPageSize = 12
oDiyItemQnAList.FCurrPage = page
oDiyItemQnAList.FRectStateDIV = statediv
oDiyItemQnAList.FRectSearchDIV = searchdiv
oDiyItemQnAList.FRectSearchTXT = searchtxt
oDiyItemQnAList.FRectMakerid = MakerID
oDiyItemQnAList.FRectmode = "list"
oDiyItemQnAList.GetDiyQnaList()
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

function fnEtcMainRelold(){
	fnAPPParentsWinReLoad();
}

function fnSearchFilterSet(callbackval){
	callbackval = Base64.decode(callbackval);
	var jsontxt = JSON.parse(callbackval);
	$("#statediv").val(jsontxt.statediv);
	$("#isFilter").val(jsontxt.filter);
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
		<div class="content bgGry" id="contents">
			<h1 class="hidden">Q&amp;A</h1>
			<% If isFilter="Y" Or searchtxt<>"" Or oDiyItemQnAList.FresultCount > 0 Then %>
			<form name="searchForm" id="searchForm" method="get" style="margin:0px;" onSubmit="fnSearchList();return false;" action="">
			<input type="hidden" name="page" id="page" value="<%=page%>">
			<input type="hidden" name="statediv" id="statediv" value="<%=statediv%>">
			<input type="hidden" name="isFilter" id="isFilter" value="<%=isFilter%>">
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
				<div class="btnFilter <%=chkIIF(isFilter="Y","filterActive","")%>">
					<button type="button" onclick="fnAPPpopupSearchFilter('<%=g_AdminURL%>/apps/academy/etc/popFilter.asp?<%=strParam%>')">필터</button>
				</div>
			</div>
			</form>
			<% End If %>
			<% if oDiyItemQnAList.FresultCount > 0 then %>
			<div class="qnaListWrap">
				<ul class="pushList">
					<% for i=0 to oDiyItemQnAList.FresultCount-1 %>
					<% if oDiyItemQnAList.FItemList(i).FanswerYN = "Y" then %>
					<li class="flagFinish">
						<a href="javascript:fnAPPpopupQnaDetail('<%=g_AdminURL%>/apps/academy/etc/qnaView.asp?gridx=<%= oDiyItemQnAList.FItemList(i).Freply_group_idx %>&makerid=<%= makerid %>');">
						<dfn>답변완료</dfn>
					<% else %>
					<li class="flagIng">
						<a href="javascript:fnAPPpopupQnaDetail('<%=g_AdminURL%>/apps/academy/etc/qnaView.asp?gridx=<%= oDiyItemQnAList.FItemList(i).Freply_group_idx %>&makerid=<%= makerid %>');">
						<dfn>답변중</dfn>
					<% end if %>
						<div><%= oDiyItemQnAList.FItemList(i).Ftitle %></div>
						<span><%= printUserId(oDiyItemQnAList.FItemList(i).Fuserid,2,"*") %><em>l</em><%= FormatDate(oDiyItemQnAList.FItemList(i).FRegdate,"0000.00.00") %></span>
						</a>
					</li>
					<% next %>
				</ul>
				<% if oDiyItemQnAList.FTotalCount>oDiyItemQnAList.FPageSize then %>
				<div class="paging">
					<%=fnDisplayPaging_New(page,oDiyItemQnAList.FTotalCount,oDiyItemQnAList.FPageSize,"jsGoPage")%>
				</div>
				<% end if %>
			</div>
			<% Else %>
			<div class="artNo" style="display:">
				<div class="linkNotice">
					<p class="fs1-5r">등록된 질문글이 없습니다.</p>
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
<form method="post" name="rfrm">
<input type="hidden" name="mode">
<input type="hidden" name="idx">
<input type="hidden" name="ridx">
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
SEt oDiyItemQnAList = Nothing
%>
<!-- #include virtual="/apps/academy/lib/pms_badge_check.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->