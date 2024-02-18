<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 응원톡"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/apps/academy/etc/talk_cls.asp" -->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<%
Dim lecturer_id
Dim i
Dim myJob , ItsMeStr
Dim searchdiv , searchtxt , isFilter , statediv
Dim oCheerTalk, page
Dim j

lecturer_id = requestCheckVar(request.cookies("partner")("userid"),32)

searchdiv		= RequestCheckVar(request("searchdiv"),1)
searchtxt		= RequestCheckVar(request("searchtxt"),32)
page	= getNumeric(requestCheckVar(request("page"),3))
isFilter		= RequestCheckVar(request("isFilter"),1)
statediv		= RequestCheckVar(request("statediv"),1)

if (statediv="") then statediv="0"
If searchdiv="" Then searchdiv=0

If (lecturer_id="") Then
	Response.Write "<script>alert('잘못된 접속 입니다.');fnAPPclosePopup();</script>"
	Response.End
End If

'작가던지 강사던지 param으로 lecturer_id가 넘어옴
'따라서 그 lecturer_id가 작가인지 강사인지 판단할 수 있는 작업 필요함..
'myJob = L : 강사, myJob = D : 작가, 그 외 X는 체크가 없던지 테이블에 아이디가 없던지..
myJob = fnWhatIsMyJob(lecturer_id)

If page = "" Then page = 1
SET oCheerTalk = new cCorner
	oCheerTalk.FCurrPage	= page
	oCheerTalk.FPageSize	= 12
	oCheerTalk.FRectGubun	= myJob		'L : 강사, D : 작가
	oCheerTalk.FRectParamid	= lecturer_id
	oCheerTalk.FRectSearchDIV = searchdiv
	oCheerTalk.FRectSearchTXT = searchtxt
	oCheerTalk.getCheerTalkCommentList()

%>
<script>
$(function() {
	// search button control
	$(".searchInput input").keyup(function () {
		$(this).parent().find('button').fadeIn();
	});

	// search box hidden scroll top auto change
	var schH = $(".artSearchTop").outerHeight();
	setTimeout(function(){
		$('html, body').animate({scrollTop:schH}, 'fast');
	}, 300);
});

//응원톡 답글//수정
function cheerTalkWrite(imode, iridx, iidx){
	var url = "<%=g_AdminURL%>/apps/academy/etc/Talk_Write.asp?gubun=<%=myJob%>&paramid=<%=lecturer_id%>&mode="+imode+"&ridx="+iridx+"&idx="+iidx;
	fnAPPpopupCheerUpWrite(url);
}

//응원톡 페이징
function fngopage(page, bfgubun, ilecturer_id){
	var vPg = "1";
	if (bfgubun=="b"){
		if(page > 1){
			vPg = page - 1;
		}else{
			alert('이전 페이지가 없습니다');
			return;
		}
	}else if(bfgubun=="f"){
		if(page < <%= CInt(oCheerTalk.FTotalPage)+1 %>){
			vPg = page++;
		}else{
			alert('다음 페이지가 없습니다');
			return;
		}
	}else{
		parent.location.reload();
	}
	$.ajax({
	   url : "/apps/academy/etc/ajax_talk.asp?gubun=<%=myJob%>&page="+vPg+"&lecturer_id="+ilecturer_id,
		dataType : "html",
		type : "get",
		success : function(result){
		    $("#cheer").empty().html(result);
			var schH = $(".artSearchTop").outerHeight();
	setTimeout(function(){
		$('html, body').animate({scrollTop:schH}, 'fast');
	}, 300);
		}
	});
}

//코멘트 삭제
function cheerTalkDel(ridx, iidx, dp){
	if(confirm("삭제하시겠습니까?")){
		document.delfrm.idx.value   = iidx;
		document.delfrm.ridx.value  = ridx;
		document.delfrm.depth.value = dp;
		document.delfrm.action		= "/apps/academy/etc/Talk_Proc.asp"
		document.delfrm.target		= "FrameCKP";
   		document.delfrm.submit();
	}
}
function onlyNumber2(event, ilecturer_id){
	var val = $("#textpage2").val();
	if(val > <%= CInt(oCheerTalk.FTotalPage) %>){
		val=<%= CInt(oCheerTalk.FTotalPage) %>;
	}else if(val < 1){
		val = 1;
	}

	event = event || window.event;
	var keyID = (event.which) ? event.which : event.keyCode;
	if( (keyID >= 48 && keyID <= 57) || (keyID >= 96 && keyID <= 105) || keyID == 8 || keyID == 46 || keyID == 37 || keyID == 39 ){
		return;
	}else if(keyID == 13){
		$.ajax({
		    url : "/apps/academy/etc/ajax_talk.asp?gubun=<%=myJob%>&page="+vPg+"&lecturer_id="+ilecturer_id,
		    dataType : "html",
		    type : "get",
		    success : function(result){
		        $("#cheer").empty().html(result);
				var schH = $(".artSearchTop").outerHeight();
	setTimeout(function(){
		$('html, body').animate({scrollTop:schH}, 'fast');
	}, 300);
		    }
		});
	}else{
		return false;
	}
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

function fnTalkListRelold(){
	document.location.reload();
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<div class="content bgGry">
			<h1 class="hidden">응원톡!</h1>
			<% If isFilter="Y" Or searchtxt<>"" Or oCheerTalk.FresultCount > 0 Then %>
			<form name="searchForm" id="searchForm" method="get" style="margin:0px;" onSubmit="fnSearchList();return false;" action="">
			<input type="hidden" name="statediv" id="statediv" value="<%=statediv%>">
			<input type="hidden" name="isFilter" id="isFilter" value="<%=isFilter%>">
			<div class="artSearchTop">
				<div class="searchInput hasOpt"><%' for dev msg : 검색의 구분선택이 있는 경우 hasOpt 클래스 붙여주세요 %>
					<span class="schSlt">
						<select name="searchdiv">
							<option value="0" <% If searchdiv=1 Then Response.write "selected"%>>구분선택</option>
							<option value="1" <% If searchdiv=1 Then Response.write "selected"%>>작성자</option>
							<option value="2" <% If searchdiv=2 Then Response.write "selected"%>>글 내용</option>
						</select>
					</span>
					<input type="Search" name="searchtxt" placeholder="작성자, 글 내용 검색" value="<%=searchtxt%>" onKeyPress="if (event.keyCode == 13){ fnSearchList(); return false;}" maxlength="32" />
					<button type="button" class="btnSearch" onClick="fnSearchList();return false;">검색</button>
				</div>
			</div>
			</form>
			<% End If %>
			<div class="boardList cmtListWrap" id="cheer">
			<% If oCheerTalk.FTotalCount > 0 Then %>
				<ul class="artList">
				<% For j = 0 To oCheerTalk.FResultCount - 1 %>
					<%
						If Trim(oCheerTalk.FItemList(j).FUserid) = Trim(lecturer_id) Then
							ItsMeStr = " teacher"
						Else
							ItsMeStr = ""
						End If
					%>
					<% If lecturer_id = oCheerTalk.FItemList(j).FUserid Then  %>
					<li class='<%= Chkiif(oCheerTalk.FItemList(j).FReply_num <> 0, "reply master", "") %>'>
						<p class="writer<%= ItsMeStr %>"><%= oCheerTalk.FItemList(j).FUserid %></p>
					<% Else %>
					<li <%= Chkiif(oCheerTalk.FItemList(j).FReply_num <> 0, "class='reply'", "") %>>
						<p class="writer<%= ItsMeStr %>"><%= printUserId(oCheerTalk.FItemList(j).FUserid, 3, "*") %></p>
					<% End If %>
						<p class="txt"><%= nl2br(oCheerTalk.FItemList(j).FComment) %></p>
						<p class="txtInfo"><%= FormatDate(oCheerTalk.FItemList(j).FRegdate,"0000.00.00") %></p>
						<div class="btnGroup">
							<% If oCheerTalk.FItemList(j).FReply_num = 0 Then %>
							<button type="button" onclick="cheerTalkWrite('reply', '<%= oCheerTalk.FItemList(j).FReply_group_idx %>', '');" class="btn btnGrn">답글</button>
							<% End If %>
						<% If lecturer_id = oCheerTalk.FItemList(j).FUserid Then %>
							<button type="button" class="btn btnWht" onclick="cheerTalkWrite('edit', '<%= oCheerTalk.FItemList(j).FReply_group_idx %>','<%= oCheerTalk.FItemList(j).Fidx %>');">수정</button>
							<button type="button" class="btn btnWht" onclick="cheerTalkDel('<%= oCheerTalk.FItemList(j).Freply_group_idx %>','<%= oCheerTalk.FItemList(j).Fidx %>','<%= oCheerTalk.FItemList(j).Freply_depth %>');">삭제</button>
						<% End If %>
						</div>
					</li>
				<% Next %>
				</ul>
				<form method="post" name="delfrm">
				<input type="hidden" name="mode" value="del">
				<input type="hidden" name="idx" value="">
				<input type="hidden" name="ridx" value="">
				<input type="hidden" name="depth" value="">
				<input type="hidden" name="gubun" value="L">
				<input type="hidden" name="paramid" value="<%= lecturer_id %>">
				</form>
				<div class="paging">
					<a href="" onclick="fngopage('<%= oCheerTalk.FCurrPage %>','b','<%= lecturer_id %>'); return false;" class="btnPrev"><span>이전 페이지</span></a>
					<span><input type="number" id="textpage2" class="pageNum" maxlength = "4" onkeydown="return onlyNumber2(event, '<%= lecturer_id %>')" style='ime-mode:disabled;' value="<%= CInt(oCheerTalk.FCurrPage) %>" />		 / <%= CInt(oCheerTalk.FTotalPage) %></span>
					<a href="" onclick="fngopage('<%= oCheerTalk.FCurrPage + 1 %>','f','<%= lecturer_id %>'); return false;" class="btnNext"><span>다음 페이지</span></a>
				</div>
			<% Else %>
				<div class="noData"><span>등록된 응원톡이 없습니다.</span></div>
			<% End If %>
			</div>
		</div>
		<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<% SET oCheerTalk = nothing %>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->