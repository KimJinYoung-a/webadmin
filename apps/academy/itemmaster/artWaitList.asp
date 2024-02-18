<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Session.codepage="65001"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 작품 관리"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<%
dim sellyn, limityn, sdiv, sortupdown, isFilter
dim page, searchtxt, owaititem, cate1, cate2

searchtxt = RequestCheckVar(request("searchtxt"),32)
sellyn  = RequestCheckVar(request("sellyn"),2)
sortupdown = RequestCheckVar(request("sortupdown"),1)
limityn = RequestCheckVar(request("limityn"),1)
sdiv = RequestCheckVar(request("sdiv"),10)
page = RequestCheckVar(request("page"),10)
cate1 = RequestCheckVar(request("cate1"),10)
cate2 = RequestCheckVar(request("cate2"),10)
isFilter = RequestCheckVar(request("isFilter"),1)

if (page="") then page=1
If (limityn="") Then limityn="A"
If (sellyn="") Then sellyn="YS"
If (sortupdown="") Then sortupdown="u"

'필터 파라메터
dim strParam
strParam = "sellyn=" & sellyn & "&sortupdown=" & sortupdown & "&limityn=" & limityn & "&sdiv=" & sdiv & "&cate1=" & getNumeric(cate1) & "&cate2=" & getNumeric(cate2)
'==============================================================================
set owaititem = new CWaitItemlist
owaititem.FPageSize = 12
owaititem.FCurrPage = page
owaititem.FRectDesignerID = request.cookies("partner")("userid")
owaititem.FRectSortUpDown = sortupdown
owaititem.FRectSearchTxt = searchtxt
owaititem.FRectLimityn = limityn
owaititem.FRectSortDiv = sdiv
owaititem.FRectCate_Large = cate1
owaititem.FRectCate_Mid = cate2
owaititem.FRectCurrState = sellyn
owaititem.WaitProductList

dim i
%>
<script>
$(function() {
	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});

	// textarea auto size
	$(".searchInput input").keyup(function () {
		$(this).parent().find('button').fadeIn();
	});

	// search box hidden scroll top auto change
	var schH = $(".artSearchTop").outerHeight();
	var tabT = $(".listTab").offset().top;
	$("body").css("min-height",schH+$(window).height());
	setTimeout(function(){
		$('html, body').animate({scrollTop:schH-tabT}, 'fast');
	}, 300);
});

function jsGoPage(iP){
	document.searchForm.page.value = iP;
	document.searchForm.submit();
}
function fnGoTapPage(param){
	if(param=="YS"){
		location.href="/apps/academy/itemmaster/artList.asp?sellyn="+param;
	}else if(param=="N"){
		location.href="/apps/academy/itemmaster/artList.asp?sellyn="+param;
	}else{
		location.href="/apps/academy/itemmaster/artWaitList.asp";
	}
}

function fnSearchFilterSet(callbackval){
	//alert(callbackval);
	var catearr = callbackval.replace(/ /g, "");
	var catearr2 = catearr.replace(/,/g, "','");
	var catearr3=eval("['" + catearr2 + "']");
	$("#cate1").val(catearr3[0]);
	$("#cate2").val(catearr3[1]);
	$("#sellyn").val(catearr3[2]);
	$("#limityn").val(catearr3[3]);
	$("#sdiv").val(catearr3[4]);
	$("#sortupdown").val(catearr3[5]);
	$("#isFilter").val(catearr3[6]);
	document.searchForm.submit();
}

function fnSearchList(){
	document.searchForm.submit();
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">작품 관리</h1>
			<div class="artManage">
				<ul class="listTab">
					<li onclick="fnGoTapPage('YS');"><div>판매중</div></li>
					<li onclick="fnGoTapPage('N');"><div>판매종료</div></li>
					<li class="current" onclick="fnGoTapPage('W');"><div>등록대기</div></li>
				</ul>
				<form name="searchForm" id="searchForm" method="get" style="margin:0px;" onSubmit="fnSearchList();return false;" action="">
				<input type="hidden" name="page" id="page" value="1">
				<input type="hidden" name="cate1" id="cate1" value="<%=cate1%>">
				<input type="hidden" name="cate2" id="cate2" value="<%=cate2%>">
				<input type="hidden" name="sellyn" id="sellyn" value="<%=sellyn%>">
				<input type="hidden" name="limityn" id="limityn" value="<%=limityn%>">
				<input type="hidden" name="sdiv" id="sdiv" value="<%=sdiv%>">
				<input type="hidden" name="sortupdown" id="sortupdown" value="<%=sortupdown%>">
				<input type="hidden" name="isFilter" id="isFilter" value="<%=isFilter%>">
				<div class="artListCont">
				<% if isFilter="Y" or searchtxt<>"" or owaititem.FresultCount > 0 then %>
					<div class="artSearchTop">
						<div class="searchInput">
							<input type="search" name="searchtxt" placeholder="작품명, 코드, 키워드 검색" value="<%=searchtxt%>" onKeyPress="if (event.keyCode == 13){ fnSearchList(); return false;}" />
							<button type="button" class="btnSearch" onClick="fnSearchList();return false;">검색</button>
							<!-- button type="button" class="btnTextDel">삭제</button -->
						</div>
						<div class="btnFilter <%=chkIIF(isFilter="Y","filterActive","")%>">
							<button type="button" onclick="fnAPPpopupSearchFilter('<%=g_AdminURL%>/apps/academy/itemmaster/popFilter.asp?div=1&<%=strParam%>')">필터</button>
						</div>
					</div>
				<% end if %>
				<% if owaititem.FresultCount > 0 then %>
					<div class="artListWrap">
						<ul class="artList">
							<% For i=0 To owaititem.FresultCount-1 %>
							<li class="<%= owaititem.FItemList(i).GetCurrStateCssClass %>"><!-- 판매중(↓ 상태표시에 따라 클래스 artFlag1 ~ artFlag8 붙습니다) //-->
								<a href="javascript:fnAPPpopupWaitItemListEdit('<%=g_AdminURL%>/apps/academy/itemmaster/artRegistEdit.asp?waititemid=<%= owaititem.FItemList(i).Fitemid %>')">
									<div class="artStatus">
										<p><span><%= FormatDate(owaititem.FItemList(i).Fregdate,"0000.00.00") %></span><span>ㅣ</span><span><%= owaititem.FItemList(i).Fitemid %></span></p>
										<p class="rt"><span class="nowStatus"><%= owaititem.FItemList(i).GetCurrStateName %></span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="<%= owaititem.FItemList(i).Fbasicimage %>" alt="" onerror="this.src='http://image.thefingers.co.kr/apps/2016/thumb_default.png'"/></div>
										<strong><% =owaititem.FItemList(i).Fitemname %></strong>
										<div class="artTxt">
											<p class="tPad1r"><span><%= FormatNumber(owaititem.FItemList(i).Fsellcash,0) %>원</span></p>
										</div>
									</div>
								</a>
							</li>
							<% Next %>
						</ul>

						<% if owaititem.FTotalCount>owaititem.FPageSize then %>
						<div class="paging">
							<%=fnDisplayPaging_New(page,owaititem.FTotalCount,owaititem.FPageSize,"jsGoPage")%>
						</div>
						<% end if %>
					</div>
				<% else %>
					<div class="artNo">
						<div class="linkNotice">
							<p class="fs1-5r">등록대기중인 작품이 없습니다.</p>
						</div>
					</div>
				<% End If %>
				</div>
				</form>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->