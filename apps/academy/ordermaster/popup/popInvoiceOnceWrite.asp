<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 송장 일괄 입력"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/ordermaster/ordercls.asp"-->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<%
dim MakerID, searchdiv, oitem, ix, searchtxt, page, CheckOrder, iy

MakerID = requestCheckVar(request.cookies("partner")("userid"),32)
searchtxt = RequestCheckVar(request("searchtxt"),32)
page = RequestCheckVar(request("page"),3)
searchdiv = requestCheckVar(request("searchdiv"),1)
If page="" Then page=1
If searchdiv="" Then searchdiv=0
'Response.write ordercheck
'Response.end

set oitem = new CJumunMaster
oitem.FPageSize = 5
oitem.FCurrPage = page
oitem.FRectDesignerID = MakerID
oitem.FRectSearchDIV = searchdiv
oitem.FRectSearchTXT = searchtxt
oitem.InvoiceBatchWriteList
iy=0
CheckOrder=""
'' 택배사 일괄적용
Sub drawSelectBoxDeliverCompanyAssign(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>택배사를 선택해 주세요</option><%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("divcd")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub
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

	//상단 고정
	var jbOffset= new Array();
	$('.artStatus').each(function(i){
		jbOffset[i]=eval("$('#FixedOrderInfo" +i +"').offset().top");
	});
	$(window).scroll(function(){
		$('.artStatus').each(function(i){
			if($(document).scrollTop()>jbOffset[i]){
				eval("$('#FixedOrderInfo" +i +"')").addClass('jbFixed');
			}
			else{
				eval("$('#FixedOrderInfo" +i +"')").removeClass('jbFixed');
			}
		});
	});


});

function fnAppCallWinConfirm(){
	if($("input[name=songjangno]").length>0){
		var checksongjang=true;
		$("input[name=songjangno]").each(function(i){
			if($("select[name=songjangdiv]:eq(" + i + ")").val()==""){
				alert("택배사를 선택해 주세요.");
				$("select[name=songjangdiv]:eq(" + i + ")").focus();
				checksongjang=false;
				return false;
			}else if($("input[name=songjangno]:eq(" + i + ")").val()==""){
				alert("운송장 번호를 입력해 주세요.");
				$("select[name=songjangdiv]:eq(" + i + ")").focus();
				checksongjang=false;
				return false;
			}
		});

		if(checksongjang){
			if(confirm("선택한 주문건을 모두 출고처리하시겠습니까?")){
				document.sfrm.action="/apps/academy/ordermaster/popup/dosongjangbatchinput.asp";
				document.sfrm.target="FrameCKP";
				document.sfrm.submit();
			}
		}
	}
}

function fnSongjangInputEnd(OrderStateNum){
	alert("선택된 주문건 모두 출고 처리되었습니다");
	setTimeout(function(){
		fnAPPChangeBadgeCount("ordercount",OrderStateNum);
	}, 300);
	setTimeout(function(){
		fnAPPParentsWinReLoad();
	}, 600);
	setTimeout(function(){
		fnAPPclosePopup();
	}, 900);
}

function jsGoPage(iP){
	document.sfrm.page.value = iP;
	document.sfrm.submit();
}

function fnSearchList(){
	frm=document.sfrm
	if(frm.searchdiv.value==0){
		alert("구분을 선택해주세요.");
		frm.searchdiv.focus();
	}else if(frm.searchtxt.value==""){
		alert("검색어를 입력해주세요.");
		frm.searchdiv.focus();
	}else{
		frm.submit();
	}
}
</script>
<style>
.artStatus {transform: translate3d(0,0,0);} /* IOS 스크롤 시 깜박임, 사라짐 방지 */
.jbFixed{position:fixed;top:0px;z-index:100}
</style>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<div class="header"></div>
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">송장 일괄 입력</h1>
			<div class="invoiceWrite">
				<form method="post" name="sfrm" onSubmit="fnSearchList();return false;" action="">
				<input type="hidden" name="page" id="page" value="1">
				<% If searchtxt<>"" Or oitem.FresultCount > 0 Then %>
				<div class="artSearchTop">
					<div class="searchInput hasOpt">
						<span class="schSlt">
							<select name="searchdiv">
								<option value="0"<% If searchdiv=0 Then Response.write " selected"%>>구분선택</option>
								<option value="1"<% If searchdiv=1 Then Response.write " selected"%>>주문번호</option>
								<option value="2"<% If searchdiv=2 Then Response.write " selected"%>>작품코드</option>
								<option value="3"<% If searchdiv=3 Then Response.write " selected"%>>구매자</option>
								<option value="4"<% If searchdiv=4 Then Response.write " selected"%>>주문작품명</option>
							</select>
						</span>
						<input type="Search" name="searchtxt" placeholder="주문번호, 작품코드, 구매자, 주문작품명 검색" value="<%=searchtxt%>" onKeyPress="if (event.keyCode == 13){ fnSearchList(); return false;}" />
						<button type="button" class="btnSearch" onClick="fnSearchList();return false;">검색</button>
					</div>
				</div>
				<% End If %>
				<% If oitem.FResultCount>0 Then %>
				<% For ix=0 To oitem.FResultCount-1 %>
				<% If CheckOrder<>oitem.FMasterItemList(ix).FOrderserial Then %>
				<% If ix<>0 Then %>
				</div>
				<% End If %>
				<div class="invoiceGrp">
					<div class="grpInfo">
						<div class="artStatus" id="FixedOrderInfo<%=iy%>">
							<p><span><%=FormatDate(oitem.FMasterItemList(ix).Fipkumdate,"0000.00.00")%></span><span class="cGy4">ㅣ</span><span><%=oitem.FMasterItemList(ix).FOrderserial%></span></p>
							<p class="rt"><span class="nowStatus"><strong><%=oitem.FMasterItemList(ix).Fbuyname%></strong></span></p>
						</div>
						<div class="invoiceAddr">[<%=oitem.FMasterItemList(ix).Freqzipcode%>] <%=oitem.FMasterItemList(ix).Freqzipaddr%> <%=oitem.FMasterItemList(ix).Freqaddress%></div>
					</div>
				<%iy=iy+1%>
				<% End If %>
				<% CheckOrder=oitem.FMasterItemList(ix).FOrderserial %>
					<ul class="artList">
						<li>
							<div class="artInfo">
								<div class="artThumb"><img src="<%=oitem.FMasterItemList(ix).FListimage%>" alt="" onerror="this.src='http://image.thefingers.co.kr/apps/2016/thumb_default.png'" /></div>
								<p class="orderNo"><%=oitem.FMasterItemList(ix).FItemid%></p>
								<strong><%=oitem.FMasterItemList(ix).FItemname%></strong>
								<div class="artTxt">
									<p><dfn><%=oitem.FMasterItemList(ix).Fitemoptionname%></dfn></p>
									<p><dfn><%=oitem.FMasterItemList(ix).Fitemno%>개</dfn></p>
								</div>
							</div>
							<% If oitem.FMasterItemList(ix).Frequiredetail<>"" Then %>
							<div class="boxUnit bdrTRtGry">
								<div class="boxHead">
									<b>주문제작 메시지</b>
								</div>
								<div class="boxCont"><%=oitem.FMasterItemList(ix).Frequiredetail%></div>
							</div>
							<% End If %>
						</li>
						<li class="selectBtn">
							<% drawSelectBoxDeliverCompanyAssign "songjangdiv", oitem.FMasterItemList(ix).Fsongjangdiv %>
							<input type="hidden" name="detailidx" value="<%=oitem.FMasterItemList(ix).Fdetailidx%>">
							<input type="hidden" name="orderserial" value="<%=oitem.FMasterItemList(ix).FOrderserial%>">
						</li>
						<li class="list">
							<dfn><b>운송장 번호</b></dfn>
							<div><input type="number" name="songjangno" id="songjangno" placeholder="운송장 번호를 입력해주세요" pattern="[0-9]*" inputmode="numeric" min="0" /></div>
						</li>
					</ul>
				<% Next %>
					<% if oitem.FTotalCount>oitem.FPageSize then %>
					<div class="paging">
						<%=fnDisplayPaging_New(page,oitem.FTotalCount,oitem.FPageSize,"jsGoPage")%>
					</div>
					<% end if %>
				<% Else %>
				<div class="artNo" style="display:">
					<div class="linkNotice">
						<p class="fs1-5r">진행중인 주문이 없습니다.</p>
					</div>
				</div>
				<% End If %>
				</form>
			</div>
		</div>
		<div class="footer"></div>
		<!--// content -->
	</div>
</div>
</body>
</html>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
	if($("input[name=songjangno]").length>0){
		fnAPPShowRightConfirmBtns();
	}
});
//-->
</script>
<%
Set oitem = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->