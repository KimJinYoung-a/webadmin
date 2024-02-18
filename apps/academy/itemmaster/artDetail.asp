<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Session.codepage="65001"
Response.ContentType="text/html;charset=UTF-8"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 작품 정보"

dim oitem, itemid, makerid, i
itemid = RequestCheckVar(request("itemid"),10)
makerid = request.cookies("partner")("userid")
If makerid="" Then makerid=RequestCheckVar(request("makerid"),32)
set oitem = new CItem
oitem.FRectMakerId = makerid
oitem.FRectItemID = itemid
if (makerid<>"") then
oitem.GetOneItem
End If

'Response.write oitem.FTotalCount
'Response.end

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

if oitem.FTotalCount<=0 then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');fnAPPclosePopup();</script>"
	Response.end
end if

'// 가격수정요청 내용 접수
Dim clsItem, arrList, editRstTotCnt, chkRstStat, oldsellcash
set clsItem = new CUpCheItemEdit
	clsItem.FRectMakerId = makerid
	clsItem.FRectItemId = itemid
	if (makerid<>"") then
		arrList = clsItem.fnGetItemPriceChangeInfo
		editRstTotCnt	= clsItem.FResultCount
		if editRstTotCnt>0 then
			oldsellcash = arrList(2,0)	'변경전 판매가
			chkRstStat = arrList(8,0)	'요청접수 상태 (N:대기, Y:완료, D:반려)
		end if
	end if
set clsItem = nothing
%>

<script>
$(function() {
	// button tab
	$(".selectBtn:not('.priceBtn') button").click(function(){
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
	setTimeout(function(){
		$('html, body').animate({scrollTop:schH-tabT}, 'fast');
	}, 300);

	fnAPPShowRightConfirmBtns();//확인 버튼 활성화 호출 함수
});

function fnStateSelect(fname,selectdata){
	if(fname!=''){
		eval("$('#"+fname+"')").val(selectdata);
	}
}

function fnAppCallWinConfirm(){
	if($("#isusing").val()=="N"){
		if(confirm("사용 안함 선택 시 판매중인 작품은 판매가 정지되며,\n 저장된 작품 정보가 제거됩니다.\n사용하지 않으시겠습니까?")){
			document.itemstate.action="/apps/academy/itemmaster/DIYItemDetailinfoEdit_Process_App.asp";
			document.itemstate.target = "FrameCKP";
			document.itemstate.submit();
		}
	}else{
		document.itemstate.action="/apps/academy/itemmaster/DIYItemDetailinfoEdit_Process_App.asp";
		document.itemstate.target = "FrameCKP";
		document.itemstate.submit();
	}
}
function fnDetailStateInfoEnd(){
	fnAPPopenerJsCallClose("fnSearchFilterSet(\"\")");
}

function fnItemPriceEditEnd(callbackdata){
	if(callbackdata==1) {
		$("#btnPriceEdit").hide();
		$("#btnPriceWait").show()
	}
}

function fnDetailPageMove(url){
	location.href=url;
}

function fnCancelPriceChng() {
	if(confirm("요청한 가격변경 요청을 취소하시겠습니까?")) {
		document.itemstate.action="/apps/academy/itemmaster/popPriceChange_Proc.asp";
		document.itemstate.hidM.value="C";
		document.itemstate.target = "FrameCKP";
		document.itemstate.submit();
	}
}

function fnConfirmPriceChng(idx) {
	document.itemstate.action="/apps/academy/itemmaster/popPriceChange_Proc.asp";
	document.itemstate.hidM.value="R";
	document.itemstate.idx.value=idx;
	document.itemstate.target = "FrameCKP";
	document.itemstate.submit();
}

</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">작품 정보</h1>
			<div class="artDetailInfo">
				<ul class="listTab">
					<li class="current" onclick="fnDetailPageMove('<%=g_AdminURL%>/apps/academy/itemmaster/artDetail.asp?itemid=<%=itemid%>&makerid=<%=makerid%>')"><div>기본 정보</div></li>
					<li onclick="fnDetailPageMove('<%=g_AdminURL%>/apps/academy/itemmaster/artItemEdit.asp?itemid=<%=itemid%>&makerid=<%=makerid%>')"><div>수정</div></li>
				</ul>
				<div class="artDetailWrap">
					<div class="artDetail">
						<ul class="artList">
							<li class="<% If oitem.FOneItem.isTempSoldOut Then %>artFlag2<% ElseIf oitem.FOneItem.IsSoldOut Then %>artFlag3<% Else %>artFlag1<% End If %>"><!-- 판매중(↓ 상태표시에 따라 클래스 artFlag1 ~ artFlag8 붙습니다) //-->
								<a href="">
									<div class="artStatus">
										<p><span><%= oitem.FOneItem.Fitemid %></span></p>
										<p class="rt"><span class="nowStatus"><strong><%= oitem.FOneItem.IsSellYnName %></strong></span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="<%= oitem.FOneItem.Flistimage %>" alt="" /></div>
										<strong><%= oitem.FOneItem.Fitemname %></strong>
										<div class="artTxt">
											<p><span><% If oitem.FOneItem.IsUpcheBeasong Then %>업체배송<% Else %><% End If %></span><% If oitem.FOneItem.Fitemdiv="06" Then %><span class="sepLine">l</span><span>주문제작 상품</span><% End If %></p>
										</div>
									</div>
								</a>
							</li>
						</ul>
						<dl class="dfCompos">
							<dt>가격 정보
							<%
								if editRstTotCnt>0 then
									if chkRstStat="Y" then
							%>
								<i class="tag2">변경 완료</i>
							<%		elseif chkRstStat="D" then %>
								<i class="tag2">요청 반려</i>
							<%		else %>
								<i class="tag1">변경 대기중</i>
							<%
									end if
								end if
							%>
							</dt>
							<dd>
								<ul class="list">
									<li class="">
										<dfn><b>소비자가</b></dfn>
										<% If oitem.FOneItem.IsSaleItem Then %>
										<div><span class="rPad0-8r"><s><% =FormatNumber(oitem.FOneItem.getOrgPrice,0) %>원</s></span><span class="rPad0-8r"><% =FormatNumber(oitem.FOneItem.getRealPrice,0) %>원</span><strong class="cOr1"><% =oitem.FOneItem.getSalePro %></strong></div>
										<% Else %>
										<div><span class="rPad0-8r"><%= FormatNumber(oitem.FOneItem.Fsellcash,0) %>원</span></div>
										<% End If %>
									</li>
									<li class="">
										<dfn><b>공급가</b></dfn>
										<div class="cGy2"><span><% =FormatNumber(oitem.FOneItem.Fsailsuplycash,0) %>원</span><span class="sepLine">l</span><span>업체</span><span class="sepLine">l</span><span>마진 <%=fnPercent(oitem.FOneItem.Forgsuplycash,oitem.FOneItem.Forgprice,1)%></span></div>
									</li>
								</ul>
							</dd>
							<% if editRstTotCnt>0 then %>
							<dd class="tPad2r <%=chkIIF(chkRstStat="D","disabled","")%>">
								<div class="boxUnit bdrTGry">
									<div class="boxHead">
										<b>가격 변경 요청</b>
										<p><span><%=formatnumber(arrList(2,0),0)%>원</span><i class="chgArw"></i><span><strong class="cOr1"><%=formatnumber(arrList(4,0),0)%>원</strong></span></p>
									</div>
									<div class="boxCont"><%=arrList(6,0)%></div>
								</div>
							</dd>
								<% if chkRstStat="D" then %>
								<dd class="tPad2r">
									<div class="boxUnit bdrTOr">
										<div class="boxHead">
											<b>반려 사유</b>
										</div>
										<div class="boxCont"><%=arrList(7,0)%></div>
									</div>
								</dd>
								<% end if %>
							<% end if %>

							<% If oitem.FOneItem.IsSaleItem Then %>
							<dd class="selectBtn priceBtn tMar2-5r">
								<div><button type="button" class="btnM1 btnGry disabled">가격 변경 요청</button></div>
							</dd>
							<% else %>
							<dd id="btnPriceWait" class="selectBtn priceBtn tMar2-5r" style="<%=chkIIF(editRstTotCnt>0 and chkRstStat<>"Y","","display:none;")%>">
								<%	if chkRstStat="D" then %>
								<div class="grid2"><button type="button" class="btnM1 btnGry selected" onclick="fnConfirmPriceChng(<%=arrList(11,0)%>);">반려 확인</button></div>
								<%	else %>
								<div class="grid2"><button type="button" class="btnM1 btnGry selected" onclick="fnAPPpopupItemPriceEdit('<%=g_AdminURL%>/apps/academy/itemmaster/popPriceChange.asp?itemid=<%= oitem.FOneItem.Fitemid %>')">가격 재변경 요청</button></div>
								<%	end if %>
								<div class="grid2"><button type="button" class="btnM1 btnGry" onclick="fnCancelPriceChng();">변경 요청 취소</button></div>
							</dd>
							<dd id="btnPriceEdit" class="selectBtn priceBtn tMar2-5r" style="<%=chkIIF(editRstTotCnt>0 and chkRstStat<>"Y","display:none;","")%>">
								<div><button type="button" class="btnM1 btnGry selected" onclick="fnAPPpopupItemPriceEdit('<%=g_AdminURL%>/apps/academy/itemmaster/popPriceChange.asp?itemid=<%= oitem.FOneItem.Fitemid %>')">가격 변경 요청</button></div>
							</dd>
							<% end if %>
						</dl>
						<dl class="dfCompos">
							<dt>재고 현황</dt>
							<dd>
								<ul class="list">
									<% if oitemoption.FResultCount>0 then %>
									<% for i=0 to oitemoption.FResultCount - 1 %>
									<li class="cGy3">
										<dfn class="cGy1"><b><%= oitemoption.FITemList(i).FItemOption %></b></dfn>
										<div class="optName"><div><%= oitemoption.FITemList(i).FOptionName %></div></div>
										<div class="rt"><%= oitemoption.FITemList(i).GetOptLimitEa %></div>
									</li>
									<% next %>
									<% Else %>
									<li class="cGy3">
									<div class="optName">재고 : <% If oitem.FOneItem.Flimityn="Y" Then %><% =FormatNumber(oitem.FOneItem.Flimitno,0) %><% End If %></div>
									</li>
									<% End If %>
								</ul>
							</dd>
						</dl>
						<form name="itemstate" method="post">
						<input type="hidden" name="sellyn" id="sellyn" value="<%= oitem.FOneItem.Fsellyn %>">
						<input type="hidden" name="isusing" id="isusing" value="<%= oitem.FOneItem.Fisusing %>">
						<input type="hidden" name="hidM" value="">
						<input type="hidden" name="idx" value="">
						<input type="hidden" name="itemid" id="itemid" value="<%= oitem.FOneItem.Fitemid %>">
						<input type="hidden" name="oldsellcash" value="<%= oldsellcash %>">
						<input type="hidden" name="makerid" value="<%= makerid %>">
						
						<dl class="dfCompos">
							<dt>판매 관리</dt>
							<dd class="selectBtn">
								<div class="grid3"><button type="button" class="btnM1 btnGry<% If oitem.FOneItem.Fsellyn="Y" Then %> selected<% End If %>" onclick="fnStateSelect('sellyn','Y')">판매</button></div>
								<div class="grid3"><button type="button" class="btnM1 btnGry<% If oitem.FOneItem.Fsellyn="S" Then %> selected<% End If %>" onclick="fnStateSelect('sellyn','S')">일시품절</button></div>
								<div class="grid3"><button type="button" class="btnM1 btnGry<% If oitem.FOneItem.Fsellyn="N" Then %> selected<% End If %>" onclick="fnStateSelect('sellyn','N')">품절</button></div>
							</dd>
						</dl>
						<dl class="dfCompos">
							<dt>사용 여부</dt>
							<dd class="selectBtn">
								<div class="grid2"><button type="button" class="btnM1 btnGry selected" onclick="fnStateSelect('isusing','Y')">사용</button></div>
								<div class="grid2"><button type="button" class="btnM1 btnGry" onclick="fnStateSelect('isusing','N')">사용안함</button></div>
							</dd>
						</dl>
						</form>
					</div>
				</div>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
<% if (application("Svr_Info")	= "Dev") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>
</body>
</html>
<%
set oitem = Nothing
set oitemoption = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->