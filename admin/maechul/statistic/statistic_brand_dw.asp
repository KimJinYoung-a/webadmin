<%@ language=vbscript %>
<% option explicit

	'스크립트 타임아웃 시간 조정 (기본 90초)
	'''Server.ScriptTimeout = 180 ''주석처리 2016/04/08 eastone
%>
<%
'####################################################
' Description :  브랜드별매출
' History : 2016.01.20 서동석 생성
'			2020.01.15 정태훈 엑셀다운로드 추가
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_dw.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSorting, chkChannel,vBrandID, rdsite	' , vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay
dim sellchnl, inc3pl, vCateL, vCateM, vCateS, vIsBanPum, vPurchasetype, v6Ago, mwdiv, dispCate, page, pagesize
Dim vTot_OrderCnt, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer
Dim vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2, vTot_upcheJungsan, vTot_avgipgoPrice, vTot_itemsku
Dim incStockAvg, groupUserLevel, imax, imin, vTot_overValueStockPrice, vstartdate, venddate, totalcolspan, isSendGift
	vstartdate = NullFillWith(requestCheckVar(request("startdate"),10),DateAdd("d",0,date()))
	venddate = NullFillWith(requestCheckVar(request("enddate"),10),date())
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")  ''beasongdate  :출고일=>주문일 2018/05/28  by eastone
	'vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	'vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	'vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	'vEYear		= NullFillWith(request("eyear"),Year(now))
	'vEMonth		= NullFillWith(request("emonth"),Month(now))
	'vEDay		= NullFillWith(request("eday"),Day(now))
	vSorting	= NullFillWith(request("sorting"),"itemcost")
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	vBrandID	= NullFillWith(request("ebrand"),"")
	dispCate = requestCheckvar(request("disp"),16)
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(request("mwdiv"),"")
	rdsite       = NullFillWith(request("rdsite"),"")
	inc3pl = request("inc3pl")
    chkChannel  = requestCheckvar(request("chkChl"),1)
	page  = requestCheckvar(request("page"),10)
	pagesize  = requestCheckvar(request("pagesize"),10)
	incStockAvg = requestCheckvar(request("incStockAvg"),10)
	groupUserLevel = requestCheckvar(request("groupUserLevel"),1)
	isSendGift	= requestCheckvar(request("isSendGift"),1)
	totalcolspan=0

if (page = "") then
	page = 1
end if

if (pagesize = "") then
	pagesize = "100"
end if

Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectSort = vSorting
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectStartdate = vstartdate		' vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = venddate		'vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	'cStatistic.FRect6MonthAgo = v6Ago
	'cStatistic.FRectChannelDiv = channelDiv
	cStatistic.FRectMakerid = vBrandID
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectRdsite = rdsite
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 추가
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectChkchannel = chkChannel
	cStatistic.FCurrPage = page
	cStatistic.FPageSize = pagesize
	cStatistic.FRectIncStockAvgPrc = (incStockAvg<>"") ''true '' 평균매입가 포함 쿼리여부.
	cStatistic.FRectGroupUserLevel = groupUserLevel
	cStatistic.FRectIsSendGift = isSendGift
	cStatistic.fStatistic_brand()

dim iTotalPage
	iTotalPage 	=  int((cStatistic.FTotalCount)/pagesize) +1

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function searchSubmit()
{
   // if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
		//if (MonthDiff(frm.syear.value + "-" + frm.smonth.value + "-" + frm.sday.value, frm.eyear.value + "-" + frm.emonth.value + "-" + frm.eday.value) >= 1) {
		//	alert("최대 1개월까지만 검색이 가능합니다.");
		//	return;
		//}

		$("#btnSubmit").prop("disabled", true);
		document.frm.target="";
		document.frm.action="";
		document.frm.submit();
	//}

/*
	if((frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>) && (frm.is6ago.checked == false))
	{
		alert("6개월전의 데이터는 6개월이전데이터를 체크하셔야 가능합니다.");
	}
	else
	{
		if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
			//if (MonthDiff(frm.syear.value + "-" + frm.smonth.value + "-" + frm.sday.value, frm.eyear.value + "-" + frm.emonth.value + "-" + frm.eday.value) >= 1) {
			//	alert("최대 1개월까지만 검색이 가능합니다.");
			//	return;
			//}

			$("#btnSubmit").prop("disabled", true);
			frm.submit();
		}
	}
*/
}

function MonthDiff(d1, d2) {
	d1 = d1.split("-");
	d2 = d2.split("-");

	d1 = new Date(d1[0], d1[1] - 1, d1[2]);
	d2 = new Date(d2[0], d2[1] - 1, d2[2]);

	var d1Y = d1.getFullYear();
	var d2Y = d2.getFullYear();
	var d1M = d1.getMonth();
	var d2M = d2.getMonth();

	return (d2M+12*d2Y)-(d1M+12*d1Y);
}

function DateCheck()
{
	var date1 = new Date(frm.syear.value,frm.smonth.value,frm.sday.value);
	var date2 = new Date(frm.eyear.value,frm.emonth.value,frm.eday.value);

	//월 비교값
	var years  = date2.getFullYear() - date1.getFullYear();
	var months = date2.getMonth() - date1.getMonth();
	var days   = date2.getDate() - date1.getDate();

	var chkmonth = years * 12 + months + (days >= 0 ? 0 : -1);

	//일 비교값
	var day   = 1000 * 3600 * 24;
	var chkday =  parseInt((date2 - date1) / day, 10);

	if(chkday > 31)
	{
		alert("날짜 검색은 1달 간격만 됩니다.");
		return false;
	}
	else
	{
		return true;
	}
}

// 엑셀 다운
function jsbrandDown(){
  document.frm.page.value = $('#selODCnt').val();
	document.frm.target="hidifr";
	document.frm.action="statistic_brand_dw_excel_download.asp";
	document.frm.submit();
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
	<td align="left">
		<table class="a" cellpadding="3">
		<tr>
			<td height="25">
				 기간:
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>주문일</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>결제일</option>
					<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>출고일</option>
					<option value="jfixeddt" <%=CHKIIF(vDateGijun="jfixeddt","selected","")%>>정산확정일</option>
				</select>
				<% 'DrawDateBoxdynamic vSYear,"syear",vEYear,"eyear",vSMonth,"smonth",vEMonth,"emonth",vSDay,"sday",vEDay,"eday" %>
				<input type="text" name="startdate" id="startdate" value="<%=vstartdate%>" style="text-align:center;height:35px;" size="10" maxlength="10" readonly>
				<strong>&nbsp;~&nbsp;</strong>
				<input type="text" name="enddate" id="enddate" value="<%=venddate%>" style="text-align:center;height:35px;" size="10" maxlength="10" readonly>
				<script type="text/javascript">
					var CAL_Start = new Calendar({
						inputField : "startdate", trigger    : "startdate",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "enddate", trigger    : "enddate",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</td>
		</tr>
		<tr>
			<td>
			    브랜드 : <input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');">
				&nbsp;&nbsp;<!-- #include virtual="/common/module/categoryselectbox.asp"-->
				&nbsp;&nbsp;전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		</td>
	</tr>
	<tr>
		<td>
			사이트:  <% Call Drawsitename("sitename", vSiteName)%>
			&nbsp;&nbsp;채널:
   			 <% drawSellChannelComboBoxGroup "sellchnl",sellchnl %>
			&nbsp;&nbsp;<b>매출처:</b> <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
			&nbsp;&nbsp;구매유형: 
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
			&nbsp;&nbsp;주문구분:
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>반품포함</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>반품제외</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>반품건만</option>
				</select>
				&nbsp;&nbsp;매입구분:
				<% Call DrawBrandMWUPCombo("mwdiv",mwdiv) %>
				&nbsp;&nbsp;
				판매처별:
				<% Call DrawRdsiteCombo("rdsite",rdsite) %>
		</td>
	</tr>
	<tr>
		<td>
				정렬: <input type="radio" name="sorting" value="itemno" <%=CHKIIF(vSorting="itemno","checked","")%>>수량순
				<input type="radio" name="sorting" value="itemcost" <%=CHKIIF(vSorting="itemcost","checked","")%>>매출순
				<input type="radio" name="sorting" value="profit" <%=CHKIIF(vSorting="profit","checked","")%>>수익순
				&nbsp;&nbsp;표시갯수:
				<select class="select" name="pagesize">
					<option value="100" <% if (pagesize = "100") then %>selected<% end if %> >100 개</option>
					<option value="500" <% if (pagesize = "500") then %>selected<% end if %> >500 개</option>
					<option value="1000" <% if (pagesize = "1000") then %>selected<% end if %> >1000 개</option>
					<option value="3000" <% if (pagesize = "3000") then %>selected<% end if %> >3000 개</option>
				</select>&nbsp;&nbsp;
				<label><input type="checkbox" name="incStockAvg" <%=CHKIIF(incStockAvg<>"","checked","")%>>평균매입가포함</label>&nbsp;&nbsp;
				<label><input type="checkbox" name="chkChl" id="chkChl" value="1" <%if chkChannel ="1" then%>checked<%end if%> onClick="$('#groupUserLevel').prop('checked',false)">채널상세보기</label>&nbsp;&nbsp;
				<label><input type="checkbox" name="groupUserLevel" id="groupUserLevel" value="1" <%if groupUserLevel ="1" then%>checked<%end if%> onClick="$('#chkChl').prop('checked',false)">회원등급별보기</label>&nbsp;&nbsp;
			    <label><input type="checkbox" name="isSendGift" value="Y" <%=CHKIIF(isSendGift<>"","checked","")%>>선물주문만 보기</label>
			</td>
		</tr>
	    </table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" id="btnSubmit" class="button_s" value="검색" onClick="searchSubmit();"></td>
</tr>
</table>
</form>
<br>
<!-- 검색 끝 -->
<table width="100%" cellpadding="0" cellspacing="0" class="a">
<tr bgcolor="#FFFFFF" >
	<td>
		* 검색 기간이 길어지면 상당히 느려집니다. 그러니 검색 버튼을 클릭한 뒤 아무 반응이 없어보인다고 재차 검색버튼을 클릭하지 마세요.<br />
		* 최대 8000개까지 표시됩니다.
	</td>
	<td align="right">
		<select name="selODCnt" id="selODCnt" class="select" style="height:25px;vertical-align:top;">
			<%for i =1 To Int(cStatistic.FTotalCount/2000)+1
					imin = ((i-1)*2000)+1
					if i <  Int(cStatistic.FTotalCount/2000)+1 then
					imax = i*2000
					else
					imax = cStatistic.FTotalCount
					end if
			%>
			<option value="<%=i%>"><%=imin%>~<%=imax%></option>
			<%Next%>
		</select>
		<input type="button" class="button" value="다운로드(엑셀)" onclick="jsbrandDown();">
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= cStatistic.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=iTotalPage%></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<% if chkChannel = "1" then %>
		<%
		' 구매유형:PB일경우
		if vPurchasetype = "3" then
		%>
			<td align="center">날짜</td>
		<% end if %>
	<% end if %>

	<td align="center">브랜드ID</td>
	<td align="center">구매유형</td>
	<%if chkChannel ="1" then%>
	<td align="center">채널</td>
	<%elseif groupUserLevel="1" then%>
	<td align="center">회원등급</td>
	<%end if%>
    <td align="center">상품수량</td>
	<%if chkChannel ="1" then%>
	<%elseif groupUserLevel="1" then%>
	<% else %>
		<td align="center">상품SKU</td>
	<%end if%>
    <% if (NOT C_InspectorUser) then %>
    <td align="center">소비자가[상품]</td>
    <td align="center">판매가[상품]<br>(할인적용)</td>
    <td align="center"><b>구매총액[상품]<br>(상품쿠폰적용)</b></td>
     <%if chkChannel ="1" then%>
    <td align="center">채널<br>점유율</td>
	<%elseif groupUserLevel="1" then%>
	<td align="center">등급<br>점유율</td>
    <%end if%>
    <td align="center"><b>보너스쿠폰<br>사용액[상품]</b></td>
    <% end if %>
    <td align="center">취급액</td>
    <td align="center">매입총액[상품]<% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %></td>
    <td align="center"><b>매출수익</b></td>
    <td align="center">수익율</td>
    <td align="center">매출수익2<br>(취급액기준)</td>
    <td align="center">수익율</td>
	<td align="center">업체<br>정산액</td>
	<td align="center"><b>회계매출</b></td>
	<td align="center">평균<br>매입가</td>
	<td align="center">재고<br>충당금</td>
    <td align="center">비고</td>
</tr>
<%
For i = 0 To cStatistic.FResultCount -1
%>
<tr bgcolor="#FFFFFF">
	<% if chkChannel = "1" then %>
		<%
		' 구매유형:PB일경우
		if vPurchasetype = "3" then
		%>
			<td align="center" <%=chkIIF(chkChannel ="1","rowspan=""7""","")%><%=chkIIF(groupUserLevel ="1","rowspan=""10""","")%>><%= cStatistic.FList(i).Fyyyymmdd %></td>
		<% end if %>
	<% end if %>

	<td align="center" <%=chkIIF(chkChannel ="1","rowspan=""7""","")%><%=chkIIF(groupUserLevel ="1","rowspan=""10""","")%>><%= cStatistic.FList(i).FMakerID %></td>
	<td align="center" <%=chkIIF(chkChannel ="1","rowspan=""7""","")%><%=chkIIF(groupUserLevel ="1","rowspan=""10""","")%>>
	<%= cStatistic.FList(i).fpurchasetypename %>
	</td>
	<%if chkChannel ="1" then%>
	<td align="center">전체</td>
	<%elseif groupUserLevel="1" then%>
	<td align="center">전체</td>
	<%end if%>
	<td align="center"><%= CDbl(cStatistic.FList(i).FItemNO) %></td>
	<%if chkChannel ="1" then%>
	<%elseif groupUserLevel="1" then%>
	<% else %>
		<td align="center"><%= CDbl(cStatistic.FList(i).Fitemsku) %></td>
	<%end if%>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></b></td>
	<%if chkChannel ="1" or groupUserLevel="1" then%>
	<td></td>
	<%end if%>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice - cStatistic.FList(i).FupcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice) %></td>
	<td  align="center" <%=chkIIF(chkChannel ="1","rowspan=""7""","")%> <%=chkIIF(groupUserLevel ="1","rowspan=""10""","")%>>
		<a href="/admin/maechul/statistic/statistic_item_dw.asp?menupos=1726&date_gijun=<%=vDateGijun%>&startdate=<%= vstartdate %>&enddate=<%= venddate %>&ebrand=<%= cStatistic.FList(i).FMakerID %>" target="_blank">[상품]</a>
		&nbsp;
		<a href="/admin/dataanalysis/chart/sellbybrand.asp?ordtype=S&startdate=<%= vstartdate %>&enddate=<%= venddate %>&pvalue=<%= cStatistic.FList(i).FMakerID %>" target="_blank">[추세]</a>
	</td>

	
</tr>
<%if chkChannel ="1" then%>
<tr bgcolor="#e3f1fb" align="Center">
    <td>www</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fwww_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fwww_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fwww_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemCost-cStatistic.FList(i).Fwww_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Fwww_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ReducedPrice-cStatistic.FList(i).Fwww_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fwww_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ReducedPrice - cStatistic.FList(i).Fwww_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_overValueStockPrice) %></td>
</tr>
<% if (FALSE) then %>
<tr bgcolor="#e3f1fb" align="Center">
    <td >모바일/App</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fma_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fma_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fma_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ItemCost-cStatistic.FList(i).Fma_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fma_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fma_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fma_ReducedPrice-cStatistic.FList(i).Fma_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Fma_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ReducedPrice - cStatistic.FList(i).Fma_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_overValueStockPrice) %></td>
</tr>
<% end if %>
<tr bgcolor="#e3f1fb" align="Center">
    <td >MOB</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fm_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fm_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fm_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ItemCost-cStatistic.FList(i).Fm_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fm_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fm_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fm_ReducedPrice-cStatistic.FList(i).Fm_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Fm_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ReducedPrice - cStatistic.FList(i).Fm_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td >MOB_제휴</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fmk_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fmk_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fmk_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ItemCost-cStatistic.FList(i).Fmk_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fmk_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ReducedPrice-cStatistic.FList(i).Fmk_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Fmk_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ReducedPrice - cStatistic.FList(i).Fmk_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td >App</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fa_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fa_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fa_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ItemCost-cStatistic.FList(i).Fa_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fa_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fa_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fa_ReducedPrice-cStatistic.FList(i).Fa_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Fa_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ReducedPrice - cStatistic.FList(i).Fa_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td >제휴</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fo_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fo_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fo_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ItemCost-cStatistic.FList(i).Fo_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fo_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fo_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fo_ReducedPrice-cStatistic.FList(i).Fo_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Fo_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ReducedPrice - cStatistic.FList(i).Fo_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td >해외몰</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Ff_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Ff_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Ff_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ItemCost-cStatistic.FList(i).Ff_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Ff_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Ff_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Ff_ReducedPrice-cStatistic.FList(i).Ff_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Ff_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ReducedPrice - cStatistic.FList(i).Ff_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_overValueStockPrice) %></td>
</tr>
<%end if%>
<% if groupUserLevel ="1" then%>
<tr bgcolor="#e3f1fb" align="Center">
    <td>WHITE</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv0_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Flv0_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv0_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ItemCost-cStatistic.FList(i).Flv0_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Flv0_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ReducedPrice-cStatistic.FList(i).Flv0_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Flv0_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ReducedPrice - cStatistic.FList(i).Flv0_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td>RED</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv1_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Flv1_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv1_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ItemCost-cStatistic.FList(i).Flv1_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Flv1_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ReducedPrice-cStatistic.FList(i).Flv1_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Flv1_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ReducedPrice - cStatistic.FList(i).Flv1_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td>VIP</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv2_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Flv2_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv2_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ItemCost-cStatistic.FList(i).Flv2_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Flv2_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ReducedPrice-cStatistic.FList(i).Flv2_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Flv2_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ReducedPrice - cStatistic.FList(i).Flv2_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td>VIP GOLD</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv3_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Flv3_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv3_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ItemCost-cStatistic.FList(i).Flv3_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Flv3_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ReducedPrice-cStatistic.FList(i).Flv3_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Flv3_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ReducedPrice - cStatistic.FList(i).Flv3_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td>VVIP</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv4_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Flv4_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv4_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ItemCost-cStatistic.FList(i).Flv4_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Flv4_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ReducedPrice-cStatistic.FList(i).Flv4_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Flv4_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ReducedPrice - cStatistic.FList(i).Flv4_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td>STAFF</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv7_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Flv7_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv7_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ItemCost-cStatistic.FList(i).Flv7_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Flv7_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ReducedPrice-cStatistic.FList(i).Flv7_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Flv7_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ReducedPrice - cStatistic.FList(i).Flv7_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td>FAMILY</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv8_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Flv8_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv8_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ItemCost-cStatistic.FList(i).Flv8_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Flv8_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ReducedPrice-cStatistic.FList(i).Flv8_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Flv8_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ReducedPrice - cStatistic.FList(i).Flv8_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td>BIZ</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv9_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Flv9_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv9_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_ItemCost-cStatistic.FList(i).Flv9_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Flv9_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv9_ReducedPrice-cStatistic.FList(i).Flv9_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Flv9_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_ReducedPrice - cStatistic.FList(i).Flv9_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Flv9_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#e3f1fb" align="Center">
    <td>비회원</td>
	<td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fnomem_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fnomem_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fnomem_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ItemCost-cStatistic.FList(i).Fnomem_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Fnomem_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ReducedPrice-cStatistic.FList(i).Fnomem_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fnomem_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ReducedPrice - cStatistic.FList(i).Fnomem_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_overValueStockPrice) %></td>
</tr>
<% end if %>
<%
	vTot_ItemNO						= vTot_ItemNO + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemNO))
	vTot_itemsku					= vTot_itemsku + CDbl(NullOrCurrFormat(cStatistic.FList(i).fitemsku))
	vTot_OrgitemCost				= vTot_OrgitemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost))
	vTot_ItemcostCouponNotApplied	= vTot_ItemcostCouponNotApplied + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied))
	vTot_ItemCost					= vTot_ItemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost))
	vTot_BonusCouponPrice			= vTot_BonusCouponPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice))
	vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice))
	vTot_BuyCash					= vTot_BuyCash + CDbl(NullOrCurrFormat(cStatistic.FList(i).FBuyCash))
	vTot_MaechulProfit				= vTot_MaechulProfit + CDbl(NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit))
	vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash))
	vTot_upcheJungsan				= vTot_upcheJungsan + CDbl(NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan))
	vTot_avgipgoPrice				= vTot_avgipgoPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice))
	vTot_overValueStockPrice		= vTot_overValueStockPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice))

Next

	vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
	vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<% if chkChannel="1" or groupUserLevel="1" then %>
		<%
		' 구매유형:PB일경우
		if chkChannel="1" and vPurchasetype = "3" then
		%>
			<% totalcolspan=4 %>
		<% else %>
			<% totalcolspan=3 %>
		<% end if %>
	<% else %>
		<% totalcolspan=2 %>
	<% end if %>
	<td align="center" colspan="<%= totalcolspan %>">총계</td>
	<td align="center"><%=vTot_ItemNO%></td>
	<%if chkChannel ="1" then%>
	<%elseif groupUserLevel="1" then%>
	<% else %>
		<td align="center"><%= vTot_itemsku %></td>
	<%end if%>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_OrgitemCost)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ItemcostCouponNotApplied)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ItemCost)%></b></td>
	<%if chkChannel="1" or groupUserLevel="1" then%><td></td><%end if%>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BonusCouponPrice)%></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ReducedPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BuyCash)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulProfit2)%></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_upcheJungsan)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ReducedPrice - vTot_upcheJungsan)%></b></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_avgipgoPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_overValueStockPrice)%></td>
	<td></td>
</tr>
<tr>
	<td align="center" colspan="30" bgcolor="#FFFFFF" height="30">
	  <%sbDisplayPaging "page", page, cStatistic.FTotalCount, pagesize, 10,menupos %>
	 </td>
</tr>
</table>
<iframe id="hidifr" src="" width="0" height="0" frameborder="0"></iframe>
<% Set cStatistic = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
