<%@ language=vbscript %>
<% option explicit

	'스크립트 타임아웃 시간 조정 (기본 90초)
	Server.ScriptTimeout = 180
%>
<%
'####################################################
' Description :  브랜드별 통계
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_datamart.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<%

	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vSorting, vCateL, vCateM, vCateS, vIsBanPum, vPurchasetype, v6Ago
	dim sellchnl, inc3pl
	Dim mwdiv
	Dim dispCate
	dim chkChannel
	dim page, pagesize


	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"beasongdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vSorting	= NullFillWith(request("sorting"),"itemcost")
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	dispCate = requestCheckvar(request("disp"),16)

	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(request("mwdiv"),"")
	inc3pl = request("inc3pl")
    chkChannel  = requestCheckvar(request("chkChl"),1)
	page  = requestCheckvar(request("page"),10)
	pagesize  = requestCheckvar(request("pagesize"),10)

	if (page = "") then
		page = 1
	end if

	if (pagesize = "") then
		pagesize = "100"
	end if


	Dim vTot_OrderCnt, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer
	Dim vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2

	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectSort = vSorting
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRect6MonthAgo = v6Ago
	'cStatistic.FRectChannelDiv = channelDiv
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 추가
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectChkchannel = chkChannel

	cStatistic.FCurrPage = page
	cStatistic.FPageSize = pagesize

	cStatistic.fStatistic_brand()


	dim iTotalPage
	iTotalPage 	=  int((cStatistic.FTotalCount)/pagesize) +1

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function searchSubmit()
{
	/*
	if(DateCheck() == false)
	{
		return;
	}
	 */

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
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
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
				</select>
				<%
					'### 년
					Response.Write "<select name=""syear"" class=""select"">"
					For i=Year(now) To 2001 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 월
					Response.Write "<select name=""smonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 일
					Response.Write "<select name=""sday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;~&nbsp;"

					'#############################

					'### 년
					Response.Write "<select name=""eyear"" class=""select"">"
					For i=Year(now) To 2001 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 월
					Response.Write "<select name=""emonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 일
					Response.Write "<select name=""eday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>"


					'### 6개월이전데이터check
					Response.Write "<input type=""checkbox"" name=""is6ago"" value=""o"" "
					If v6Ago = "o" Then
						Response.Write "checked"
					End If
					Response.Write ">6개월이전데이터"
				%>

			</td>
		</tr>
		<tr>
			<td>
				<!-- #include virtual="/common/module/categoryselectbox.asp"-->
				&nbsp;&nbsp;전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		</td>
	</tr>
	<tr>
		<td>
			사이트:  <% Call Drawsitename("sitename", vSiteName)%>
			&nbsp;&nbsp;채널:
   			 <% drawSellChannelComboBox "sellchnl",sellchnl %>
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
				<% Call DrawBrandMWUCombo("mwdiv",mwdiv) %>
				&nbsp;&nbsp;
				<input type="checkbox" name="chkChl" value="1" <%if chkChannel ="1" then%>checked<%end if%>>채널상세보기
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
				</select>
			</td>
		</tr>
	    </table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" id="btnSubmit" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
* 검색 기간이 길어지면 상당히 느려집니다. 그러니 검색 버튼을 클릭한 뒤 아무 반응이 없어보인다고 재차 검색버튼을 클릭하지 마세요.
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= cStatistic.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=iTotalPage%></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">브랜드ID</td>
	<%if chkChannel ="1" then%>
	<td align="center">채널</td>
	<%end if%>
    <td align="center">상품수량</td>
    <% if (NOT C_InspectorUser) then %>
    <td align="center">소비자가[상품]</td>
    <td align="center">판매가[상품]<br>(할인적용)</td>
    <td align="center"><b>구매총액[상품]<br>(상품쿠폰적용)</b></td>
     <%if chkChannel ="1" then%>
    <td>채널<br>점유율</td>
    <%end if%>
    <td align="center"><b>보너스쿠폰<br>사용액[상품]</b></td>
    <% end if %>
    <td align="center">취급액</td>
    <td align="center">매입총액[상품]<% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %></td>
    <td align="center"><b>매출수익</b></td>
    <td align="center">수익율</td>
    <td align="center">매출수익2<br>(취급액기준)</td>
    <td align="center">수익율</td>
    <td align="center">비고</td>
</tr>
<%
	For i = 0 To cStatistic.FResultCount -1
%>
<tr bgcolor="#FFFFFF">
	<td align="center" <%if chkChannel ="1" then%>rowspan="3"<%end if%>><%= cStatistic.FList(i).FMakerID %></td>
	<%if chkChannel ="1" then%>
	<td align="center">전체</td>
	<%end if%>
	<td align="center"><%= CDbl(cStatistic.FList(i).FItemNO) %></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></b></td>
	<%if chkChannel ="1" then%>
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
	<td  align="center" <%if chkChannel ="1" then%>rowspan="3"<%end if%>><a href="/admin/maechul/statistic/statistic_item_datamart.asp?menupos=1726&date_gijun=beasongdate&syear=<%=vSYear%>&smonth=<%=vSMonth%>&sday=<%=vSDay%>&eyear=<%=vEYear%>&emonth=<%=vEMonth%>&eday=<%=vEDay%>&ebrand=<%= cStatistic.FList(i).FMakerID %>" target="_blank">[상품상세]</a></td>
</tr>
<%if chkChannel ="1" then%>
<tr bgcolor="#FAECC5" align="Center">
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
</tr>
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
</tr>
<%end if%>
<%
	vTot_ItemNO						= vTot_ItemNO + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemNO))
	vTot_OrgitemCost				= vTot_OrgitemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost))
	vTot_ItemcostCouponNotApplied	= vTot_ItemcostCouponNotApplied + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied))
	vTot_ItemCost					= vTot_ItemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost))
	vTot_BonusCouponPrice			= vTot_BonusCouponPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice))
	vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice))
	vTot_BuyCash					= vTot_BuyCash + CDbl(NullOrCurrFormat(cStatistic.FList(i).FBuyCash))
	vTot_MaechulProfit				= vTot_MaechulProfit + CDbl(NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit))
	vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash))

	Next

	vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
	vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" <%if chkChannel ="1" then%>colspan="2"<%end if%>>총계</td>
	<td align="center"><%=vTot_ItemNO%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_OrgitemCost)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ItemcostCouponNotApplied)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ItemCost)%></b></td>
	<%if chkChannel ="1" then%><td></td><%end if%>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BonusCouponPrice)%></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ReducedPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BuyCash)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulProfit2)%></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
	<td></td>
</tr>
<tr>
	<td align="center" colspan="30" bgcolor="#FFFFFF" height="30">
	  <%sbDisplayPaging "page", page, cStatistic.FTotalCount, pagesize, 10,menupos %>
	 </td>
</tr>
</table>
<% Set cStatistic = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
