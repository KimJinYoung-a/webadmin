<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 온라인 바코드 출력
' Hieditor : 2016.01.20 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'###########################################################
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

<%
	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vIsBanPum, vPurchasetype, v6Ago, vmakerid
	dim sellchnl, inc3pl, vSorting, dispCate, maxDepth
	Dim mwdiv, chkShowGubun,itemid, showsuply
	Dim incStockAvg, isSendGift
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")  ''beasongdate  :출고일=>주문일 2018/05/28  by eastone
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",-13,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	v6Ago		= NullFillWith(request("is6ago"),"")
	dispCate	= requestCheckVar(request("disp"),20)
	maxDepth = "1"
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	vmakerid    = NullFillWith(request("makerid"),"")
	mwdiv       = NullFillWith(request("mwdiv"),"")
	itemid      = requestCheckvar(request("itemid"),255)
	inc3pl      = request("inc3pl")
	chkShowGubun = request("chkShowGubun")
	vSorting	= NullFillWith(request("sorting"),"yyyymmddD")
	showsuply   = requestCheckvar(request("showsuply"),10)
	incStockAvg = requestCheckvar(request("incStockAvg"),10)
	isSendGift	= requestCheckvar(request("isSendGift"),1)

	Dim vTot_countOrder, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer
	Dim vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
	dim vTot_upcheJungsan, vTot_avgipgoPrice, vTot_overValueStockPrice

   if itemid<>"" then
    	dim iA ,arrTemp,arrItemid
    	itemid = replace(itemid,",",chr(10))
      	itemid = replace(itemid,chr(13),"")
    	arrTemp = Split(itemid,chr(10))

    	iA = 0
    	do while iA <= ubound(arrTemp)
    		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
    			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
    		end if
    		iA = iA + 1
    	loop

    	if len(arrItemid)>0 then
    		itemid = left(arrItemid,len(arrItemid)-1)
    	else
    		if Not(isNumeric(itemid)) then
    			itemid = ""
    		end if
    	end if
    end if

	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	''cStatistic.FRect6MonthAgo = v6Ago
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectMakerid = vmakerid
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 추가
	cStatistic.FRectChkShowGubun = chkShowGubun					'// 2015-10-22, skyer9
	cStatistic.FRectIncStockAvgPrc = (incStockAvg<>"") ''true '' 평균매입가 포함 쿼리여부.
	cStatistic.FRectItemid   = itemid  '/2016-03-18 추가
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectSort = vSorting
	cStatistic.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
	cStatistic.FRectIsSendGift = isSendGift
	cStatistic.fStatistic_daily_item()
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function jstrSort(vsorting){
	var tmpSorting = document.getElementById("img"+vsorting)

	if (-1 < tmpSorting.src.indexOf("_alpha")){
		frm.sorting.value= vsorting+"D";
	}else if (-1 < tmpSorting.src.indexOf("_bot")){
		frm.sorting.value= vsorting+"A";
	}else{
		frm.sorting.value= vsorting+"D";
	}
	searchSubmit();
}

function searchSubmit(){
	//if((frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>) && (frm.is6ago.checked == false))
	//{
	//	alert("6개월전의 데이터는 6개월이전데이터를 체크하셔야 가능합니다.");
	//}
	//else
	//{
		if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
			frm.submit();
		}
	//}
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sorting" value="<%= vsorting %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색</td>
	<td align="left">
		<table class="a" border="0">
		<tr>
			<td>
				* 기간 :
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>주문일</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>결제일</option>
					<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>상품출고일</option>
					<option value="jfixeddt" <%=CHKIIF(vDateGijun="jfixeddt","selected","")%>>정산확정일</option>
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
					'Response.Write "<input type=""checkbox"" name=""is6ago"" value=""o"" "
					'If v6Ago = "o" Then
					'	Response.Write "checked"
					'End If
					'Response.Write ">6개월이전데이터"
				%>
				&nbsp;
                * 채널구분 :
                <% drawSellChannelComboBox "sellchnl",sellchnl %>
        	    &nbsp;
				* 주문구분 :
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>반품포함</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>반품제외</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>반품건만</option>
				</select>
			</td>
		</tr>
		<tr>
		    <td>
				* 매입구분 :
				<% Call DrawBrandMWUPCombo("mwdiv",mwdiv) %>
        	    &nbsp;
        	    * 구매유형 : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
        	    &nbsp;
				* 사이트:
				&nbsp;
				* 브랜드 : <% drawSelectBoxDesigner "makerid",vmakerid %>
				&nbsp;
				* 전시카테고리 :
				<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
			</td>
		 </tr>
		 <tr>
			<td>
				* 사이트구분 : <% Call Drawsitename("sitename", vSiteName) %>
				&nbsp;
				* 매출처 :
        	    <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
        	    &nbsp;
		        <label><input type="checkbox" name="chkShowGubun" value="Y" <% if (chkShowGubun = "Y") then %>checked<% end if %> > 채널구분,매입구분 표시</label>
		        <!--&nbsp;* 상품코드 : <textarea rows="3" cols="10" name="itemid" id="itemid"><%'=replace(itemid,",",chr(10))%></textarea>-->
			    &nbsp;
			    <label><input type="checkbox" name="showsuply" value="on" <%= CHKIIF(showsuply="on","checked","") %> >공급가로 표시</label>
			    &nbsp;&nbsp;
			    <label><input type="checkbox" name="incStockAvg" <%=CHKIIF(incStockAvg<>"","checked","")%>>평균매입가포함</label>
				&nbsp;&nbsp;
			    <label><input type="checkbox" name="isSendGift" value="Y" <%=CHKIIF(isSendGift<>"","checked","")%>>선물주문만 보기</label>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
* 검색 기간이 길어지면 상당히 느려집니다. 그러니 검색 버튼을 클릭한 뒤 아무 반응이 없어보인다고 재차 검색버튼을 클릭하지 마세요.<br>
* 배송비 매출 제외
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<% if (chkShowGubun = "Y") then %>
		<td align="center" onClick="jstrSort('beadaldiv'); return false;" style="cursor:hand;">
			채널<br>구분
			<img src="/images/list_lineup<%=CHKIIF(vSorting="beadaldivD","_bot","_top")%><%=CHKIIF(instr(vSorting,"beadaldiv")>0,"_on","")%>.png" id="imgbeadaldiv">
		</td>
		<td align="center" onClick="jstrSort('omwdiv'); return false;" style="cursor:hand;">
			매입<br>구분
			<img src="/images/list_lineup<%=CHKIIF(vSorting="omwdivD","_bot","_top")%><%=CHKIIF(instr(vSorting,"omwdiv")>0,"_on","")%>.png" id="imgomwdiv">
		</td>
	<% end if %>

	<td align="center" colspan="2" onClick="jstrSort('yyyymmdd'); return false;" style="cursor:hand;">
		기간
		<img src="/images/list_lineup<%=CHKIIF(vSorting="yyyymmddD","_bot","_top")%><%=CHKIIF(instr(vSorting,"yyyymmdd")>0,"_on","")%>.png" id="imgyyyymmdd">
	</td>
    <td align="center" onClick="jstrSort('countOrder'); return false;" style="cursor:hand;">
    	주문수
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="countOrderD","_bot","_top")%><%=CHKIIF(instr(vSorting,"countOrder")>0,"_on","")%>.png" id="imgcountOrder">
    </td>
    <td align="center" onClick="jstrSort('itemno'); return false;" style="cursor:hand;">
    	판매수량
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemnoD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemno")>0,"_on","")%>.png" id="imgitemno">
    </td>

    <% if (NOT C_InspectorUser) then %>
	    <td align="center" onClick="jstrSort('orgitemcost'); return false;" style="cursor:hand;">
	    	소비자가[상품]
	    	<img src="/images/list_lineup<%=CHKIIF(vSorting="orgitemcostD","_bot","_top")%><%=CHKIIF(instr(vSorting,"orgitemcost")>0,"_on","")%>.png" id="imgorgitemcost">
	    </td>
	    <td align="center" onClick="jstrSort('itemcostcouponnotapplied'); return false;" style="cursor:hand;">
	    	판매가[상품]<br>(할인적용)
	    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemcostcouponnotappliedD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemcostcouponnotapplied")>0,"_on","")%>.png" id="imgitemcostcouponnotapplied">
	    </td>
	    <td align="center" onClick="jstrSort('itemcost1'); return false;" style="cursor:hand;">
	    	구매총액[상품]<br>(상품쿠폰적용)
	    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemcost1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemcost1")>0,"_on","")%>.png" id="imgitemcost1">
	    </td>
	    <td align="center" onClick="jstrSort('itemcostnotreducedprice'); return false;" style="cursor:hand;">
	    	보너스쿠폰<br>사용액[상품]
	    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemcostnotreducedpriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemcostnotreducedprice")>0,"_on","")%>.png" id="imgitemcostnotreducedprice">
	    </td>
    <% end if %>

    <td align="center" onClick="jstrSort('reducedPrice'); return false;" style="cursor:hand;">
    	취급액
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="reducedPriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"reducedPrice")>0,"_on","")%>.png" id="imgreducedPrice">
    </td>
	<td align="center" onClick="jstrSort('upchejungsan1'); return false;" style="cursor:hand;">
		업체<br>정산액
		<img src="/images/list_lineup<%=CHKIIF(vSorting="upchejungsan1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"upchejungsan1")>0,"_on","")%>.png" id="imgupchejungsan1">
	</td>
	<td align="center" onClick="jstrSort('reducedpricenotupchejungsan'); return false;" style="cursor:hand;">
		<b>회계매출</b>
		<img src="/images/list_lineup<%=CHKIIF(vSorting="reducedpricenotupchejungsanD","_bot","_top")%><%=CHKIIF(instr(vSorting,"reducedpricenotupchejungsan")>0,"_on","")%>.png" id="imgreducedpricenotupchejungsan">
	</td>
	<td align="center" onClick="jstrSort('avgipgoprice'); return false;" style="cursor:hand;">
		평균<br>매입가
		<img src="/images/list_lineup<%=CHKIIF(vSorting="avgipgopriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"avgipgoprice")>0,"_on","")%>.png" id="imgavgipgoprice">
	</td>
	<td align="center" onClick="jstrSort('overvaluestockprice'); return false;" style="cursor:hand;">
		재고<br>충당금
		<img src="/images/list_lineup<%=CHKIIF(vSorting="overvaluestockpriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"overvaluestockprice")>0,"_on","")%>.png" id="imgovervaluestockprice">
	</td>
    <td align="center" onClick="jstrSort('buycash'); return false;" style="cursor:hand;">
    	매입총액[상품]<% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %>
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="buycashD","_bot","_top")%><%=CHKIIF(instr(vSorting,"buycash")>0,"_on","")%>.png" id="imgbuycash">
    </td>
    <td align="center" onClick="jstrSort('maechulprofit1'); return false;" style="cursor:hand;">
    	<b>매출수익</b>
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofit1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofit1")>0,"_on","")%>.png" id="imgmaechulprofit1">
    </td>
    <td align="center">수익율</td>
    <td align="center" onClick="jstrSort('maechulprofit2'); return false;" style="cursor:hand;">
    	매출수익2<br>(취급액기준)
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofit2D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofit2")>0,"_on","")%>.png" id="imgmaechulprofit2">
    </td>
    <td align="center">수익율</td>
    <td align="center">비고</td>
</tr>
<% if cStatistic.FTotalCount > 0 then %>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<% if (chkShowGubun = "Y") then %>
	<td align="center"><%= getSellChannelName(cStatistic.flist(i).Fbeadaldiv) %></td>
	<td align="center"><%= cStatistic.flist(i).Fomwdiv %></td>
	<% end if %>
	<td align="center">
		<% if right(FormatDateTime(cStatistic.flist(i).FRegdate,1),3) = "토요일" then %>
			<font color="blue"><%= cStatistic.flist(i).FRegdate %></font>
		<% elseif right(FormatDateTime(cStatistic.flist(i).FRegdate,1),3) = "일요일" then %>
			<font color="red"><%= cStatistic.flist(i).FRegdate %></font>
		<% else %>
			<%= cStatistic.flist(i).FRegdate %>
		<% end if %>
	</td>
	<td align="center"><%= DateToWeekName(DatePart("w",cStatistic.FList(i).FRegdate)) %></td>
	<td align="center"><%= CDbl(cStatistic.FList(i).FcountOrder) %></td>
	<td align="center"><%= CDbl(cStatistic.FList(i).FItemNO) %></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice) %></td>

	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice - cStatistic.FList(i).FupcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer2 %>%</td>
	<td align="center" >[<a href="/admin/upchejungsan/upcheselllist.asp?datetype=jumunil&yyyy1=<%=Year(cStatistic.FList(i).FRegdate)%>&mm1=<%=TwoNumber(Month(cStatistic.FList(i).FRegdate))%>&dd1=<%=TwoNumber(Day(cStatistic.FList(i).FRegdate))%>&yyyy2=<%=Year(cStatistic.FList(i).FRegdate)%>&mm2=<%=TwoNumber(Month(cStatistic.FList(i).FRegdate))%>&dd2=<%=TwoNumber(Day(cStatistic.FList(i).FRegdate))%>&disp=<%=dispCate%>&delivertype=all&inc3pl=<%= inc3pl %>&isSendGift=<%=isSendGift%>" target="_blank">상세</a>]</td>
</tr>
<%
	vTot_countOrder					= vTot_countOrder + CLng(NullOrCurrFormat(cStatistic.FList(i).FcountOrder))
	vTot_ItemNO						= vTot_ItemNO + CLng(NullOrCurrFormat(cStatistic.FList(i).FItemNO))
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
	<% if (chkShowGubun = "Y") then %>
	<td colspan="2"></td>
	<% end if %>
	<td align="center" colspan="2">총계</td>
	<td align="center"><%=vTot_countOrder%></td>
	<td align="center"><%=vTot_ItemNO%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_OrgitemCost)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ItemcostCouponNotApplied)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ItemCost)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BonusCouponPrice)%></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ReducedPrice)%></td>

	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_upcheJungsan)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ReducedPrice - vTot_upcheJungsan)%></b></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_avgipgoPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_overValueStockPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BuyCash)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulProfit2)%></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
	<td></td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="26">등록된 내용이 없습니다.</td>
</tr>
<%END IF%>
</table>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
