<%@ language=vbscript %>
<% option explicit

	'스크립트 타임아웃 시간 조정 (기본 90초)
	''Server.ScriptTimeout = 180
%>
<%
'####################################################
' Description :  상품별 매출통계
' History : 2016.10.10 정윤정 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_analisys_diary.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vSorting
dim sellchnl, inc3pl, vCateL, vCateM, vCateS, vIsBanPum, vPurchasetype, v6Ago, mwdiv
dim iCurrPage,iPageSize,iTotalPage,iTotCnt, dispCate,vBrandID, chkImg ,itemid, sVType
dim  vTotwww_ItemNO,vTotwww_ItemCost,vTotwww_MaechulProfit,vTotwww_BuyCash,vTotma_ItemNO,vTotma_ItemCost,vTotma_MaechulProfit
dim vTotma_BuyCash,vTotout_ItemNO,vTotout_ItemCost,vTotout_MaechulProfit	,vTotout_BuyCash			
dim vTotwww_MaechulProfitPer ,vTotma_MaechulProfitPer ,vTotout_MaechulProfitPer 
Dim vTot_OrderCnt, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit
Dim vTot_MaechulProfitPer, vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
dim vTot_upcheJungsan, vTot_avgipgoPrice, vTot_overValueStockPrice
dim chkcate
dim dy ,diaryyear

	iPageSize = 100
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vSorting	= NullFillWith(request("sorting"),"itemcost")
	vBrandID	= NullFillWith(request("ebrand"),"")
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	dispCate = requestCheckvar(request("disp"),16)
	itemid      = requestCheckvar(request("itemid"),255)
	chkImg		= requestCheckvar(request("chkImg"),1)
	chkcate		= requestCheckvar(request("chkcate"),1)
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(request("mwdiv"),"")
	inc3pl = request("inc3pl")
	iCurrPage =requestCheckVar(request("iC"),4)
	sVType      = requestCheckvar(request("rdoVType"),1)
	diaryyear = requestCheckvar(request("selDDy"),4)  
  if diaryyear ="" then
  	diaryyear = year(dateadd("yyyy",1,now()))
  end if
 
if iCurrPage = "" then iCurrPage = 1
if chkImg ="" then chkImg = 0	
	if chkcate ="" then chkcate = 0	
if sVType ="" then sVType = 1

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
	'cStatistic.FRect6MonthAgo = v6Ago 
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectMakerid = vBrandID
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 추가
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectItemid   = itemid 
	cStatistic.FRectVType = sVType
	cStatistic.FPageSize = iPageSize
	cStatistic.FCurrPage = iCurrPage
	cStatistic.FRectIncStockAvgPrc = true '' 평균매입가 포함 쿼리여부. ''잠시 뺌 eastone
 cStatistic.FRectdiaryyear = diaryyear

	if sVType=3 then
	    cStatistic.fStatistic_item_channel()
    else    
		cStatistic.fStatistic_item()
    end if

	iTotCnt = cStatistic.FResultCount
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

 
function searchSubmit()
{
    document.frm.target = "_self"; 
    document.frm.action = "statistic_item_analisys.asp";  
	
	if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
		//if (MonthDiff(frm.syear.value + "-" + frm.smonth.value + "-" + frm.sday.value, frm.eyear.value + "-" + frm.emonth.value + "-" + frm.eday.value) >= 1) {
		//	alert("최대 1개월까지만 검색이 가능합니다.");
		//	return;
		//}

		$("#btnSubmit").prop("disabled", true);
		frm.submit(); 
	}

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

 function jsexceldown(){ 
  
    var icurrpage = $('#selDCnt').val();
    document.frm.target =  "XLdown"; 
    document.frm.iC.value =icurrpage;
    document.frm.action = "statistic_item_analisys_xls.asp";  
	document.frm.submit(); 
	
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>"> 
<input type="hidden" name="iC" value=""> 
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"  >
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
	<td align="left">
		<table class="a" cellpadding="3" border="0" cellspacing="0" width=1200>
		<tr>
			<td height="25" colspan="4">
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
					'Response.Write "<input type=""checkbox"" name=""is6ago"" value=""o"" "
					'If v6Ago = "o" Then
					'	Response.Write "checked"
					'End If
					'Response.Write ">6개월이전데이터"
				%>
				&nbsp;&nbsp;다이어리스토리 기준: <select name="selDDy">
					<%For dy=year(dateadd("yyyy",1,now())) To 2012 step -1 %>
					<option value="<%=dy%>" <%if Cint(dy)=Cint(diaryyear) then%>selected<%end if%>><%=dy%></option>
					<%Next%>
				</select> 
			</td>
		</tr>
		<tr>
			<td colspan="4">
				<!-- #include virtual="/common/module/categoryselectbox.asp"-->
				&nbsp;&nbsp;전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		</td>
	</tr> 
	<tr>
		<td colspan="4">
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
				<% Call DrawBrandMWUPCombo("mwdiv",mwdiv) %>
		</td>
	</tr>
	<tr>
		<td width="300"> 브랜드 : <input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');"></td>
		     
		<td align="right">상품코드 :</td>
		<td rowspan="2" align="left" width="800"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea></td>
		 
   </tr> 
	<tr>
		<td colspan="4">
				정렬: <input type="radio" name="sorting" value="itemno" <%=CHKIIF(vSorting="itemno","checked","")%>>수량순
				<input type="radio" name="sorting" value="itemcost" <%=CHKIIF(vSorting="itemcost","checked","")%>>매출순
				<input type="radio" name="sorting" value="profit" <%=CHKIIF(vSorting="profit","checked","")%>>수익순 
				
		</td>
	</tr>
	<tr>		
		<td colspan="4"> 리스트타입:
		    <input type="radio" name="rdoVType" value="1" <%=CHKIIF(sVType="1","checked","")%>>상품별 
		    <input type="radio" name="rdoVType" value="2" <%=CHKIIF(sVType="2","checked","")%>>날짜별 
		    <input type="radio" name="rdoVType" value="3" <%=CHKIIF(sVType="3","checked","")%>>판매채널별 
		    
		    &nbsp;&nbsp;
		    <input type="checkbox" name="chkImg" value="1" <%if chkImg = 1 then%>checked<%end if%>>상품이미지 보기
		    &nbsp; 
		    <input type="checkbox" name="chkcate" value="1" <%if chkcate = 1 then%>checked<%end if%>>카테고리 보기
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
<div style="width:100%;text-align:right;">
엑셀다운:
	<% dim iDownCnt, imaxDCnt, iminDCnt 
 	%> 
	<select name="selDCnt" id="selDCnt" class="select">
	    <option value="0">--페이지 선택--</option>
	    <%
	    if iTotCnt >0 then
	        iDownCnt =  Int(iTotCnt/5000)+1 
	        imaxDCnt = 0
	    for i=1 to iDownCnt 
	        iminDCnt = imaxDCnt + 1
	        if iDownCnt = 1 then
	            imaxDCnt = iTotCnt
	        else    
	            imaxDCnt = 5000*i
	        end if    
	    %>
	    <option value="<%=i%>"><%=iminDCnt%>~<%=imaxDCnt%></option>
	    <%next%>
	    <%end if%> 
	</select>
    <a href="javascript:jsexceldown();"><image src="/images/btn_excel.gif" border="0" align="absmiddle"></a> 
</div>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" >
	<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="25">
				검색결과 : <b><%=iTotCnt%></b>
				&nbsp;
				페이지 : <b><%= iCurrPage %> / <%=iTotalPage%></b>
			</td> 
		</tr>
		 <%IF sVType = 3  then%>
			 <tr bgcolor="<%= adminColor("tabletop") %>"> 
		        <td align="center" rowspan="2">날짜</td>
		        <td align="center" rowspan="2">상품코드</td>
		        <td align="center"  rowspan="2">상품명</td>
    			<%IF chkImg = 1 then%>
    			<td align="center" rowspan="2">이미지</td>
    			<%END IF%>	
    			<%IF chkcate =1 then%>
    			<td align="center" rowspan="2">카테고리</td>		
    			<%END IF%>
    			<td align="center" rowspan="2">브랜드</td> 
    			<td align="center" colspan="4">TOTAL</td>
    			<td align="center" colspan="4">WEB</td>
    			<td align="center" colspan="4">MOB+APP</td>
    			<td align="center" colspan="4">OUT</td>
    			<!--<td align="center" rowspan="2">위시수</td>-->
				<td align="center" rowspan="2">업체<br>정산액</td>
				<td align="center" rowspan="2"><b>회계매출</b></td>
				<td align="center" rowspan="2">평균<br>매입가</td>
				<td align="center" rowspan="2">재고<br>충당금</td>
	    	</tr>
	    	<tr bgcolor="<%= adminColor("tabletop") %>">
    	        <td align="center">상품수량</td>
    	        <td align="center">구매총액</td>
    	        <td align="center">매출수익</td>
    	        <td align="center">수익율</td>
    	        <td align="center">상품수량</td>
    	        <td align="center">구매총액</td> 
    	        <td align="center">매출수익</td>
    	        <td align="center">수익율</td>
    	        <td align="center">상품수량</td>
    	        <td align="center">구매총액</td> 
    	        <td align="center">매출수익</td>
    	        <td align="center">수익율</td>
    	        <td align="center">상품수량</td>
    	        <td align="center">구매총액</td> 
    	        <td align="center">매출수익</td> 
    	        <td align="center">수익율</td>
	    	</tr>		
    	<%else%>		
			<tr bgcolor="<%= adminColor("tabletop") %>">  
    		    <%IF sVType = 2  then%>
    			<td align="center">날짜</td>
    			<%END IF%>			
    			<td align="center">상품코드</td>
    			<td align="center">상품명</td>
    			<%IF chkImg = 1 then%>
    			<td align="center">이미지</td>
    			<%END IF%>			
    			<%IF chkcate =1 THEN%>
    			<td align="center">카테고리</td>
    			<%END IF%>
    			<td align="center">브랜드</td>
    		    <td align="center">상품수량</td>
    		    <% if (NOT C_InspectorUser) then %>
    		    <td align="center">소비자가[상품]</td>
    		    <td align="center">판매가[상품]<br>(할인적용)</td>
    		    <td align="center"><b>구매총액[상품]<br>(상품쿠폰적용)</b></td>
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
			</tr> 
        <%end if%>    
		  <%IF sVType = 3  then%>
		  <%
			For i = 0 To cStatistic.FTotalCount -1
		%>
		 <tr bgcolor="#FFFFFF"> 
		    <td align="center"><%= cStatistic.FList(i).Fddate %></td>
		    <td align="center"><a href="<%=vwwwUrl%>/shopping/category_prd.asp?itemid=<%= cStatistic.FList(i).FitemID %>" target="_blank"><%= cStatistic.FList(i).FitemID %></a></td>
		    <td align="center"><%= cStatistic.FList(i).FItemName %></td>
			<%IF chkImg = 1 then%>
			<td align="center"><img src="<%= cStatistic.FList(i).FSmallImage %>" width="50" height="50" border="0"></td>
			<%END IF%>
			<%IF chkcate = 1 then%>
			<td align="left">&nbsp;&nbsp;<%=cStatistic.FList(i).FCateFullName%></td> 
			<%END IF%>
			<td align="center"><%=cStatistic.FList(i).FMakerID%></td> 
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemNo) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulProfitper) %>%</td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_itemno) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_itemcost) %></td> 
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_maechulprofit) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_maechulprofitper) %>%</td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_itemno) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_itemcost) %></td> 
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_maechulprofit) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_maechulprofitper) %>%</td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Foutmall_itemno) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Foutmall_itemcost) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Foutmall_maechulprofit) %></td> 
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Foutmall_maechulprofitper) %>%</td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan) %></td>
			<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice - cStatistic.FList(i).FupcheJungsan) %></b></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice) %></td>
		</tr>	
		  <% 
		    vTot_ItemNO						= vTot_ItemNO + CLng(NullOrCurrFormat(cStatistic.FList(i).FItemNO))
		  	vTot_ItemCost					= vTot_ItemCost + CLng(NullOrCurrFormat(cStatistic.FList(i).FItemCost))
		  	vTot_MaechulProfit				= vTot_MaechulProfit + CLng(NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit))
		  	vTot_BuyCash					= vTot_BuyCash + CLng(NullOrCurrFormat(cStatistic.FList(i).FBuyCash))
		  	vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice))
		  	vTotwww_ItemNO					= vTotwww_ItemNO + CLng(NullOrCurrFormat(cStatistic.FList(i).Fwww_itemno))
		  	vTotwww_ItemCost				= vTotwww_ItemCost + CLng(NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemCost))
		  	vTotwww_MaechulProfit			= vTotwww_MaechulProfit + CLng(NullOrCurrFormat(cStatistic.FList(i).Fwww_MaechulProfit))
		  	vTotwww_BuyCash					= vTotwww_BuyCash + CLng(NullOrCurrFormat(cStatistic.FList(i).Fwww_BuyCash))
		  	
		  	vTotma_ItemNO					= vTotma_ItemNO + CLng(NullOrCurrFormat(cStatistic.FList(i).Fma_ItemNO))
		  	vTotma_ItemCost					= vTotma_ItemCost + CLng(NullOrCurrFormat(cStatistic.FList(i).Fma_ItemCost))
		  	vTotma_MaechulProfit			= vTotma_MaechulProfit + CLng(NullOrCurrFormat(cStatistic.FList(i).Fma_MaechulProfit))
		  	vTotma_BuyCash					= vTotma_BuyCash + CLng(NullOrCurrFormat(cStatistic.FList(i).Fma_BuyCash))
		  	
		  	vTotout_ItemNO					= vTotout_ItemNO + CLng(NullOrCurrFormat(cStatistic.FList(i).Foutmall_itemno))
		  	vTotout_ItemCost				= vTotout_ItemCost + CLng(NullOrCurrFormat(cStatistic.FList(i).Foutmall_ItemCost))
		  	vTotout_MaechulProfit			= vTotout_MaechulProfit + CLng(NullOrCurrFormat(cStatistic.FList(i).Foutmall_MaechulProfit))
		  	vTotout_BuyCash					= vTotout_BuyCash + CLng(NullOrCurrFormat(cStatistic.FList(i).Foutmall_BuyCash))
			vTot_upcheJungsan				= vTot_upcheJungsan + CDbl(NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan))
			vTot_avgipgoPrice				= vTot_avgipgoPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice))
			vTot_overValueStockPrice		= vTot_overValueStockPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice))
		  Next 
		     vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
		     vTotwww_MaechulProfitPer = Round(((vTotwww_ItemCost - vTotwww_BuyCash)/CHKIIF(vTotwww_ItemCost=0,1,vTotwww_ItemCost))*100,2)
		     vTotma_MaechulProfitPer = Round(((vTotma_ItemCost - vTotma_BuyCash)/CHKIIF(vTotma_ItemCost=0,1,vTotma_ItemCost))*100,2)
		     vTotout_MaechulProfitPer = Round(((vTotout_ItemCost - vTotout_BuyCash)/CHKIIF(vTotout_ItemCost=0,1,vTotout_ItemCost))*100,2)
		  %>
        <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
            <td colspan="4" align="center">총계</td>
            <td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTot_ItemNO) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTot_ItemCost)%></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTot_MaechulProfit) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTot_MaechulProfitPer) %>%</td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotwww_ItemNO) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotwww_ItemCost) %></td> 
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotwww_MaechulProfit) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotwww_MaechulProfitPer) %>%</td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotma_ItemNO) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotma_ItemCost) %></td> 
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotma_MaechulProfit) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotma_MaechulProfitPer) %>%</td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotout_ItemNO) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotout_ItemCost) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotout_MaechulProfit) %></td> 
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(vTotout_MaechulProfitPer) %>%</td>
			<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_upcheJungsan)%></td>
			<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ReducedPrice - vTot_upcheJungsan)%></b></td>
			<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_avgipgoPrice)%></td>
			<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_overValueStockPrice)%></td>
        </tr>
		  <%
		  ELSE%>
		  <%
			For i = 0 To cStatistic.FTotalCount -1
		%>
		<tr bgcolor="#FFFFFF">
		  
		     <%IF sVType = 2 then%>
			<td align="center"><%= cStatistic.FList(i).Fddate %></td>
			<%END IF%>	
			<td align="center"><a href="<%=vwwwUrl%>/shopping/category_prd.asp?itemid=<%= cStatistic.FList(i).FitemID %>" target="_blank"><%= cStatistic.FList(i).FitemID %></a></td>
			<td align="center"><%= cStatistic.FList(i).FItemName %></td>
			<%IF chkImg = 1 then%>
			<td align="center"><img src="<%= cStatistic.FList(i).FSmallImage %>" width="50" height="50" border="0"></td>
			<%END IF%>
			<%if chkcate = 1 then%>
			<td align="left">&nbsp;&nbsp;<%=cStatistic.FList(i).FCateFullName%></td> 
			<%end if%>
			<td align="center"><%=cStatistic.FList(i).FMakerID%></td>
			<td align="center"><%= CDbl(cStatistic.FList(i).FItemNO) %></td>
			<% if (NOT C_InspectorUser) then %>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost) %></td>
			<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied) %></td>
			<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></b></td>
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
		</tr>
		<%
			vTot_ItemNO						= vTot_ItemNO + CLng(NullOrCurrFormat(cStatistic.FList(i).FItemNO))
			vTot_OrgitemCost				= vTot_OrgitemCost + CLng(NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost))
			vTot_ItemcostCouponNotApplied	= vTot_ItemcostCouponNotApplied + CLng(NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied))
			vTot_ItemCost					= vTot_ItemCost + CLng(NullOrCurrFormat(cStatistic.FList(i).FItemCost))
			vTot_BonusCouponPrice			= vTot_BonusCouponPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice))
			vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice))
			vTot_BuyCash					= vTot_BuyCash + CLng(NullOrCurrFormat(cStatistic.FList(i).FBuyCash))
			vTot_MaechulProfit				= vTot_MaechulProfit + CLng(NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit))
			vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash))
			vTot_upcheJungsan				= vTot_upcheJungsan + CDbl(NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan))
			vTot_avgipgoPrice				= vTot_avgipgoPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice))
			vTot_overValueStockPrice		= vTot_overValueStockPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice))
		Next
		
			vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
			vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
		%>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td align="center" colspan="<%IF chkImg = 1 then%><%if sVType ="2" then%>5<%else%>4<%end if%><%else%><%if sVType ="2" then%>4<%else%>3<%end if%><%end if%>">총계</td>  
					<%if chkcate = 1 then%><td></td><%end if%>
			<td align="center"><%=vTot_ItemNO%></td>
			<% if (NOT C_InspectorUser) then %>
			<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_OrgitemCost)%></td>
			<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ItemcostCouponNotApplied)%></td>
			<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ItemCost)%></b></td>
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
		</tr>
		    <%END IF%>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
	  <%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
	 </td>
</tr> 
</table>
<% Set cStatistic = Nothing %>
<iframe id="XLdown" name="XLdown" src="about:blank" frameborder="0" width="110" height="110"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
