<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  카테고리별매출
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<%
	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vCateL, vCateM, vCateS, vIsBanPum, vBrandID, vCateGubun, vPurchasetype, vParam, vbizsec
	dim sellchnl, mwdiv, categbn
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	vBrandID	= NullFillWith(request("ebrand"),"")
	vbizsec     = NullFillWith(request("bizsec"),"")
	mwdiv       = NullFillWith(request("mwdiv"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	categbn     = NullFillWith(request("categbn"),"")
	
	vCateGubun = "L"
	If vCateL <> "" Then
		vCateGubun = "M"
	End IF
	If vCateM <> "" Then
		vCateGubun = "S"
	End IF
	if (categbn="") then
        categbn="D"
    end if
	vParam = CurrURL() & "?menupos="&Request("menupos")&"&date_gijun="&vDateGijun&"&syear="&vSYear&"&smonth="&vSMonth&"&sday="&vSDay&"&eyear="&vEYear&"&emonth="&vEMonth&"&eday="&vEDay&"&sitename="&vSiteName&"&isBanpum="&vIsBanPum&"&ebrand="&vBrandID&"&purchasetype="&vPurchasetype&"&mwdiv="&mwdiv&"&categbn="&categbn&"&sellchnl="&sellchnl
	'<!-- //-->
	
	Dim vTot_OrderCnt, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer
	Dim vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
	
	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectCateGubun = vCateGubun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectMakerID = vBrandID
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectBizSectionCd = vbizsec
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectSellChannelDiv = sellchnl
 
	if (categbn="M") then
	    cStatistic.fStatistic_category()
	else
    	cStatistic.fStatistic_DispCategory  ''2013/10/17 추가
    end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function searchSubmit()
{
	if(DateCheck() == false)
	{
		return;
	}

	if(frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>)
	{
		alert("6개월전까지만 실시간검색이 가능합니다.");
	}
	else
	{
		frm.submit();
	}
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

function popViewItem(cdl,cdm,cds) {
	<% if categbn="M" then %>
	alert("전시카테고리만 상품 확인이 가능합니다.");
	<% else %>
	var frm = document.frm;
	var param = 'cdl='+cdl+'&cdm='+cdm+'&cds='+cds
	param += '&'+ $(frm).serialize();
	var itemView = window.open('statistic_category_itemList.asp?'+param,'itemView','width=1024,height=768,scrollbars=yes,resizable=yes');
	itemView.focus();
	<% end if %>
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* 기간 :&nbsp;
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>주문일</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>결제일</option>
				</select>
				<%
					'### 년
					Response.Write "<select name=""syear"" class=""select"">"
					For i=Year(now) To Year(v6MonthDate) Step -1
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
					For i=Year(now) To Year(v6MonthDate) Step -1
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
					
					
					'### 사이트구분
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* 사이트구분 : "
					Call Drawsitename("sitename", vSiteName)
					
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* 기본 매출부서 : "
					Call DrawBizSectionGain("O,T","bizsec", vbizsec,"")
				%>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				* 주문구분 :
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>반품포함</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>반품제외</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>반품건만</option>
				</select>
				&nbsp;&nbsp;
				* 매입구분 : 
				<% Call DrawBrandMWUCombo("mwdiv",mwdiv) %>
				<br>
                	* 채널구분
                	<% drawSellChannelComboBox "sellchnl",sellchnl %>
                	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				* 브랜드 : <input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				* 구매유형 : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
				&nbsp;&nbsp;
				<input type="radio" name="categbn" value="D" <%=CHKIIF(categbn="D","checked","")%> >전시카테고리
				<input type="radio" name="categbn" value="M" <%=CHKIIF(categbn="M","checked","")%> >관리카테고리
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
※ 실시간 데이터는 최근 6개월까지 데이터만 검색 가능합니다.
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
	<td><%=CateGubun(vCateGubun)%>카테고리</td>
    <td>상품수량</td>
    <% if (NOT C_InspectorUser) then %>
    <td>소비자가[상품]</td>
    <td>판매가[상품]<br>(할인적용)</td>
    <td><b>구매총액[상품]<br>(상품쿠폰적용)</b></td>
    <td><b>보너스쿠폰<br>사용액[상품]</b></td>
    <% end if %>
    <td>취급액</td>
    <td>매입총액[상품]<% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %></td>
    <td><b>매출수익</b></td>
    <td>수익율</td>
    <td>매출수익2<br>(취급액기준)</td>
    <td>수익율</td>
</tr>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<td  style="padding-left:5px;">
		<%= cStatistic.FList(i).FCategoryName %>&nbsp;
		<%
			If vCateGubun = "L" Then
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"""><font color=""gray"">[중]</font></a>"
			ElseIf vCateGubun = "M" Then
				Response.Write "<a href="""&vParam&"""><font color=""gray"">[대]</font></a>"
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"&cdm="&cStatistic.FList(i).FCateM&"""><font color=""gray"">[소]</font></a>"
			ElseIf vCateGubun = "S" Then
				Response.Write "<a href="""&vParam&"""><font color=""gray"">[대]</font></a>"
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"""><font color=""gray"">[중]</font></a>"
				if (categbn="D") then
                Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"&cdm="&cStatistic.FList(i).FCateM&"&cds="&cStatistic.FList(i).FCateS&"""><font color=""gray"">[세]</font></a>"
                end if
			End IF
		%>
	</td>
	<td align="center"><a href="#" onclick="popViewItem('<%=cStatistic.FList(i).FCateL%>','<%=cStatistic.FList(i).FCateM%>','<%=cStatistic.FList(i).FCateS%>')"><%= CDbl(cStatistic.FList(i).FItemNO) %></a></td>
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
</tr>
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
	<td align="center">총계</td>
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
</tr>
</table>
<% Set cStatistic = Nothing

Function CateGubun(g)
	If g = "L" Then
		CateGubun = "대"
	ElseIf vCateGubun = "M" Then
		CateGubun = "중"
	ElseIf vCateGubun = "S" Then
		CateGubun = "소"
	End IF
End Function
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->