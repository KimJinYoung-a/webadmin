<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 매출집계-브랜드별매출
' History : 2016.09.21 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/maechul/statisticCls.asp" -->

<%
dim menupos : menupos = getNumeric(requestcheckvar(request("menupos"),16))
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vSorting,vBrandID
dim sellchnl, vCateL, vCateM, vCateS, vIsBanPum, v6Ago, mwdiv, dispCate, page, pagesize, lec_cdl, lec_cdm
Dim vTot_OrderCnt, vTot_ItemNO, vTot_couponNotAsigncost, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer
Dim vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2, vTot_upcheJungsan
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= RequestCheckvar(request("sitename"),16)
	vDateGijun	= NullFillWith(RequestCheckvar(request("date_gijun"),16),"regdate")
	vSYear		= NullFillWith(RequestCheckvar(request("syear"),4),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(RequestCheckvar(request("smonth"),2),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(RequestCheckvar(request("sday"),2),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(RequestCheckvar(request("eyear"),4),Year(now))
	vEMonth		= NullFillWith(RequestCheckvar(request("emonth"),2),Month(now))
	vEDay		= NullFillWith(RequestCheckvar(request("eday"),2),Day(now))
	vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"itemcostD")
	vCateL		= NullFillWith(RequestCheckvar(request("cdl"),3),"")
	vCateM		= NullFillWith(RequestCheckvar(request("cdm"),3),"")
	vCateS		= NullFillWith(RequestCheckvar(request("cds"),3),"")
	vBrandID	= NullFillWith(RequestCheckvar(request("ebrand"),32),"")
	dispCate = requestCheckvar(request("disp"),16)
	vIsBanPum	= NullFillWith(RequestCheckvar(request("isBanpum"),16),"all")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(RequestCheckvar(request("mwdiv"),1),"")
	page  = requestCheckvar(request("page"),10)
	pagesize  = requestCheckvar(request("pagesize"),10)
	lec_cdl = RequestCheckvar(request("lec_cdl"),3)
	lec_cdm = RequestCheckvar(request("lec_cdm"),3)

if (page = "") then
	page = 1
end if

if (pagesize = "") then
	pagesize = 5000
end if

Set cStatistic = New cacademyStatic_list
	cStatistic.FRectlec_cdl = lec_cdl
	cStatistic.FRectlec_cdm = lec_cdm
	cStatistic.FRectSort = vSorting
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectMakerid = vBrandID
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectDispCate = dispCate
	cStatistic.FCurrPage = page
	cStatistic.FPageSize = pagesize
	cStatistic.fStatistic_brand()

dim iTotalPage
	iTotalPage 	=  int((cStatistic.FTotalCount)/pagesize) +1

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>

<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= cStatistic.FTotalCount %></b>
		※ 최대 5천건 까지 보여 집니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>
		브랜드ID
	</td>
    <td>
    	상품수량
    </td>

    <% if (NOT C_InspectorUser) then %>
	    <td>
	    	판매가[상품]<br>(할인적용)
	    </td>
	    <td>
	    	<b>구매총액[상품]<br>(상품쿠폰적용)</b>
	    </td>
	    <td>
	    	<b>보너스쿠폰<br>사용액[상품]</b>
	    </td>
    <% end if %>

    <td>
    	취급액
    </td>
    <td>
    	매입총액[상품]<% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %>
    </td>
    <td>
    	<b>매출수익</b>
    </td>
    <td>
    	수익율1
    </td>
    <td>
    	매출수익2<br>(취급액기준)
    </td>
    <td>
    	수익율2
    </td>
	<td>
		업체<br>정산액
	</td>
	<td>
		<b>회계매출</b>
	</td>
    <td>
    	비고
    </td>
</tr>

<% if cStatistic.FResultCount > 0 then %>
	<% For i = 0 To cStatistic.FResultCount -1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= cStatistic.FItemList(i).FMakerID %></td>
		<td align="center"><%= CDbl(cStatistic.FItemList(i).FItemNO) %></td>
	
		<% if (NOT C_InspectorUser) then %>
			<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcouponNotAsigncost,0) %></td>
			<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= FormatNumber(cStatistic.FItemList(i).FItemCost,0) %></b></td>
			<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FItemCost-cStatistic.FItemList(i).FReducedPrice,0) %></td>
	    <% end if %>
	
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FReducedPrice,0) %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FBuyCash,0) %></td>
		<td align="right" style="padding-right:5px;"><b><%= FormatNumber(cStatistic.FItemList(i).FMaechulProfit,0) %></b></td>
		<td align="right" style="padding-right:5px;"><%= cStatistic.FItemList(i).FMaechulProfitPer %>%</td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FReducedPrice-cStatistic.FItemList(i).FBuyCash,0) %></td>
		<td align="right" style="padding-right:5px;"><%= cStatistic.FItemList(i).FMaechulProfitPer2 %>%</td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FupcheJungsan,0) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= FormatNumber(cStatistic.FItemList(i).FReducedPrice - cStatistic.FItemList(i).FupcheJungsan,0) %></b></td>
		<td  align="center"></td>
	</tr>

	<%
	vTot_ItemNO						= vTot_ItemNO + CLng(FormatNumber(cStatistic.FItemList(i).FItemNO,0))
	vTot_couponNotAsigncost	= vTot_couponNotAsigncost + CLng(FormatNumber(cStatistic.FItemList(i).FcouponNotAsigncost,0))
	vTot_ItemCost					= vTot_ItemCost + CLng(FormatNumber(cStatistic.FItemList(i).FItemCost,0))
	vTot_BonusCouponPrice			= vTot_BonusCouponPrice + CDbl(FormatNumber(cStatistic.FItemList(i).FItemCost-cStatistic.FItemList(i).FReducedPrice,0))
	vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(FormatNumber(cStatistic.FItemList(i).FReducedPrice,0))
	vTot_BuyCash					= vTot_BuyCash + CLng(FormatNumber(cStatistic.FItemList(i).FBuyCash,0))
	vTot_MaechulProfit				= vTot_MaechulProfit + CLng(FormatNumber(cStatistic.FItemList(i).FMaechulProfit,0))
	vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(FormatNumber(cStatistic.FItemList(i).FReducedPrice-cStatistic.FItemList(i).FBuyCash,0))
	vTot_upcheJungsan				= vTot_upcheJungsan + CDbl(FormatNumber(cStatistic.FItemList(i).FupcheJungsan,0))
	%>
	<% Next %>
	<%
	vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
	vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
	%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td align="center">총계</td>
		<td align="center"><%=vTot_ItemNO%></td>
	
		<% if (NOT C_InspectorUser) then %>
			<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_couponNotAsigncost,0)%></td>
			<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_ItemCost,0)%></b></td>
			<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_BonusCouponPrice,0)%></td>
	    <% end if %>
	
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_ReducedPrice,0)%></td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_BuyCash,0)%></td>
		<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_MaechulProfit,0)%></b></td>
		<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_MaechulProfit2,0)%></td>
		<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_upcheJungsan,0)%></td>
		<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_ReducedPrice - vTot_upcheJungsan,0)%></b></td>
		<td></td>
	</tr>

<% ELSE %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="25">등록된 내용이 없습니다.</td>
	</tr>
<% end if %>

</table>

<iframe id="view" name="view" src="" width=1000 height=300 frameborder="0" scrolling="no"></iframe>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->