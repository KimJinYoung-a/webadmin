<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드별매출 엑셀다운로드
' Hieditor : 2020.01.15 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_dw.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSorting, chkChannel,vBrandID, rdsite    ' , vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay
dim sellchnl, inc3pl, vCateL, vCateM, vCateS, vIsBanPum, vPurchasetype, v6Ago, mwdiv, dispCate, page, pagesize
Dim vTot_OrderCnt, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer, vTot_itemsku
Dim vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2, vTot_upcheJungsan, vTot_avgipgoPrice, vTot_overValueStockPrice, vstartdate, venddate
Dim incStockAvg, groupUserLevel, isSendGift
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
	cStatistic.FPageSize = 2000
	cStatistic.FRectIncStockAvgPrc = (incStockAvg<>"") ''true '' 평균매입가 포함 쿼리여부.
	cStatistic.FRectGroupUserLevel = groupUserLevel
    cStatistic.FRectIsSendGift = isSendGift
	cStatistic.fStatistic_brand()

dim iTotalPage
	iTotalPage 	=  int((cStatistic.FTotalCount)/pagesize) +1

%>
<%
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_Brand_DW" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="<%= adminColor("tabletop") %>" align="center">
        <td>브랜드ID</td>
        <td>구매유형</td>
        <%if chkChannel ="1" then%>
        <td>채널</td>
        <%elseif groupUserLevel="1" then%>
        <td>회원등급</td>
        <%end if%>
        <td>상품수량</td>
	<%if chkChannel ="1" then%>
	<%elseif groupUserLevel="1" then%>
	<% else %>
		<td align="center">상품SKU</td>
	<%end if%>
        <% if (NOT C_InspectorUser) then %>
        <td>소비자가[상품]</td>
        <td>판매가[상품]<br>(할인적용)</td>
        <td><b>구매총액[상품]<br>(상품쿠폰적용)</b></td>
        <%if chkChannel ="1" then%>
        <td>채널<br>점유율</td>
        <%elseif groupUserLevel="1" then%>
        <td>등급<br>점유율</td>
        <%end if%>
        <td><b>보너스쿠폰<br>사용액[상품]</b></td>
        <% end if %>
        <td>취급액</td>
        <td>매입총액[상품]<% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %></td>
        <td><b>매출수익</b></td>
        <td>수익율</td>
        <td>매출수익2<br>(취급액기준)</td>
        <td>수익율</td>
        <td>업체<br>정산액</td>
        <td><b>회계매출</b></td>
        <td>평균<br>매입가</td>
        <td>재고<br>충당금</td>
    </tr>
<% if cStatistic.FresultCount>0 then %>
	<% For i = 0 To cStatistic.FResultCount -1 %>
    <tr bgcolor="#FFFFFF">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
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
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></b></td>
        <%if chkChannel ="1" or groupUserLevel="1" then%>
        <td></td>
        <%end if%>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FBuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit) %></b></td>
        <td align="right"><%= cStatistic.FList(i).FMaechulProfitPer %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).FMaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice - cStatistic.FList(i).FupcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice) %></td>
    </tr>
    <%if chkChannel ="1" then%>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td>www</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fwww_ItemNO))%></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Fwww_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fwww_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemCost-cStatistic.FList(i).Fwww_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_MaechulProfit) %></b></td>
        <td align="right"><%=cStatistic.FList(i).Fwww_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ReducedPrice-cStatistic.FList(i).Fwww_BuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).Fwww_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ReducedPrice - cStatistic.FList(i).Fwww_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_overValueStockPrice) %></td>
    </tr>
    <% if (FALSE) then %>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td >모바일/App</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fma_ItemNO)) %></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Fma_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fma_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ItemCost-cStatistic.FList(i).Fma_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fma_MaechulProfit)%></b></td>
        <td align="right"><%= cStatistic.FList(i).Fma_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Fma_ReducedPrice-cStatistic.FList(i).Fma_BuyCash) %></td>
        <td align="right"> <%= cStatistic.FList(i).Fma_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fma_ReducedPrice - cStatistic.FList(i).Fma_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fma_overValueStockPrice) %></td>
    </tr>
    <% end if %>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td >MOB</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fm_ItemNO)) %></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Fm_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fm_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ItemCost-cStatistic.FList(i).Fm_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fm_MaechulProfit)%></b></td>
        <td align="right"><%= cStatistic.FList(i).Fm_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Fm_ReducedPrice-cStatistic.FList(i).Fm_BuyCash) %></td>
        <td align="right"> <%= cStatistic.FList(i).Fm_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ReducedPrice - cStatistic.FList(i).Fm_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td >MOB_제휴</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fmk_ItemNO)) %></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Fmk_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fmk_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ItemCost-cStatistic.FList(i).Fmk_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_MaechulProfit)%></b></td>
        <td align="right"><%= cStatistic.FList(i).Fmk_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ReducedPrice-cStatistic.FList(i).Fmk_BuyCash) %></td>
        <td align="right"> <%= cStatistic.FList(i).Fmk_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_ReducedPrice - cStatistic.FList(i).Fmk_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fmk_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td >App</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fa_ItemNO)) %></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Fa_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fa_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ItemCost-cStatistic.FList(i).Fa_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fa_MaechulProfit)%></b></td>
        <td align="right"><%= cStatistic.FList(i).Fa_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Fa_ReducedPrice-cStatistic.FList(i).Fa_BuyCash) %></td>
        <td align="right"> <%= cStatistic.FList(i).Fa_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ReducedPrice - cStatistic.FList(i).Fa_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td >제휴</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fo_ItemNO)) %></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Fo_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fo_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ItemCost-cStatistic.FList(i).Fo_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fo_MaechulProfit)%></b></td>
        <td align="right"><%= cStatistic.FList(i).Fo_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Fo_ReducedPrice-cStatistic.FList(i).Fo_BuyCash) %></td>
        <td align="right"> <%= cStatistic.FList(i).Fo_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ReducedPrice - cStatistic.FList(i).Fo_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td >해외몰</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Ff_ItemNO)) %></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Ff_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Ff_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ItemCost-cStatistic.FList(i).Ff_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Ff_MaechulProfit)%></b></td>
        <td align="right"><%= cStatistic.FList(i).Ff_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Ff_ReducedPrice-cStatistic.FList(i).Ff_BuyCash) %></td>
        <td align="right"> <%= cStatistic.FList(i).Ff_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ReducedPrice - cStatistic.FList(i).Ff_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_overValueStockPrice) %></td>
    </tr>
    <%end if%>
    <% if groupUserLevel ="1" then%>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td>WHITE</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv0_ItemNO))%></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Flv0_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv0_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ItemCost-cStatistic.FList(i).Flv0_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_MaechulProfit) %></b></td>
        <td align="right"><%=cStatistic.FList(i).Flv0_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ReducedPrice-cStatistic.FList(i).Flv0_BuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).Flv0_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_ReducedPrice - cStatistic.FList(i).Flv0_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv0_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td>RED</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv1_ItemNO))%></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Flv1_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv1_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ItemCost-cStatistic.FList(i).Flv1_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_MaechulProfit) %></b></td>
        <td align="right"><%=cStatistic.FList(i).Flv1_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ReducedPrice-cStatistic.FList(i).Flv1_BuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).Flv1_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_ReducedPrice - cStatistic.FList(i).Flv1_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv1_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td>VIP</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv2_ItemNO))%></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Flv2_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv2_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ItemCost-cStatistic.FList(i).Flv2_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_MaechulProfit) %></b></td>
        <td align="right"><%=cStatistic.FList(i).Flv2_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ReducedPrice-cStatistic.FList(i).Flv2_BuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).Flv2_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_ReducedPrice - cStatistic.FList(i).Flv2_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv2_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td>VIP GOLD</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv3_ItemNO))%></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Flv3_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv3_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ItemCost-cStatistic.FList(i).Flv3_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_MaechulProfit) %></b></td>
        <td align="right"><%=cStatistic.FList(i).Flv3_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ReducedPrice-cStatistic.FList(i).Flv3_BuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).Flv3_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_ReducedPrice - cStatistic.FList(i).Flv3_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv3_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td>VVIP</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv4_ItemNO))%></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Flv4_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv4_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ItemCost-cStatistic.FList(i).Flv4_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_MaechulProfit) %></b></td>
        <td align="right"><%=cStatistic.FList(i).Flv4_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ReducedPrice-cStatistic.FList(i).Flv4_BuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).Flv4_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_ReducedPrice - cStatistic.FList(i).Flv4_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv4_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td>STAFF</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv7_ItemNO))%></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Flv7_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv7_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ItemCost-cStatistic.FList(i).Flv7_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_MaechulProfit) %></b></td>
        <td align="right"><%=cStatistic.FList(i).Flv7_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ReducedPrice-cStatistic.FList(i).Flv7_BuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).Flv7_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_ReducedPrice - cStatistic.FList(i).Flv7_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv7_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td>FAMILY</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Flv8_ItemNO))%></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Flv8_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Flv8_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ItemCost-cStatistic.FList(i).Flv8_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_MaechulProfit) %></b></td>
        <td align="right"><%=cStatistic.FList(i).Flv8_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ReducedPrice-cStatistic.FList(i).Flv8_BuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).Flv8_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_ReducedPrice - cStatistic.FList(i).Flv8_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Flv8_overValueStockPrice) %></td>
    </tr>
    <tr bgcolor="#e3f1fb" align="Center">
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
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
        <td align="center"><%= cStatistic.FList(i).FMakerID %></td>
        <td align="center"><%= cStatistic.FList(i).fpurchasetypename %></td>
        <td>비회원</td>
        <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fnomem_ItemNO))%></td>
        <% if (NOT C_InspectorUser) then %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_OrgitemCost) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ItemcostCouponNotApplied) %></td>
        <td align="right" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ItemCost) %></b></td>
        <td align="right"><%if cStatistic.FList(i).Fnomem_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fnomem_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ItemCost-cStatistic.FList(i).Fnomem_ReducedPrice) %></td>
        <% end if %>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ReducedPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_BuyCash) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_MaechulProfit) %></b></td>
        <td align="right"><%=cStatistic.FList(i).Fnomem_MaechulProfitper%>%</td>
        <td align="right"> <%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ReducedPrice-cStatistic.FList(i).Fnomem_BuyCash) %></td>
        <td align="right"><%= cStatistic.FList(i).Fnomem_MaechulProfitPer2 %>%</td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_upcheJungsan) %></td>
        <td align="right"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_ReducedPrice - cStatistic.FList(i).Fnomem_upcheJungsan) %></b></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_avgipgoPrice) %></td>
        <td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).Fnomem_overValueStockPrice) %></td>
    </tr>
    <% end if %>
    <%
        if (i mod 500)=0 then
            Response.Flush          '버퍼리플래쉬
        end if
    next
    %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="22" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->