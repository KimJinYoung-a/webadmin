<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 매출집계- 관심등록전환매출
' History : 2016.10.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/maechul/statisticCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<%
dim menupos : menupos = getNumeric(requestcheckvar(request("menupos"),10))
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vSorting
dim sellchnl, vCateL, vCateM, vCateS, vIsBanPum, mwdiv
dim iCurrPage,iPageSize,iTotalPage,iTotCnt, dispCate,vBrandID ,itemid
dim  vTotwww_ItemNO,vTotwww_ItemCost,vTotwww_MaechulProfit,vTotwww_BuyCash,vTotma_ItemNO,vTotma_ItemCost,vTotma_MaechulProfit
dim vTotma_BuyCash,vTotout_ItemNO,vTotout_ItemCost,vTotout_MaechulProfit	,vTotout_BuyCash			
dim vTotwww_MaechulProfitPer ,vTotma_MaechulProfitPer ,vTotout_MaechulProfitPer 
Dim vTot_OrderCnt, vTot_ItemNO, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit
Dim vTot_MaechulProfitPer, vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
dim vTot_upcheJungsan, lec_cdl, lec_cdm, chkImg
	iPageSize = 5000
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= RequestCheckvar(request("sitename"),16)
	vDateGijun	= NullFillWith(RequestCheckvar(request("date_gijun"),16),"regdate")
	vSYear		= NullFillWith(RequestCheckvar(request("syear"),4),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(RequestCheckvar(request("smonth"),2),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(RequestCheckvar(request("sday"),2),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(RequestCheckvar(request("eyear"),4),Year(now))
	vEMonth		= NullFillWith(RequestCheckvar(request("emonth"),2),Month(now))
	vEDay		= NullFillWith(RequestCheckvar(request("eday"),2),Day(now))
	vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"itemsellcntD")
	vBrandID	= NullFillWith(RequestCheckvar(request("ebrand"),32),"")
	vCateL		= NullFillWith(RequestCheckvar(request("cdl"),3),"")
	vCateM		= NullFillWith(RequestCheckvar(request("cdm"),3),"")
	vCateS		= NullFillWith(RequestCheckvar(request("cds"),3),"")
	dispCate = requestCheckvar(request("disp"),16)
	itemid      = requestCheckvar(request("itemid"),255)
	vIsBanPum	= NullFillWith(RequestCheckvar(request("isBanpum"),16),"all")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(RequestCheckvar(request("mwdiv"),1),"")
	iCurrPage =requestCheckVar(request("iC"),4)
	lec_cdl = RequestCheckvar(request("lec_cdl"),3)
	lec_cdm = RequestCheckvar(request("lec_cdm"),3)
	chkImg		= requestCheckvar(request("chkImg"),1)
  	if itemid <> "" then
		if checkNotValidHTML(itemid) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
if chkImg ="" then chkImg = 0	
if iCurrPage = "" then iCurrPage = 1
if vSiteName = "" then vSiteName = "diyitem"

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
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectMakerid = vBrandID
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectItemid   = itemid 
	cStatistic.FPageSize = iPageSize
	cStatistic.FCurrPage = iCurrPage
	cStatistic.FRectIncStockAvgPrc = true '' 평균매입가 포함 쿼리여부.
	cStatistic.fStatistic_wish()

	iTotCnt = cStatistic.FResultCount
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
	
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
		검색결과 : <b><%=iTotCnt%></b>
		※ 최대 5천건 까지 보여 집니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td></td>

	<% IF chkImg = 1 then %>
		<td></td>
	<% END IF %>

	<td></td>
    <td></td>
    <td></td>
	<td>A</td>
	<td>B</td>
    <!--<td>C</td>-->
    <td>D</td>
    <td>E</td>
    <td>F</td>
    <td>G</td>
    <td>H</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품코드</td>

	<% IF chkImg = 1 then %>
		<td>이미지</td>
	<% END IF %>

	<td>브랜드</td>
    <td>전시카테고리</td>
    <td>상품명</td>
	<td>
		판매가
	</td>
	<td>
		매입가
	</td>
    <!--<td>
    	총담은수
    	<Br>D+E
    </td>-->
    <td>
    	판매전환수
    	<br>(판매건수)
    </td>
    <td>
    	위시
    	<br>담긴건수
    </td>
    <td>
    	판매전환율
    </td>
    <td>
    	전체매출
    </td>
    <td>
    	최근위시수1일
    </td>
</tr>
<% if cStatistic.FTotalCount>0 then %>
	<%
	dim tot_totwishcnt, tot_itemsellcnt, tot_itemwishcnt, tot_itemsellconversrate, tot_itemsellsum, tot_recentfavcount

	For i = 0 To cStatistic.FTotalCount -1

	'tot_totwishcnt = tot_totwishcnt + cStatistic.FItemList(i).ftotwishcnt
	tot_itemsellcnt = tot_itemsellcnt + cStatistic.FItemList(i).fitemsellcnt
	tot_itemwishcnt = tot_itemwishcnt + cStatistic.FItemList(i).fitemwishcnt
	tot_itemsellconversrate = tot_itemsellconversrate + cStatistic.FItemList(i).fitemsellconversrate
	tot_itemsellsum = tot_itemsellsum + cStatistic.FItemList(i).fitemsellsum
	tot_recentfavcount = tot_recentfavcount + cStatistic.FItemList(i).frecentfavcount
	%>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
		<td>
			<%= cStatistic.FItemList(i).FitemID %>
		</td>

		<% IF chkImg = 1 then %>
			<td><img src="<%= cStatistic.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<% END IF %>

		<td>
			<%= cStatistic.FItemList(i).FMakerID %>
		</td>
		<td align="left">
			<% if cStatistic.FItemList(i).Fcode_large_nm<>"" then %>
				<%= Replace(cStatistic.FItemList(i).Fcode_large_nm,"^^"," >> ") %>
			<% end if %>

			<% if cStatistic.FItemList(i).Fcode_mid_nm <> "" then %>
				>> <%= cStatistic.FItemList(i).Fcode_mid_nm %>
			<% end if %>
		</td>
		<td align="left"><%= cStatistic.FItemList(i).fitemname %></td>
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).fsellcash) %></td>
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).fbuycash) %></td>
		<!--<td align="right"><%'= CurrFormat(cStatistic.FItemList(i).ftotwishcnt) %></td>-->
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).fitemsellcnt) %></td>
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).fitemwishcnt) %></td>
		<td align="right"><%= round(CurrFormat(cStatistic.FItemList(i).fitemsellconversrate),1) %>%</td>
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).fitemsellsum) %></td>
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).frecentfavcount) %></td>
	</tr>
	<% Next %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="<% IF chkImg = 1 then %>7<% else %>6<% end if %>">총계</td>
		<!--<td align="right"><%= CurrFormat(tot_totwishcnt) %></td>-->
		<td align="right"><%= CurrFormat(tot_itemsellcnt) %></td>
		<td align="right"><%= CurrFormat(tot_itemwishcnt) %></td>
		<td align="right"><%= round(CurrFormat(tot_itemsellconversrate/cStatistic.FTotalCount),1) %>%</td>
		<td align="right"><%= CurrFormat(tot_itemsellsum) %></td>
		<td align="right"><%= CurrFormat(tot_recentfavcount) %></td>
	</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">검색결과가 없습니다.</td>
	</tr>
<% end if %>
</table>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->