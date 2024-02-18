<%@ language=vbscript %>
<% option explicit

	'스크립트 타임아웃 시간 조정 (기본 90초)
	'Server.ScriptTimeout = 180
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_analisys.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
 
<% 
 
	dim  vTotwww_ItemNO,vTotwww_ItemCost,vTotwww_MaechulProfit,vTotwww_BuyCash,vTotma_ItemNO,vTotma_ItemCost,vTotma_MaechulProfit,vTotma_BuyCash,vTotout_ItemNO,vTotout_ItemCost,vTotout_MaechulProfit	,vTotout_BuyCash			
	dim vTotwww_MaechulProfitPer ,vTotma_MaechulProfitPer ,vTotout_MaechulProfitPer 
	dim  bufStr 
	iPageSize = 5000
	
	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vSorting, vCateL, vCateM, vCateS
    dim vIsBanPum, vPurchasetype, v6Ago, sellchnl, inc3pl, mwdiv, dispCate,vBrandID, chkImg ,itemid, iCurrPage,iPageSize,iTotalPage,iTotCnt
    dim syyyy, smm, sdd, eyyyy, emm, edd, reloading, date_gijun
    
	syyyy		= NullFillWith(request("syyyy"),Year(DateAdd("d",0,now())))
	smm		= NullFillWith(request("smm"),Month(DateAdd("d",0,now())))
	sdd		= NullFillWith(request("sdd"),Day(DateAdd("d",0,now())))
	eyyyy		= NullFillWith(request("eyyyy"),Year(now))
	emm		= NullFillWith(request("emm"),Month(now))
	edd		= NullFillWith(request("edd"),Day(now))
	 
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vSorting	= NullFillWith(request("sorting"),"itemsellcntD")
	vBrandID	= NullFillWith(request("ebrand"),"")
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	dispCate = requestCheckvar(request("disp"),16)
	itemid      = requestCheckvar(request("itemid"),255)
	chkImg		= requestCheckvar(request("chkImg"),1)
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(request("mwdiv"),"")
	inc3pl = request("inc3pl")
	iCurrPage =requestCheckVar(request("iC"),4)
	reloading    = requestCheckVar(request("reloading"),2)
	iCurrPage =requestCheckVar(request("iC"),4)
 
	 
	
	if iCurrPage = "" or iCurrPage ="0" then 
	    %>
	<script type="text/javascript">
	//    alert("다운받을 내용이 없습니다. 페이지 선택을 해주세요 ");
	  //  window.close();
	</script>
<% ' response.end
    end if
    
	if chkImg ="" then chkImg = 0	
    if reloading="" and vSiteName="" then vSiteName="10x10"
	    
	Dim vTot_OrderCnt, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer
	Dim vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2


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
	'cStatistic.FRectmaechulStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	'cStatistic.FRectmaechulEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectStartdate = syyyy & "-" & TwoNumber(smm) & "-" & TwoNumber(sdd)
	cStatistic.FRectEndDate = eyyyy & "-" & TwoNumber(emm) & "-" & TwoNumber(edd)
	cStatistic.FRectSiteName = vSiteName
	'cStatistic.FRect6MonthAgo = v6Ago 
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectMakerid = vBrandID
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 추가
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectItemid   = itemid 
	cStatistic.FPageSize = iPageSize
	cStatistic.FCurrPage = iCurrPage
	cStatistic.fStatistic_baguni()
    
   iTotCnt = cStatistic.FResultCount
 
Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=장바구니전환매출.csv"
Response.CacheControl = "public"
  
response.write "상품코드, 브랜드, 카테고리, 상품명, 판매가, 매입가, 총담은수, 판매전환수, 장바구니건수, 판매전환율, 전체매출, 총위시수, 최근위시수1일"& VbCrlf 
 
 
 
			For i = 0 To cStatistic.FTotalCount -1
			 bufStr = ""  
			 
		    bufStr = bufStr & cStatistic.FList(i).FitemID      
		    bufStr = bufStr & "," & cStatistic.FList(i).FMakerID
            bufStr = bufStr & "," & cStatistic.FList(i).fcatename
            bufStr = bufStr & "," & replace(cStatistic.FList(i).fitemname,",","") 
            bufStr = bufStr & "," & cStatistic.FList(i).fsellcash  
            bufStr = bufStr & "," & cStatistic.FList(i).fbuycash 
            bufStr = bufStr & "," & cStatistic.FList(i).ftotbagunicnt 
            bufStr = bufStr & "," & cStatistic.FList(i).fitemsellcnt 
            bufStr = bufStr & "," & cStatistic.FList(i).fitembagunicnt 
            bufStr = bufStr & "," & round(cStatistic.FList(i).fitemsellconversrate,1)  
            bufStr = bufStr & "," & cStatistic.FList(i).fitemsellsum  
            bufStr = bufStr & "," & cStatistic.FList(i).ffavcount 
            bufStr = bufStr & "," & cStatistic.FList(i).frecentfavcount 
	        
	        response.write bufStr & VbCrlf
            NEXT
            
	  
 Set cStatistic = Nothing %>
 
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
