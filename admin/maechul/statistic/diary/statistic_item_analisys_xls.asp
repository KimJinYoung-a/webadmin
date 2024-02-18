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
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_analisys_diary.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<%


	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vSorting, vCateL, vCateM, vCateS, vIsBanPum, vPurchasetype, v6Ago
	dim sellchnl, inc3pl
	Dim mwdiv 
	Dim dispCate,vBrandID, chkImg ,itemid
	dim iCurrPage,iPageSize,iTotalPage,iTotCnt
	dim sVType
	dim  vTotwww_ItemNO,vTotwww_ItemCost,vTotwww_MaechulProfit,vTotwww_BuyCash,vTotma_ItemNO,vTotma_ItemCost,vTotma_MaechulProfit,vTotma_BuyCash,vTotout_ItemNO,vTotout_ItemCost,vTotout_MaechulProfit	,vTotout_BuyCash			
	dim vTotwww_MaechulProfitPer ,vTotma_MaechulProfitPer ,vTotout_MaechulProfitPer 
	dim  bufStr 
	dim dy ,diaryyear
	
	iPageSize = 5000
	
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
  
	if iCurrPage = "" or iCurrPage ="0" then 
	    %>
	<script type="text/javascript">
	    alert("다운받을 내용이 없습니다. 페이지 선택을 해주세요 ");
	    window.close();
	</script>
<%response.end
    end if
    
	if chkImg ="" then chkImg = 0	
	if sVType ="" then sVType = 1
	    
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
	cStatistic.FRectdiaryyear = diaryyear
	
	if sVType=3 then
	    cStatistic.fStatistic_item_channel()
    else    
	    cStatistic.fStatistic_item()
    end if
    
    iTotCnt = cStatistic.FResultCount
 
 Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=[다이어리]상품별매출통계.csv"
Response.CacheControl = "public"

 IF sVType = 1 THEN   
    response.write "상품코드,상품명,카테고리,브랜드,상품수량,소비자가[상품],판매가[상품],구매총액[상품],보너스쿠폰사용액[상품],취급액,매입총액[상품],매출수익,수익율,매출수익2,수익율" & VbCrlf
 ELSEIF sVType = 2 THEN   
    response.write "날짜,상품코드,상품명,카테고리,브랜드,상품수량,소비자가[상품],판매가[상품],구매총액[상품],보너스쿠폰사용액[상품],취급액,매입총액[상품],매출수익,수익율,매출수익2,수익율" & VbCrlf
 ELSEIF  sVType = 3 THEN   
    response.write "날짜,상품코드,상품명,카테고리,브랜드,[Total]상품수량,[Total]구매총액,[Total]매출수익,[Total]수익율,[WWW]상품수량,[WWW]구매총액,[WWW]매출수익,[WWW]수익율,[MOB+APP]상품수량,[MOB+APP]구매총액,[MOB+APP]매출수익,[MOB+APP]수익율,[OUT]상품수량,[OUT]구매총액,[OUT]매출수익,[OUT]수익율"& VbCrlf 
 END IF
 
 
			For i = 0 To cStatistic.FTotalCount -1
			 bufStr = "" 
			 
			IF sVType = 3  then
			    
			bufStr = bufStr & cStatistic.FList(i).Fddate
			bufStr = bufStr & "," & cStatistic.FList(i).FitemID
			bufStr = bufStr & "," & replace(cStatistic.FList(i).Fitemname,",","")
			bufStr = bufStr & "," & cStatistic.FList(i).FCateFullName
			bufStr = bufStr & "," & cStatistic.FList(i).FMakerID
			bufStr = bufStr & "," & cStatistic.FList(i).FItemNo
			bufStr = bufStr & "," & cStatistic.FList(i).FItemCost 
			bufStr = bufStr & "," & cStatistic.FList(i).FMaechulProfit  
			bufStr = bufStr & "," & cStatistic.FList(i).FMaechulProfitper  &"%"
			bufStr = bufStr & "," & cStatistic.FList(i).Fwww_itemno    
			bufStr = bufStr & "," & cStatistic.FList(i).Fwww_itemcost  
			bufStr = bufStr & "," & cStatistic.FList(i).Fwww_maechulprofit 
			bufStr = bufStr & "," & cStatistic.FList(i).Fwww_maechulprofitper  &"%"
			bufStr = bufStr & "," & cStatistic.FList(i).Fma_itemno  
			bufStr = bufStr & "," & cStatistic.FList(i).Fma_itemcost  
			bufStr = bufStr & "," & cStatistic.FList(i).Fma_maechulprofit 
			bufStr = bufStr & "," & cStatistic.FList(i).Fma_maechulprofitper  &"%"
			bufStr = bufStr & "," & cStatistic.FList(i).Foutmall_itemno 
			bufStr = bufStr & "," & cStatistic.FList(i).Foutmall_itemcost  
			bufStr = bufStr & "," & cStatistic.FList(i).Foutmall_maechulprofit  
			bufStr = bufStr & "," & cStatistic.FList(i).Foutmall_maechulprofitper &"%" 
		 
		    ELSE
		        
		     IF sVType = 2 then 
		     bufStr = bufStr & cStatistic.FList(i).Fddate   
		     bufStr = bufStr & "," &  cStatistic.FList(i).FitemID
		    else
		     bufStr = bufStr & cStatistic.FList(i).FitemID     
		     END IF
		    		bufStr = bufStr & "," & replace(cStatistic.FList(i).Fitemname,",","")
            bufStr = bufStr & "," & cStatistic.FList(i).FCateFullName
            bufStr = bufStr & "," & cStatistic.FList(i).FMakerID
            bufStr = bufStr & "," & CDbl(cStatistic.FList(i).FItemNO) 
            bufStr = bufStr & "," & cStatistic.FList(i).FOrgitemCost              
            bufStr = bufStr & "," & cStatistic.FList(i).FItemcostCouponNotApplied 
            bufStr = bufStr & "," & cStatistic.FList(i).FItemCost 
            bufStr = bufStr & "," & cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice 
            bufStr = bufStr & "," & cStatistic.FList(i).FReducedPrice  
            bufStr = bufStr & "," & cStatistic.FList(i).FBuyCash                        
            bufStr = bufStr & "," & cStatistic.FList(i).FMaechulProfit        
            bufStr = bufStr & "," & cStatistic.FList(i).FMaechulProfitPer   &"%"                          
            bufStr = bufStr & "," & cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash 
            bufStr = bufStr & "," & cStatistic.FList(i).FMaechulProfitPer2 &"%"   
            
	        END IF
	        
	        response.write bufStr & VbCrlf
            NEXT
            
	  
 Set cStatistic = Nothing %>
 
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
