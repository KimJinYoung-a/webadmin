<%@ language=vbscript %>
<% option explicit

	'��ũ��Ʈ Ÿ�Ӿƿ� �ð� ���� (�⺻ 90��)
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
	    alert("�ٿ���� ������ �����ϴ�. ������ ������ ���ּ��� ");
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
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 �߰�
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
Response.AddHeader "Content-Disposition", "attachment; filename=[���̾]��ǰ���������.csv"
Response.CacheControl = "public"

 IF sVType = 1 THEN   
    response.write "��ǰ�ڵ�,��ǰ��,ī�װ�,�귣��,��ǰ����,�Һ��ڰ�[��ǰ],�ǸŰ�[��ǰ],�����Ѿ�[��ǰ],���ʽ���������[��ǰ],��޾�,�����Ѿ�[��ǰ],�������,������,�������2,������" & VbCrlf
 ELSEIF sVType = 2 THEN   
    response.write "��¥,��ǰ�ڵ�,��ǰ��,ī�װ�,�귣��,��ǰ����,�Һ��ڰ�[��ǰ],�ǸŰ�[��ǰ],�����Ѿ�[��ǰ],���ʽ���������[��ǰ],��޾�,�����Ѿ�[��ǰ],�������,������,�������2,������" & VbCrlf
 ELSEIF  sVType = 3 THEN   
    response.write "��¥,��ǰ�ڵ�,��ǰ��,ī�װ�,�귣��,[Total]��ǰ����,[Total]�����Ѿ�,[Total]�������,[Total]������,[WWW]��ǰ����,[WWW]�����Ѿ�,[WWW]�������,[WWW]������,[MOB+APP]��ǰ����,[MOB+APP]�����Ѿ�,[MOB+APP]�������,[MOB+APP]������,[OUT]��ǰ����,[OUT]�����Ѿ�,[OUT]�������,[OUT]������"& VbCrlf 
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
