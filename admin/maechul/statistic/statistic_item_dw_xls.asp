<%@ language=vbscript %>
<% option explicit

	'��ũ��Ʈ Ÿ�Ӿƿ� �ð� ���� (�⺻ 90��)
	'Server.ScriptTimeout = 180
%>
<%
'####################################################
' Description :  ��ǰ�� �������
' History : 2016.01.20 ������ ����
'			2016.06.10 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_dw.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSorting		' , vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay
dim sellchnl, inc3pl, vCateL, vCateM, vCateS, vIsBanPum, vPurchasetype, v6Ago, mwdiv, rdsite
dim iCurrPage,iPageSize,iTotalPage,iTotCnt, dispCate,vBrandID, chkImg ,itemid, sVType
dim  vTotwww_ItemNO,vTotwww_ItemCost,vTotwww_MaechulProfit,vTotwww_BuyCash,vTotm_ItemNO,vTotm_ItemCost,vTotm_MaechulProfit, vTotm_BuyCash
dim  vTota_ItemNO,vTota_ItemCost,vTota_MaechulProfit,vTota_BuyCash,vTotf_ItemNO,vTotf_ItemCost,vTotf_MaechulProfit, vTotf_BuyCash
dim vTotout_ItemNO,vTotout_ItemCost,vTotout_MaechulProfit	,vTotout_BuyCash
dim vTotwww_MaechulProfitPer ,vTotma_MaechulProfitPer ,vTotout_MaechulProfitPer, vTotm_MaechulProfitPer,vTota_MaechulProfitPer,vTotf_MaechulProfitPer
Dim vTot_OrderCnt, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit
Dim vTot_MaechulProfitPer, vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
dim vTot_upcheJungsan, vTot_avgipgoPrice, vTot_overValueStockPrice, vstartdate, venddate
dim chkcate,chkchn, showsuply, crect, groupid
Dim incStockAvg
	vstartdate = NullFillWith(requestCheckVar(request("startdate"),10),DateAdd("d",0,date()))
	venddate = NullFillWith(requestCheckVar(request("enddate"),10),date())
	iPageSize = 100000
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")  ''beasongdate  :�����=>�ֹ��� 2018/05/28  by eastone
	'vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	'vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	'vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	'vEYear		= NullFillWith(request("eyear"),Year(now))
	'vEMonth		= NullFillWith(request("emonth"),Month(now))
	'vEDay		= NullFillWith(request("eday"),Day(now))
	vSorting	= NullFillWith(request("sorting"),"itemcost")
	vBrandID	= NullFillWith(request("ebrand"),"")
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	dispCate = requestCheckvar(request("disp"),16)
	itemid      = requestCheckvar(request("itemid"),1000)
	chkImg		= requestCheckvar(request("chkImg"),1)
	chkcate		= requestCheckvar(request("chkcate"),1)
	chkchn     = requestCheckvar(request("chkchn"),1)
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(request("mwdiv"),"")
	rdsite		= NullFillWith(request("rdsite"),"")
	inc3pl = request("inc3pl")
	iCurrPage =requestCheckVar(request("iC"),4)
	sVType      = requestCheckvar(request("rdoVType"),1)
	showsuply   = requestCheckvar(request("showsuply"),10)
	crect       = RequestCheckVar(request("crect"),32)
	groupid     = RequestCheckVar(request("groupid"),32)
	incStockAvg = requestCheckvar(request("incStockAvg"),10)

	if iCurrPage = "" or iCurrPage ="0" then
	    %>
	<script type="text/javascript">
	    alert("�ٿ���� ������ �����ϴ�. ������ ������ ���ּ��� ");
	    window.close();
	</script>
<%response.end
    end if

if chkImg ="" then chkImg = 0
	if chkcate ="" then chkcate = 0
if sVType ="" then sVType = 1
if chkchn ="" then chkchn = 0
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
	cStatistic.FRectStartdate = vstartdate		' vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = venddate		' vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	'cStatistic.FRect6MonthAgo = v6Ago
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectRdsite = rdsite
	cStatistic.FRectMakerid = vBrandID
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 �߰�
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectItemid   = itemid
	cStatistic.FRectVType = sVType
	cStatistic.FPageSize = iPageSize
	cStatistic.FCurrPage = iCurrPage
	cStatistic.FRectIncStockAvgPrc = (incStockAvg<>"") ''true '' ��ո��԰� ���� ��������.
	cStatistic.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
	cStatistic.FRectGroupid = groupid
	cStatistic.FRectCompanyname = crect

	if chkchn="1" then
	    cStatistic.fStatistic_item_channel()
    else
	    cStatistic.fStatistic_item()
    end if

    iTotCnt = cStatistic.FResultCount

dim bufStr

Response.Buffer = true    '���ۻ�뿩��
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=��ǰ���������.csv"
Response.CacheControl = "public"

IF chkchn ="1" then
	IF sVType = 1 THEN
		response.write "��ǰ�ڵ�,��ǰ��,ī�װ�,�귣��,���Ա���,��������,[Total]��ǰ����,[Total]�����Ѿ�,[Total]��޾�,[Total]�������,[Total]������,[WWW]��ǰ����,[WWW]�����Ѿ�,[WWW]��޾�,[WWW]�������,[WWW]������,[MOB]��ǰ����,[MOB]�����Ѿ�,[MOB]��޾�,[MOB]�������,[MOB]������,[APP]��ǰ����,[APP]�����Ѿ�,[APP]��޾�,[APP]�������,[APP]������,[����]��ǰ����,[����]�����Ѿ�,[����]��޾�,[����]�������,[����]������,[�ؿܸ�]��ǰ����,[�ؿܸ�]�����Ѿ�,[�ؿܸ�]��޾�,[�ؿܸ�]�������,[�ؿܸ�]������,��ո��԰�,�������"& VbCrlf
	ELSEIF sVType = 2 THEN
		response.write "��¥,��ǰ�ڵ�,��ǰ��,ī�װ�,�귣��,���Ա���,��������,[Total]��ǰ����,[Total]�����Ѿ�,[Total]�������,[Total]������,[WWW]��ǰ����,[WWW]�����Ѿ�,[WWW]�������,[WWW]������,[MOB]��ǰ����,[MOB]�����Ѿ�,[MOB]�������,[MOB]������,[APP]��ǰ����,[APP]�����Ѿ�,[APP]�������,[APP]������,[����]��ǰ����,[����]�����Ѿ�,[����]�������,[����]������,[�ؿܸ�]��ǰ����,[�ؿܸ�]�����Ѿ�,[�ؿܸ�]�������,[�ؿܸ�]������,��ո��԰�,�������"& VbCrlf
	ELSEIF sVType = 3 THEN
		response.write "��ǰ����,��ǰ�ڵ�,�ɼ��ڵ�,���ڵ�,��ǰ��,ī�װ�,�귣��,���Ա���,��������,[Total]��ǰ����,[Total]�����Ѿ�,[Total]�������,[Total]������,[WWW]��ǰ����,[WWW]�����Ѿ�,[WWW]�������,[WWW]������,[MOB]��ǰ����,[MOB]�����Ѿ�,[MOB]�������,[MOB]������,[APP]��ǰ����,[APP]�����Ѿ�,[APP]�������,[APP]������,[����]��ǰ����,[����]�����Ѿ�,[����]�������,[����]������,[�ؿܸ�]��ǰ����,[�ؿܸ�]�����Ѿ�,[�ؿܸ�]�������,[�ؿܸ�]������,��ո��԰�,�������"& VbCrlf
	END IF
ELSE	
	IF sVType = 1 THEN
	    response.write "��ǰ�ڵ�,��ǰ��,ī�װ�,�귣��,���Ա���,��������,��ǰ����,�Һ��ڰ�[��ǰ],�ǸŰ�[��ǰ](��������),�����Ѿ�[��ǰ](��ǰ��������),���ʽ���������[��ǰ],��޾�,��ü�����1(��ǰ��������),�������1(�����Ѿױ���),������,�������2(��޾ױ���),������,��ü�����2(�����������),ȸ�����,��ո��԰�,�������" & VbCrlf
	ELSEIF sVType = 2 THEN
	    response.write "��¥,��ǰ�ڵ�,��ǰ��,ī�װ�,�귣��,���Ա���,��������,��ǰ����,�Һ��ڰ�[��ǰ],�ǸŰ�[��ǰ](��������),�����Ѿ�[��ǰ](��ǰ��������),���ʽ���������[��ǰ],��޾�,��ü�����1(��ǰ��������),�������1(�����Ѿױ���),������,�������2(��޾ױ���),������,��ü�����2(�����������),ȸ�����,��ո��԰�,�������" & VbCrlf
	ELSEIF sVType = 3 THEN
		response.write "��ǰ����,��ǰ�ڵ�,�ɼ��ڵ�,���ڵ�,��ǰ��,ī�װ�,�귣��,���Ա���,��������,��ǰ����,�Һ��ڰ�[��ǰ],�ǸŰ�[��ǰ](��������),�����Ѿ�[��ǰ](��ǰ��������),���ʽ���������[��ǰ],��޾�,��ü�����1(��ǰ��������),�������1(�����Ѿױ���),������,�������2(��޾ױ���),������,��ü�����2(�����������),ȸ�����,��ո��԰�,�������" & VbCrlf
	END IF
END IF	

For i = 0 To cStatistic.FTotalCount -1
	bufStr = ""

	IF chkchn ="1" then
		IF sVType = 3 then
			bufStr = bufStr & "10,"
		END IF
		IF sVType = 2  then 
			bufStr = bufStr & cStatistic.FList(i).Fddate & ","
		END IF	
		bufStr = bufStr & cStatistic.FList(i).FitemID
		IF sVType = 3 then
			bufStr = bufStr & "," & cStatistic.FList(i).Fitemoption
			bufStr = bufStr & "," & BF_MakeTenBarcode("10", cStatistic.FList(i).FitemID, cStatistic.FList(i).Fitemoption)
		END IF
		bufStr = bufStr & "," & replace(cStatistic.FList(i).Fitemname,",","")
		bufStr = bufStr & "," & cStatistic.FList(i).FCateFullName
		bufStr = bufStr & "," & cStatistic.FList(i).FMakerID
		bufStr = bufStr & "," & cStatistic.FList(i).Fomwdiv
		bufStr = bufStr & "," & mwdivName(cStatistic.FList(i).Fomwdiv)
		bufStr = bufStr & "," & vatIncludeName(cStatistic.FList(i).fvatinclude)
		bufStr = bufStr & "," & cStatistic.FList(i).FItemNo
		bufStr = bufStr & "," & cStatistic.FList(i).FItemCost
		bufStr = bufStr & "," & cStatistic.FList(i).freducedprice
		bufStr = bufStr & "," & cStatistic.FList(i).FMaechulProfit
		bufStr = bufStr & "," & cStatistic.FList(i).FMaechulProfitper  &"%"
		bufStr = bufStr & "," & cStatistic.FList(i).Fwww_itemno
		bufStr = bufStr & "," & cStatistic.FList(i).Fwww_itemcost
		bufStr = bufStr & "," & cStatistic.FList(i).fwww_reducedprice
		bufStr = bufStr & "," & cStatistic.FList(i).Fwww_maechulprofit
		bufStr = bufStr & "," & cStatistic.FList(i).Fwww_maechulprofitper  &"%"
		bufStr = bufStr & "," & cStatistic.FList(i).Fm_itemno
		bufStr = bufStr & "," & cStatistic.FList(i).Fm_itemcost
		bufStr = bufStr & "," & cStatistic.FList(i).fm_reducedprice
		bufStr = bufStr & "," & cStatistic.FList(i).Fm_maechulprofit
		bufStr = bufStr & "," & cStatistic.FList(i).Fm_maechulprofitper  &"%"
		bufStr = bufStr & "," & cStatistic.FList(i).Fa_itemno
		bufStr = bufStr & "," & cStatistic.FList(i).Fa_itemcost
		bufStr = bufStr & "," & cStatistic.FList(i).fa_reducedprice
		bufStr = bufStr & "," & cStatistic.FList(i).Fa_maechulprofit
		bufStr = bufStr & "," & cStatistic.FList(i).Fa_maechulprofitper  &"%"
		bufStr = bufStr & "," & cStatistic.FList(i).Foutmall_itemno
		bufStr = bufStr & "," & cStatistic.FList(i).Foutmall_itemcost
		bufStr = bufStr & "," & cStatistic.FList(i).foutmall_reducedprice
		bufStr = bufStr & "," & cStatistic.FList(i).Foutmall_maechulprofit
		bufStr = bufStr & "," & cStatistic.FList(i).Foutmall_maechulprofitper &"%"
		bufStr = bufStr & "," & cStatistic.FList(i).Ff_itemno
		bufStr = bufStr & "," & cStatistic.FList(i).Ff_itemcost
		bufStr = bufStr & "," & cStatistic.FList(i).ff_reducedprice
		bufStr = bufStr & "," & cStatistic.FList(i).Ff_maechulprofit
		bufStr = bufStr & "," & cStatistic.FList(i).Ff_maechulprofitper  &"%"
		bufStr = bufStr & "," & cStatistic.FList(i).FupcheJungsan
		bufStr = bufStr & "," & cStatistic.FList(i).FReducedPrice - cStatistic.FList(i).FupcheJungsan
		bufStr = bufStr & "," & cStatistic.FList(i).FavgipgoPrice
		bufStr = bufStr & "," & cStatistic.FList(i).FoverValueStockPrice

	ELSE

		IF sVType = 3 then
			bufStr = bufStr & "10,"
		END IF
		IF sVType = 2 then
		    bufStr = bufStr & cStatistic.FList(i).Fddate & ","
		END IF
	    bufStr = bufStr & cStatistic.FList(i).FitemID
		IF sVType = 3 then
			bufStr = bufStr & "," & cStatistic.FList(i).Fitemoption
			bufStr = bufStr & "," & BF_MakeTenBarcode("10", cStatistic.FList(i).FitemID, cStatistic.FList(i).Fitemoption)
		END IF
		bufStr = bufStr & "," & replace(cStatistic.FList(i).Fitemname,",","")
        bufStr = bufStr & "," & cStatistic.FList(i).FCateFullName
        bufStr = bufStr & "," & cStatistic.FList(i).FMakerID
		bufStr = bufStr & "," & mwdivName(cStatistic.FList(i).Fomwdiv)
		bufStr = bufStr & "," & vatIncludeName(cStatistic.FList(i).fvatinclude)
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
        bufStr = bufStr & "," & cStatistic.FList(i).FupcheJungsan
        bufStr = bufStr & "," & cStatistic.FList(i).FReducedPrice - cStatistic.FList(i).FupcheJungsan
		bufStr = bufStr & "," & cStatistic.FList(i).FavgipgoPrice
		bufStr = bufStr & "," & cStatistic.FList(i).FoverValueStockPrice

	END IF

	if i mod 3000 = 0 then
		Response.Flush		' ���۸��÷���
	end if
	response.write bufStr & VbCrlf
NEXT


 Set cStatistic = Nothing %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
