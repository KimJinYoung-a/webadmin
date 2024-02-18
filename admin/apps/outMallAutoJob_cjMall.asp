<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/admin/etc/cjmall/cjmallitemcls.asp"-->
<!-- #include virtual="/admin/etc/cjmall/incCJmallFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false
    dim VaildIP
    VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72","61.252.133.70")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    dbget.Close()
    response.end
end if

dim act     : act = requestCheckVar(request("act"),32)
dim param1  : param1 = requestCheckVar(request("param1"),32)
dim param2  : param2 = requestCheckVar(request("param2"),32)
dim param3  : param3 = requestCheckVar(request("param3"),32)
dim param4  : param4 = requestCheckVar(request("param4"),32)
dim sqlStr, i, paramData, retVal
dim retCnt : retCnt = 0

Dim cnt
Dim OutMallOrderSerialArr
Dim OrgDetailKeyArr
Dim songjangDivArr
Dim songjangNoArr, sendReqCntArr, beasongdateArr, outmallGoodsIDArr
Dim cjMall, itemidArr, ArrRows

select Case act

    Case "outmallSongJangIp" ''제휴사 송장입력
    'response.end
        sqlStr = "select top 40 T.orderserial, T.OutMallOrderSerial"
        sqlStr = sqlStr & " ,T.OrgDetailKey, IsNULL(T.sendState,0) as sendState"
        sqlStr = sqlStr & " ,D.songjangDiv, D.songjangNo, D.itemNo, D.beasongdate, T.outMallGoodsNo"
        sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master M"
        sqlStr = sqlStr & " 	on T.orderserial=M.orderserial"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D"
        sqlStr = sqlStr & " 	on T.orderserial=D.orderserial"
        ''sqlStr = sqlStr & " 	and T.matchitemid=D.itemid"
        ''sqlStr = sqlStr & " 	and T.matchitemoption=D.itemoption"
		sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid"					'// 기존 주문에 합쳐진 경우(빨강1개,파랑1개 -> 파랑2개)
		sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption"
        sqlStr = sqlStr & " 	and D.currstate=7"
        sqlStr = sqlStr & " 	left join db_order.dbo.tbl_songjang_div V"
        sqlStr = sqlStr & " 	on D.songjangDiv=V.divcd"
        sqlStr = sqlStr & " where datediff(m,T.regdate,getdate())<7"    ''20130304 추가
        sqlStr = sqlStr & " and T.sellsite='"&param1&"'"
        sqlStr = sqlStr & " and T.OrgDetailKey is Not NULL"             ''디테일키 입력 주문건만..
        sqlStr = sqlStr & " and IsNULL(T.sendState,0)=0"
        sqlStr = sqlStr & " and T.sendReqCnt<3"                         ''여러번 시도 안되도록. 추가.
        sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"      ''교환 취소 반품 제외.
        sqlStr = sqlStr & " order by D.beasongdate desc"
        rsget.Open sqlStr,dbget,1
        cnt = rsget.RecordCount
        ReDim TenOrderserial(cnt)
        ReDim OutMallOrderSerialArr(cnt)
        ReDim OrgDetailKeyArr(cnt)
        ReDim songjangDivArr(cnt)
        ReDim songjangNoArr(cnt)
        Redim sendReqCntArr(cnt)
        Redim beasongdateArr(cnt)
        Redim outmallGoodsIDArr(cnt)
        i = 0
        if Not rsget.Eof then
            do until rsget.eof
            TenOrderserial(i) = rsget("orderserial")
            OutMallOrderSerialArr(i) = rsget("OutMallOrderSerial")
            OrgDetailKeyArr(i) = rsget("OrgDetailKey")
			songjangDivArr(i) = rsget("songjangDiv")
			songjangNoArr(i) = rsget("songjangNo")
			sendReqCntArr(i) = rsget("itemNo")
			beasongdateArr(i) = rsget("beasongdate")
			outmallGoodsIDArr(i) = rsget("outMallGoodsNo")
            i=i+1
            rsget.MoveNext
    		loop
        end if
        rsget.close

        if (cnt<1) then
            response.Write "S_NONE.."
            dbget.Close() : response.end
        else
            rw "CNT="&CNT
            for i=LBound(OutMallOrderSerialArr) to UBound(OutMallOrderSerialArr)
                if (OutMallOrderSerialArr(i)<>"") then
					IF (LCASE(param1)="cjmall") then
                        ''var params = "ten_ord_no="+tenorderserial+"&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
                        ''var popwin=window.open('/admin/etc/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');

                        paramData = "redSsnKey=system&ten_ord_no="&TenOrderserial(i)&"&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2cjMallDlvCode(songjangDivArr(i))&"&inv_no="&server.URLEncode(songjangNoArr(i))
                        rw paramData

                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCJmallSongjangInputProc.asp",paramData)
                             rw retVal
                        end if
                    end if
                end if
            next
        end if
    Case "cjmallCheckRDItem" '' cjmall 승인,판매상태 확인 Batch
        paramData = "redSsnKey=system&cmdparam=confirmItemAuto&subcmd="&param2
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        response.write retVal&VbCRLF

    Case "cjmallExpireItem"   '' 품절 처리 요망 (사용안함, 조건배송등)
        Set cjMall = new CCjmall
    	cjMall.FCurrPage					= 1
    	cjMall.FPageSize					= 10
    	cjMall.FRectExtNotReg				= "D"
    	cjMall.FRectMatchCate				= ""
    	cjMall.FRectPrdDivMatch				= ""
    	cjMall.FRectSellYn					= ""
    	cjMall.FRectLimitYn					= ""
    	cjMall.FRectonlyValidMargin 		= ""
    	cjMall.FRectMinusMargin 			= ""
    	cjMall.FRectFailCntExists			= ""
    	cjMall.FRectCjSell10x10Soldout      = ""
    	cjMall.FRectExpensive10x10          = ""
        cjMall.FRectExtSellYn               = "Y"

    	cjMall.getCjmallreqExpireItemList

    	IF (cjMall.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set cjMall= Nothing
            response.end
        ENd IF

        for i=0 to cjMall.FResultCount - 1
            itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
        next
        set cjMall= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr
        paramData = "redSsnKey=system&cmdparam=EditSellYn&subcmd=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal

    Case "cjmallSoldOutItem" '' 품절처리 상품. (10x10 품절, cj판매중)

        Set cjMall = new CCjmall
    	cjMall.FCurrPage					= 1
    	cjMall.FPageSize					= 10
    	cjMall.FRectExtNotReg				= "D"
    	cjMall.FRectMatchCate				= ""
    	cjMall.FRectPrdDivMatch				= ""
    	cjMall.FRectSellYn					= ""
    	cjMall.FRectLimitYn					= ""
    	cjMall.FRectonlyValidMargin 		= ""
    	cjMall.FRectMinusMargin 			= ""
    	cjMall.FRectFailCntExists			= ""
    	cjMall.FRectCjSell10x10Soldout      = "Y"
    	cjMall.FRectExpensive10x10          = ""

    	cjMall.GetCjmallRegedItemList

        IF (cjMall.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set cjMall= Nothing
            response.end
        ENd IF

        for i=0 to cjMall.FResultCount - 1
            itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
        next
        set cjMall= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=EditSellYn&subcmd=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "cjmallexpensive10x10" '' 가격비싼상품 수정

        Set cjMall = new CCjmall
    	cjMall.FCurrPage					= 1
    	cjMall.FPageSize					= 10
    	cjMall.FRectExtNotReg				= "D"
    	cjMall.FRectMatchCate				= ""
    	cjMall.FRectPrdDivMatch				= ""
    	cjMall.FRectSellYn					= "" ''"Y"
    	cjMall.FRectLimitYn					= ""
    	cjMall.FRectonlyValidMargin 		= ""
    	cjMall.FRectMinusMargin 			= ""
    	cjMall.FRectFailCntExists			= ""
    	cjMall.FRectCjSell10x10Soldout      = ""
    	cjMall.FRectExtSellYn               = "" ''"Y"
    	cjMall.FRectExpensive10x10          = "Y"

    	cjMall.GetCjmallRegedItemList

        IF (cjMall.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set cjMall= Nothing
            response.end
        ENd IF

        for i=0 to cjMall.FResultCount - 1
            itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
        next
        set cjMall= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "cjmallpriceDiff" '' 가격 상이 수정

        Set cjMall = new CCjmall
    	cjMall.FCurrPage					= 1
    	cjMall.FPageSize					= 10
    	cjMall.FRectExtNotReg				= "D"
    	cjMall.FRectMatchCate				= ""
    	cjMall.FRectPrdDivMatch				= "Y"
    	cjMall.FRectSellYn					= "Y"
    	cjMall.FRectLimitYn					= ""
    	cjMall.FRectonlyValidMargin 		= "Y"
    	cjMall.FRectMinusMargin 			= ""
    	cjMall.FRectFailCntExists			= ""
    	cjMall.FRectCjSell10x10Soldout      = ""
    	cjMall.FRectExpensive10x10          = ""
        cjMall.FRectdiffPrc 				= "Y"
    	cjMall.GetCjmallRegedItemList

        IF (cjMall.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set cjMall= Nothing
            response.end
        ENd IF

        for i=0 to cjMall.FResultCount - 1
            itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
        next
        set cjMall= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr
        paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal

    Case "cjmallpriceEdit2" '' 가격 수정(옵션추가금액 관련)
        sqlStr = " select top 10 ro.itemid"
        sqlStr = sqlStr & " from  db_item.dbo.tbl_cjmall_regItem r"
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_outmall_regedoption ro"
        sqlStr = sqlStr & " 	on ro.itemid=r.itemid"
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item_option o"
        sqlStr = sqlStr & " 	on ro.itemid=o.itemid"
        sqlStr = sqlStr & " 	and ro.itemoption=o.itemoption"
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i"
        sqlStr = sqlStr & " 	on r.itemid=i.itemid"
        sqlStr = sqlStr & " where ro.mallid='cjmall'"
        sqlStr = sqlStr & " and r.optaddPrcCnt>0"
        sqlStr = sqlStr & " and r.cjmallprice+o.optAddprice<>ro.outmallAddPrice"
        sqlStr = sqlStr & " and r.cjmallsellyn='Y'"
        sqlStr = sqlStr & " order by r.lastStatCheckDate"

        rsget.Open sqlStr,dbget,1
        if not rsget.Eof then
            ArrRows = rsget.getRows()
        end if
        rsget.close

        itemidArr = ""
        if isArray(ArrRows) then
            For i =0 To UBound(ArrRows,2)
                itemidArr = itemidArr + CStr(ArrRows(0,i)) + ","
            Next
        else
            rw "S_NONE"
            dbget.Close() : response.end
        end if

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr
        paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "cjmallpriceEdit3" '' 가격 수정(2014-12-19 진영 추가)
        sqlStr = " select top 20 itemid from "
        sqlStr = sqlStr & " db_item.dbo.tbl_cjmall_regitem "
        sqlStr = sqlStr & " where optAddPrcCnt > 0  "
        sqlStr = sqlStr & " and cjmallsellyn = 'Y' "
        sqlStr = sqlStr & " and cjmallStatCd = '7' "
        sqlStr = sqlStr & " order by lastpriceCheckDate asc "
        rsget.Open sqlStr,dbget,1
        if not rsget.Eof then
            ArrRows = rsget.getRows()
        end if
        rsget.close

        itemidArr = ""
        if isArray(ArrRows) then
            For i =0 To UBound(ArrRows,2)
                itemidArr = itemidArr + CStr(ArrRows(0,i)) + ","
            Next
        else
            rw "S_NONE"
            dbget.Close() : response.end
        end if

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr
        paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "cjmallmarginItem" '' 역마진 상품이면(보통 할인들어간 것) sellcash를 orgprice로
    	sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 10 "
		sqlStr = sqlStr & " i.itemid, i.itemname, i.smallImage , i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash , i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt, c.mapCnt , J.cjmallRegdate, J.cjmallLastUpdate, J.cjmallPrdNo, J.cjmallPrice, J.cjmallSellYn, J.regUserid, IsNULL(J.cjmallStatCd,-9) as cjmallStatCd , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr, PD.infodiv, PD.cdmKey, PD.cddkey, UC.defaultfreeBeasongLimit  "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_item_contents as t on i.itemid = t.itemid "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='cjmall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and t.infodiv = PD.infodiv "
		sqlStr = sqlStr & " WHERE 1 = 1 and i.isusing='Y' and i.basicimage is not null and i.itemdiv < 50 and i.itemdiv not in ('08','09') "
		sqlStr = sqlStr & " and i.cate_large <> '' "
		sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) "
		sqlStr = sqlStr & " and i.sellcash >= 1000 and i.itemdiv<>'06' and uc.isExtUsing='Y' "
		sqlStr = sqlStr & " and not (i.deliverytype='9' and uc.defaultfreeBeasongLimit < 10000) "
		sqlStr = sqlStr & " and J.cjmallStatCd = 7 and J.cjmallPrdNo is Not Null and i.sellYn='Y' "
		sqlStr = sqlStr & " and J.cjmallSellYn='Y' and i.sellcash<>0 "
		sqlStr = sqlStr & " and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)<15 "
		sqlStr = sqlStr & " and J.cjmallSellYn= 'Y' and isNULL(J.optAddPrcCnt,0)=0 and i.isExtUsing='Y' "
		sqlStr = sqlStr & " and i.deliverytype not in ('7') and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000))) "
		sqlStr = sqlStr & " and i.cate_large <> '075' and (i.cate_large + i.cate_mid <> '110010') "
		sqlStr = sqlStr & " and (i.cate_large + i.cate_mid <> '110030') and (i.cate_large + i.cate_mid <> '110040') "
		sqlStr = sqlStr & " and (i.cate_large + i.cate_mid <> '110060') and (i.cate_large + i.cate_mid <> '110050') "
		sqlStr = sqlStr & " and i.orgprice <> J.cjmallPrice "
		sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "

        rsget.Open sqlStr,dbget,1
        if not rsget.Eof then
            ArrRows = rsget.getRows()
        end if
        rsget.close

        itemidArr = ""
        if isArray(ArrRows) then
            For i =0 To UBound(ArrRows,2)
                itemidArr = itemidArr + CStr(ArrRows(0,i)) + ","
            Next
        else
            rw "S_NONE"
            dbget.Close() : response.end
        end if

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=EditPriceSelect&subcmd=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
   Case "cjmallmarginNotSaleItem" ''역마진 중에서 세일N인 것들 품절처리
        Set cjMall = new CCjmall
    	cjMall.FCurrPage					= 1
    	cjMall.FPageSize					= 10
    	cjMall.FRectExtNotReg				= "D"
    	cjMall.FRectSellYn					= "A"
    	cjMall.FRectSailYn					= "N"
    	cjMall.FRectCjshowminusmagin		= "on"
		cjMall.FRectExtSellYn               = "Y"
    	cjMall.GetCjmallRegedItemList

        IF (cjMall.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set cjMall= Nothing
            response.end
        ENd IF

        for i=0 to cjMall.FResultCount - 1
            itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
        next
        set cjMall= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=EditSellYn&subcmd=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "cjmallEditItemOptChk" '' cj 상품수정(옵션관련)
        sqlStr = " select top 10 ro.itemid"
        sqlStr = sqlStr & " from  db_item.dbo.tbl_cjmall_regItem r"
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_outmall_regedoption ro"
        sqlStr = sqlStr & " 	on ro.itemid=r.itemid"
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item_option o"
        sqlStr = sqlStr & " 	on ro.itemid=o.itemid"
        sqlStr = sqlStr & " 	and ro.itemoption=o.itemoption"
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i"
        sqlStr = sqlStr & " 	on r.itemid=i.itemid"
        sqlStr = sqlStr & " where ro.mallid='cjmall'"
        sqlStr = sqlStr & " and i.optionCnt>0"
        sqlStr = sqlStr & " and r.cjmallsellyn='Y'"
        sqlStr = sqlStr & " and r.cjmallStatCd>3"   ''2013/06/20 추가
        ''sqlStr = sqlStr & " and o.optsellyn='N'"
        sqlStr = sqlStr & " and (o.optsellyn='N' or (o.optsellyn='Y' and o.optlimityn='Y' and (o.optlimitno-o.optlimitsold<1)))"
        sqlStr = sqlStr & " and ro.outmallsellyn='Y'"
        sqlStr = sqlStr & " group by ro.itemid,r.cjmallLastUpdate,lastStatCheckDate,i.lastupdate"
        sqlStr = sqlStr & " order by r.lastStatCheckDate"

        rsget.Open sqlStr,dbget,1
        if not rsget.Eof then
            ArrRows = rsget.getRows()
        end if
        rsget.close

        itemidArr = ""
        if isArray(ArrRows) then
            For i =0 To UBound(ArrRows,2)
                itemidArr = itemidArr + CStr(ArrRows(0,i)) + ","
            Next
        else
            rw "S_NONE"
            dbget.Close() : response.end
        end if

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr
        paramData = "redSsnKey=system&cmdparam=EditSelect2&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal

    ''---------------------------------------------------------------------------------------
    Case "cjmallEditItem" '' cj 상품수정
        Set cjMall = new CCjmall
    	cjMall.FCurrPage					= 1
    	cjMall.FPageSize					= 8
    	cjMall.FRectExtNotReg				= "R"
    	cjMall.FRectMatchCate				= "Y"
    	cjMall.FRectPrdDivMatch				= "Y"
    	cjMall.FRectSellYn					= ""
    	cjMall.FRectLimitYn					= param3
    	cjMall.FRectonlyValidMargin 		= "Y"
    	cjMall.FRectMinusMargin 			= ""
    	cjMall.FRectFailCntExists			= ""
    	cjMall.FRectCjSell10x10Soldout      = ""
    	cjMall.FRectExpensive10x10          = ""
        cjMall.FRectOrdType = param2

    	cjMall.GetCjmallRegedItemList

        IF (cjMall.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set cjMall= Nothing
            response.end
        ENd IF

        for i=0 to cjMall.FResultCount - 1
            itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
        next
        set cjMall= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=EditSelect2&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCjMallReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal

    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
