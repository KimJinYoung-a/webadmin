<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/admin/etc/kaffa/kaffaCls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->

<%
dim webImgUrl : webImgUrl		= "http://webimage.10x10.co.kr"				'웹이미지

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","110.93.128.114","110.93.128.113")
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
Dim oitem, itemidArr

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
                    IF (LCASE(param1)="lottecom") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2LotteDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)
                        'response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actLotteSongjangInputProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="lotteimall") then
                        paramData = "redSsnKey=system&cmdparam=songjangip&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&sendQnt="&sendReqCntArr(i)&"&sendDate="&replace(Left(beasongdateArr(i),10),"-","")&"&outmallGoodsID="&outmallGoodsIDArr(i)&"&hdc_cd="&TenDlvCode2LotteiMallDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)
                        'rw paramData
                        'response.end
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotteimall/actLotteiMallReq.asp",paramData)
                             rw retVal
                        end if
                    ELSEIF (LCASE(param1)="interpark") then
                        ''IF (TenDlvCode2InterParkDlvCode(songjangDivArr(i))<>"") then
                            paramData = "redSsnKey=system&ordclmNo="&OutMallOrderSerialArr(i)&"&ordSeq="&OrgDetailKeyArr(i)
                            paramData = paramData&"&delvDt="&replace(Left(beasongdateArr(i),10),"-","")&"&delvEntrNo="&TenDlvCode2InterParkDlvCode(songjangDivArr(i))&"&invoNo="&songjangNoArr(i)
                            paramData = paramData&"&optPrdTp=01&optOrdSeqList="&OrgDetailKeyArr(i)
                            'rw paramData
                            'response.end
                            if (application("Svr_Info")<>"Dev") then
                                 retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/interparkXML/actInterparkSongjangInputProc.asp",paramData)
                                 rw retVal
                            end if
                        ''ELSe
                        ''    rw "택배사코드 매칭오류 ["&songjangDivArr(i)&"]"
                        ''END If
                    ELSEIF (LCASE(param1)="cjmall") then


                        ''var params = "ten_ord_no="+tenorderserial+"&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
                        ''var popwin=window.open('/admin/etc/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');

                        paramData = "redSsnKey=system&ten_ord_no="&TenOrderserial(i)&"&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2cjMallDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)
                        rw paramData

                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/cjmall/actCJmallSongjangInputProc.asp",paramData)
                             rw retVal
                        end if
                    end if
                end if
            next
        end if
    Case "cn10x10CheckRDItem" '' cn10x10 판매상태 확인 Batch
        paramData = "redSsnKey=system&cmdparam=CheckItemStatAuto&subcmd="&param2
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/kaffa/actKaffaReq.asp",paramData)

        response.write retVal&VbCRLF

    Case "cn10x10ExpireItem"   '' 품절 처리 요망 (업배 및 무게차이 큰것)
        Set oitem = new cKaffaItem
        oitem.FPageSize						= 10
    	oitem.FCurrPage         			= 1
    	oitem.FRectKaffaUseYN				= "Y"
    	oitem.FRectItemID					= ""
    	oitem.FRectSellYn					= ""
    	oitem.FRectLimitYn					= ""
    	oitem.FRectonlyValidMargin 			= ""
    	oitem.FRectFailCntExists			= ""
    	oitem.FRectoptAddprcExists			= ""
    	oitem.FRectoptAddprcExistsExcept	= ""
    	oitem.FRectoptExists				= ""
    	oitem.FRectKAFFASell10x10Soldout   	= "on"
    	oitem.FRectExpensive10x10          	= ""
    	oitem.FRectdiffPrc 					= ""
    	oitem.FRectExtSellYn  				= "Y"

    	oitem.getKaffaReqExpireItemList

        IF (oitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oitem= Nothing
            response.end
        ENd IF

        for i=0 to oitem.FResultCount - 1
            itemidArr = itemidArr & oitem.FItemList(i).FItemID &","
        next
        set oitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=product_sale&subcmd=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/kaffa/actKaffaReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal

    Case "cn10x10SoldOutItem" '' 품절처리 상품. (10x10 품절, cj판매중)
        Set oitem = new cKaffaItem
        oitem.FPageSize						= 10
    	oitem.FCurrPage         			= 1
    	oitem.FRectKaffaUseYN				= ""
    	oitem.FRectItemID					= ""
    	oitem.FRectSellYn					= ""
    	oitem.FRectLimitYn					= ""
    	oitem.FRectonlyValidMargin 			= ""
    	oitem.FRectFailCntExists			= ""
    	oitem.FRectoptAddprcExists			= ""
    	oitem.FRectoptAddprcExistsExcept	= ""
    	oitem.FRectoptExists				= ""
    	oitem.FRectKAFFASell10x10Soldout   	= "on"
    	oitem.FRectExpensive10x10          	= ""
    	oitem.FRectdiffPrc 					= ""
    	oitem.FRectExtSellYn  				= ""

    	oitem.GetAllKaffaItemList_USESCM

        IF (oitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oitem= Nothing
            response.end
        ENd IF

        for i=0 to oitem.FResultCount - 1
            itemidArr = itemidArr & oitem.FItemList(i).FItemID &","
        next
        set oitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=product_sale&subcmd=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/kaffa/actKaffaReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "cn10x10expensive10x10" '' 가격비싼상품 수정

        Set oitem = new cKaffaItem
        oitem.FPageSize						= 10
    	oitem.FCurrPage         			= 1
    	oitem.FRectKaffaUseYN				= "y"
    	oitem.FRectItemID					= ""
    	oitem.FRectSellYn					= "Y"
    	oitem.FRectLimitYn					= ""
    	oitem.FRectonlyValidMargin 			= ""
    	oitem.FRectFailCntExists			= ""
    	oitem.FRectoptAddprcExists			= ""
    	oitem.FRectoptAddprcExistsExcept	= ""
    	oitem.FRectoptExists				= ""
    	oitem.FRectKAFFASell10x10Soldout   	= ""
    	oitem.FRectExpensive10x10          	= "on"
    	oitem.FRectdiffPrc 					= ""
    	oitem.FRectExtSellYn  				= "Y"

    	oitem.GetAllKaffaItemList_USESCM

        IF (oitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oitem= Nothing
            response.end
        ENd IF

        for i=0 to oitem.FResultCount - 1
            itemidArr = itemidArr & oitem.FItemList(i).FItemID &","
        next
        set oitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=productstock&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/kaffa/actKaffaReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "cn10x10priceDiff" '' 가격 상이 수정

        Set oitem = new cKaffaItem
        oitem.FPageSize						= 10
    	oitem.FCurrPage         			= 1
    	oitem.FRectKaffaUseYN				= "y"
    	oitem.FRectItemID					= ""
    	oitem.FRectSellYn					= "Y"
    	oitem.FRectLimitYn					= ""
    	oitem.FRectonlyValidMargin 			= "Y"
    	oitem.FRectFailCntExists			= ""
    	oitem.FRectoptAddprcExists			= ""
    	oitem.FRectoptAddprcExistsExcept	= ""
    	oitem.FRectoptExists				= ""
    	oitem.FRectKAFFASell10x10Soldout   	= ""
    	oitem.FRectExpensive10x10          	= ""
    	oitem.FRectdiffPrc 					= "on"
    	oitem.FRectExtSellYn  				= "Y"

    	oitem.GetAllKaffaItemList_USESCM

        IF (oitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oitem= Nothing
            response.end
        ENd IF

        for i=0 to oitem.FResultCount - 1
            itemidArr = itemidArr & oitem.FItemList(i).FItemID &","
        next
        set oitem= Nothing


        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=productstock&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/kaffa/actKaffaReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    ''---------------------------------------------------------------------------------------
    Case "cn10x10DiffMultiPrc" '' 가격 상이 수정(해외기준가<>kaffa)

        Set oitem = new cKaffaItem
        oitem.FPageSize						= 10
    	oitem.FCurrPage         			= 1
    	oitem.FRectKaffaUseYN				= "y"
    	oitem.FRectItemID					= ""
    	oitem.FRectSellYn					= "Y"
    	oitem.FRectLimitYn					= ""
    	oitem.FRectonlyValidMargin 			= "Y"
    	oitem.FRectFailCntExists			= ""
    	oitem.FRectoptAddprcExists			= ""
    	oitem.FRectoptAddprcExistsExcept	= ""
    	oitem.FRectoptExists				= ""
    	oitem.FRectKAFFASell10x10Soldout   	= ""
    	oitem.FRectExpensive10x10          	= ""
    	oitem.FRectdiffPrc 					= ""
    	oitem.FRectdiffMultiPrc             = "on"
    	oitem.FRectExtSellYn  				= "Y"
        oitem.FRectextdispyn                = "Y"
        oitem.FRectMWDiv                    = "MW"
    	oitem.GetAllKaffaItemList_USESCM

        IF (oitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oitem= Nothing
            response.end
        ENd IF

        for i=0 to oitem.FResultCount - 1
            itemidArr = itemidArr & oitem.FItemList(i).FItemID &","
        next
        set oitem= Nothing


        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr

        paramData = "redSsnKey=system&cmdparam=productstock&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/kaffa/actKaffaReq.asp",paramData)

        response.Write "<br>"&retVal
    ''---------------------------------------------------------------------------------------

    Case "cn10x10EditItem" ''  상품수정
        Set oitem = new cKaffaItem
        oitem.FPageSize						= 10
    	oitem.FCurrPage         			= 1
    	oitem.FRectKaffaUseYN				= "m"                   ''수정요망으로  R
    	oitem.FRectItemID					= ""
    	oitem.FRectSellYn					= "Y"                            ''판매상품만
    	oitem.FRectLimitYn					= CHKIIF(param2="0","Y","")      ''한정상품먼저
    	oitem.FRectonlyValidMargin 			= "Y"                            ''마진ok
   	    oitem.FRectFailCntExists			= ""
    	oitem.FRectoptAddprcExists			= ""
    	oitem.FRectoptAddprcExistsExcept	= ""
    	oitem.FRectoptExists				= ""
    	oitem.FRectKAFFASell10x10Soldout   	= ""
    	oitem.FRectExpensive10x10          	= ""
    	oitem.FRectdiffPrc 					= ""
    	oitem.FRectExtSellYn  				= ""
        oitem.FRectMWDiv                    = "MW"

''       한정5개인경우 관련
'        if (param2="1") then                                            ''등록될때 기본적으로 판매중지로 되는듯.
'            oitem.FRectKaffaUseYN				= "y"                   ''수정요망으로  R
'            oitem.FRectLimitYn                  = ""
'            oitem.FRectExtSellYn  				= "N"
'        end if

    	oitem.GetAllKaffaItemList_USESCM

        IF (oitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oitem= Nothing
            response.end
        ENd IF

        for i=0 to oitem.FResultCount - 1
            itemidArr = itemidArr & oitem.FItemList(i).FItemID &","
        next
        set oitem= Nothing


        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
rw itemidArr
        paramData = "redSsnKey=system&cmdparam=productstock&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/kaffa/actKaffaReq.asp",paramData)

        'response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal

    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
