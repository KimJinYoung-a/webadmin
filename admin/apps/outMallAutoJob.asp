<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->

<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72", "61.252.133.67", "61.252.133.70")
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
    'rw ref
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
dim oLotteitem, itemidArr

select Case act

    Case "outmallSongJangIp" ''제휴사 송장입력  갯수 수정 40=>6*N
    'response.end

        sqlStr = "select top 6 T.orderserial, T.OutMallOrderSerial"
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

        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
     
        cnt = rsget.RecordCount
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
                          '   retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actLotteSongjangInputProc.asp",paramData)
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

                    ELSEIF (LCASE(param1)="gseshop") then
						paramData = "redSsnKey=system&ordclmNo="&OutMallOrderSerialArr(i)&"&ordSeq="&OrgDetailKeyArr(i)
						paramData = paramData&"&delvDt="&replace(Left(beasongdateArr(i),10),"-","")&"&delvEntrNo="&TenDlvCode2GSShopDlvCode(songjangDivArr(i))&"&invoNo="&songjangNoArr(i)
						'rw paramData
						'response.end
						if (application("Svr_Info")<>"Dev") then
							 retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actGSShopSongjangInputProc.asp",paramData)
							 rw retVal
						end if

                    end if
                end if
            next
        end if

    Case "lotteCheckRDItem" '' 임시상품 전시등록 일괄확인.
        paramData = "redSsnKey=system&param2="&param2
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actLotteCheckRDItem.asp",paramData)

        response.write retVal&VbCRLF
    CASE "lotteCsQ" ''롯데 닷컴 CS 쿼리
        ''http://wapi.10x10.co.kr/outmall/proc/xSiteCsOrder_Process.asp?mode=getxsitecslist&sellsite=lotteCom 여기 바로 호출
        ''paramData = "redSsnKey=system&mode=getxsitecslist&sellsite=lotteCom"
        ''retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/orderinput/xSiteCSOrder_Process.asp?",paramData)

        response.write retVal&VbCRLF
    Case "lotteRegItem" '' 롯데 상품등록
        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 5
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "W" ''등록예정
        oLotteitem.FRectMatchCate  = "Y" ''카테고리매칭
        oLotteitem.FRectSellYn  = "Y" ''판매중인상품
        'oLotteitem.FRectLotteYes10x10No = "on" ''LotteYes10x10No
        'oLotteitem.FRectMinusMigin = showminusmagin
        'oLotteitem.FRectExpensive10x10 = expensive10x10
        'oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
        'oLotteitem.FRectOnreginotmapping = onreginotmapping
        'oLotteitem.FRectdiffPrc = diffPrc
        oLotteitem.FRectonlyValidMargin = "on"      '' 마진 되는거만.
        oLotteitem.FRectOrdType = "B"               '' 베스트 셀러순
        oLotteitem.FRectoptAddprcExistsExcept= "on" '' 옵션 추가금액 제외.
        oLotteitem.FRectFailCntOverExcept="3"       '' 3회 이상 실패내역 제낌.
        oLotteitem.FRectLimitOver="5"               ''추가조건 필요 : 옵션 한정<5 미만 제외.
        oLotteitem.GetLotteRegedItemList

        IF (oLotteitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oLotteitem= Nothing
            response.end
        ENd IF

        for i=0 to oLotteitem.FResultCount - 1
            itemidArr = itemidArr & oLotteitem.FItemList(i).FItemID &","
        next
        set oLotteitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        paramData = "redSsnKey=system&mode=RegSelect&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "lotteEditItem" '' 롯데 상품수정
        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 10
        oLotteitem.FCurrPage       = param2
        oLotteitem.FRectLotteNotReg  = "R" ''수정요망
        oLotteitem.FRectMatchCate  = "Y" ''카테고리매칭
        oLotteitem.FRectSellYn  = "Y" ''판매중인상품
        'oLotteitem.FRectLotteYes10x10No = "on" ''LotteYes10x10No
        'oLotteitem.FRectMinusMigin = showminusmagin
        'oLotteitem.FRectExpensive10x10 = expensive10x10
        'oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
        'oLotteitem.FRectOnreginotmapping = onreginotmapping
        'oLotteitem.FRectdiffPrc = diffPrc
        oLotteitem.FRectOrdType = param3                '' 베스트 셀러순"B"

        if param4<>"" then                                  ''한정판매
            oLotteitem.FRectLimitYn="Y"
        else
            oLotteitem.FRectonlyValidMargin = "on"          '' 마진 되는거만.           :: 차후 이조건 품절처리
            '''oLotteitem.FRectoptAddprcExistsExcept= "on"     '' 옵션 추가금액 제외.      :: 차후 이조건 품절처리 (주석제거 2013/01/21)
        end if

        oLotteitem.FRectFailCntOverExcept="5"       '' 3회 이상 실패내역 제낌.(3=>5 2013/01/21)

        oLotteitem.GetLotteRegedItemList

        IF (oLotteitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oLotteitem= Nothing
            response.end
        ENd IF

        for i=0 to oLotteitem.FResultCount - 1
            itemidArr = itemidArr & oLotteitem.FItemList(i).FItemID &","
        next
        set oLotteitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        paramData = "redSsnKey=system&mode=EditSelect&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "lotteExpireItem", "lotteKillItem" '' 롯데 품절처리(판매금지)
        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 10
        oLotteitem.FCurrPage       = param2
        'oLotteitem.FRectLotteNotReg  = "R" ''수정요망
        'oLotteitem.FRectMatchCate  = "Y" ''카테고리매칭
        'oLotteitem.FRectSellYn  = "Y" ''판매중인상품
        'oLotteitem.FRectLotteYes10x10No = "on" ''LotteYes10x10No
        'oLotteitem.FRectMinusMigin = showminusmagin
        'oLotteitem.FRectExpensive10x10 = expensive10x10
        'oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
        'oLotteitem.FRectOnreginotmapping = onreginotmapping
        'oLotteitem.FRectdiffPrc = diffPrc
        'oLotteitem.FRectonlyValidMargin = "on"      '' 마진 되는거만.
        'oLotteitem.FRectOrdType = param3               '' 베스트 셀러순"B"
        'oLotteitem.FRectoptAddprcExistsExcept= "on" '' 옵션 추가금액 제외.
        if (act="lotteKillItem") then
            oLotteitem.FRectExtSellYn  = "YN"
            oLotteitem.FRectOnlyNotUsingCheck ="on"
        else
            oLotteitem.FRectExtSellYn  = "Y"            '' 판매중인상품
        end if
        oLotteitem.FRectFailCntOverExcept="3"       '' 3회 이상 실패내역 제낌.

        oLotteitem.getLottereqExpireItemList

        IF (oLotteitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oLotteitem= Nothing
            response.end
        ENd IF

        for i=0 to oLotteitem.FResultCount - 1
            itemidArr = itemidArr & oLotteitem.FItemList(i).FItemID &","
        next
        set oLotteitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        ''paramData = "redSsnKey=system&mode=EditSellYn&chgSellYn=X&cksel="&itemidArr
        if (act="lotteKillItem") then
            paramData = "redSsnKey=system&mode=EditSellYn&chgSellYn=X&cksel="&itemidArr
        else
            paramData = "redSsnKey=system&mode=EditSellYn&chgSellYn=N&cksel="&itemidArr
        end if
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "lotteSoldOutItem" '' 품절처리 상품.

        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 20
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D" ''LotteNotReg
        ''oLotteitem.FRectMatchCate  = "Y" ''MatchCate              ''매핑 상관없이 품절처리
        oLotteitem.FRectSellYn  = "A" ''sellyn
        oLotteitem.FRectLotteYes10x10No = "on" ''LotteYes10x10No
        'oLotteitem.FRectMinusMigin = showminusmagin
        'oLotteitem.FRectExpensive10x10 = expensive10x10
        'oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
        'oLotteitem.FRectOnreginotmapping = onreginotmapping
        'oLotteitem.FRectdiffPrc = diffPrc
        'oLotteitem.FRectonlyValidMargin = onlyValidMargin

        oLotteitem.GetLotteRegedItemList

        IF (oLotteitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oLotteitem= Nothing
            response.end
        ENd IF

        for i=0 to oLotteitem.FResultCount - 1
            itemidArr = itemidArr & oLotteitem.FItemList(i).FItemID &","
        next
        set oLotteitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        paramData = "redSsnKey=system&mode=EditSellYn&chgSellYn=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "lotteinfoDivNoItem" '' 품목정보 없는상품 품절처리

        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 10
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D" ''LotteNotReg            '' 차후 D로 변경
        ''oLotteitem.FRectMatchCate  = "Y" ''MatchCate              ''매핑 상관없이 품절처리
        oLotteitem.FRectSellYn  = "A" ''sellyn
        oLotteitem.FRectExtSellYn  = "Y"
        oLotteitem.FRectInfoDivYn = "N"
 		oLotteitem.FRectFailCntOverExcept="5"       '' 5회 이상 실패내역 제낌
        oLotteitem.GetLotteRegedItemList

        IF (oLotteitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oLotteitem= Nothing
            response.end
        ENd IF

        for i=0 to oLotteitem.FResultCount - 1
            itemidArr = itemidArr & oLotteitem.FItemList(i).FItemID &","
        next
        set oLotteitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        paramData = "redSsnKey=system&mode=EditSellYn&chgSellYn=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "lotteoptAddPrcSoldout" '' 옵션추가금액 존재상품 품절처리
        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 10
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D" ''LotteNotReg            '' 전시
        oLotteitem.FRectoptAddprcExists= "on"                       '' 옵션추가금액존재
        oLotteitem.FRectoptAddPrcRegTypeNone = "on"                 ''옵션추가금액상품 미설정 상품.
        oLotteitem.FRectSellYn  = "Y"                               ''sellyn
        oLotteitem.FRectExtSellYn  = "Y"
        ''oLotteitem.FRectInfoDivYn = "N"

        oLotteitem.GetLotteRegedItemList

        IF (oLotteitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oLotteitem= Nothing
            response.end
        ENd IF

        for i=0 to oLotteitem.FResultCount - 1
            itemidArr = itemidArr & oLotteitem.FItemList(i).FItemID &","
        next
        set oLotteitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        paramData = "redSsnKey=system&mode=EditSellYn&chgSellYn=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
	Case "lottemarginItem"
		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 10 r.itemid, (i.buycash)/R.lotteprice*100 as margin, i.buycash, i.orgprice, i.sellcash, r.lotteprice, r.lottesellyn  "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_lotte_regItem as r "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on r.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE r.lottestatcd = '30' "
		sqlStr = sqlStr & " and r.lottegoodNo is Not Null "
		sqlStr = sqlStr & " and R.lotteprice<>0 "
		sqlStr = sqlStr & " and (i.buycash)/R.lotteprice*100>85 "
		sqlStr = sqlStr & " and r.lottesellyn = 'Y' "
		sqlStr = sqlStr & " and i.orgprice <> R.lotteprice "
		sqlStr = sqlStr & " and r.accFailCNT<5 "						'실패횟수 5회
		sqlStr = sqlStr & " ORDER BY (i.buycash)/R.lotteprice*100 "
        rsget.Open sqlStr,dbget,1
        cnt = rsget.RecordCount
		If (cnt < 1) Then
			response.Write "S_NONE"
			response.end
		Else
	        For i = 0 to cnt - 1
	            itemidArr = itemidArr & rsget("itemid") &","
				rsget.MoveNext
	        Next
		End If
        rsget.close

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
        paramData = "redSsnKey=system&mode=EditPriceSelect&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
	Case "lottemarginNotSaleItem"			'역마진 세일N인것 품절처리
		set oLotteitem = new CLotte
		oLotteitem.FPageSize       = 10
		oLotteitem.FCurrPage       = 1
		oLotteitem.FRectLotteNotReg  = "D" ''LotteNotReg            '' 전시
		oLotteitem.FRectMatchCate  = "Y"
		oLotteitem.FRectSellYn  = "A"                               ''sellyn
		oLotteitem.FRectSailYn  = "N"                               ''sailyn
		oLotteitem.FRectMinusMigin = "on"
		oLotteitem.GetLotteRegedItemList

		IF (oLotteitem.FResultCount<1) then
		    response.Write "S_NONE"
		    dbget.Close()
		    set oLotteitem= Nothing
		    response.end
		ENd IF

		for i=0 to oLotteitem.FResultCount - 1
		    itemidArr = itemidArr & oLotteitem.FItemList(i).FItemID &","
		next
		set oLotteitem= Nothing

		IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

		paramData = "redSsnKey=system&mode=EditSellYn&chgSellYn=N&cksel="&itemidArr
		retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

		response.Write "itemidArr="&itemidArr
		response.Write "<br>"&retVal
    Case "lotteexpensive10x10" '' 롯데가격<텐바이텐 가격

        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 20
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D"  ''LotteNotReg(등록)
        oLotteitem.FRectMatchCate  = "Y"    ''MatchCate
        oLotteitem.FRectSellYn  = "Y"       ''sellyn
        oLotteitem.FRectExtSellYn  = "Y"
        oLotteitem.FRectOrdType = "B"       ''베스트순
        'oLotteitem.FRectLotteYes10x10No = "on" ''LotteYes10x10No
        'oLotteitem.FRectMinusMigin = showminusmagin
        oLotteitem.FRectExpensive10x10 = "on"
        'oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
        'oLotteitem.FRectOnreginotmapping = onreginotmapping
        'oLotteitem.FRectdiffPrc = diffPrc
        'oLotteitem.FRectonlyValidMargin = onlyValidMargin
        oLotteitem.FRectFailCntOverExcept="3"       '' 3회 이상 실패내역 제낌.

        oLotteitem.GetLotteRegedItemList

        IF (oLotteitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oLotteitem= Nothing
            response.end
        ENd IF

        for i=0 to oLotteitem.FResultCount - 1
            itemidArr = itemidArr & oLotteitem.FItemList(i).FItemID &","
        next
        set oLotteitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        paramData = "redSsnKey=system&mode=EditSelect&cksel="&itemidArr                             ''가격및내용수정
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "lotteEditItemLastupdate" '' 상품최종수정일 기준 상품수정

        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 20
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D"  ''LotteNotReg(등록)
        oLotteitem.FRectMatchCate  = "Y"    ''MatchCate
        oLotteitem.FRectSellYn  = "Y"       ''sellyn
        oLotteitem.FRectExtSellYn  = "Y"
        oLotteitem.FRectOrdType = "LU"       ''아이템테이블 상품최근 수정일 기준
        oLotteitem.FRectFailCntOverExcept="3"       '' 3회 이상 실패내역 제낌.

        oLotteitem.GetLotteRegedItemList

        IF (oLotteitem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oLotteitem= Nothing
            response.end
        ENd IF

        for i=0 to oLotteitem.FResultCount - 1
            itemidArr = itemidArr & oLotteitem.FItemList(i).FItemID &","
        next
        set oLotteitem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        paramData = "redSsnKey=system&mode=EditSelect3&cksel="&itemidArr                             ''가격및내용수정
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    CASE "ChkItemStatauto"   ''판매상태,가격 Check : tbl_lotte_regItem 과 비교
        paramData = "redSsnKey=system&mode=CheckItemStatAuto"
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "<br>"&retVal
    CASE "CheckItemNmAuto"    ''상품명수정 2013/03/20
        paramData = "redSsnKey=system&mode=CheckItemNmAuto"
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "<br>"&retVal
    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
