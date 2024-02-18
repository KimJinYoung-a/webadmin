<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/admin/etc/LotteiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/incLotteiMallFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71")
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

''response.end
dim act     : act = requestCheckVar(request("act"),32)
dim param1  : param1 = requestCheckVar(request("param1"),32)
dim param2  : param2 = requestCheckVar(request("param2"),32)
dim sqlStr, i, paramData, retVal
dim retCnt : retCnt = 0

Dim cnt
Dim OutMallOrderSerialArr
Dim OrgDetailKeyArr
Dim songjangDivArr
Dim songjangNoArr, sendReqCntArr, beasongdateArr, outmallGoodsIDArr
dim oLotteitem, itemidArr
dim oiMallItem, ierrStr
dim iSuccCNT, isValidDel

select Case act

    Case "outmallSongJangIp" ''제휴사 송장입력
        sqlStr = "select top 20 T.orderserial, T.OutMallOrderSerial"
        sqlStr = sqlStr & " ,T.OrgDetailKey, IsNULL(T.sendState,0) as sendState"
        sqlStr = sqlStr & " ,D.songjangDiv, D.songjangNo, D.itemNo, D.beasongdate, T.outMallGoodsNo"
        sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master M"
        sqlStr = sqlStr & " 	on T.orderserial=M.orderserial"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D"
        sqlStr = sqlStr & " 	on T.orderserial=D.orderserial"
        sqlStr = sqlStr & " 	and T.matchitemid=D.itemid"
        sqlStr = sqlStr & " 	and T.matchitemoption=D.itemoption"
        sqlStr = sqlStr & " 	and D.currstate=7"
        sqlStr = sqlStr & " 	left join db_order.dbo.tbl_songjang_div V"
        sqlStr = sqlStr & " 	on D.songjangDiv=V.divcd"
        sqlStr = sqlStr & " where T.sellsite='"&param1&"'"
        sqlStr = sqlStr & " and T.OrgDetailKey is Not NULL"
        sqlStr = sqlStr & " and IsNULL(T.sendState,0)=0"
        sqlStr = sqlStr & " and T.sendReqCnt<3"                     ''여러번 시도 안되도록. 추가.
        sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"      ''교환 취소 제외.
        sqlStr = sqlStr & " order by D.beasongdate desc"

        rsget.Open sqlStr,dbget,1
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
            response.Write "S_NONE"
            dbget.Close() : response.end
        else
            for i=LBound(OutMallOrderSerialArr) to UBound(OutMallOrderSerialArr)
                if (OutMallOrderSerialArr(i)<>"") then
                    IF (LCASE(param1)="lotteimall") then
                        paramData = "redSsnKey=system&cmdparam=songjangip&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&sendQnt="&sendReqCntArr(i)&"&sendDate="&replace(Left(beasongdateArr(i),10),"-","")&"&outmallGoodsID="&outmallGoodsIDArr(i)&"&hdc_cd="&TenDlvCode2LotteiMallDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)
                        ''rw paramData
                        ''response.end
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotteimall/actLotteiMallReq.asp",paramData)
                             rw retVal
                        end if

                    end if
                end if
            next
        end if
    Case "imallSoldOutItem" '' 품절처리 상품.

        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 10
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "D" ''LotteNotReg
        oiMallItem.FRectMatchCate    = "Y" ''MatchCate
        oiMallItem.FRectSellYn       = "A" ''sellyn
        oiMallItem.FRectLotteYes10x10No = "on"
        oiMallItem.FRectOrdType = "B"

        oiMallItem.getLTiMallRegedItemList

        IF (oiMallItem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oiMallItem= Nothing
            response.end
        ENd IF

        for i=0 to oiMallItem.FResultCount - 1
            itemidArr = itemidArr & oiMallItem.FItemList(i).FItemID &","
        next
        set oiMallItem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        rw "itemidArr="&itemidArr
        itemidArr = split(itemidArr,",")

        for i=LBound(itemidArr) to UBound(itemidArr)
            if (itemidArr(i)<>"") then
                ierrStr=""
                Call chkLotteiMallOneItem("EditSellYn", itemidArr(i), ierrStr, iSuccCNT, isValidDel)  ''2013/03/27 추가 아이몰 판매상태 check
                ierrStr=""
                CALL editSOLDOUTLotteiMallOneItem(itemidArr(i), ierrStr)
'                if (ierrStr<>"") then
'                    retVal = retVal & "["&itemidArr(i)&"] 품절 처리 실패 : "&ierrStr
'                else
'                    retVal = retVal & "["&itemidArr(i)&"] 품절 처리 성공"
'                end if
            end if
        next

        response.Write "<br>"&retVal
    Case "imallSoldOutItem2" '' 품절처리 상품.(제휴몰 사용안함등)

        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 20
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "" ''LotteNotReg
        oiMallItem.FRectMatchCate    = "" ''MatchCate
        oiMallItem.FRectSellYn       = "A" ''sellyn
        oiMallItem.FRectExtSellYn  = "Y"

        oiMallItem.getLottereqExpireItemList

        IF (oiMallItem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oiMallItem= Nothing
            response.end
        ENd IF

        for i=0 to oiMallItem.FResultCount - 1
            itemidArr = itemidArr & oiMallItem.FItemList(i).FItemID &","
        next
        set oiMallItem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        rw "itemidArr="&itemidArr
        itemidArr = split(itemidArr,",")

        for i=LBound(itemidArr) to UBound(itemidArr)
            if (itemidArr(i)<>"") then
                ierrStr=""
                Call chkLotteiMallOneItem("EditSellYn", itemidArr(i), ierrStr, iSuccCNT, isValidDel)  ''2013/03/27 추가 아이몰 판매상태 check
                ierrStr=""
                CALL editSOLDOUTLotteiMallOneItem(itemidArr(i), ierrStr)
'                if (ierrStr<>"") then
'                    retVal = retVal & "["&itemidArr(i)&"] 품절 처리 실패 : "&ierrStr
'                else
'                    retVal = retVal & "["&itemidArr(i)&"] 품절 처리 성공"
'                end if
            end if
        next

        response.Write "<br>"&retVal
    Case "iMallexpensive10x10" '' 롯데가격<텐바이텐 가격

        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 5
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "D" ''LotteNotReg
        oiMallItem.FRectMatchCate    = "Y" ''MatchCate
        oiMallItem.FRectSellYn       = "A" ''sellyn
        oiMallItem.FRectExpensive10x10 = "on"
        oiMallItem.FRectOrdType = "B"
        oiMallItem.FRectFailCntOverExcept="3"       '' 3회 이상 실패내역 제낌.

        oiMallItem.getLTiMallRegedItemList

        IF (oiMallItem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oiMallItem= Nothing
            response.end
        ENd IF

        for i=0 to oiMallItem.FResultCount - 1
            itemidArr = itemidArr & oiMallItem.FItemList(i).FItemID &","
        next
        set oiMallItem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
        rw "itemidArr="&itemidArr
        itemidArr = split(itemidArr,",")

        for i=LBound(itemidArr) to UBound(itemidArr)
            if (itemidArr(i)<>"") then
                ierrStr=""
                CALL editLotteiMallOneItem(itemidArr(i), ierrStr)

                Call chkLotteiMallOneItem("CheckItemStatAuto", itemidArr(i), ierrStr, iSuccCNT, isValidDel)  ''2013/03/28 추가 아이몰 판매상태 check
'                if (ierrStr<>"") then
'                    retVal = retVal & "["&itemidArr(i)&"] 품절 처리 실패 : "&ierrStr
'                else
'                    retVal = retVal & "["&itemidArr(i)&"] 품절 처리 성공"
'                end if
            end if
        next
        response.Write "<br>"&retVal
    Case "iMallEditItem" '' 롯데iMall 상품수정

        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 6
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "R" ''수정요망
        oiMallItem.FRectMatchCate    = "Y" ''MatchCate
        oiMallItem.FRectSellYn       = "Y" ''sellyn
        ''oiMallItem.FRectOrdType = "B" ''"BM"      ''느림.
        oiMallItem.FRectonlyValidMargin = "on"
        oiMallItem.FRectFailCntOverExcept="3"       '' 3회 이상 실패내역 제낌.

        oiMallItem.getLTiMallRegedItemList

        IF (oiMallItem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oiMallItem= Nothing
            response.end
        ENd IF

        for i=0 to oiMallItem.FResultCount - 1
            itemidArr = itemidArr & oiMallItem.FItemList(i).FItemID &","
        next
        set oiMallItem= Nothing

        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        rw "itemidArr="&itemidArr
        itemidArr = split(itemidArr,",")

        for i=LBound(itemidArr) to UBound(itemidArr)
            if (itemidArr(i)<>"") then
                ierrStr=""
                CALL editLotteiMallOneItem(itemidArr(i), ierrStr)

                Call chkLotteiMallOneItem("CheckItemStatAuto", itemidArr(i), ierrStr, iSuccCNT, isValidDel)  ''2013/03/28 추가 아이몰 판매상태 check
            end if
        next


        response.Write "<br>"&retVal
    Case "iMallregWait" '' 등록예정상품 등록

'        response.Write "롯데iMall 상품등록 일시중지"
'        dbget.Close()
'        response.end

        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 3
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "V" ''LotteNotReg  W=>V
        oiMallItem.FRectMatchCate    = "Y" ''MatchCate
        oiMallItem.FRectSellYn       = "Y" ''sellyn
        oiMallItem.FRectonlyValidMargin = "on"
        oiMallItem.FRectOrdType = "B"
        oiMallItem.FRectFailCntOverExcept="3"       '' 3회 이상 실패내역 제낌.

        oiMallItem.getLTiMallRegedItemList

        IF (oiMallItem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oiMallItem= Nothing
            response.end
        ENd IF

        for i=0 to oiMallItem.FResultCount - 1
            itemidArr = itemidArr & oiMallItem.FItemList(i).FItemID &","
        next
        set oiMallItem= Nothing


        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)

        rw "itemidArr="&itemidArr
        itemidArr = split(itemidArr,",")

        for i=LBound(itemidArr) to UBound(itemidArr)
            if (itemidArr(i)<>"") then
                ierrStr=""
                CALL regLotteiMallOneItem(itemidArr(i), ierrStr)
'                if (ierrStr<>"") then
'                    retVal = retVal & "["&itemidArr(i)&"] 품절 처리 실패 : "&ierrStr
'                else
'                    retVal = retVal & "["&itemidArr(i)&"] 품절 처리 성공"
'                end if
            end if
        next

        response.Write "<br>"&retVal
    Case "imallCheckItem" '' 임시상품 전시등록 일괄확인.
        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 50
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "F" ''LotteNotReg
        ''oiMallItem.FRectMatchCate    = "Y" ''MatchCate
        if (param2="0") then
            oiMallItem.FRectSellYn       = "N" ''품절상품 먼저확인
        end if
        ''oiMallItem.FRectonlyValidMargin = "on"
        ''oiMallItem.FRectOrdType = "B"

        oiMallItem.getLTiMallRegedItemList

        IF (oiMallItem.FResultCount<1) then
            response.Write "S_NONE"
            dbget.Close()
            set oiMallItem= Nothing
            response.end
        ENd IF

        for i=0 to oiMallItem.FResultCount - 1
            itemidArr = itemidArr & oiMallItem.FItemList(i).FItemID &","
        next
        set oiMallItem= Nothing


        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
        rw "itemidArr="&itemidArr
        paramData = "redSsnKey=system&cmdparam=getconfirmList&cksel="&itemidArr
        ''rw paramData
        ''response.end

        if (application("Svr_Info")<>"Dev") then
             retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotteimall/actLotteiMallReq.asp",paramData)
             ''rw retVal
        end if

        ''
        response.Write "<br>"&retVal
    CASE "CheckItemStatAuto" ''판매상태 체크
        paramData = "redSsnKey=system&cmdparam=CheckItemStatAuto"
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotteimall/actLotteiMallReq.asp",paramData)
        response.Write "<br>"&retVal
    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
