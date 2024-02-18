<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.12","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","61.252.133.75","110.93.128.114","110.93.128.113","61.252.133.72")
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

    Case "outmallSongJangIp" ''���޻� �����Է�
        sqlStr = "select top 10 T.orderserial, T.OutMallOrderSerial"
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
        sqlStr = sqlStr & " and T.sendReqCnt<3"                     ''������ �õ� �ȵǵ���. �߰�.
        sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"      ''��ȯ ��� ����.
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

                        if InStr(OrgDetailKeyArr(i),"-")>0 then  '' �� API
                            paramData = "redSsnKey=system&cmdparam=songjangip&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&sendQnt="&sendReqCntArr(i)&"&sendDate="&replace(Left(beasongdateArr(i),10),"-","")&"&outmallGoodsID="&outmallGoodsIDArr(i)&"&hdc_cd="&TenDlvCode2LotteiMallNewDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)
                            if (application("Svr_Info")<>"Dev") then
                                 retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotteimall/actLotteiMallReq.asp",paramData)
                                 rw retVal
                            end if
                        else
                            paramData = "redSsnKey=system&mode=sendsongjang&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&sendQnt="&sendReqCntArr(i)&"&sendDate="&replace(Left(beasongdateArr(i),10),"-","")&"&outmallGoodsID="&outmallGoodsIDArr(i)&"&hdc_cd="&TenDlvCode2LotteiMallNewDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)

                            if (application("Svr_Info")<>"Dev") then
                                 retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/orderInput/xSiteCSOrder_lotteimall_Process.asp",paramData)
                                 rw retVal&"-"&OutMallOrderSerialArr(i)
                            end if
                        end if
                    'rw paramData
                    end if
                end if
            next
        end if
    Case "imallSoldOutItem" '' ǰ��ó�� ��ǰ.

        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 10
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "D" ''LotteNotReg
        oiMallItem.FRectMatchCate    = "" ''MatchCate  Y
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

        paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr

        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "iMallexpensive10x10" '' �Ե�����<�ٹ����� ����

        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 5
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "D" ''LotteNotReg
        oiMallItem.FRectMatchCate    = "Y" ''MatchCate
        oiMallItem.FRectSellYn       = "A" ''sellyn
        oiMallItem.FRectExpensive10x10 = "on"
        oiMallItem.FRectOrdType = "B"
        ''oiMallItem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
        oiMallItem.FRectExtSellYn	= "Y" ''���޻� �Ǹ����ΰŸ�.
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

        paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr

        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "iMallEditItem" '' �Ե�iMall ��ǰ����

        set oiMallItem = new CLotteiMall
        if (param2="0") then
            oiMallItem.FPageSize       = 10
            oiMallItem.FCurrPage       = 1
            oiMallItem.FRectLotteNotReg  = "R" ''�������
            oiMallItem.FRectMatchCate    = "" ''MatchCate Y ::��Ī�� ������� ����
            oiMallItem.FRectSellYn       = "Y" ''sellyn
            oiMallItem.FRectOrdType      = "BM" ''"B" ''"BM"      ''����.
            oiMallItem.FRectLimitYn      = "Y"
            oiMallItem.FRectoptExists    = "Y"
            oiMallItem.FRectonlyValidMargin = "on"
            oiMallItem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
       	elseif (param2="1") then
            oiMallItem.FPageSize       = 10
            oiMallItem.FCurrPage       = 1
            oiMallItem.FRectLotteNotReg  = "R" ''�������
            oiMallItem.FRectMatchCate    = "" ''MatchCate Y ::��Ī�� ������� ����
            oiMallItem.FRectSellYn       = "Y" ''sellyn
            oiMallItem.FRectOrdType      = "BM" ''"B" ''"BM"      ''����.
            oiMallItem.FRectLimitYn      = "N"
            oiMallItem.FRectoptExists    = "Y"
            oiMallItem.FRectonlyValidMargin = "on"
            oiMallItem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
        elseif (param2="2") then
            oiMallItem.FPageSize       = 10
            oiMallItem.FCurrPage       = 1
            oiMallItem.FRectLotteNotReg  = "R" ''�������
            oiMallItem.FRectMatchCate    = "" ''MatchCate Y ::��Ī�� ������� ����
            oiMallItem.FRectSellYn       = "Y" ''sellyn
            oiMallItem.FRectOrdType      = "BM" ''"B" ''"BM"      ''����.
            oiMallItem.FRectLimitYn      = "Y"
            oiMallItem.FRectonlyValidMargin = "on"
            oiMallItem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
        elseif (param2="3") then					'2014-07-10 ������ ���� / ��ǰ�̰� �Ե�iMallǰ��&�ٹ������ǸŰ���, �����Ϲ�
            oiMallItem.FPageSize       = 10
            oiMallItem.FCurrPage       = 1
            oiMallItem.FRectLotteNotReg  = "D"
            oiMallItem.FRectMatchCate    = "Y" ''MatchCate Y
            oiMallItem.FRectoptnotExists = "on"
            oiMallItem.FRectLotteNo10x10Yes = "on"
            oiMallItem.FRectLimitYn      = "N"
            oiMallItem.FRectOrdType      = "BM"
            oiMallItem.FRectonlyValidMargin = "on"
            oiMallItem.FRect10000_Over = "on"
            oiMallItem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
        else
            oiMallItem.FPageSize       = 10
            oiMallItem.FCurrPage       = 1
            oiMallItem.FRectLotteNotReg  = "R" ''�������
            oiMallItem.FRectMatchCate    = "" ''MatchCate Y ::��Ī�� ������� ����
            oiMallItem.FRectSellYn       = "Y" ''sellyn
            oiMallItem.FRectOrdType      = "BM" ''"B" ''"BM"      ''����.
            oiMallItem.FRectonlyValidMargin = "on"
            oiMallItem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
        end if
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

        paramData = "redSsnKey=system&cmdparam=EditSelect&cksel="&itemidArr

        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "imallSoldOutItem2" '' ǰ��ó�� ��ǰ.(���޸� �����Ե�)

        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 20
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "" ''LotteNotReg
        oiMallItem.FRectMatchCate    = "" ''MatchCate
        oiMallItem.FRectSellYn       = "A" ''sellyn
        oiMallItem.FRectExtSellYn  = "Y"

        oiMallItem.getLtiMallreqExpireItemList ''getLottereqExpireItemList

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

        paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr

        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal


    Case "imalloptAddPrcSoldout" '' �ɼ��߰��ݾ� �����ǰ ǰ��ó��
        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 10
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "D" ''LotteNotReg            '' ����
        oiMallItem.FRectoptAddprcExists= "on"                       '' �ɼ��߰��ݾ�����
        oiMallItem.FRectoptAddPrcRegTypeNone = "on"                 ''�ɼ��߰��ݾ׻�ǰ �̼��� ��ǰ.
        oiMallItem.FRectSellYn  = "Y"                               ''sellyn
        oiMallItem.FRectExtSellYn  = "Y"
        ''oiMallItem.FRectInfoDivYn = "N"

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

        paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
	Case "iMallmarginNotSaleItem" ''������ ����N�ΰ� ǰ��ó��
		set oiMallItem = new CLotteiMall
		oiMallItem.FPageSize       = 10
		oiMallItem.FCurrPage       = 1
		oiMallItem.FRectLotteNotReg  = "D" ''LotteNotReg            '' ����
		oiMallItem.FRectSellYn       = "A" ''sellyn
		oiMallItem.FRectSailYn       = "N" ''sailyn
		oiMallItem.FRectMinusMigin	 = "on"
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

        paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)
        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "iMallmarginItem" ''������ ���ݼ���

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 10 r.itemid, (i.buycash)/R.ltimallprice*100 as margin, i.buycash, i.orgprice, i.sellcash, r.ltimallprice, r.ltimallsellyn  "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_ltimall_regitem as r "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on r.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE r.ltimallstatcd = '7' "
		sqlStr = sqlStr & " and r.ltimallgoodNo is Not Null "
		sqlStr = sqlStr & " and R.ltimallprice<>0 "
		sqlStr = sqlStr & " and (i.buycash)/R.ltimallprice*100>85.1 "
		sqlStr = sqlStr & " and r.ltimallsellyn = 'Y' "
		sqlStr = sqlStr & " and i.orgprice <> R.ltimallprice "
		sqlStr = sqlStr & " ORDER BY (i.buycash)/R.ltimallprice*100 "
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

        paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal

    Case "imallCheckItem" '' �ӽû�ǰ ���õ�� �ϰ�Ȯ��.
        set oiMallItem = new CLotteiMall
        oiMallItem.FPageSize       = 30 ''50
        oiMallItem.FCurrPage       = 1
        oiMallItem.FRectLotteNotReg  = "F" ''LotteNotReg
        ''oiMallItem.FRectMatchCate    = "Y" ''MatchCate
        if (param2="0") then
            oiMallItem.FRectSellYn       = "N" ''ǰ����ǰ ����Ȯ��
        end if
        ''oiMallItem.FRectonlyValidMargin = "on"
        oiMallItem.FRectOrdType = "LS"
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

        ''paramData = "redSsnKey=system&cmdparam=CheckItemStat&cksel="&itemidArr		''2013-07-24 ������ �ּ�// CheckItemStat�� ���û�ǰ �ǸŻ��� Ȯ���̶�...
       ''paramData = "redSsnKey=system&cmdparam=CheckItemStat&cksel="&itemidArr  ''2013/10/18 �ٽ� �ٲ� eastone �ǸŻ��� Ȯ�ο� ���Ȯ�γ��� �ִµ�(X)
        paramData = "redSsnKey=system&cmdparam=getconfirmList&cksel="&itemidArr			''2013-07-24 ������ �߰�// getconfirmList�� ���û�ǰ ���Ȯ��

        if (application("Svr_Info")<>"Dev") then
             retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)
             ''rw retVal
        end if

        rw "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "iMallregWait_OLD" '' ��Ͽ�����ǰ ���

'        response.Write "�Ե�iMall ��ǰ��� �Ͻ�����"
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
        oiMallItem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.

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
'                    retVal = retVal & "["&itemidArr(i)&"] ǰ�� ó�� ���� : "&ierrStr
'                else
'                    retVal = retVal & "["&itemidArr(i)&"] ǰ�� ó�� ����"
'                end if
            end if
        next

        response.Write "<br>"&retVal
    CASE "CheckItemStatAuto" ''�ǸŻ��� üũ
        paramData = "redSsnKey=system&cmdparam=CheckItemStatAuto"
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)
        response.Write "<br>"&retVal
    CASE "CheckItemNmAuto"    ''��ǰ����� 2013/07/02
        paramData = "redSsnKey=system&cmdparam=CheckItemNmAuto"
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/LtiMall/actLotteiMallReq.asp",paramData)

        response.Write "<br>"&retVal
    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
