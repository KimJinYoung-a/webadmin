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

    Case "outmallSongJangIp" ''���޻� �����Է�  ���� ���� 40=>6*N
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
		sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid"					'// ���� �ֹ��� ������ ���(����1��,�Ķ�1�� -> �Ķ�2��)
		sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption"
        sqlStr = sqlStr & " 	and D.currstate=7"
        sqlStr = sqlStr & " 	left join db_order.dbo.tbl_songjang_div V"
        sqlStr = sqlStr & " 	on D.songjangDiv=V.divcd"
        sqlStr = sqlStr & " where datediff(m,T.regdate,getdate())<7"    ''20130304 �߰�
        sqlStr = sqlStr & " and T.sellsite='"&param1&"'"
        sqlStr = sqlStr & " and T.OrgDetailKey is Not NULL"             ''������Ű �Է� �ֹ��Ǹ�..
        sqlStr = sqlStr & " and IsNULL(T.sendState,0)=0"
        sqlStr = sqlStr & " and T.sendReqCnt<3"                         ''������ �õ� �ȵǵ���. �߰�.
        sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"      ''��ȯ ��� ��ǰ ����.
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
                        ''    rw "�ù���ڵ� ��Ī���� ["&songjangDivArr(i)&"]"
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

    Case "lotteCheckRDItem" '' �ӽû�ǰ ���õ�� �ϰ�Ȯ��.
        paramData = "redSsnKey=system&param2="&param2
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actLotteCheckRDItem.asp",paramData)

        response.write retVal&VbCRLF
    CASE "lotteCsQ" ''�Ե� ���� CS ����
        ''http://wapi.10x10.co.kr/outmall/proc/xSiteCsOrder_Process.asp?mode=getxsitecslist&sellsite=lotteCom ���� �ٷ� ȣ��
        ''paramData = "redSsnKey=system&mode=getxsitecslist&sellsite=lotteCom"
        ''retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/orderinput/xSiteCSOrder_Process.asp?",paramData)

        response.write retVal&VbCRLF
    Case "lotteRegItem" '' �Ե� ��ǰ���
        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 5
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "W" ''��Ͽ���
        oLotteitem.FRectMatchCate  = "Y" ''ī�װ���Ī
        oLotteitem.FRectSellYn  = "Y" ''�Ǹ����λ�ǰ
        'oLotteitem.FRectLotteYes10x10No = "on" ''LotteYes10x10No
        'oLotteitem.FRectMinusMigin = showminusmagin
        'oLotteitem.FRectExpensive10x10 = expensive10x10
        'oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
        'oLotteitem.FRectOnreginotmapping = onreginotmapping
        'oLotteitem.FRectdiffPrc = diffPrc
        oLotteitem.FRectonlyValidMargin = "on"      '' ���� �Ǵ°Ÿ�.
        oLotteitem.FRectOrdType = "B"               '' ����Ʈ ������
        oLotteitem.FRectoptAddprcExistsExcept= "on" '' �ɼ� �߰��ݾ� ����.
        oLotteitem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
        oLotteitem.FRectLimitOver="5"               ''�߰����� �ʿ� : �ɼ� ����<5 �̸� ����.
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
    Case "lotteEditItem" '' �Ե� ��ǰ����
        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 10
        oLotteitem.FCurrPage       = param2
        oLotteitem.FRectLotteNotReg  = "R" ''�������
        oLotteitem.FRectMatchCate  = "Y" ''ī�װ���Ī
        oLotteitem.FRectSellYn  = "Y" ''�Ǹ����λ�ǰ
        'oLotteitem.FRectLotteYes10x10No = "on" ''LotteYes10x10No
        'oLotteitem.FRectMinusMigin = showminusmagin
        'oLotteitem.FRectExpensive10x10 = expensive10x10
        'oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
        'oLotteitem.FRectOnreginotmapping = onreginotmapping
        'oLotteitem.FRectdiffPrc = diffPrc
        oLotteitem.FRectOrdType = param3                '' ����Ʈ ������"B"

        if param4<>"" then                                  ''�����Ǹ�
            oLotteitem.FRectLimitYn="Y"
        else
            oLotteitem.FRectonlyValidMargin = "on"          '' ���� �Ǵ°Ÿ�.           :: ���� ������ ǰ��ó��
            '''oLotteitem.FRectoptAddprcExistsExcept= "on"     '' �ɼ� �߰��ݾ� ����.      :: ���� ������ ǰ��ó�� (�ּ����� 2013/01/21)
        end if

        oLotteitem.FRectFailCntOverExcept="5"       '' 3ȸ �̻� ���г��� ����.(3=>5 2013/01/21)

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
    Case "lotteExpireItem", "lotteKillItem" '' �Ե� ǰ��ó��(�Ǹű���)
        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 10
        oLotteitem.FCurrPage       = param2
        'oLotteitem.FRectLotteNotReg  = "R" ''�������
        'oLotteitem.FRectMatchCate  = "Y" ''ī�װ���Ī
        'oLotteitem.FRectSellYn  = "Y" ''�Ǹ����λ�ǰ
        'oLotteitem.FRectLotteYes10x10No = "on" ''LotteYes10x10No
        'oLotteitem.FRectMinusMigin = showminusmagin
        'oLotteitem.FRectExpensive10x10 = expensive10x10
        'oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
        'oLotteitem.FRectOnreginotmapping = onreginotmapping
        'oLotteitem.FRectdiffPrc = diffPrc
        'oLotteitem.FRectonlyValidMargin = "on"      '' ���� �Ǵ°Ÿ�.
        'oLotteitem.FRectOrdType = param3               '' ����Ʈ ������"B"
        'oLotteitem.FRectoptAddprcExistsExcept= "on" '' �ɼ� �߰��ݾ� ����.
        if (act="lotteKillItem") then
            oLotteitem.FRectExtSellYn  = "YN"
            oLotteitem.FRectOnlyNotUsingCheck ="on"
        else
            oLotteitem.FRectExtSellYn  = "Y"            '' �Ǹ����λ�ǰ
        end if
        oLotteitem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.

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
    Case "lotteSoldOutItem" '' ǰ��ó�� ��ǰ.

        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 20
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D" ''LotteNotReg
        ''oLotteitem.FRectMatchCate  = "Y" ''MatchCate              ''���� ������� ǰ��ó��
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
    Case "lotteinfoDivNoItem" '' ǰ������ ���»�ǰ ǰ��ó��

        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 10
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D" ''LotteNotReg            '' ���� D�� ����
        ''oLotteitem.FRectMatchCate  = "Y" ''MatchCate              ''���� ������� ǰ��ó��
        oLotteitem.FRectSellYn  = "A" ''sellyn
        oLotteitem.FRectExtSellYn  = "Y"
        oLotteitem.FRectInfoDivYn = "N"
 		oLotteitem.FRectFailCntOverExcept="5"       '' 5ȸ �̻� ���г��� ����
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
    Case "lotteoptAddPrcSoldout" '' �ɼ��߰��ݾ� �����ǰ ǰ��ó��
        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 10
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D" ''LotteNotReg            '' ����
        oLotteitem.FRectoptAddprcExists= "on"                       '' �ɼ��߰��ݾ�����
        oLotteitem.FRectoptAddPrcRegTypeNone = "on"                 ''�ɼ��߰��ݾ׻�ǰ �̼��� ��ǰ.
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
		sqlStr = sqlStr & " and r.accFailCNT<5 "						'����Ƚ�� 5ȸ
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
	Case "lottemarginNotSaleItem"			'������ ����N�ΰ� ǰ��ó��
		set oLotteitem = new CLotte
		oLotteitem.FPageSize       = 10
		oLotteitem.FCurrPage       = 1
		oLotteitem.FRectLotteNotReg  = "D" ''LotteNotReg            '' ����
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
    Case "lotteexpensive10x10" '' �Ե�����<�ٹ����� ����

        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 20
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D"  ''LotteNotReg(���)
        oLotteitem.FRectMatchCate  = "Y"    ''MatchCate
        oLotteitem.FRectSellYn  = "Y"       ''sellyn
        oLotteitem.FRectExtSellYn  = "Y"
        oLotteitem.FRectOrdType = "B"       ''����Ʈ��
        'oLotteitem.FRectLotteYes10x10No = "on" ''LotteYes10x10No
        'oLotteitem.FRectMinusMigin = showminusmagin
        oLotteitem.FRectExpensive10x10 = "on"
        'oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
        'oLotteitem.FRectOnreginotmapping = onreginotmapping
        'oLotteitem.FRectdiffPrc = diffPrc
        'oLotteitem.FRectonlyValidMargin = onlyValidMargin
        oLotteitem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.

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

        paramData = "redSsnKey=system&mode=EditSelect&cksel="&itemidArr                             ''���ݹ׳������
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "lotteEditItemLastupdate" '' ��ǰ���������� ���� ��ǰ����

        set oLotteitem = new CLotte
        oLotteitem.FPageSize       = 20
        oLotteitem.FCurrPage       = 1
        oLotteitem.FRectLotteNotReg  = "D"  ''LotteNotReg(���)
        oLotteitem.FRectMatchCate  = "Y"    ''MatchCate
        oLotteitem.FRectSellYn  = "Y"       ''sellyn
        oLotteitem.FRectExtSellYn  = "Y"
        oLotteitem.FRectOrdType = "LU"       ''���������̺� ��ǰ�ֱ� ������ ����
        oLotteitem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.

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

        paramData = "redSsnKey=system&mode=EditSelect3&cksel="&itemidArr                             ''���ݹ׳������
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    CASE "ChkItemStatauto"   ''�ǸŻ���,���� Check : tbl_lotte_regItem �� ��
        paramData = "redSsnKey=system&mode=CheckItemStatAuto"
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "<br>"&retVal
    CASE "CheckItemNmAuto"    ''��ǰ����� 2013/03/20
        paramData = "redSsnKey=system&mode=CheckItemNmAuto"
        retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/lotte/actRegLotteItem.asp",paramData)

        response.Write "<br>"&retVal
    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
