<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 900 %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
const C_APIKEY="$2a$08$ik.RQbF9tGCZibk7JnPueuG/8AIeuTDd.lgCP/fYuuZX7dnNuJRe6"
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),30)
Dim cksel : cksel = request("cksel")
Dim subcmd : subcmd = requestCheckVar(request("subcmd"),30)
Dim sqlStr, AssignedRow, iErrStr, retStr
Dim SuccCNT, FailCNT, alertMsg
Dim ArrRows, bufStr, buf, bufStr2
Dim i,iitemid, mode, kaffareqSellYN
Dim kaffaItemStat, kaffaItemSellYN, kaffaItemUpdate, kaffaOptionStock
Dim objXML,iMessage
Dim strParam, SaleStatCd, GoodsViewCount
Dim iRbody, strObj, target_list, itemdata

Function GetKaffaLmtQty(limityn, limitno, limitsold)
	CONST CLIMIT_SOLDOUT_NO = 5

	If (limityn="Y") then
		If (limitno - limitsold) < CLIMIT_SOLDOUT_NO Then
			GetKaffaLmtQty = 0
		Else
			GetKaffaLmtQty = limitno - limitsold - CLIMIT_SOLDOUT_NO
		End If
	Else
		GetKaffaLmtQty = 999
	End If
End Function

Function getOptionLimitNo(optsellyn, optlimityn, optlimitno, optlimitsold, optisusing)
	CONST CLIMIT_SOLDOUT_NO = 5

	If (optsellyn="N") or (optisusing="N") or ((optlimityn="Y") and (optlimitno - optlimitsold < CLIMIT_SOLDOUT_NO)) Then
		getOptionLimitNo = 0
	Else
		If (optlimityn = "Y") Then
			If (optlimitno - optlimitsold < CLIMIT_SOLDOUT_NO) Then
				getOptionLimitNo = 0
			Else
				getOptionLimitNo = optlimitno - optlimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			getOptionLimitNo = 999
		End if
	End If
End Function

if (cmdparam="CheckItemStatAuto") then ''�ǸŻ��� Ȯ��
    SuccCNT = 0
    FailCNT = 0

    if (subcmd="0") then                            ''��Ȯ�ΰ��� ǰ���̳� ���ݻ��� ���� ����
        sqlStr = "select top 20 r.itemid "
        sqlStr = sqlStr & "	from db_item.dbo.tbl_kaffa_reg_item r"
        sqlStr = sqlStr & "	Join db_item.dbo.tbl_item i"
	    sqlStr = sqlStr & "	on r.itemid=i.itemid"
        sqlStr = sqlStr & "	where r.kaffaGoodno is NULL"
        sqlStr = sqlStr & "	and (i.sellyn<>'Y' or i.orgprice>isNULL(r.kaffaprice,0))"
        sqlStr = sqlStr & "	order by r.lastStatCheckDate, i.sellyn desc, (CASE WHEN r.kaffasellyn='X' THEN '0' ELSE r.kaffasellyn END),  r.LastUpdate , r.itemid desc"
    else
        sqlStr = "select top 20 r.itemid "
        sqlStr = sqlStr & "	from db_item.dbo.tbl_kaffa_reg_item r"
        sqlStr = sqlStr & "	where 1=1"
        sqlStr = sqlStr & "	order by r.lastStatCheckDate, (CASE WHEN r.kaffasellyn='X' THEN '0' ELSE r.kaffasellyn END),  r.LastUpdate , r.itemid desc"
    end if

    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
        ArrRows = rsget.getRows()
    end if
    rsget.close

    if isArray(ArrRows) then

        For i =0 To UBound(ArrRows,2)
            retStr = ""
            iitemid = CStr(ArrRows(0,i))

            if (iitemid<>"") then
               '' rw iitemid
                kaffaItemStat = CheckKaffaItemStat(iitemid,retStr)

                if (kaffaItemStat) then
                    SuccCNT = SuccCNT+1
                else
                    FailCNT = FailCNT+1
               end if

                bufStr = bufStr + retStr
            end if
        next
    end if
    rw bufStr
ELSEIF (cmdparam="CheckItemStat") then 			''���û�ǰ �ǸŻ��� Ȯ��
    SuccCNT = 0
    FailCNT = 0

    cksel = split(cksel,",")
    For i=0 To UBound(cksel)
        iitemid=Trim(cksel(i))
        retStr =""

        kaffaItemStat = CheckKaffaItemStat(iitemid,retStr)

        bufStr = bufStr + retStr
    next
    rw bufStr
ElseIf (cmdparam="product_sale") Then 			''�ǸŻ��� ���� ��(�ɼ� �ǸŻ��� ��������� Ȯ�� ��)
	cksel = split(cksel,",")
	For i = 0 To UBound(cksel)
		iitemid = Trim(cksel(i))
        retStr =""
        kaffaItemSellYN = CheckKaffaItemSellYN(iitemid, subcmd, retStr)
        bufStr = bufStr + retStr &"<br>"
	Next
	rw bufStr
ElseIf (cmdparam="set_product") Then 			''���û�ǰ ���� ����(2013-06-06 ���� ���� ��, ��ǰ����/����...�ɼ�����/���� ��������� Ȯ�� ��)
	cksel = split(cksel,",")
	For i = 0 To UBound(cksel)
		iitemid = Trim(cksel(i))
        retStr =""
        kaffaItemUpdate = CheckKaffaItemUpdate(iitemid, retStr)
        bufStr = bufStr + retStr &"<br>"
	Next
	rw bufStr
ElseIf (cmdparam="stock_fix") Then				''���û�ǰ��ǰ/�ǸŻ��� ����
	cksel = split(cksel,",")
	For i = 0 To UBound(cksel)
		iitemid = Trim(cksel(i))
        retStr =""
        kaffaOptionStock = CheckKaffaOptionStock(iitemid, retStr)
        bufStr = bufStr + retStr &"<br>"
	Next
	rw bufStr
ElseIf (cmdparam="productstock") Then		''���û�ǰ ���� ���� + ��ǰ ���� (����,�ǸŻ��� ���� �߰�)
	cksel = split(cksel, ",")
	For i=0 To UBound(cksel)
		iitemid = Trim(cksel(i))

		if IsRequreUpdateSellStatUpdate(iitemid,kaffareqSellYN) then
		    retStr =""
            kaffaItemSellYN = CheckKaffaItemSellYN(iitemid, kaffareqSellYN, retStr)
            bufStr = bufStr + retStr &"<br>"
	    end if

        retStr =""
        kaffaItemUpdate = CheckKaffaItemUpdate(iitemid, retStr)
        bufStr = bufStr + retStr &"<br>"

        retStr =""
        kaffaOptionStock = CheckKaffaOptionStock(iitemid, retStr)
        bufStr = bufStr + retStr &"<br>"

        retStr =""
        kaffaItemStat = CheckKaffaItemStat(iitemid,retStr)  ''/2013/07/24 �߰�
        bufStr = bufStr + retStr &"<br>"
	Next
	rw bufStr
Else
    rw "������-"&cmdparam
End If

function CheckKaffaItemStat(iitemid,byRef iretStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim strParam, SaleStatCd, GoodsViewCount
    Dim iRbody, jsResult, strResult
    dim idataArr
    dim sqlStr

    CheckKaffaItemStat = false

    strParam = "api_key="&C_APIKEY&"&product_code="&iitemid

''  On Error Resume Next

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", "http://10x10shop.com/api/call/get_product", false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send(strParam)

iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")

'rw iRbody
'' data : json ���� ���ڵ��� �� ��ü �ѹ��� ���ڵ��ȵ�
dim iResultcode, iResultmessage
dim isDataExists : isDataExists =false


set jsResult = JSON.parse(iRbody)
    iResultcode = jsResult.code
    iResultmessage = jsResult.message
    idataArr = jsResult.data

    on Error Resume Next                            ''dataArr ������ �ִ°��� ���°�� ���� �Ұ�
    site_product_id = jsResult.data.site_product_id
    IF Err then isDataExists = true
    on Error Goto 0
set jsResult = Nothing


if (iResultcode<>"1001") then
    iretStr = "["&iitemid&"] "&iResultmessage&"("&iResultcode&")"
    exit function
end if

if Not isDataExists then
    iretStr = "["&iitemid&"] ��ȸ���� ����"

    sqlStr = " update db_item.dbo.tbl_kaffa_reg_item"
    sqlStr = sqlStr & " SET lastStatCheckDate=getdate()"
    sqlStr = sqlStr & " where itemid="&iitemid&VbCRLF

    dbget.Execute sqlStr
    exit function
end if

dim site_product_id, site_brand_id, is_deleted
dim is_sale,is_display, supply_price,  sale_price, consumer_price
dim discount_price, discount_rate, discount_begin_datetime, discount_end_datetime
dim minimum, maximum
dim reg_datetime, is_shipping_free, weight, kaffaitemname
dim AssignedRow

set jsResult = JSON.parse(idataArr)
    rw "site_product_id:"&jsResult.site_product_id

  ''default
    site_product_id = jsResult.default.site_product_id          ''kaffa ��ǰ��ȣ
    site_brand_id   = jsResult.default.site_brand_id            ''kaffa �귣��ID
    is_deleted      = jsResult.default.is_deleted                '' ��������??
    is_sale         = jsResult.default.is_sale                      ''�Ǹſ���
    is_display      = jsResult.default.is_display                ''��������
    ''rw "is_used:"&jsResult.default.is_used                      '' �߰��ǰ����
    supply_price    = jsResult.default.supply_price         '' ���ް�
    sale_price      = jsResult.default.sale_price                '' �ǸŰ�
    consumer_price  = jsResult.default.consumer_price        '' �Һ�
    discount_price  = jsResult.default.discount_price                        '' ���ΰ���      --���� ����� �ʿ� ������
    discount_rate   = jsResult.default.discount_rate                          '' ������
    discount_begin_datetime = jsResult.default.discount_begin_datetime      '' ���ν�����
    discount_end_datetime   = jsResult.default.discount_end_datetime          '' ����������
    minimum = jsResult.default.minimum                      ''�ּұ��ż���
    maximum = jsResult.default.maximum                      ''�ִ뱸�ż���
    reg_datetime = jsResult.default.reg_datetime            ''��ǰ�����
    is_shipping_free= jsResult.default.is_shipping_free    ''�����ۿ���

  ''additional
    'rw "user_code1:"&jsResult.additional.user_code1             '' ������ڵ� - ��ǰ�ڵ�(TEN)
    'rw "user_code2:"&jsResult.additional.user_code2             ''
    'rw "hs_code:"&jsResult.additional.hs_code                   '' hs_code ?
    weight = jsResult.additional.weight                     '' ����
    'rw "production_date:"&jsResult.additional.production_date   ''������
    'rw "limit_date:"&jsResult.additional.limit_date             ''��ȿ��

  ''languages
    kaffaitemname = jsResult.languages.name.ko                       '' ��ǰ��

set jsResult = Nothing

if (site_product_id<>"0") and (iitemid<>"") then
    sqlStr = " update db_item.dbo.tbl_kaffa_reg_item"
    sqlStr = sqlStr & " SET useyn='y'"&VbCRLF                          ''����� y�� ��?
    sqlStr = sqlStr & ", kaffamakerid="&site_brand_id&VbCRLF
    sqlStr = sqlStr & ", kaffaGoodNo="&site_product_id&VbCRLF
    sqlStr = sqlStr & ", kaffaPrice="&sale_price&VbCRLF
    sqlStr = sqlStr & ", kaffaSellyn='"&CHKIIF(is_sale="1","Y","N")&"'"&VbCRLF
    sqlStr = sqlStr & ", kaffaIsDisplay='"&is_display&"'"&VbCRLF
    sqlStr = sqlStr & ", kaffaIsDeleted='"&is_deleted&"'"&VbCRLF
    sqlStr = sqlStr & ", kaffaSuplyPrice="&supply_price&VbCRLF
    sqlStr = sqlStr & ", kaffaConsumerPrice="&consumer_price&VbCRLF
    sqlStr = sqlStr & ", kaffaDiscountPrice="&discount_price&VbCRLF
    sqlStr = sqlStr & ", kaffaDiscountRate="&discount_rate&VbCRLF
    if (discount_begin_datetime="0000-00-00 00:00:00") then
        sqlStr = sqlStr & ", kaffaDiscount_Begin_DateTime=NULL"&VbCRLF
    else
        sqlStr = sqlStr & ", kaffaDiscount_Begin_DateTime='"&discount_begin_datetime&"'"&VbCRLF
    end if
    if (discount_end_datetime="0000-00-00 00:00:00") then
        sqlStr = sqlStr & ", kaffaDiscount_End_DateTime=NULL"&VbCRLF
    else
        sqlStr = sqlStr & ", kaffaDiscount_End_DateTime='"&discount_end_datetime&"'"&VbCRLF
    end if
    sqlStr = sqlStr & ", kaffaMinimum="&minimum&VbCRLF
    sqlStr = sqlStr & ", kaffaMaxium="&maximum&VbCRLF
    sqlStr = sqlStr & ", kaffaRegDateTime='"&reg_datetime&"'"&VbCRLF
    if IsNumeric(weight) and (weight<>"") then
        sqlStr = sqlStr & ", kaffaWeight="&weight*1000&VbCRLF
    else
        rw "[ERR]:weight:"&weight
    end if
    sqlStr = sqlStr & ", kaffaIsShippingfree="&is_shipping_free&VbCRLF
    sqlStr = sqlStr & ", lastStatCheckDate=getdate()"&VbCRLF
    sqlStr = sqlStr & " where itemid="&iitemid&VbCRLF
    ''rw sqlStr
    dbget.Execute sqlStr,AssignedRow

    CheckKaffaItemStat = (AssignedRow=1)
end if

end function

Function CheckKaffaItemSellYN(iitemid, subcmd, byRef iretStr)
    Dim objXML, strParam
    Dim iRbody, jsResult, ukaffaSellyn, kaffaIsDisplay
	Dim iResultcode, iResultmessage, idataArr
	mode = "ItemSellYN"
    CheckKaffaItemSellYN = false

	sqlStr = ""
	sqlStr = sqlStr & " SELECT itemid, isNull(kaffaGoodNo,'') as kaffaGoodNo, isNULL(kaffaIsDisplay,-1) as kaffaIsDisplay FROM db_item.dbo.tbl_kaffa_reg_item " & VBCRLF
	sqlStr = sqlStr & " WHERE itemid in ("&iitemid&") " & VBCRLF
    rsget.Open sqlStr,dbget,1
    If not rsget.Eof Then
		target_list = rsget("kaffaGoodNo")
		kaffaIsDisplay = rsget("kaffaIsDisplay")
    End If
    rsget.close

	If (target_list = "") Then
	    iretStr = "["&iitemid&"]�� KAFFA ��ǰ�ڵ� ����"
	    Exit Function
	End If

	If subcmd = "Y" Then
		ukaffaSellyn = 1
	Else
		ukaffaSellyn = 0
	End If

    ''���ô� �߱��� ���? (������)
	''���ð� �⺻ 0 ���� ������ :: �Ǹŷ� ����� ���� �ǰ� ����
	if (FALSE) and (ukaffaSellyn=1) and (kaffaIsDisplay=0) then
	    strParam = "api_key="&C_APIKEY&"&Command=product_display"&"&value=1&target_list="&target_list		'value : 0 ���þ��� / 1 ����
        Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
    		objXML.Open "POST", "http://10x10shop.com/api/call/product_display", false
    		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    		objXML.Send(strParam)
    		iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
    		''rw iRbody
    		Set jsResult = JSON.parse(iRbody)
    		    iResultcode = jsResult.code
    		    iResultmessage = jsResult.message
    		    idataArr = jsResult.data
    		Set jsResult = Nothing

    		If (iResultcode = "1001") AND (target_list <> "") Then
        		iretStr = "["&iitemid&"] "&iResultmessage&"("&iResultcode&")["&subcmd&"]�� �������� ����<br>"
        	end if

    	Set objXML = Nothing
    end if

    strParam = "api_key="&C_APIKEY&"&Command=product_sale"&"&value="&ukaffaSellyn&"&target_list="&target_list		'value : 0 �Ǹ����� / 1 �Ǹ�
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", "http://10x10shop.com/api/call/product_sale", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
		''rw iRbody
		Set jsResult = JSON.parse(iRbody)
		    iResultcode = jsResult.code
		    iResultmessage = jsResult.message
		    idataArr = jsResult.data
		Set jsResult = Nothing
	Set objXML = Nothing

	If (iResultcode = "1001") AND (target_list <> "") Then
		iretStr = iretStr&"["&iitemid&"] "&iResultmessage&"("&iResultcode&")["&subcmd&"]�� �ǸŻ��� ����"
		Call saveCommonItemResult(mode, iitemid, subcmd, "")
	Else
	    iretStr = iretStr&"["&iitemid&"] "&iResultmessage&"("&iResultcode&")"
	    Call Fn_AcctFailTouch(iitemid, iResultmessage)
	End if
End Function

function getKaffaDiscountInfo(iitemid,orgprice,byref discount_price,byref discount_rate,byref discount_begin_datetime,byref discount_end_datetime)
    ''�������Ұ��
    discount_price = ""
    discount_rate  = ""
    discount_begin_datetime = ""
    discount_end_datetime = ""

    dim sqlStr, dateExists : dateExists= false
    sqlStr = " exec db_item.dbo.[sp_Ten_KaFFA_DiscountInfo] "&iitemid&","&orgprice&""

    rsget.CursorLocation = adUseClient
	rsget.CursorType = adOpenStatic
	rsget.LockType = adLockOptimistic
    rsget.Open sqlStr,dbget,1

    if not rsget.Eof then
        discount_price = rsget("discountPrice")
        discount_rate  = rsget("disrate")
        discount_begin_datetime = rsget("stDT")
        discount_end_datetime = rsget("edDT")
        dateExists = true
    end if

    rsget.close()

    if (Not dateExists) then
        exit function
    end if

    IF IsNULL(discount_price) then ''���� ���� ����.
        discount_price = ""
        discount_rate  = ""
        discount_begin_datetime = ""
        discount_end_datetime = ""
    elseif IsNULL(discount_begin_datetime) then ''����
        discount_price = orgprice
        discount_rate  = "0.00"
        discount_begin_datetime = "null"
        discount_end_datetime = "null"
    else

    end if

    ''���� ���̺� �ִ��� ���� üũ
'    if iitemid=895065 then
'        ''���ιݿ���
'        discount_price = "22500"
'        discount_rate  = "15"
'        discount_begin_datetime = "2013-07-10 00:00:00"
'        discount_end_datetime = "2013-07-26 23:59:59"
'
'        ''���������
'        'discount_price = orgprice
'        'discount_rate  = "0.00"
'        'discount_begin_datetime = "null"
'        'discount_end_datetime = "null"
'    end if
end function

Function CheckKaffaItemUpdate(iitemid, byRef iretStr)		'��ǰ ���� �� �ٽ� ������� ���ư��� ��� ���� ã�� �� �� ex)��۷� ���ἳ���� 0->1 �� �ٲ� �� �ٽ� 1->0���� �ٲܶ�
    Dim objXML, strParam
    Dim iRbody, jsResult, ukaffaSellyn
	Dim iResultcode, iResultmessage, idataArr
	Dim isDataExists : isDataExists =false
	Dim obj
	Dim kaffaPrice, imakerid

	Dim discount_price,discount_rate,discount_begin_datetime,discount_end_datetime

	mode = "ItemUpdate"
	CheckKaffaItemUpdate = false

	sqlStr = ""
	sqlStr = sqlStr & " SELECT isNull(s.kaffaGoodNo,'') as kaffaGoodNo, s.useyn" & VbCRLF
	sqlStr = sqlStr & " , i.itemid, i.makerid, i.cate_large, i.cate_mid, i.cate_small" & VbCRLF
    sqlStr = sqlStr & " , i.itemdiv, i.itemgubun, i.itemname, i.sellcash, i.buycash" & VbCRLF
    sqlStr = sqlStr & " , isNULL(p.orgprice,i.orgprice) as orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash, i.mileage" & VbCRLF ''2013/07/16 i.orgprice => isNULL(p.orgprice,i.orgprice)
    sqlStr = sqlStr & " , i.regdate, i.lastupdate, i.sellEndDate, i.sellyn, i.limityn, i.danjongyn" & VbCRLF
    sqlStr = sqlStr & " , i.sailyn, i.isusing, i.isextusing, i.mwdiv, i.specialuseritem, i.vatinclude" & VbCRLF
    sqlStr = sqlStr & " , i.deliverytype, i.availPayType, i.deliverarea, i.deliverfixday, i.ismobileitem" & VbCRLF
    sqlStr = sqlStr & " , i.pojangok, i.limitno, i.limitsold, i.evalcnt, i.evalCnt_photo, i.optioncnt" & VbCRLF
    sqlStr = sqlStr & " , i.itemrackcode, i.upchemanagecode, i.reipgodate, i.brandname, i.titleimage" & VbCRLF
    sqlStr = sqlStr & " , i.mainimage, i.smallimage, i.listimage, i.listimage120, i.basicimage, i.icon1image" & VbCRLF
    sqlStr = sqlStr & " , i.icon2image, i.itemcouponyn, i.curritemcouponidx, i.itemcoupontype, i.itemcouponvalue" & VbCRLF
    sqlStr = sqlStr & " , i.itemscore, i.itemWeight, i.deliverOverseas, i.tenOnlyYn, i.basicimage600, i.mainImage2" & VbCRLF
    sqlStr = sqlStr & " , i.frontMakerid, i.reserveItemTp" & VbCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VbCRLF
	sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_kaffa_reg_item as s on i.itemid = s.itemid " & VbCRLF
    sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_multiLang_price P on P.itemid=i.itemid and sitename='CHNWEB'and currencyunit='WON'" & VbCRLF ''2013/07/16 �߰�
	sqlStr = sqlStr & " where 1=1 " & VbCRLF
	sqlStr = sqlStr & " and s.useyn = 'y'"
	sqlStr = sqlStr & " and isNull(s.kaffaGoodNo, '') <> '' " & VBCRLF
	sqlStr = sqlStr & " and i.itemid in ("&iitemid&") "
	rsget.Open sqlStr,dbget,1
	If not rsget.Eof then
		ArrRows = rsget.getRows()
	End If
	rsget.close

	If isArray(ArrRows) Then
		itemdata = ""
		imakerid   = ArrRows(3,0) ''makerid
		kaffaPrice = ArrRows(12,0) ''sellcash ==> orgprice 2013/07/04

'' �ؿ��ǸŰ� �������� �纯�� 1.5 �ּ�ó�� 2013/07/24
'		if (LCASE(imakerid)="ithinkso") or (LCASE(imakerid)="antennashop") then
'		    kaffaPrice = CLNG(kaffaPrice * 1.5)
'		end if

        discount_price          = ""
        discount_rate           = ""
        discount_begin_datetime = ""
        discount_end_datetime   = ""
		Call getKaffaDiscountInfo(ArrRows(2,0),kaffaPrice,discount_price,discount_rate,discount_begin_datetime,discount_end_datetime)
rw kaffaPrice&"|"&discount_price&"|"&discount_rate&"|"&discount_begin_datetime&"|"&discount_end_datetime
		Set obj = jsObject()
		Set obj(""&ArrRows(0,0)&"") = jsObject()
			obj(""&ArrRows(0,0)&"")("is_shipping_free") = "0"			'��۷� ���ἳ�� : is_shipping_free	(0�Ǵ�1)
			obj(""&ArrRows(0,0)&"")("supply_price") = CLNG(ArrRows(10,0)*0.9) ''"0"					'���ް� : supply_price	(numeric type) 90%�� ���� 0 ���� ���� �ȵ�?
			obj(""&ArrRows(0,0)&"")("sale_price") = kaffaPrice		'�ǸŰ� : sale_price	(numeric type)
			obj(""&ArrRows(0,0)&"")("consumer_price") = kaffaPrice	'�Һ��ڰ� : consumer_price	(numeric type)

			obj(""&ArrRows(0,0)&"")("discount_price") = discount_price ''""				'���ΰ� : discount_price	(numeric type)
			obj(""&ArrRows(0,0)&"")("discount_rate") = discount_rate ''""				'������ : discount_rate	(decimal)
			obj(""&ArrRows(0,0)&"")("discount_begin_datetime") = discount_begin_datetime ''""		'���αⰣ �����Ͻ� : discount_begin_datetime	(datetime)
			obj(""&ArrRows(0,0)&"")("discount_end_datetime") = discount_end_datetime ''""		'���αⰣ �����Ͻ� : discount_end_datetime	(datetime)

			obj(""&ArrRows(0,0)&"")("minimum") = 1						'�ּұ��ż��� : minimum	(numeric type)
			obj(""&ArrRows(0,0)&"")("maximum") = CHKIIF(ArrRows(36,0) > 0 OR ArrRows(36,0) = 0, 30, ArrRows(36,0))		'�ִ뱸�ż��� : maximum	(numeric type)
		itemdata = obj.jsString
		strParam = "api_key="&C_APIKEY&"&Command=set_product"&"&data="&itemdata
		Set obj = nothing
''rw  itemdata
''response.end

	    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "POST", "http://10x10shop.com/api/call/set_product", false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send(strParam)
			iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
			Set jsResult = JSON.parse(iRbody)
			    iResultcode = jsResult.code
			    iResultmessage = jsResult.message
			    idataArr = jsResult.data
			Set jsResult = Nothing
		Set objXML = Nothing
	Else
	    iretStr = "["&iitemid&"] �̵�� ��ǰ�̰ų� KAFFA ��ǰ�ڵ尡 ����"
	    Exit Function
	End If
	If (iResultcode = "1001") AND (ArrRows(0,0) <> "0") Then
		iretStr = "["&iitemid&"] "&iResultmessage&" ����/���� ����"
		Call saveCommonItemResult(mode, iitemid, "", kaffaPrice)
	Else
	    iretStr = "["&iitemid&"] "&iResultmessage&"("&iResultcode&")"
	    Call Fn_AcctFailTouch(iitemid, iResultmessage)
	End if
End Function

Function CheckKaffaOptionStock(iitemid, byRef iretStr)
    Dim objXML, strParam
    Dim iRbody, jsResult, ukaffaSellyn
	Dim iResultcode, iResultmessage, idataArr
	Dim isDataExists : isDataExists =false
	Dim obj, k, i
	Dim itemSu, itemoption, optaddprice, optsellyn, optlimityn, optlimitno, optlimitsold, optisusing
	mode = "OptionStock"
	CheckKaffaOptionStock = false

	sqlStr = ""
	sqlStr = sqlStr & " SELECT i.limityn, i.limitno, i.limitsold, o.* " & VbCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VbCRLF
	sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid " & VbCRLF
	sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_kaffa_reg_item as s on i.itemid = s.itemid " & VbCRLF
	sqlStr = sqlStr & " where 1=1 " & VbCRLF
	sqlStr = sqlStr & " and s.useyn = 'y'"
	sqlStr = sqlStr & " and isNull(s.kaffaGoodNo, '') <> '' " & VBCRLF
	sqlStr = sqlStr & " and i.itemid in ("&iitemid&") "
	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		itemdata = ""
		Set obj = jsObject()
		For i = 1 to rsget.RecordCount
			If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''���ϻ�ǰ
				itemoption		= "0000"
				itemSu			= GetKaffaLmtQty(rsget("limityn"), rsget("limitno"), rsget("limitsold"))
				obj(""&iitemid&"-"&itemoption&"") = itemSu
			Else
				itemoption 		= rsget("itemoption")
				optsellyn 		= rsget("optsellyn")
				optlimityn 		= rsget("optlimityn")
				optlimitno 		= rsget("optlimitno")
				optlimitsold 	= rsget("optlimitsold")
				optisusing      = rsget("isusing")          ''2013/08/08 �߰�
				itemSu 			= getOptionLimitNo(optsellyn, optlimityn, optlimitno, optlimitsold, optisusing)
				obj(""&iitemid&"-"&itemoption&"") = itemSu
			End If
			rsget.MoveNext
rw "["&iitemid&"-"&itemoption&"]:"&itemSu

		Next
		itemdata = obj.jsString
		strParam = "api_key="&C_APIKEY&"&Command=stock_fix"&"&data="&itemdata
		Set obj = nothing
	End If
	rsget.Close

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", "http://10x10shop.com/api/call/stock_fix", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
'rw iRbody
		Set jsResult = JSON.parse(iRbody)
		    iResultcode = jsResult.code
		    iResultmessage = jsResult.message
'rw iResultcode
'rw iResultmessage
            if (iResultcode="2001") then ''API ���� ����
                set idataArr = Nothing
            else
		        set idataArr = jsResult.data
            end if
		   ' rw jsResult.data.[0].product_item_code
		   ' rw jsResult.data.[0].result
		   ' rw jsResult.data.[0].message

		Set jsResult = Nothing
	Set objXML = Nothing


    ''�ɼ��� ���°�� ���ϰ��� ����.
'    for i=0 to idataArr.length-1
'        rw idataArr.get(i).product_item_code
'        rw idataArr.get(i).result
'        rw idataArr.get(i).message
'    next

	If (iResultcode = "1001") Then
		iretStr = "["&iitemid&"] "&iResultmessage&" ��ǰ���� ����"
		Call saveCommonItemResult(mode, iitemid, "", idataArr)
	Else
	    iretStr = "["&iitemid&"] "&iResultmessage&"("&iResultcode&")"
	    Call Fn_AcctFailTouch(iitemid, iResultmessage)
	End if
End Function

Function saveCommonItemResult(mode, iitemid, subcmd, idataArr)
    dim i
    dim product_item_code,result,message

	If mode = "ItemSellYN" Then						'���û�ǰ �ǸŻ��� ����
		sqlStr = ""
		sqlStr = sqlStr &" UPDATE R " &VbCRLF
		sqlStr = sqlStr &" SET accFailCnt = 0" & VBCRLF
		sqlStr = sqlStr &" ,kaffaSellyn = '"&subcmd&"'" & VBCRLF
		sqlStr = sqlStr &" FROM db_item.dbo.tbl_kaffa_reg_item as R" & VBCRLF
		sqlStr = sqlStr &" WHERE itemid = "&iitemid
		dbget.Execute(sqlStr)
	ElseIf mode = "ItemUpdate" Then					'���û�ǰ ���� ����
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE R" & VBCRLF
		sqlStr = sqlStr & " SET accFailCNT=0" & VBCRLF
		if (idataArr<>"") then
		    sqlStr = sqlStr & " ,kaffaprice = "&idataArr & VBCRLF
		else
		    sqlStr = sqlStr & " ,kaffaprice = i.orgprice" & VBCRLF ''sellcash=>orgprice
	    end if
		sqlStr = sqlStr & " ,kaffaConSumerPrice = i.orgprice" & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_kaffa_reg_item as R" & VBCRLF
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid" & VBCRLF
		sqlStr = sqlStr & " WHERE R.itemid = "&iitemid
		dbget.Execute(sqlStr)
	ElseIf mode = "OptionStock" Then				'���û�ǰ��ǰ/�ǸŻ��� ����
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE R" & VBCRLF
		sqlStr = sqlStr & " SET lastupdate = getdate()" & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_kaffa_reg_item as R" & VBCRLF
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid" & VBCRLF
		sqlStr = sqlStr & " WHERE R.itemid = "&iitemid
		dbget.Execute(sqlStr)

        if Not idataArr is Nothing Then
            for i=0 to idataArr.length-1
                product_item_code = idataArr.get(i).product_item_code
                result = idataArr.get(i).result
                message = idataArr.get(i).message

                if ((LCase(result)="true") or ((LCase(result)="false") and (message="�Է��� ��� ������ ���� ��� ������ �����ϴ�."))) then
                    sqlStr = "exec [dbo].[sp_Ten_OutMall_regedOption_Update_cn10x10] "&SplitValue(product_item_code,"-",0)&",'"&SplitValue(product_item_code,"-",1)&"' "

                    dbget.Execute sqlStr

                else
                    rw product_item_code
                    rw result
                    rw message
                end if
            next

            sqlStr = ""
    	    sqlStr = sqlStr & " update R"
            ''sqlStr = sqlStr & " set kaffasellyn=(CASE WHEN T.SellCNT>0 THEN 'Y' ELSE 'N' END)"
            sqlStr = sqlStr & " set regedOptCnt=isNULL(T.regedOptCnt,0)"
            sqlStr = sqlStr & " from db_item.dbo.tbl_kaffa_reg_item R"
            sqlStr = sqlStr & " 	Join ("
            sqlStr = sqlStr & " 		select itemid, count(*) as optCNT"
            sqlStr = sqlStr & " 		, sum(CASE WHEN outmallsellyn='Y' THEN 1 ELSE 0 END) as SellCNT"
            sqlStr = sqlStr & " 		, sum(CASE WHEN outmallsellyn='Y' and itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
            sqlStr = sqlStr & " 		from db_item.dbo.tbl_OutMall_regedoption"
            sqlStr = sqlStr & " 		where itemid="&iitemid&""
            sqlStr = sqlStr & " 		and mallid='cn10x10'"
            sqlStr = sqlStr & " 		group by itemid"
            sqlStr = sqlStr & " 	) T on R.itemid=T.itemid"
            sqlStr = sqlStr & " where R.itemid="&iitemid&""

            dbget.Execute sqlStr
        end if
	End If
End Function

Function Fn_AcctFailTouch(iitemid, iLastErrStr)
	sqlStr = ""
	sqlStr = sqlStr &" UPDATE R " &VbCRLF
	sqlStr = sqlStr &" SET accFailCnt = accFailCnt + 1" & VBCRLF
	sqlStr = sqlStr &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
	sqlStr = sqlStr &" FROM db_item.dbo.tbl_kaffa_reg_item as R" & VBCRLF
	sqlStr = sqlStr &" WHERE itemid = "&iitemid
	dbget.Execute(sqlStr)
End Function

function IsRequreUpdateSellStatUpdate(iitemid,kaffareqSellYN)
    dim sqlStr

    IsRequreUpdateSellStatUpdate = false
    sqlStr = " select CASE "
    'sqlStr = sqlStr & "		WHEN i.optionCnt>0 and convert(varchar(10),kaffaRegDateTime,21)<'2013-05-28' and kaffaSellyn='Y' THEN 'N'"        ''�ɼǵ�� ������ �ɼ� �ִ� ��ǰ �ӽ÷� ǰ�� ó��(2013-04-03 ��ϻ�ǰ) : 2013/06/10 ���� �̰� �ּ�����
    'sqlStr = sqlStr & "		WHEN i.optionCnt>0 and convert(varchar(10),kaffaRegDateTime,21)<'2013-05-28' and kaffaSellyn='N' THEN ''"
    sqlStr = sqlStr & "		WHEN (i.sellyn<>'Y' or i.isusing='N' or (i.limityn='Y' and i.limitno-5-i.limitsold<1)) and kaffaSellyn='Y' THEN 'N'"
    sqlStr = sqlStr & "		WHEN (i.sellyn<>'Y' or i.isusing='N' or (i.limityn='Y' and i.limitno-5-i.limitsold<1)) and kaffaSellyn='N' THEN ''"
	sqlStr = sqlStr & "		WHEN (i.mwdiv='U' and kaffaSellyn='Y') THEN 'N'"
	sqlStr = sqlStr & "		WHEN (i.mwdiv='U' and kaffaSellyn='N') THEN ''"
	sqlStr = sqlStr & "		WHEN (i.sellyn='Y' and (kaffaSellyn<>'Y' or kaffaisDisplay=0)) THEN 'Y'"
	sqlStr = sqlStr & "		ELSE '' END as kaffareqSellYN"
    sqlStr = sqlStr & " from db_item.dbo.tbl_item i"
    sqlStr = sqlStr & "	    Join db_item.dbo.tbl_kaffa_reg_item r"
    sqlStr = sqlStr & "	    on i.itemid=r.itemid"
    sqlStr = sqlStr & " where i.itemid="&iitemid

    rsget.Open sqlStr,dbget,1
    If not rsget.Eof Then
		kaffareqSellYN = rsget("kaffareqSellYN")
    End If
    rsget.close

    if (kaffareqSellYN<>"") then
        IsRequreUpdateSellStatUpdate = true
    end if
end function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->