<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
Function CheckVaildIP(ref)
	CheckVaildIP = false
	Dim VaildIP
	VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72","61.252.133.70")
	Dim i
	For i=0 to UBound(VaildIP)
		If (VaildIP(i)=ref) then
			CheckVaildIP = true
			Exit Function
		End If
	Next
End Function

Dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

If (Not CheckVaildIP(ref)) Then
    dbget.Close()
    response.end
End If

Dim act     : act = requestCheckVar(request("act"),32)
Dim param1  : param1 = requestCheckVar(request("param1"),32)
Dim param2  : param2 = requestCheckVar(request("param2"),32)
Dim param3  : param3 = requestCheckVar(request("param3"),32)
Dim param4  : param4 = requestCheckVar(request("param4"),32)
Dim param5  : param5 = requestCheckVar(request("param5"),32)
Dim sqlStr, i, paramData, retVal
Dim retCnt : retCnt = 0

Dim cnt
Dim OutMallOrderSerialArr
Dim OrgDetailKeyArr
Dim songjangDivArr
Dim songjangNoArr, sendReqCntArr, beasongdateArr, outmallGoodsIDArr
Dim oGSShopitem, itemidArr

Select Case act
    Case "gsSoldOutItem"	'ǰ��ó�� ��ǰ.
		Set oGSShopitem = new CGSShop
			oGSShopitem.FPageSize				= 15
			oGSShopitem.FCurrPage				= 1
			oGSShopitem.FRectGSShopNotReg		= "G"
			oGSShopitem.FRectSellYn				= "A"
			oGSShopitem.FRectGSShopYes10x10No	= "on"
	        oGSShopitem.getGSShopRegedItemList

			If (oGSShopitem.FResultCount < 1) Then
				response.Write "S_NONE"
				dbget.Close()
				Set oGSShopitem = Nothing
				response.end
			End If

			For i = 0 to oGSShopitem.FResultCount - 1
				itemidArr = itemidArr & oGSShopitem.FItemList(i).FItemID &","
			Next
			Set oGSShopitem = Nothing

			If (Right(itemidArr,1) = ",") Then itemidArr=Left(itemidArr,Len(itemidArr)-1)

			paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actgsshopReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
	Case "gsexpensive10x10"	'GSShop ���� < �ٹ����� ����
		Set oGSShopitem = new CGSShop
			oGSShopitem.FPageSize				= 15
			oGSShopitem.FCurrPage				= 1
			oGSShopitem.FRectGSShopNotReg		= "G"
			oGSShopitem.FRectMatchCate			= "Y"	'MatchCate
			oGSShopitem.FRectSellYn				= "Y"
			oGSShopitem.FRectExtSellYn			= "Y"
			oGSShopitem.FRectOrdType			= "B"	'����Ʈ��
			oGSShopitem.FRectExpensive10x10		= "on"
			oGSShopitem.FRectFailCntOverExcept	= "3"	' 3ȸ �̻� ���г��� ����.
			oGSShopitem.getGSShopRegedItemList
			If (oGSShopitem.FResultCount < 1) Then
				response.Write "S_NONE"
				dbget.Close()
				Set oGSShopitem= Nothing
				response.end
			End If

			For i = 0 to oGSShopitem.FResultCount - 1
				itemidArr = itemidArr & oGSShopitem.FItemList(i).FItemID &","
			Next
			Set oGSShopitem = Nothing

			If (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr) - 1)
			paramData = "redSsnKey=system&cmdparam=EditPrice&cksel="&itemidArr			'���� ����
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actgsshopReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
    Case "gsshopmarginItem"		'������ ���ݼ���
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 10 r.itemid, (i.buycash)/r.GSShopprice*100 as margin, i.buycash, i.orgprice, i.sellcash, r.GSShopprice, r.GSShopsellyn "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_gsshop_regitem as r "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on r.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE (r.GSShopstatcd = '3' OR r.GSShopstatcd = '7') "
		sqlStr = sqlStr & " and r.GSShopgoodNo is Not Null "
		sqlStr = sqlStr & " and r.GSShopprice<>0 "
		sqlStr = sqlStr & " and (i.buycash)/R.GSShopprice*100>85.1 "
		sqlStr = sqlStr & " and r.GSShopsellyn = 'Y' "
		sqlStr = sqlStr & " and i.orgprice <> R.GSShopprice "
		sqlStr = sqlStr & " ORDER BY (i.buycash)/R.GSShopprice*100 "
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
        paramData = "redSsnKey=system&cmdparam=EditPrice&cksel="&itemidArr
		retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actgsshopReq.asp",paramData)
        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
	Case "gsshopmarginNotSaleItem"			'������ ����N�ΰ� ǰ��ó��
		Set oGSShopitem = new CGSShop
			oGSShopitem.FPageSize			= 10
			oGSShopitem.FCurrPage			= 1
			oGSShopitem.FRectGSShopNotReg	= "G"
			oGSShopitem.FRectMatchCate		= "Y"
			oGSShopitem.FRectSellYn			= "A"
			oGSShopitem.FRectSailYn			= "N"
			oGSShopitem.FRectMinusMigin		= "on"
			oGSShopitem.getGSShopRegedItemList

			If (oGSShopitem.FResultCount < 1) Then
				response.Write "S_NONE"
				dbget.Close()
				set oGSShopitem= Nothing
				response.end
			End If
	
			For i = 0 to oGSShopitem.FResultCount - 1
			    itemidArr = itemidArr & oGSShopitem.FItemList(i).FItemID &","
			Next
			Set oGSShopitem= Nothing
	
			If (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr)-1)
			paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actgsshopReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
	Case "gsshopEditItem"	'GSShop ��ǰ����
		Set oGSShopitem = new CGSShop
			oGSShopitem.FPageSize			= 3
			oGSShopitem.FCurrPage			= param2
			oGSShopitem.FRectGSShopNotReg	= param5
			oGSShopitem.FRectMatchCate		= "Y"
			oGSShopitem.FRectSellYn			= "Y"
			oGSShopitem.FRectOrdType		= param3		'����Ʈ ������"B"
			If param4 <> "" Then							'�����Ǹ�
				oGSShopitem.FRectLimitYn = "Y"
			Else
				oGSShopitem.FRectonlyValidMargin = "on"		'���� �Ǵ°Ÿ�.           :: ���� ������ ǰ��ó��
			End If
			oGSShopitem.FRectFailCntOverExcept = "5"		'5ȸ �̻� ���г��� ����
			oGSShopitem.getGSShopRegedItemList

			If (oGSShopitem.FResultCount < 1) Then
				response.Write "S_NONE"
				dbget.Close()
				Set oGSShopitem = Nothing
				response.end
			End If
		
		
			For i=0 to oGSShopitem.FResultCount - 1
				itemidArr = itemidArr & oGSShopitem.FItemList(i).FItemID &","
			Next
		'response.write 	oGSShopitem.FResultCount&"��" &itemidArr
			Set oGSShopitem= Nothing
        
		'response.end
		
			If (Right(itemidArr, 1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr) - 1)
			paramData = "redSsnKey=system&cmdparam=EditOPT&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actgsshopReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
	Case "gsshopExpireItem"	'GS��ϵ��� ���ƾ� �� ��ǰ
		Set oGSShopitem = new CGSShop
			oGSShopitem.FPageSize       = 10
			oGSShopitem.FCurrPage       = param2
			oGSShopitem.FRectExtSellYn  = "Y"		'�Ǹ����λ�ǰ
			oGSShopitem.FRectFailCntOverExcept="3"	'3ȸ �̻� ���г��� ����.
			oGSShopitem.getGSShopreqExpireItemList

			If (oGSShopitem.FResultCount < 1) Then
				response.Write "S_NONE"
				dbget.Close()
				set oGSShopitem = Nothing
				response.end
			End If

			For i = 0 to oGSShopitem.FResultCount - 1
				itemidArr = itemidArr & oGSShopitem.FItemList(i).FItemID &","
			Next
			Set oGSShopitem= Nothing
			If (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr)-1)
			paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actgsshopReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
	Case "gsshopEditItemLastupdate"						'��ǰ���������� ���� ��ǰ����
		Set oGSShopitem = new CGSShop
			oGSShopitem.FPageSize				= 3
			oGSShopitem.FCurrPage				= 1
			oGSShopitem.FRectGSShopNotReg		= "G"
			oGSShopitem.FRectMatchCate			= "Y"
			oGSShopitem.FRectSellYn				= "Y"
			oGSShopitem.FRectExtSellYn			= "Y"
			oGSShopitem.FRectOrdType			= "LU"	'���������̺� ��ǰ�ֱ� ������ ����
			oGSShopitem.FRectFailCntOverExcept	= "3"
			oGSShopitem.getGSShopRegedItemList
			If (oGSShopitem.FResultCount < 1) Then
				response.Write "S_NONE"
				dbget.Close()
				Set oGSShopitem= Nothing
				response.end
			End If

			For i = 0 to oGSShopitem.FResultCount - 1
				itemidArr = itemidArr & oGSShopitem.FItemList(i).FItemID &","
			Next

			Set oGSShopitem= Nothing
			If (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr) - 1)
			paramData = "redSsnKey=system&cmdparam=EditOPT&cksel="&itemidArr                             ''���ݹ׳������
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actgsshopReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
    CASE "CheckItemNmAuto"
        paramData = "redSsnKey=system&cmdparam=CheckItemNmAuto"
		retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actgsshopReq.asp",paramData)
        response.Write "<br>"&retVal
    Case "gsshopLimitBrand"		'Ư�� �귣�� �����ٸ�
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 10 r.itemid, i.makerid, r.gsshoplastupdate "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_gsshop_regitem as r "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on r.itemid = i.itemid "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_outmall_limt_brand as b on i.makerid = b.makerid "
		sqlStr = sqlStr & " WHERE b.isusing ='Y' "
		sqlStr = sqlStr & " and (r.GSShopStatCd=3 OR r.GSShopStatCd=7) "
		sqlStr = sqlStr & " and r.GSShopGoodNo is Not Null "
		sqlStr = sqlStr & " order by r.gsshoplastupdate asc "
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
        paramData = "redSsnKey=system&cmdparam=EditOPT&cksel="&itemidArr
		retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/gsshop/actgsshopReq.asp",paramData)
        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case Else
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->