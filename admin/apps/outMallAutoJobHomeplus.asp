<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/admin/etc/homeplus/homepluscls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
Function CheckVaildIP(ref)
	CheckVaildIP = false
	Dim VaildIP
	VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
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
Dim oHomeplus, itemidArr

Select Case act
	Case "outmallSongJangIp"
    'response.end
        '' �����Է��� �߰����� :: wAPI�� �Űܾ� �ҵ�. :�ʹ� ����.
        dim mayTime : mayTime = replace(LEFT(FormatDateTime(Now(), 4),2),":","")
        if (mayTime>8) and (mayTime<18) then
            response.write "mayTime:"&mayTime
        else
    		sqlStr = "select top 5 T.orderserial, T.OutMallOrderSerial"
    		sqlStr = sqlStr & " ,T.OrgDetailKey, IsNULL(T.sendState,0) as sendState"
    		sqlStr = sqlStr & " ,D.songjangDiv, D.songjangNo, D.itemNo, D.beasongdate, T.outMallGoodsNo"
    		sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
    		sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master M"
    		sqlStr = sqlStr & " 	on T.orderserial=M.orderserial"
    		sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D"
    		sqlStr = sqlStr & " 	on T.orderserial=D.orderserial"
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
    				    IF (LCASE(param1)="homeplus") then
    				        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&Server.URLEncode(TenDlvCode2HomeplusDlvCode(songjangDivArr(i)))&"&inv_no="&songjangNoArr(i)
    				        if (application("Svr_Info")<>"Dev") then
    							retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/actHomeplusSongjangInputProc.asp",paramData)
    							rw retVal
    				        else
    							retVal = SendReq("http://testwebadmin.10x10.co.kr/admin/etc/homeplus/actHomeplusSongjangInputProc.asp",paramData)
    							rw retVal
    				        end if
    				    end if
    				end if
    			next
            end if
        end if
    Case "homeplusSoldOutItem"	'ǰ��ó�� ��ǰ.
		Set oHomeplus = new CHomeplus
			oHomeplus.FPageSize					= 20
			oHomeplus.FCurrPage					= 1
			oHomeplus.FRectHomeplusNotReg		= "D"
			oHomeplus.FRectHomeplusYes10x10No	= "on"
	        oHomeplus.getHomeplusRegedItemList

			If (oHomeplus.FResultCount < 1) Then
				response.Write "S_NONE"
				dbCTget.Close()
				Set oHomeplus = Nothing
				response.end
			End If

			For i = 0 to oHomeplus.FResultCount - 1
				itemidArr = itemidArr & oHomeplus.FItemList(i).FItemID &","
			Next
			Set oHomeplus = Nothing

			If (Right(itemidArr,1) = ",") Then itemidArr=Left(itemidArr,Len(itemidArr)-1)

			paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/acthomeplusReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
	Case "homeplusexpensive10x10"	'homeplus ���� < �ٹ����� ����
		Set oHomeplus = new CHomeplus
			oHomeplus.FPageSize					= 20
			oHomeplus.FCurrPage					= 1
			oHomeplus.FRectMatchCate			= "Y"
			oHomeplus.FRectHomeplusNotReg		= "D"
			oHomeplus.FRectSellYn				= "Y"
			oHomeplus.FRectExtSellYn			= "Y"
			oHomeplus.FRectOrdType				= "B"	'����Ʈ��
			oHomeplus.FRectExpensive10x10		= "on"
			oHomeplus.FRectFailCntOverExcept	= "3"	' 3ȸ �̻� ���г��� ����.
			oHomeplus.getHomeplusRegedItemList
			If (oHomeplus.FResultCount < 1) Then
				response.Write "S_NONE"
				dbCTget.Close()
				Set oHomeplus= Nothing
				response.end
			End If

			For i = 0 to oHomeplus.FResultCount - 1
				itemidArr = itemidArr & oHomeplus.FItemList(i).FItemID &","
			Next
			Set oHomeplus = Nothing

			If (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr) - 1)
			paramData = "redSsnKey=system&cmdparam=EditItemSelect&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/acthomeplusReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
    Case "homeplusmarginItem"		'������ ���ݼ���
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 10 r.itemid, (i.buycash)/r.homeplusprice*100 as margin, i.buycash, i.orgprice, i.sellcash, r.homeplusprice, r.homeplussellyn "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_homeplus_regitem as r "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as i on r.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE r.homeplusstatcd = '7' "
		sqlStr = sqlStr & " and r.homeplusgoodNo is Not Null "
		sqlStr = sqlStr & " and r.homeplusprice<>0 "
		sqlStr = sqlStr & " and (i.buycash)/R.homeplusprice*100>85.1 "
		sqlStr = sqlStr & " and r.homeplussellyn = 'Y' "
		sqlStr = sqlStr & " and i.orgprice <> R.homeplusprice "
		sqlStr = sqlStr & " ORDER BY (i.buycash)/R.homeplusprice*100 "
        rsCTget.Open sqlStr,dbCTget,1
        cnt = rsCTget.RecordCount
		If (cnt < 1) Then
			response.Write "S_NONE"
			response.end
		Else
	        For i = 0 to cnt - 1
	            itemidArr = itemidArr & rsCTget("itemid") &","
				rsCTget.MoveNext
	        Next
		End If
        rsCTget.close
        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
        paramData = "redSsnKey=system&cmdparam=EditItemSelect&cksel="&itemidArr
		retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/acthomeplusReq.asp",paramData)
        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
	Case "homeplusmarginNotSaleItem"			'������ ����N�ΰ� ǰ��ó��
		Set oHomeplus = new CHomeplus
			oHomeplus.FPageSize				= 10
			oHomeplus.FCurrPage				= 1
			oHomeplus.FRectMatchCate		= "Y"
			oHomeplus.FRectHomeplusNotReg	= "D"
			oHomeplus.FRectSellYn			= "Y"
			oHomeplus.FRectSailYn			= "N"
			oHomeplus.FRectMinusMigin		= "on"
			oHomeplus.getHomeplusRegedItemList

			If (oHomeplus.FResultCount < 1) Then
				response.Write "S_NONE"
				dbCTget.Close()
				set oHomeplus= Nothing
				response.end
			End If
	
			For i = 0 to oHomeplus.FResultCount - 1
			    itemidArr = itemidArr & oHomeplus.FItemList(i).FItemID &","
			Next
			Set oHomeplus= Nothing
	
			If (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr)-1)
			paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/acthomeplusReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
	Case "homeplusEditItem"	'Homeplus ��ǰ����
		Set oHomeplus = new CHomeplus
			oHomeplus.FPageSize				= 10
			oHomeplus.FCurrPage				= param2
			oHomeplus.FRectHomeplusNotReg	= param5
			oHomeplus.FRectMatchCate		= "Y"
			oHomeplus.FRectSellYn			= "Y"
			oHomeplus.FRectOrdType			= param3		'����Ʈ ������"B"
			If param4 <> "" Then							'�����Ǹ�
				oHomeplus.FRectLimitYn = "Y"
			Else
				oHomeplus.FRectonlyValidMargin = "on"		'���� �Ǵ°Ÿ�.           :: ���� ������ ǰ��ó��
			End If
			oHomeplus.FRectFailCntOverExcept = "5"		'5ȸ �̻� ���г��� ����
			oHomeplus.getHomeplusRegedItemList

			If (oHomeplus.FResultCount < 1) Then
				response.Write "S_NONE"
				dbCTget.Close()
				Set oHomeplus = Nothing
				response.end
			End If
			
			For i=0 to oHomeplus.FResultCount - 1
				itemidArr = itemidArr & oHomeplus.FItemList(i).FItemID &","
			Next
			Set oHomeplus= Nothing

			If (Right(itemidArr, 1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr) - 1)
			paramData = "redSsnKey=system&cmdparam=EditItemSelect&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/acthomeplusReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
	Case "homeplusExpireItem"	'Homeplus��ϵ��� ���ƾ� �� ��ǰ
		Set oHomeplus = new CHomeplus
			oHomeplus.FPageSize   			    = 10
			oHomeplus.FCurrPage   			    = param2
			oHomeplus.FRectExtSellYn  			= "Y"		'�Ǹ����λ�ǰ
			oHomeplus.FRectFailCntOverExcept	="3"		'3ȸ �̻� ���г��� ����.
			oHomeplus.getHomeplusreqExpireItemList

			If (oHomeplus.FResultCount < 1) Then
				response.Write "S_NONE"
				dbCTget.Close()
				set oHomeplus = Nothing
				response.end
			End If

			For i = 0 to oHomeplus.FResultCount - 1
				itemidArr = itemidArr & oHomeplus.FItemList(i).FItemID &","
			Next
			Set oHomeplus= Nothing
			If (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr)-1)
			paramData = "redSsnKey=system&cmdparam=EditSellYn&chgSellYn=N&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/acthomeplusReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
	Case "homeplusEditItemLastupdate"						'��ǰ���������� ���� ��ǰ����
		Set oHomeplus = new CHomeplus
			oHomeplus.FPageSize					= 30
			oHomeplus.FCurrPage					= 1
			oHomeplus.FRectHomeplusNotReg		= "D"
			oHomeplus.FRectMatchCate			= "Y"
			oHomeplus.FRectSellYn				= "Y"
			oHomeplus.FRectExtSellYn			= "Y"
			oHomeplus.FRectOrdType				= "LU"	'���������̺� ��ǰ�ֱ� ������ ����
			oHomeplus.FRectFailCntOverExcept	= "3"
			oHomeplus.getHomeplusRegedItemList
			If (oHomeplus.FResultCount < 1) Then
				response.Write "S_NONE"
				dbCTget.Close()
				Set oHomeplus= Nothing
				response.end
			End If

			For i = 0 to oHomeplus.FResultCount - 1
				itemidArr = itemidArr & oHomeplus.FItemList(i).FItemID &","
			Next

			Set oHomeplus= Nothing
			If (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr) - 1)
			paramData = "redSsnKey=system&cmdparam=EditItemSelect&cksel="&itemidArr
			retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/acthomeplusReq.asp",paramData)
			response.Write "itemidArr="&itemidArr
			response.Write "<br>"&retVal
    Case "CheckItemNmAuto"
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 10 r.itemid, r.homeplusGoodNo, i.ItemName "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_homeplus_regItem r "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item i on r.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE r.regitemname is Not NULL "
		sqlStr = sqlStr & " and r.homeplusStatCd=7 "
		sqlStr = sqlStr & " and r.homeplusGoodNo is Not Null "
		sqlStr = sqlStr & " and r.regitemname <> i.itemname "
		sqlStr = sqlStr & " ORDER BY r.regdate DESC "
        rsCTget.Open sqlStr,dbCTget,1
        cnt = rsCTget.RecordCount
		If (cnt < 1) Then
			response.Write "S_NONE"
			response.end
		Else
	        For i = 0 to cnt - 1
	            itemidArr = itemidArr & rsCTget("itemid") &","
				rsCTget.MoveNext
	        Next
		End If
        rsCTget.close
        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
        paramData = "redSsnKey=system&cmdparam=EditSelect&cksel="&itemidArr
		retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/acthomeplusReq.asp",paramData)
        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case "homeplusLimitBrand"		'Ư�� �귣�� �����ٸ�
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 10 r.itemid, i.makerid, r.homepluslastupdate "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_homeplus_regitem as r "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as i on r.itemid = i.itemid "
		sqlStr = sqlStr & " JOIN db_outmall.dbo.tbl_outmall_limt_brand as b on i.makerid = b.makerid "
		sqlStr = sqlStr & " WHERE b.isusing ='Y' "
		sqlStr = sqlStr & " and r.homeplusStatCd=7 "
		sqlStr = sqlStr & " and r.homeplusGoodNo is Not Null "
		sqlStr = sqlStr & " order by r.homepluslastupdate asc "
        rsCTget.Open sqlStr,dbCTget,1
        cnt = rsCTget.RecordCount
		If (cnt < 1) Then
			response.Write "S_NONE"
			response.end
		Else
	        For i = 0 to cnt - 1
	            itemidArr = itemidArr & rsCTget("itemid") &","
				rsCTget.MoveNext
	        Next
		End If
        rsCTget.close
        IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
        paramData = "redSsnKey=system&cmdparam=EditItemSelect&cksel="&itemidArr
		retVal = SendReq("http://webadmin.10x10.co.kr/admin/etc/homeplus/acthomeplusReq.asp",paramData)
        response.Write "itemidArr="&itemidArr
        response.Write "<br>"&retVal
    Case Else
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->