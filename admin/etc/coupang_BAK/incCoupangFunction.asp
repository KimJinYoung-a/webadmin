<%
'############################################## ���� �����ϴ� API �Լ� ���� ���� ############################################
'����� ���
Public Function fnCoupangDeliveryReg(iMakerid, iMaeipdiv, iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, isRegYn, strObj
    isRegYn = "N"
	If iMaeipdiv = "U" Then
	    istrParam = "makerID="&iMakerid
		'/////// �츮DB�� �ϴ� ����.. ���� ���� �ִٸ� ������� ���� ����ó�� ///////
		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Coupang_deliveryInfo_Add] '"&iMakerid&"' "
		dbget.Execute strSql

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as cnt "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " WHERE makerid = '"&iMakerid&"' "
		strSql = strSql & " and isNull(companyContactNumber, '') <> '' "
		strSql = strSql & " and isNull(phoneNumber2, '') <> '' "
		strSql = strSql & " and isNull(returnZipCode, '') <> '' "
		strSql = strSql & " and isNull(returnAddress, '') <> '' "
		strSql = strSql & " and isNull(returnAddressDetail, '') <> '' "
		strSql = strSql & " and isNull(deliveryCode, '') <> '' "
		rsget.Open strSql,dbget,1
		If rsget("cnt") > 0 Then
			isRegYn = "Y"
		End If
		rsget.Close
		'//////////////////////////////////////////////////////////////////////
		If isRegYn = "Y" Then
			On Error Resume Next
			Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
				objXML.open "POST", "http://xapi.10x10.co.kr:8080/Delivery/Coupang/regoutbound", false
				objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				objXML.Send(istrParam)

				If Err.number <> 0 Then
					iErrStr = "ERR||"&iMakerid&"||����[�����] " & Err.Description
					Exit Function
				End If

				If objXML.Status = "200" OR objXML.Status = "201" Then
					iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
					'response.write iRbody
					Set strObj = JSON.parse(iRbody)
						'rw strObj.outboundShippingPlaceCode �̰ɷ� DB������Ʈ �Ϸ��ߴ� ��, �̹� API�������� �����ѵ�..
						iErrStr = "OK||"&iMakerid&"||����[�����]"
					Set strObj = nothing
				Else
					iErrStr = "ERR||"&iMakerid&"||����[�����] ��ſ���"
				End If
			Set objXML = nothing
		Else
			iErrStr = "ERR||"&iMakerid&"||����[�����] ��������"
		End If
	Else		'���� or Ư���̶�� ������� ������ ����
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping WHERE makerid='"&iMakerid&"' )"
		strSql = strSql & " BEGIN "
		strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " (makerid, vendorId, deliveryCode, companyContactNumber, notJeju, outboundShippingPlaceCode, regdate ) VALUES "
		strSql = strSql & " ('"&iMakerid&"', '', 'CJGLS', '1644-6035', '3000', '122412', getdate()) END "
		dbget.Execute strSql
		iErrStr = "OK||"&iMakerid&"||����[�����]"
	End If
End Function

'��ǰ ���
Public Function fnCoupangItemReg(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Product/Coupang/registryproduct", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ���] " & Err.Description
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				strSql = " EXEC db_etcmall.[dbo].[usp_API_Coupang_RegItemInfo_Upd] '"&iitemid&"', 'I' "
				dbget.execute strSql

				iErrStr = "OK||"&iitemid&"||����[��ǰ���]"
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||����[��ǰ���] ��ſ���"
		End If
	Set objXML = nothing
End Function

'��ǰ ��ȸ
Public Function fnCoupangStatChk(iitemid, icoupangGoodno, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
    Dim productId, regedItemname, statusName, coupangStatcd, strObjitems
    Dim vItemoption, vendorItemId, sellerProductItemId, vOptionName, vItemsu
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Product/Coupang/getsingleproduct?sellerProductId="&icoupangGoodno, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ȸ] " & Err.Description
			Exit Function
		End If
		'rw objXML.Status
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		'response.write iRbody
		'response.end
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				statusName		= strObj.data.statusName
				productId		= strObj.data.productId
				regedItemname	= strObj.data.sellerProductName

				If retCode = "SUCCESS" Then
					Select Case statusName
						Case "���οϷ�"					coupangStatcd = 7
						Case "���ιݷ�"					coupangStatcd = 2
						Case "���δ����","���ο�û"		coupangStatcd = 3
					End Select

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regitem " & VbCRLF
					strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
					strSql = strSql & "	,coupangStatcd='"&coupangStatcd&"' "
					strSql = strSql & " ,productId='" & productId & "' "
					strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(strSql)
					If coupangStatcd = 7 Then
						set strObjitems = strObj.data.items
						strSql = ""
						strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_coupang_regedoption WHERE itemid = '"&iitemid&"' "
						dbget.Execute(strSql)
							For i=0 to strObjitems.length-1
								vendorItemId = strObjitems.get(i).vendorItemId
								vItemoption	= Split(strObjitems.get(i).externalVendorSku, "_")(1)
								sellerProductItemId = strObjitems.get(i).sellerProductItemId
								vItemsu		= strObjitems.get(i).maximumBuyCount

								If vItemoption <> "0000" Then
									'vOptionName = Trim(replace(strObjitems.get(i).itemName, regedItemname, ""))
									strSql = ""
									strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_RegItemOption_Add] '"&iitemid&"', '"&vItemoption&"', '"&vendorItemId&"', '"&sellerProductItemId&"', '"&vItemsu&"', 'Y' "
									dbget.Execute(strSql)
								Else
									'vOptionName = strObjitems.get(i).itemName
									strSql = ""
									strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_RegItemOption_Add] '"&iitemid&"', '"&vItemoption&"', '"&vendorItemId&"', '"&sellerProductItemId&"', '"&vItemsu&"', 'N' "
									dbget.Execute(strSql)
								End If
							Next
							strSql = ""
							strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_RegItemOptionCnt_Upd] '"&iitemid&"' "
							dbget.Execute(strSql)
						set strObjitems = nothing
					End If
					iErrStr = "OK||"&iitemid&"||����[��ȸ("&statusName&")]"
				Else					
					iErrStr = "ERR||"&iitemid&"||����[��ȸ]NOT SUCCESS"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||����[��ȸ]��ſ���"
		End If
	Set objXML = nothing
End Function

'��ǰ ���� ����
Public Function fnCoupangSellyn(iitemid, ichgSellyn, ivendorItemId, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
	istrParam = "vendorItemId="&ivendorItemId
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If ichgSellyn = "Y" Then
			objXML.open "POST", "http://xapi.10x10.co.kr:8080/Product/Coupang/productagainsell", false
		ElseIf ichgSellyn = "N" Then
			objXML.open "POST", "http://xapi.10x10.co.kr:8080/Product/Coupang/itemstop", false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				If retCode <> "SUCCESS" Then
					iErrStr = ivendorItemId
				End If
			Set strObj = nothing
		Else
			iErrStr = ivendorItemId
		End If
	Set objXML = nothing
End Function

'��ǰ ����
Public Function fnCoupangDelete(iitemid, icoupangGoodno, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
	On Error Resume Next
	istrParam = "sellerProductId="&icoupangGoodno
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Product/Coupang/deleteproduct", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ����] " & Err.Description
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				If retCode <> "SUCCESS" Then
					iErrStr = "ERR||"&iitemid&"||����[��ǰ����]NOT SUCCESS"
				Else
					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_coupang_regitem " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_coupang_regedoption " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_outmall_API_Que " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' " & vbcrlf
					strSql = strSql & " and mallid = '"&CMALLNAME&"' " & vbcrlf
					dbget.Execute(strSql)
				End If
				iErrStr = "OK||"&iitemid&"||����[��ǰ����]"
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||����[��ǰ����] ��ſ���"
		End If
	Set objXML = nothing
End Function

'��ǰ ���� ����
Public Function fnCoupangPrice(iitemid, ivendorItemId, imustprice, imustOptionprice, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
	istrParam = "vendorItemId="&ivendorItemId&"&Price="&imustOptionprice
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Product/Coupang/updateitemprice", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				If retCode <> "SUCCESS" Then
					iErrStr = ivendorItemId
				End If
			Set strObj = nothing
		Else
			iErrStr = ivendorItemId
		End If
	Set objXML = nothing
End Function

'��ǰ ��� ����
Public Function fnCoupangQuantity(iitemid, ivendorItemId, iquantity, isNameDiff, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
    If iquantity < 0 OR isNameDiff = 1 Then
    	iquantity = 0
    End If

	istrParam = "vendorItemId="&ivendorItemId&"&quantity="&iquantity
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Product/Coupang/productqtychange", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				If retCode <> "SUCCESS" Then
					iErrStr = ivendorItemId
				End If
			Set strObj = nothing
		Else
			iErrStr = ivendorItemId
		End If
	Set objXML = nothing
End Function

'��ǰ ����
Public Function fnCoupangItemEdit(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Product/Coupang/updateproduct", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ����] " & Err.Description
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			strSql = ""
			strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Coupang_RegItemInfo_Upd] '"&iitemid&"', 'R' "
			dbget.Execute(strSql)
			iErrStr = "OK||"&iitemid&"||����[��ǰ����]"
		Else
			iErrStr = "ERR||"&iitemid&"||����[��ǰ����] ��ſ���"
		End If
	Set objXML = nothing
End Function

Function fnBrandmaeipdiv(iMakerid)
	Dim strSql
	strSql = strSql & " SELECT TOP 1 maeipdiv "
	strSql = strSql & " FROM db_user.dbo.tbl_user_c "
	strSql = strSql & " WHERE userid = '"& iMakerid &"' "
    rsget.Open strSql,dbget,1
    if (Not rsget.EOF) then
	    fnBrandmaeipdiv = rsget("maeipdiv")
	end if
    rsget.Close
End Function

Function getCoupangGoodno(iitemid)
	Dim strSql
	strSql = strSql & " SELECT TOP 1 isnull(coupangGoodno, '') as coupangGoodno "
	strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regitem "
	strSql = strSql & " WHERE itemid = '"& iitemid &"' "
    rsget.Open strSql,dbget,1
    if (Not rsget.EOF) then
	    getCoupangGoodno = rsget("coupangGoodno")
	end if
    rsget.Close
End Function

Function getCoupangVendorItemidList(iitemid)
	Dim strSql
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_VendorItemIdList_Get] '"&iitemid&"' "
    rsget.Open strSql,dbget,1
    if (Not rsget.EOF) then
	    getCoupangVendorItemidList = rsget.getRows
	end if
    rsget.Close
End Function

Function ArrErrStrInfo(iaction, iValue, iitemid, ierrVendorItemId)
	Dim ErrStrComma, strSql
	If iaction = "EditSellYn" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||����[���º���] " & ErrStrComma
		Else
			If iValue = "N" Then
				strSql = ""
				strSql = strSql & " UPDATE R"
				strSql = strSql & "	Set coupangSellYn = 'N'"
				strSql = strSql & "	,accFailCnt = 0"
				strSql = strSql & "	,coupangLastUpdate = getdate()"
				strSql = strSql & "	From db_etcmall.dbo.tbl_coupang_regitem  R"
				strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
				dbget.Execute(strSql)
				ArrErrStrInfo = "OK||"&iitemid&"||ǰ��ó��[���º���]"
			Else
				strSql = ""
				strSql = strSql & " UPDATE R"
				strSql = strSql & "	Set coupangSellYn = 'Y'"
				strSql = strSql & "	,coupangLastUpdate = getdate()"
				strSql = strSql & "	From db_etcmall.dbo.tbl_coupang_regitem  R"
				strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
				dbget.Execute(strSql)
				ArrErrStrInfo = "OK||"&iitemid&"||�Ǹ�[���º���]"
			End If
		End If
	ElseIf iaction = "PRICE" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||����[���ݼ���] " & ErrStrComma
		Else
		    strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regitem " & VbCRLF
			strSql = strSql & "	SET coupangLastUpdate = getdate() " & VbCRLF
			strSql = strSql & "	, coupangPrice = " & iValue & VbCRLF
			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
			strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
			dbget.Execute(strSql)
			ArrErrStrInfo =  "OK||"&iitemid&"||����[���ݼ���]"
		End If
	ElseIf iaction = "QTY" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||����[������] " & ErrStrComma
		Else
		    strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regedoption "
			strSql = strSql & " SET outmalllimitno = CASE WHEN (A.outmallOptName <> '���ϻ�ǰ') AND (A.outmalloptName <> isNULL(B.optionname,''))  THEN 0 "
			strSql = strSql & " WHEN isnull(B.itemoption, '') = '' AND i.limityn = 'Y' THEN i.limitno - i.limitsold - 5 "
			strSql = strSql & " WHEN isnull(B.itemoption, '') = '' AND i.limityn = 'N' THEN 9999 "
			strSql = strSql & " WHEN isnull(B.itemoption, '') <> '' AND i.limityn = 'Y' THEN B.optlimitno - B.optlimitsold - 5 "
			strSql = strSql & " WHEN isnull(B.itemoption, '') <> '' AND i.limityn = 'N' THEN 9999 END "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regedoption as A "
			strSql = strSql & " JOIN db_item.dbo.tbl_item as i on A.itemid = i.itemid "
			strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_option as B on A.itemid = B.itemid and A.itemoption = B.itemoption "
			strSql = strSql & " WHERE A.itemid = '"&iitemid&"' "
			dbget.Execute(strSql)
			ArrErrStrInfo =  "OK||"&iitemid&"||����[������]"
		End If
	End If
End Function
%>