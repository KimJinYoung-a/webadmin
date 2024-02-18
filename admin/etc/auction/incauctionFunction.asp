<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
Public Function fnebayCommonCode(iccd, goodsGrpCd)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message, strSql, datalist, i
	Dim action
	Dim SDCategoryCode, SDCategoryName, IsLeafCategory
	Dim Gmkt, Iac
	Dim infoCodes, details, isUseTenOpt

	Select Case iccd
		Case "brand"						action = "BRAND"
		Case "maker"						action = "MANUFACTURER"
		Case "category"						action = "CATEGORY"
		Case "matchcategory"				action = "MATCH_CATEGORY"
		Case "address"						action = "ADDRESS"
		Case "locaddress"					action = "LOC_ADDRESS"
		Case "editNamebysitecategory"		action = "IS_EDIT_NAME_CATEGORY"
		Case "mastercode"					action = "GET_MASTER_CODE_BY_GOODS_NO"
		Case "sitecode"						action = "GET_GOODS_NO_BY_MASTER_CODE"
		Case "optionPolicy"					action = "OPTION_POLICY"
		Case "rcmdOption"					action = "RECOMMENDED_OPTION"
	End Select

	Set obj = jsObject()
		obj("action") = action
		obj("mallId") = CMALLNAME
		Select Case action
			Case "CATEGORY"						obj("goodsGrpCd") = "0"
			Case "BRAND", "MANUFACTURER"		obj("goodsGrpCd") = URLEncodeUTF8(goodsGrpCd)
			Case Else							obj("goodsGrpCd") = goodsGrpCd
		End Select
		istrParam = obj.jsString
	Set obj = nothing
' rw istrParam
' response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/code", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/code", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			status			= strObj.status
			code			= strObj.code
			If action = "CATEGORY" Then
				If IsNull(strObj.result.resultCode) Then
					Set datalist = strObj.result.sdCategoryTree
						strSql = ""
						strSql = strSql & " DELETE FROM db_temp.dbo.tbl_ebay_esmCategory "
						dbget.execute(strSql)
						For i=0 to datalist.length-1
							SDCategoryCode = datalist.get(i).SDCategoryCode
							SDCategoryName = datalist.get(i).SDCategoryName
							If datalist.get(i).IsLeafCategory = "True" Then
								IsLeafCategory = "Y"
							Else
								IsLeafCategory = "N"
							End If
							strSql = ""
							strSql = strSql & " INSERT INTO db_temp.dbo.tbl_ebay_esmCategory "
							strSql = strSql & " (SDCategoryCode, SDCategoryName, IsLeafCategory) "
							strSql = strSql & " VALUES ('" & SDCategoryCode & "', '" & SDCategoryName & "', '" & IsLeafCategory & "') "
							dbget.execute(strSql)
							If (i mod 1000) = 0 Then
								rw "저장 중 : " & i
								response.flush
							End If
						Next
					Set datalist = nothing

					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_ebay_esmCategory "
					dbget.execute(strSql)

					strSql = ""
					strSql = strSql & " SELECT SDCategoryCode "
					strSql = strSql & " , CASE WHEN (right(SDCategoryCode, 16) = '0000000000000000') THEN '0' "
					strSql = strSql & " 	   WHEN (right(SDCategoryCode, 12) = '000000000000') and (right(SDCategoryCode, 16) <> '0000000000000000') THEN left(SDCategoryCode, 4) + '0000000000000000' "
					strSql = strSql & " 	   WHEN (right(SDCategoryCode,  8) = '00000000')     and (right(SDCategoryCode, 12) <> '000000000000')     THEN left(SDCategoryCode, 8) + '000000000000' "
					strSql = strSql & " 	   WHEN (right(SDCategoryCode,  4) = '0000')         and (right(SDCategoryCode, 8) <> '00000000')		   THEN left(SDCategoryCode,12) + '00000000' "
					strSql = strSql & " else left(SDCategoryCode, 16) + '0000' "
					strSql = strSql & " end as parentSDCategoryCode "
					strSql = strSql & " , SDCategoryName, IsLeafCategory "
					strSql = strSql & " INTO #CleanTable "
					strSql = strSql & " FROM db_temp.dbo.tbl_ebay_esmCategory  "
					dbget.execute(strSql)

					strSql = ""
					strSql = strSql & " ;WITH CTETABLE(SDCategoryCode, parentSDCategoryCode, SDCategoryName, SDCategoryName2, LV, IsLeafCategory) as ( "
					strSql = strSql & " 	SELECT A.SDCategoryCode, A.parentSDCategoryCode "
					strSql = strSql & " 	, convert(varchar(300), A.SDCategoryName) as SDCategoryName "
					strSql = strSql & " 	, SDCategoryName as SDCategoryName2 "
					strSql = strSql & " 	, 1 "
					strSql = strSql & " 	, A.IsLeafCategory "
					strSql = strSql & " 	FROM #CleanTable A "
					strSql = strSql & " 	WHERE A.parentSDCategoryCode = '0' "
					strSql = strSql & " 	UNION ALL "
					strSql = strSql & " 	SELECT B.SDCategoryCode, B.parentSDCategoryCode "
					strSql = strSql & " 	, convert(varchar(300), C.SDCategoryName + ' > ' + B.SDCategoryName) as SDCategoryName "
					strSql = strSql & " 	, B.SDCategoryName as SDCategoryName2 "
					strSql = strSql & " 	, (C.LV + 1) LV "
					strSql = strSql & " 	, B.IsLeafCategory "
					strSql = strSql & " 	FROM #CleanTable B, "
					strSql = strSql & " 	CTETABLE C "
					strSql = strSql & " 	WHERE B.parentSDCategoryCode = C.SDCategoryCode "
					strSql = strSql & " ) "
					strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_ebay_esmCategory ( SDCategoryCode, parentSDCategoryCode, SDCategoryName, SDCategoryName2, LV, regdate) "
					strSql = strSql & " SELECT SDCategoryCode, parentSDCategoryCode, SDCategoryName, SDCategoryName2, LV, getdate() "
					strSql = strSql & " FROM CTETABLE "
					strSql = strSql & " where IsLeafCategory = 'Y' "
					strSql = strSql & " ORDER BY SDCategoryName, LV "
					dbget.execute(strSql)
					rw "OK"
				Else
					rw strObj.result.message
				End If
			ElseIf action = "MATCH_CATEGORY" Then
				If IsNull(strObj.result.resultCode) Then
					strSql = ""
					strSql = strSql &  " DELETE FROM [db_etcmall].[dbo].[tbl_ebay_matchCategory] WHERE SDCategoryCode = '"& goodsGrpCd &"' "
					dbget.execute(strSql)

					Set Gmkt = strObj.result.MatchedCategory.Gmkt
						For i = 0 to Gmkt.length - 1
							strSql = ""
							strSql = strSql &  " INSERT INTO [db_etcmall].[dbo].[tbl_ebay_matchCategory] (SDCategoryCode, siteCateCode, gubun, regdate) VALUES "
							strSql = strSql &  " ('"& goodsGrpCd &"', '"& Gmkt.get(i).catCode &"', 'gmarket1010', GETDATE()) "
							dbget.execute(strSql)
						Next
					Set Gmkt = nothing

					Set Iac = strObj.result.MatchedCategory.Iac
						For i = 0 to Iac.length - 1
							strSql = ""
							strSql = strSql &  " INSERT INTO [db_etcmall].[dbo].[tbl_ebay_matchCategory] (SDCategoryCode, siteCateCode, gubun, regdate) VALUES "
							strSql = strSql &  " ('"& goodsGrpCd &"', '"& Iac.get(i).catCode &"', 'auction1010', GETDATE()) "
							dbget.execute(strSql)
						Next
					Set Iac = nothing
				Else
					rw strObj.result.message
				End If
			ElseIf action = "IS_EDIT_NAME_CATEGORY" Then
				If IsNull(strObj.result.resultCode) Then
					strSql = ""
					strSql = strSql &  " UPDATE db_etcmall.dbo.tbl_ebay_siteCategory "
					strSql = strSql &  " SET isEditName = '"& CHKIIF(strObj.result.isEditable = true, "Y", "N") &"' "
					strSql = strSql &  " WHERE cateCode = '"& goodsGrpCd &"' "
					strSql = strSql &  " and gubun = '"& CMALLNAME &"' "
					dbget.execute(strSql)
				Else
					rw strObj.result.message
				End If
			ElseIf action = "OPTION_POLICY" Then
				If IsNull(strObj.result.resultCode) Then
					strSql = ""
					strSql = strSql &  " UPDATE db_etcmall.dbo.tbl_ebay_siteCategory "
					strSql = strSql &  " SET optionUseYn = '"& strObj.result.optionUseYn &"' "
					strSql = strSql &  " ,customOptionUseYn = '"& strObj.result.customOptionUseYn &"' "
					strSql = strSql &  " ,genrlUseYn = '"& strObj.result.optionType.genrlUseYn &"' "
					strSql = strSql &  " ,twoCmbtUseYn = '"& strObj.result.optionType.twoCmbtUseYn &"' "
					strSql = strSql &  " ,threeCmbtUseYn = '"& strObj.result.optionType.threeCmbtUseYn &"' "
					strSql = strSql &  " ,textUseYn = '"& strObj.result.optionType.textUseYn &"' "
					strSql = strSql &  " ,calcUseYn = '"& strObj.result.optionType.calcUseYn &"' "
					strSql = strSql &  " ,addAmntUseYn = '"& strObj.result.addAmnt.addAmntUseYn &"' "
					strSql = strSql &  " WHERE cateCode = '"& goodsGrpCd &"' "
					strSql = strSql &  " and gubun = '"& CMALLNAME &"' "
					dbget.execute(strSql)
				Else
					rw strObj.result.message
				End If
			ElseIf action = "RECOMMENDED_OPTION" Then
				If IsNull(strObj.result.resultCode) Then
					Set details = strObj.result.details
						isUseTenOpt = "N"
						For i = 0 to details.length - 1
							If details.get(i).recommendedOptNo = 0 AND details.get(i).recommendedOptTypeName = "선택형" Then
								isUseTenOpt = "Y"
								Exit For
							End If
						Next
						strSql = ""
						strSql = strSql &  " UPDATE db_etcmall.dbo.tbl_ebay_siteCategory "
						strSql = strSql &  " SET isUseTenOpt = '"& isUseTenOpt &"' "
						strSql = strSql &  " WHERE cateCode = '"& goodsGrpCd &"' "
						strSql = strSql &  " and gubun = '"& CMALLNAME &"' "
						dbget.execute(strSql)
					Set details = nothing
				Else
					rw strObj.result.message
				End If
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Public Function fnebaySiteCategoryCode(vdepth, vcateCode, vGubun)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message, strSql, datalist, i
	Dim action, goodsGrpCd
	Dim catCode, catName, isLeaf

	If vdepth = "1" Then
		goodsGrpCd = "0"
	Else
		goodsGrpCd = vcateCode
	End If

	Set obj = jsObject()
		obj("action") = "DISPLAYCATEGORY"
		obj("mallId") = CMALLNAME
		obj("goodsGrpCd") = goodsGrpCd
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/code", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/code", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

		Set strObj = JSON.parse(iRbody)
			status			= strObj.status
			code			= strObj.code

			If code = "0" Then
				If IsNull(strObj.result.resultCode) Then
					Set datalist = strObj.result.subCats
						For i = 0 to datalist.length - 1
							catCode = datalist.get(i).catCode
							catName = datalist.get(i).catName
							'IsLeafCategory = subCats.get(i).isLeaf
							If datalist.get(i).isLeaf = "True" Then
								isLeaf = "Y"
							Else
								isLeaf = "N"
							End If
							strSql = ""
							strSql = strSql & " INSERT INTO db_temp.dbo.tbl_ebay_siteCategory "
							strSql = strSql & " (gubun, depth, parentCatCode, catCode, catName, isLeaf) "
							strSql = strSql & " VALUES ('"& vGubun &"', '"& vDepth &"', '"& goodsGrpCd &"', '" & catCode & "', '" & catName & "', '" & isLeaf & "') "
							dbget.execute(strSql)
						Next
					Set datalist = nothing
				Else
					rw strObj.result.message
				End If
			End If

'			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
'			End If
		Set strObj = nothing
	Set objXML= nothing
End Function


%>