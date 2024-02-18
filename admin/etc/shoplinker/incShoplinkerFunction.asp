<%
Dim isShoplinker_DebugMode : isShoplinker_DebugMode = True
Dim ShoplinkerAPIURL
ShoplinkerAPIURL = "http://apiweb.shoplinker.co.kr/ShoplinkerApi"

Public XMLURL

Function saveCommonItemResult(retDoc, mode, prdno)
	Dim errorMsg, result
	Dim sqlStr
	Dim AssignedRow, successYn
	Dim Titemid, Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu, Tsellcash
	Dim Nodes, OneNode, SubNodes
	Dim typeCD, itemCD_ZIP, newUnitRetail, newUnitCost, packInd
	Dim unitCd, strDt, endDt, availSupQty, notitemId, notmakerid
    Dim ierrStr, shoplinkerPrdno
    Dim product_id
    Dim MustPrice

	successYn = false

	If mode = "MDT" Then
		Set Nodes = retDoc.getElementsByTagName("ResultMessage")
		If (Not (retDoc is Nothing)) Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT sellcash, buycash, orgprice, makerid FROM db_item.dbo.tbl_item where itemid = "&prdno&"  " & VBCRLF
			rsget.Open sqlStr,dbget,1
			If Not(rsget.EOF or rsget.BOF) Then
				If CLng(10000 - rsget("buycash") / rsget("sellcash") * 100 * 100) / 100 < CMAXMARGIN Then
					MustPrice = rsget("orgprice")
				Else
					MustPrice = rsget("sellcash")
				End If

				If rsget("makerid") = "KLING" Then
					MustPrice = rsget("sellcash")
				End If
			End If
			rsget.Close

			For each SubNodes in Nodes
			    result		= SubNodes.getElementsByTagName("result").item(0).text
				errorMsg	= SubNodes.getElementsByTagName("message").item(0).text
				product_id	= SubNodes.getElementsByTagName("product_id").item(0).text

			    If result = "true" Then
					sqlStr = ""
					sqlStr = sqlStr & " UPDATE R" & VBCRLF
					sqlStr = sqlStr & " SET accFailCNT=0" & VBCRLF                 ''실패회수 초기화
				    sqlStr = sqlStr & " , shoplinkerLastUpdate = getdate()" & VBCRLF
					sqlStr = sqlStr & " ,shoplinkerPrice = '"&MustPrice&"'" & VBCRLF
					sqlStr = sqlStr & " ,shoplinkerSellYn = i.sellyn" & VBCRLF              ''MDT 일경우 결과로 저장할지 확인.
					sqlStr = sqlStr & " FROM db_item.dbo.tbl_shoplinker_regItem as R" & VBCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid" & VBCRLF
					sqlStr = sqlStr & " WHERE R.itemid = "&prdno&""   & VBCRLF
					dbget.Execute sqlStr,AssignedRow

					sqlStr = ""
					sqlStr = sqlStr & "UPDATE db_item.dbo.tbl_Shoplinker_Outmall SET lastupdate = getdate() WHERE itemid = '"&prdno&"' and mall_product_id = '"&product_id&"' "
					dbget.Execute sqlStr
				Else
					Call Fn_AcctFailTouch(CMALLNAME, prdno, errorMsg)
		    	End if

				If (isShoplinker_DebugMode) Then
					rw prdno &"_"&mode&"_"&errorMsg
				End If
			Next
			saveCommonItemResult=true
		End If
		Set Nodes = Nothing
	Else
		If (Not (retDoc is Nothing)) Then
			result = retDoc.getElementsByTagName("result").item(0).text
			errorMsg = retDoc.getElementsByTagName("message").item(0).text
		    If result = "true" Then
				successYn= true
		    Else
				successYn= false
	    	End if
		End If
	End If

	If mode <> "MDT" Then
		If (successYn = true) Then
			If mode = "REG" Then
				shoplinkerPrdno = retDoc.getElementsByTagName("product_id").item(0).text
	
				sqlStr = ""
				sqlStr = sqlStr & " SELECT i.itemid, i.limitno ,i.limitsold, i.sellcash, o.itemoption, o.optionname, o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, o.optaddprice, i.sellyn, i.limityn " & VBCRLF
				sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
				sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
				sqlStr = sqlStr & " WHERE i.itemid = "&prdno&" " & VBCRLF
				sqlStr = sqlStr & " ORDER BY o.itemoption ASC "
				rsget.Open sqlStr,dbget,1
				If Not(rsget.EOF or rsget.BOF) Then
					If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then
						Titemid			= rsget("itemid")
						Titemoption 	= "0000"
						Toptionname		= "단일상품"
						Tlimitno		= rsget("limitno")
						Tlimitsold		= rsget("limitsold")
						Tlimityn		= rsget("limityn")
						Tsellyn			= rsget("sellyn")
	
						If (Tlimityn="Y") then
							If (Tlimitno - Tlimitsold) < 5 Then
								Titemsu = 0
							Else
								Titemsu = Tlimitno - Tlimitsold - 5
							End If
						Else
							Titemsu = 900
						End If
	
						sqlStr = ""
						sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						sqlStr = sqlStr & " VALUES " & VBCRLF
						sqlStr = sqlStr & " ('"&Titemid&"',  '"&Titemoption&"', '"&CMALLNAME&"', '', '"&html2db(Toptionname)&"', '"&rsget("sellyn")&"', '"&rsget("limityn")&"', '"&Titemsu&"', '0', getdate()) "
						dbget.Execute sqlStr
					Else
						For i = 1 to rsget.RecordCount
							Titemid			= rsget("itemid")
							Tlimitno		= rsget("limitno")
							Tlimitsold		= rsget("limitsold")
							Titemoption		= rsget("itemoption")
							Toptionname		= rsget("optionname")
							Toptlimitno		= rsget("optlimitno")
							Toptlimitsold	= rsget("optlimitsold")
							Toptsellyn		= rsget("optsellyn")
							Toptlimityn		= rsget("optlimityn")
							Toptaddprice	= rsget("optaddprice")
							Tsellcash		= rsget("sellcash")
	
							If (Toptlimityn="Y") then
								If (Toptlimitno - Toptlimitsold) < 5 Then
									Titemsu = 0
								Else
									Titemsu = Toptlimitno - Toptlimitsold - 5
								End If
							Else
								Titemsu = 900
							End If
	
							If Left(Titemoption, 1) = "Z" Then
								Toptionname = Replace(Toptionname, ",", "/")
							End If
	
							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
							sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
							sqlStr = sqlStr & " VALUES " & VBCRLF
							sqlStr = sqlStr & " ('"&Titemid&"',  '"&Titemoption&"', '"&CMALLNAME&"', '', '"&html2db(Toptionname)&"', '"&Toptsellyn&"', '"&Toptlimityn&"', '"&Titemsu&"', '"&Toptaddprice+Tsellcash&"', getdate()) "
							dbget.Execute sqlStr
							rsget.MoveNext
						Next
					End If
				End If
				rsget.Close
			End If
	
			If (prdno <> "") Then
				sqlStr = ""
				sqlStr = sqlStr & " SELECT sellcash, buycash, orgprice,makerid FROM db_item.dbo.tbl_item where itemid = "&prdno&"  " & VBCRLF
				rsget.Open sqlStr,dbget,1
				If Not(rsget.EOF or rsget.BOF) Then
					If CLng(10000 - rsget("buycash") / rsget("sellcash") * 100 * 100) / 100 < CMAXMARGIN Then
						MustPrice = rsget("orgprice")
					Else
						MustPrice = rsget("sellcash")
					End If

					If rsget("makerid") = "KLING" Then
						MustPrice = rsget("sellcash")
					End If
				End If
				rsget.Close
	
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE R" & VBCRLF
				sqlStr = sqlStr & " SET accFailCNT=0" & VBCRLF                 ''실패회수 초기화
	
				If (mode = "REG") or (mode = "MDT") or (mode = "EDT") or (mode = "REGP") Then
				    sqlStr = sqlStr & " , shoplinkerLastUpdate = getdate()" & VBCRLF
					If (mode = "REGP") Then
						sqlStr = sqlStr & " , insert_infoCD = 'Y' " & VBCRLF
					End If
	            End If
	
				If (mode = "REG") Then
					sqlStr = sqlStr & " ,shoplinkerGoodNo='"&shoplinkerPrdno&"'"
					sqlStr = sqlStr & " ,shoplinkerStatCd=(CASE WHEN isNULL(shoplinkerStatCd, -1) < 3 then 3 ELSE shoplinkerStatCd END)"        ''등록완료(외부몰 미연결)
					sqlStr = sqlStr & " ,shoplinkerRegdate=isNULL(shoplinkerRegdate,getdate())" & VbCrlf
				End If
	
				If (mode = "MDT") or (mode = "REG") Then
					sqlStr = sqlStr & " ,shoplinkerPrice = '"&MustPrice&"'" & VBCRLF
				End If
	
				If (mode = "SLD") Then
					sqlStr = sqlStr & " ,shoplinkerSellYn = 'N'" & VBCRLF
				Else
					If (mode = "MDT") or (mode = "REG") Then
						sqlStr = sqlStr & " ,shoplinkerSellYn = i.sellyn" & VBCRLF              ''MDT 일경우 결과로 저장할지 확인.
					End If
				End If
				sqlStr = sqlStr & " FROM db_item.dbo.tbl_shoplinker_regItem as R" & VBCRLF
				sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid" & VBCRLF
				sqlStr = sqlStr & " WHERE R.itemid = "&prdno&""   & VBCRLF
				dbget.Execute sqlStr,AssignedRow
	
				If (mode = "MDT") Then
					sqlStr = ""
					sqlStr = sqlStr & "UPDATE db_item.dbo.tbl_Shoplinker_Outmall SET lastupdate = getdate() WHERE itemid = '"&prdno&"' "
					dbget.Execute sqlStr
				End If
	
				saveCommonItemResult=true
			End If
		Else
			Call Fn_AcctFailTouch(CMALLNAME, prdno, errorMsg)
			sqlStr = ""
			sqlStr = sqlStr & " INSERT into db_log.dbo.tbl_interparkEdit_log" & VBCRLF
			sqlStr = sqlStr & " (itemid, interParkPrdNo, sellcash, buycash, sellyn, ErrMsg, logdate, mallid)" & VBCRLF
			sqlStr = sqlStr & " SELECT "&prdno & VBCRLF
			sqlStr = sqlStr & " ,'' ,i.sellcash, i.buycash, i.sellyn" & VBCRLF
			sqlStr = sqlStr & " ,convert(varchar(100), '"&html2db(errorMsg)&"')" & VBCRLF
			sqlStr = sqlStr & " ,getdate(), '"&CMALLNAME&"'" & VBCRLF
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i WHERE i.itemid="&prdno&"" & VBCRLF
			dbget.Execute sqlStr
	
			If (mode = "REG") Then
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE R" & VBCRLF
				sqlStr = sqlStr & " SET shoplinkerStatCd = -1"                   '''등록실패
				sqlStr = sqlStr & " FROM db_item.dbo.tbl_shoplinker_regitem as R" & VBCRLF
				sqlStr = sqlStr & " WHERE R.itemid = "&prdno&"" & VBCRLF
				sqlStr = sqlStr & " and shoplinkerStatCd = 1"                    ''전송
				dbget.Execute sqlStr
			End If
		End If

		If (isShoplinker_DebugMode) Then
			rw prdno &"_"&mode&"_"&errorMsg
		End If
	End If


End Function

Function saveSearchItemResult(retDoc, mode, prdno)
	Dim errorMsg, count_all, i
	Dim AssignedRow, successYn
	Dim Nodes, OneNode, SubNodes
	Dim partner_product_id, shoplinker_id, product_id, mall_name, mall_user_id, mall_product_id, mall_product_name, supply_price, sale_price
	Dim sqlStr, ControlMakerid
	AssignedRow = 0
	successYn = false

	If (Not (retDoc is Nothing)) Then
		count_all = retDoc.getElementsByTagName("count_all").item(0).text
	    If count_all > 0 Then
			successYn= true
	    Else
			successYn= false
			rw iitemid &":"& " 연결된 외부몰이 없습니다"
			saveSearchItemResult = False
			Exit Function
    	End if
	End If

	If (successYn = true) Then
		For i = 0 To count_all - 1
			partner_product_id	= retDoc.getElementsByTagName("partner_product_id").item(i).text		'고객사 상품코드
			shoplinker_id		= retDoc.getElementsByTagName("shoplinker_id").item(i).text				'샵링커 아이디
			product_id			= retDoc.getElementsByTagName("product_id").item(i).text				'샵링커 상품코드
			mall_name			= retDoc.getElementsByTagName("mall_name").item(i).text					'쇼핑몰명
			mall_user_id		= retDoc.getElementsByTagName("mall_user_id").item(i).text				'쇼핑몰 아이디
			mall_product_id		= retDoc.getElementsByTagName("mall_product_id").item(i).text			'쇼핑몰 상품코드
			mall_product_name	= retDoc.getElementsByTagName("mall_product_name").item(i).text			'쇼핑몰 상품명
			supply_price		= retDoc.getElementsByTagName("supply_price").item(i).text				'공급가
			sale_price			= retDoc.getElementsByTagName("sale_price").item(i).text				'판매가

			sqlStr = ""
			sqlStr = sqlStr & " SELECT TOP 1 makerid FROM db_item.dbo.tbl_Shoplinker_OutmallControl where mall_user_id = '"&mall_user_id&"' and mall_name = '"&mall_name&"' "
			rsget.open sqlStr, dbget, 1
			If Not(rsget.EOF or rsget.BOF) Then
				ControlMakerid = rsget("makerid")
			Else
				response.write "<script language='javascript'>alert('"&prdno&":"&mall_name&"이 설정되어있지 않습니다.');</script>"
				saveSearchItemResult=false
				rw prdno & ":" & mall_name&"이 설정되어있지 않습니다."
'				exit function
			End If
			rsget.close

			sqlStr = ""
			sqlStr = sqlStr & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_Shoplinker_Outmall where mall_user_id = '"&mall_user_id&"' and itemid = '"&partner_product_id&"' and mall_product_id = '"&mall_product_id&"' )"
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_Shoplinker_Outmall "
		    sqlStr = sqlStr & " (itemid, mall_user_id, makerid, mall_name, shoplinker_id, product_id, mall_product_id, mall_product_name, supply_price, sale_price, regdate)"
		    sqlStr = sqlStr & " VALUES ('"&prdno&"', '"&mall_user_id&"', '"&ControlMakerid&"', '"&mall_name&"','"&shoplinker_id&"', '"&product_id&"', '"&mall_product_id&"', '"&mall_product_name&"', '"&supply_price&"', '"&sale_price&"', getdate())"
			sqlStr = sqlStr & " END "
		    dbget.Execute sqlStr
		     rw iitemid&" : "&mall_name&"코드 : "&mall_product_id
		Next

		If ControlMakerid <> "" Then
			sqlStr = ""
			sqlStr = sqlStr & " Update db_item.dbo.tbl_Shoplinker_regItem SET ShoplinkerOutMallConnect = 'Y', ShoplinkerStatCd = '7' WHERE itemid = '"&iitemid&"' "
			dbget.Execute sqlStr
			saveSearchItemResult=true
		End If
	End If
End Function

Function saveSellYNItemResult(retDoc, mode, iitemid, cmd)
	Dim errorMsg, count_all, i
	Dim AssignedRow, successYn
	Dim sqlStr, result, cmdval
	AssignedRow = 0
	successYn = false

	Select Case cmd
		Case "N"		cmdval = "품절"
		Case "Y"		cmdval = "판매중"
	End Select

	If (Not (retDoc is Nothing)) Then
		result = retDoc.getElementsByTagName("result").item(0).text
		errorMsg = retDoc.getElementsByTagName("message").item(0).text
	    If result = "true" Then
			successYn= true
	    Else
			successYn= false
    	End if
	End If

	If (successYn = true) Then
		If mode = "SLD" Then			'선택상품 상태수정
            sqlStr = ""                                                                 ''2013/06/20추가
            sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_Shoplinker_regItem" & VBCRLF
            sqlStr = sqlStr & " SET shoplinkerLastUpdate = getdate()" & VBCRLF
			If cmd = "N" Then
				sqlStr = sqlStr & " ,shoplinkerSellyn = 'N' " & VBCRLF
			ElseIf cmd = "Y" Then
				sqlStr = sqlStr & " ,shoplinkerSellyn = 'Y' " & VBCRLF
			End If
            sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"'  " & VBCRLF
            dbget.Execute sqlStr,AssignedRow
            rw iitemid&" : 판매상태 "&cmdval&"(으)로 변경"
		End If
		saveSellYNItemResult = true
	Else
		Call Fn_AcctFailTouch(CMALLNAME, iitemid, errorMsg)
		rw iitemid &"_"&mode&"_"&errorMsg
	End If
End Function

Function regShoplinkerOneItem(byval iitemid, byRef ierrStr, subcmd)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "REG"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, subcmd)
	Dim cause
	If (xmlStr = "") Then
		ierrStr = "등록불가"
		sqlStr = ""
		sqlStr = sqlStr & " SELECT i.itemid, isNULL(R.shoplinkerStatCD, -9) as shoplinkerStatCD " & VBCRLF
		sqlStr = sqlStr & " ,i.sellyn, i.limityn, i.limitno, i.limitsold, isnull(N.itemid,'') as Nitemid " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_Shoplinker_regItem as R on i.itemid = R.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_jaehyumall_not_in_itemid as N on i.itemid=N.itemid and N.mallgubun = '"&CMALLNAME&"' " & VBCRLF
		sqlStr = sqlStr & " WHERE i.itemid = "&iitemid
		rsget.Open sqlStr,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
			If (rsget("shoplinkerStatCD") >= 3) Then
				ierrStr = ierrStr & " - 기등록상품"&" :: 상태["&rsget("shoplinkerStatCD")&"]"
			End If

			If (rsget("sellyn") <> "Y") Then
			    ierrStr = ierrStr & " - 품절상태"
			End If

			If (rsget("limityn") = "Y") and (rsget("limitno") - rsget("limitsold") < CMAXLIMITSELL) Then
				ierrStr = ierrStr & " - 한정수량 부족 ("&rsget("limitno")-rsget("limitsold")&") 개 남음"
				cause = "limitErr"
			End If

			If (rsget("Nitemid") <> "0") Then
			    ierrStr = ierrStr & " - 등록제외상품"
			End If
	    Else
			ierrStr = ierrStr & " - 상품조회불가"
	    End If
		rsget.Close

		''불가 사유를 못찾을 경우
		If (ierrStr = "등록불가") Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT itemid" & VBCRLF
			sqlStr = sqlStr & " ,count(*) as optCNT" & VBCRLF
			sqlStr = sqlStr & " ,sum(CASE WHEN optAddPrice > 0 then 1 ELSE 0 END) as optAddCNT" & VBCRLF
			sqlStr = sqlStr & " ,sum(CASE WHEN (optsellyn = 'N') or (optlimityn = 'Y' and (optlimitno - optlimitsold < 1)) then 1 ELSE 0 END) as optNotSellCnt" & VBCRLF
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option" & VBCRLF
			sqlStr = sqlStr & " WHERE itemid="&iitemid & VBCRLF
			sqlStr = sqlStr & " and isusing='Y'" & VBCRLF
			sqlStr = sqlStr & " GROUP BY itemid"
			rsget.Open sqlStr,dbget,1
			If Not(rsget.EOF or rsget.BOF) Then
'				If (rsget("optAddCNT") > 0) Then
'					ierrStr = ierrStr & " - 옵션추가 금액 존재상품 등록불가"
'					cause = "optAddPrcExist"
'				End If

				If (rsget("optCnt") - rsget("optNotSellCnt") < 1) Then
					ierrStr = ierrStr & " - 옵션 판매가능상품 없음."
					cause = "noValidOpt"
				End If
			End If
			rsget.Close
		End If

		If (ierrStr <> "등록불가") Then
			ierrStr = iitemid &":"& ierrStr
		End If
		regShoplinkerOneItem = False
		Exit Function
    End If

    IF (isShoplinker_DebugMode) Then
        CALL XMLFileSave(xmlStr, mode, iitemid)
    End If

	sqlStr = ""
	sqlStr = sqlStr & " IF Exists(SELECT * FROM db_item.dbo.tbl_Shoplinker_regItem where itemid="&iitemid&")"
	sqlStr = sqlStr & " BEGIN"& VbCRLF
	sqlStr = sqlStr & " UPDATE R" & VbCRLF
	sqlStr = sqlStr & "	SET ShoplinkerLastUpdate = getdate() "  & VbCRLF
	sqlStr = sqlStr & "	, shoplinkerStatCD='1'"& VbCRLF
	sqlStr = sqlStr & "	FROM db_item.dbo.tbl_Shoplinker_regItem R"& VbCRLF
	sqlStr = sqlStr & " WHERE R.itemid='" & iitemid & "'"
	sqlStr = sqlStr & " END ELSE "
	sqlStr = sqlStr & " BEGIN"& VbCRLF
	sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_Shoplinker_regItem "
    sqlStr = sqlStr & " (itemid, regdate, reguserid, shoplinkerStatCD, insert_infoCD, ShoplinkerOutMallConnect)"
    sqlStr = sqlStr & " VALUES ("&iitemid&", getdate(), '"&session("SSBctID")&"', '1', 'N', 'N')"
	sqlStr = sqlStr & " END "
    dbget.Execute sqlStr

	AssignedRow = 0
	Dim retDoc, sURL
	sURL = shoplinkerAPIURL&"/Product/xmlInsert.php"

	SET retDoc = xmlSend(sURL, XMLURL)
	    If (isShoplinker_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML,"RET_"&mode, iitemid)
	    End If
	    regshoplinkerOneItem = saveCommonItemResult(retDoc, mode, iitemid)
	SET retDoc = Nothing
End Function

Function edtShoplinkerOneItem(byval iitemid, byRef ierrStr, subcmd)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "EDT"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, subcmd)
	Dim cause
	If (xmlStr = "") Then
		ierrStr = "수정불가"
		ierrStr = iitemid &":"& ierrStr
		edtShoplinkerOneItem = False
		Exit Function
    End If

    IF (isShoplinker_DebugMode) Then
        CALL XMLFileSave(xmlStr, mode, iitemid)
    End If

	AssignedRow = 0
	Dim retDoc, sURL
	sURL = shoplinkerAPIURL&"/Product/xmlInsert.php"

	SET retDoc = xmlSend(sURL, XMLURL)
	    If (isShoplinker_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML,"RET_"&mode, iitemid)
	    End If
	    edtShoplinkerOneItem = saveCommonItemResult(retDoc, mode, iitemid)
	SET retDoc = Nothing
End Function

Function regShoplinkerPoomOK(byval iitemid, byRef ierrStr)
	Dim sqlStr, shoplinkerPrdno
	shoplinkerPrdno = getShoplinkerPrdno(iitemid)

	If shoplinkerPrdno = "" Then
		ierrStr = "["&iitemid&"]등록불가 - Shoplinker 상품코드가 없습니다."
		Exit Function
	End If

	Dim mode : mode = "REGP"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, shoplinkerPrdno)
	Dim cause

    IF (isShoplinker_DebugMode) Then
        CALL XMLFileSave(xmlStr, mode, iitemid)
    End If

	Dim retDoc, sURL
	sURL = shoplinkerAPIURL&"/Product/goods_info_reg.php"

	SET retDoc = xmlSend(sURL, XMLURL)
	    If (isShoplinker_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML,"RET_"&mode, iitemid)
	    End If
	    regShoplinkerPoomOK = saveCommonItemResult(retDoc, mode, iitemid)
	SET retDoc = Nothing
End Function

Function editSellStatusShoplinkerOneItem(byval iitemid, byRef ierrStr, cmd)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "SLD"
	Dim xmlStr : xmlStr = getXMLSellyn(iitemid,mode,cmd)

	If (xmlStr = "") Then
		ierrStr = "수정불가"
		editSellStatusShoplinkerOneItem = False
		Exit Function
	End If

	If (isShoplinker_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

	Dim retDoc, sURL
	sURL = shoplinkerAPIURL&"/Product/Shopmall_soldout.html"
	Set retDoc = xmlSend (sURL, XMLURL)
		If (isShoplinker_DebugMode) Then
			CALL XMLFileSave(retDoc.XML, "RET_"&mode, iitemid)
		End If
		Call saveSellYNItemResult(retDoc, mode, iitemid, cmd)
	Set retDoc = Nothing
End Function

Function editOutmallShoplinkerOneItem(byval iitemid, byRef ierrStr, byval mallprdid)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "MDT"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, mallprdid)
	Dim cause
	If (xmlStr = "") Then
		ierrStr = "수정불가"
		ierrStr = iitemid &":"& ierrStr
		editOutmallShoplinkerOneItem = False
		Exit Function
    End If

    IF (isShoplinker_DebugMode) Then
        CALL XMLFileSave(xmlStr, mode, iitemid)
    End If

	AssignedRow = 0
	Dim retDoc, sURL
	sURL = shoplinkerAPIURL&"/Product/OpenMarket_soldout.html"

	SET retDoc = xmlSend(sURL, XMLURL)
	    If (isShoplinker_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML,"RET_"&mode, iitemid)
	    End If
	    editOutmallShoplinkerOneItem = saveCommonItemResult(retDoc, mode, iitemid)
	SET retDoc = Nothing
End Function

Function ShoplinkerSearchItem(byval iitemid, byRef ierrStr)
	Dim sqlStr, AssignedRow, shoplinkerPrdno
	shoplinkerPrdno = getShoplinkerPrdno(iitemid)

	Dim mode : mode = "SCH"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, shoplinkerPrdno)

	If (xmlStr = "") Then
		ierrStr = "검색불가"
		ShoplinkerSearchItem = False
		Exit Function
	End If

	If (isShoplinker_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

	Dim retDoc, sURL
	sURL = shoplinkerAPIURL&"/Product/mall_product_list.php"
	Set retDoc = xmlSend (sURL, XMLURL)
		If (isShoplinker_DebugMode) Then
			CALL XMLFileSave(retDoc.XML, "RET_"&mode, iitemid)
		End If
		ShoplinkerSearchItem = saveSearchItemResult(retDoc, mode, iitemid)
	Set retDoc = Nothing
End Function

Function getXMLSellyn(byval iitemid, mode, cmd)
	Dim oShoplinkerItem
	Dim strRst
	SET oShoplinkerItem = new CShoplinker
		oShoplinkerItem.FRectMode = mode
		oShoplinkerItem.FRectItemID = iitemid
		oShoplinkerItem.FRectMatchCateNotCheck = "on"
		oShoplinkerItem.getShoplinkerEditedItemList
		If (oShoplinkerItem.FREsultCount > 0) Then
			oShoplinkerItem.FItemList(0).FSellYN = cmd
			getXMLSellyn = oShoplinkerItem.FItemList(0).getshoplinkerItemSellStatusDTXML()
		End If
	SET oShoplinkerItem = Nothing
End Function

Function getXMLString(byval iitemid, mode, paramA)
	Dim oShoplinkerItem
	Dim strRst, bufRET, buf1, notitemId, notmakerid

	SET oShoplinkerItem = new CShoplinker
		oShoplinkerItem.FRectMode = mode
		oShoplinkerItem.FRectItemID = iitemid

		'2013-09-23 김지웅 과장님 요청 // 특정브랜드만 한정 -5 풀어달라고 함 snurk추가위해 oShoplinkerItem.getShoplinkerNot5RegItemList 새로 생성
		'2013-11-01 김지웅 과장님 요청 // 특정브랜드만 한정 -5 풀어달라고 함 KLING추가
		'2013-11-11 홍영주님 요청 // 특정브랜드만 한정 -5 풀어달라고 함 sushi추가
		'2013-12-03 13:42분 김진영 수정..예외브랜드 추가..예외브랜드란:한정5개미만이라도 등록가능..
		'위 요청에따라 snurk, KLING, sushi 김진영 예외브랜드 등록
		Dim strSQL, etcCnt
		strSQL = ""
		strSQL = strSQL & " SELECT count(m.makerid) as etcCnt "
		strSQL = strSQL & " FROM db_item.dbo.tbl_item as i "
		strSQL = strSQL & " JOIN [db_temp].dbo.tbl_shoplinker_not_in_makerid as M on i.makerid = M.makerid "
		strSQL = strSQL & " WHERE M.isusing = 'Y' "
		strSQL = strSQL & " and M.mallgubun = 'shoplinker' "
		strSQL = strSQL & " and i.itemid in ("&iitemid&") "
		rsget.OPEN strSQL, dbget, 1
			etcCnt = rsget("etcCnt")
		rsget.CLOSE

		If (mode = "REG") Then
			If etcCnt > 0 Then
				oShoplinkerItem.getShoplinkerNot5RegItemList
			Else
				oShoplinkerItem.getShoplinkerNotRegItemList
			End If

			If (oShoplinkerItem.FResultCount > 0) Then
				getXMLString = oShoplinkerItem.FItemList(0).getShoplinkerItemRegXML(paramA)
			End If
		ElseIf (mode = "EDT") Then
			oShoplinkerItem.getShoplinkerEditedItemList
			If (oShoplinkerItem.FResultCount > 0) Then
				getXMLString = oShoplinkerItem.FItemList(0).getShoplinkerItemRegXML(paramA)
			End If
		ElseIf (mode = "REGP") Then
			getXMLString = getShoplinkerPoomOK(iitemid, paramA)
		ElseIf (mode = "SCH") Then
			getXMLString = getSearchITEM(iitemid, paramA)
		ElseIf (mode = "MDT") Then
			If etcCnt > 0 Then
				rw "itemid:" & iitemid & "는 예외브랜드에 속하므로 수정하지 않습니다."
				Exit Function
				response.end
			Else
				oShoplinkerItem.getShoplinkerEditedItemList
			End If

			If (oShoplinkerItem.FResultCount > 0) Then
				getXMLString = oShoplinkerItem.FItemList(0).getOutmallItemEdtXML(paramA)
			End If
		End If
	SET oShoplinkerItem = Nothing
	If (mode = "SCHCate") Then
		getXMLString = getSearchCate()
	End If
End Function

Function XMLFileSave(xmlStr, mode, iitemid)
'   Exit function  ''로그 안남김

	Dim fso,tFile
	Dim opath
	Select Case mode
		Case "REG", "RET_REG", "REGP", "RET_REGP"
			opath = "/admin/etc/shoplinker/xmlFiles/INSERT/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "SCH", "RET_SCH"
			opath = "/admin/etc/shoplinker/xmlFiles/SELECT/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "EDT", "RET_EDT", "MDT", "RET_MDT"
			opath = "/admin/etc/shoplinker/xmlFiles/UPDATE/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "SLD", "RET_SLD"
			opath = "/admin/etc/shoplinker/xmlFiles/UPDATE_SellStatus/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	End Select

	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName
	If mode = "LIST" or mode = "SCHCate" Then
		fileName = mode &"_"& getCurrDateTimeFormat& ".xml"
	Else
		fileName = mode &"_"& getCurrDateTimeFormat&"_"&iitemid&".xml"
	End If

	Select Case mode
		Case "REG", "RET_REG", "REGP", "RET_REGP"
			IF application("Svr_Info")="Dev" THEN
				XMLURL = "?iteminfo_url=http://61.252.133.2:8888"&opath&FileName
			Else
				XMLURL = "?iteminfo_url=http://webadmin.10x10.co.kr"&opath&FileName
			End If
		Case "SCH", "RET_SCH"
			IF application("Svr_Info")="Dev" THEN
				XMLURL = "?iteminfo_url=http://61.252.133.2:8888"&opath&FileName
			Else
				XMLURL = "?iteminfo_url=http://webadmin.10x10.co.kr"&opath&FileName
			End If
		Case "EDT", "RET_EDT", "MDT", "RET_MDT"
			IF application("Svr_Info")="Dev" THEN
				XMLURL = "?iteminfo_url=http://61.252.133.2:8888"&opath&FileName
			Else
				XMLURL = "?iteminfo_url=http://webadmin.10x10.co.kr"&opath&FileName
			End If
		Case "SLD", "RET_SLD"
			IF application("Svr_Info")="Dev" THEN
				XMLURL = "?iteminfo_url=http://61.252.133.2:8888"&opath&FileName
			Else
				XMLURL = "?iteminfo_url=http://webadmin.10x10.co.kr"&opath&FileName
			End If
	End Select

	CALL CheckFolderCreate(defaultPath)
	''debug
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.Write(xmlStr)
			tFile.Close
		Set tFile = nothing
	Set fso = nothing
End Function

Function XMLSend(url, xmlStr)
	Dim poster, SendDoc, retDoc

	Set poster= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		poster.open "GET", url&xmlStr, false
		poster.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'		poster.setTimeouts 5000,90000,90000,90000  ''2013/07/22 추가
		poster.Send()

	Set retDoc = server.createobject("MSXML2.DomDocument.3.0")
		retDoc.async = False
		retDoc.LoadXML BinaryToText(poster.ResponseBody, "euc-kr")

	Set XMLSend = retDoc
	Set SendDoc = Nothing
	Set poster = Nothing
End Function

Function getCurrDateTimeFormat()
    dim nowtimer : nowtimer= timer()
    getCurrDateTimeFormat = left(now(),10)&"_"&nowtimer
End Function

Sub CheckFolderCreate(sFolderPath)
	Dim objfile
	Set objfile=Server.CreateObject("Scripting.FileSystemObject")
		IF NOT  objfile.FolderExists(sFolderPath) THEN
			objfile.CreateFolder sFolderPath
		END IF
	Set objfile=Nothing
End Sub

Function Fn_AcctFailTouch(iMallID, iitemid, iLastErrStr)
	Dim strSql
	iLastErrStr = html2db(iLastErrStr)

	If (iMallID = "shoplinker") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_item.dbo.tbl_shoplinker_regItem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	End If
End Function

Function getShoplinkerPoomOK(iitemid, shoplinerGoodno)
	Dim strRst, strSQL
	Dim mallinfoDiv, mallinfoCdAll, mallinfoCd, infoCDVal, infocd
	Dim getsafetyYn, getsafetyDiv, getsafetyNum

	strSql = ""
	strSql = strSql & "select top 1 safetyYn, safetyDiv, safetyNum from db_item.dbo.tbl_item_contents where itemid = '"&iitemid&"'"
	rsget.Open strSql,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		getsafetyYn		= rsget("safetyYn")
		getsafetyDiv	= rsget("safetyDiv")
		getsafetyNum	= rsget("safetyNum")
	End If
	rsget.Close

	strSql = ""
	strSql = strSql & " SELECT top 100 M.* " & vbcrlf
	strSql = strSql & " ,isNULL(CASE WHEN M.infocd='00000' then 'N' "
	strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='N' THEN '해당없음' " & vbcrlf
	strSql = strSql & " 	  WHEN c.infotype='P' THEN replace(c.infoDesc,'1644-6030','1644-6035') " & vbcrlf
	strSql = strSql & "  ELSE F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
	strSql = strSql & "  END,'') as infoCDVal " & vbcrlf
	strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
	strSql = strSql & " Join db_item.dbo.tbl_item_contents IC on IC.infoDiv=M.mallinfoDiv " & vbcrlf
	strSql = strSql & " LEFT Join db_item.dbo.tbl_item_infoCode c on M.infocd=c.infocd " & vbcrlf
	strSql = strSql & " LEFT Join db_item.dbo.tbl_item_infoCont F on M.infocd=F.infocd and F.itemid=" & iitemid & vbcrlf
	strSql = strSql & " LEFT join db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd=F2.infocd and F2.itemid=" & iitemid & vbcrlf
	strSql = strSql & " WHERE M.mallid = '"&CMALLNAME&"' AND IC.itemid=" & iitemid
	strSql = strSql & " ORDER BY M.mallinfoCd"
	rsget.Open strSql,dbget,1
	If Not(rsget.EOF or rsget.BOF) then
		mallinfoDiv = rsget("mallinfoDiv")
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst &"<openMarket>"
		strRst = strRst &"<goodsinfo>"
		strRst = strRst &"<customer_id>"&CUSTOMERID&"</customer_id>"					'샵링커 고객사 번호
		strRst = strRst &"<product_id>"&shoplinerGoodno&"</product_id>"					'샵링커 상품코드
		strRst = strRst &"<partner_product_id>"&iitemid&"</partner_product_id>"			'고객사(자체) 상품코드
		strRst = strRst &"<lclass_id>i"&mallinfoDiv&"</lclass_id>"						'샵링커 품목 대분류코드
		Do until rsget.EOF
			mallinfoCd = rsget("mallinfoCd")
			infoCDVal = rsget("infoCDVal")
			infocd	  = rsget("infocd")
			If infocd = "00000" Then
				If (getsafetyYn="Y" and getsafetyDiv<>0) Then
					If (getsafetyDiv=10) Then
						infoCDVal = "1<**>"&getsafetyNum
					Elseif (getsafetyDiv=20) Then
						infoCDVal = "2<**>"&getsafetyNum
					Elseif (getsafetyDiv=30) Then
						infoCDVal = "3<**>"&getsafetyNum
					Elseif (getsafetyDiv=40) Then
						infoCDVal = "4<**>"&getsafetyNum
					Elseif (getsafetyDiv=50) Then
						infoCDVal = "5<**>"&getsafetyNum
					End if
				End If
			End If
			strRst = strRst &"<item>"
			strRst = strRst &"	<item_seq>"&mallinfoCd&"</item_seq>"
			strRst = strRst &"	<item_info><![CDATA["&infoCDVal&"]]></item_info>"
			strRst = strRst &"</item>"
			rsget.MoveNext
		Loop
		strRst = strRst &"</goodsinfo>"
		strRst = strRst &"</openMarket>"
	End If
	rsget.Close
	getShoplinkerPoomOK = strRst
End Function

Function getSearchITEM(iitemid, shoplinerGoodno)
	Dim strRst
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst &"<Shoplinker>"
		strRst = strRst &"<MessageHeader>"
		strRst = strRst &"	<sendID>10x10</sendID>"
		strRst = strRst &"	<senddate>"&replace(date(),"-","")&"</senddate>"
		strRst = strRst &"</MessageHeader>"
		strRst = strRst &"<productInfo>"
		strRst = strRst &"<Product>"
		strRst = strRst &"<customer_id>"&CUSTOMERID&"</customer_id>"						'샵링커 고객사 번호
		strRst = strRst &"<st_date>"&replace(dateadd("m", -6, date()), "-", "")&"</st_date>"'6개월전으로 지정
		strRst = strRst &"<ed_date>"&replace(date(),"-","")&"</ed_date>"					'현재날짜
		strRst = strRst &"<partner_product_id>"&iitemid&"</partner_product_id>"
'		strRst = strRst &"<partner_product_id>885499</partner_product_id>"
'		strRst = strRst &"<page>1</page>"
'		strRst = strRst &"<mall_id>APISHOP_0000</mall_id>"
		strRst = strRst &"</Product>"
		strRst = strRst &"</productInfo>"
		strRst = strRst &"</Shoplinker>"
		getSearchITEM = strRst
End Function

Function getShoplinkerPrdno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT top 1 isnull(shoplinkerGoodNo,'') as shoplinkerGoodNo FROM db_item.dbo.tbl_shoplinker_regItem WHERE itemid = '"&iitemid&"' "
	rsget.Open strSql,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		getShoplinkerPrdno = rsget("shoplinkerGoodNo")
	End If
	rsget.Close
End Function

Function getMallPrdid(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT mall_product_id, mall_name FROM db_item.dbo.tbl_Shoplinker_Outmall WHERE itemid = '"&iitemid&"' "
	rsget.Open strSql,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		getMallPrdid = rsget.getRows() 
	End If
	rsget.Close
End Function
%>