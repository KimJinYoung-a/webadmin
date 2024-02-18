<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
''롯데아이몰 단품정보 조회
Function LtiMallOneItemCheckStock(iitemid,byRef iErrStr)
	Dim ilottegoods_no
	Dim objXML,xmlDOM,strRst, iMessage
	Dim ProdCount, buf, AssignedRow, oneProdInfo, strParam
	Dim GoodNo,ItemNo,OptDesc,DispYn,SaleStatCd,StockQty, bufopt
	Dim strSql, actCnt, CorpItemNo, getRegOptCD

	On Error Resume Next
	LtiMallOneItemCheckStock = False
	ilottegoods_no = getLtiMallItemIdByTenItemID(iitemid)
	If (ilottegoods_no="") then
	    iErrStr = "["&iitemid&"] 롯데 아이몰 코드 없음."
	    Exit Function
	End If

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		strParam = "?subscriptionId=" & ltiMallAuthNo				'롯데아이몰 인증번호	(*)
		strParam = strParam & "&search_type=goods_no"
		strParam = strParam & "&search_value=" & ilottegoods_no		'롯데아이몰 상품번호	(*)

		objXML.Open "GET", ltiMallAPIURL & "/openapi/searchStockList.lotte"&strParam, false
''rw ltiMallAPIURL & "/openapi/searchStockList.lotte"&strParam
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML replace(buf,"&","＆")
				If Err <> 0 Then
					iErrStr =  "롯데아이몰 결과 분석 중에 오류가 발생 LtiMallOneItemCheckStock[" & iitemid & "]:"
					Set objXML = Nothing
					Set xmlDOM = Nothing
					Exit function
				End If
				ProdCount   = Trim(xmlDOM.getElementsByTagName("GoodsCount").item(0).text)   '' 단품 갯수

				'If (ProdCount = "1") Then                                                   ''2013/07/03 주석처리
				'	If (getItemOptionCount(iitemid) < 1) Then ProdCount = ""
				'End If
				If (ProdCount <> "") Then
			        Set oneProdInfo = xmlDOM.getElementsByTagName("GoodsInfoList")
			        strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption='')"
			        strSql = strSql & " BEGIN"
			        strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption=''"
			        strSql = strSql & " END"
			        dbget.Execute strSql

			        ''2013/05/30 추가
			        strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and Len(outmalloptCode)>6)"
			        strSql = strSql & " BEGIN"
			        strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and Len(outmalloptCode)>6"
			        strSql = strSql & " END"
			        dbget.Execute strSql
					For each SubNodes in oneProdInfo
						GoodNo	    = Trim(SubNodes.getElementsByTagName("GoodNo").item(0).text)
					    ItemNo	    = Trim(SubNodes.getElementsByTagName("ItemNo").item(0).text)        '' 단품코드 (숫자 0,1,2,)
					    OptDesc	    = Trim(SubNodes.getElementsByTagName("OptDesc").item(0).text)
					    ''DispYn	    = Trim(SubNodes.getElementsByTagName("DispYn").item(0).text)         ''N:안함 Y:전시            ''2013/07/03 제대로 안넘어오는거 같아 SaleStatCd 로변경
					    SaleStatCd	= Trim(SubNodes.getElementsByTagName("SaleStatCd").item(0).text) ''판매진행, 판매종료, 품절'
					    StockQty	= Trim(SubNodes.getElementsByTagName("StockQty").item(0).text)
                        CorpItemNo  = Trim(SubNodes.getElementsByTagName("CorpItemNo").item(0).text)  '' 상품코드_옵션코드

						getRegOptCD = Split(CorpItemNo,"_")(1)
					    OptDesc = replace(OptDesc, "＆", "&")
					    If (SaleStatCd <> "10") Then
					        DispYn = "N"
					    else
					        DispYn = "Y"
					    End If

					    If (StockQty = "null") Then
					        StockQty = "0"
					    End If

					    bufopt = OptDesc
						If InStr(bufopt,",") > 0 then
						    If (splitValue(bufopt,",",0) <> "") Then
						        OptDesc = splitValue(splitValue(bufopt,",",0),":",1)
						    End If

						    If (splitValue(bufopt,",",1) <> "") Then
						        OptDesc = OptDesc+","+splitValue(splitValue(bufopt,",",1),":",1)
						    End If

						    If (splitValue(bufopt,",",2)<>"") Then
						        OptDesc = OptDesc+","+splitValue(splitValue(bufopt,",",2),":",1)
						    End If
						Else
							OptDesc = splitValue(OptDesc,":",1)
						End If

						'################  2014-02-21 14:34 김진영 ######################
						'OptDesc = replace(OptDesc, ",,", ",")라인 추가..
						'이유 : [db_item].[dbo].tbl_item_option_Multiple 에 optionTypeName에 ,가 들어가 있는 경우가 있음..
 						'ex)이니셜각인(5,000원)..이렇게 되어있을 때 ,를 split시킴에 따라 ,,를 ,로 치환함
 						OptDesc = replace(OptDesc, ",,", ",")
 						'#################################################################

						rw GoodNo&"|"&ItemNo&"|"&CorpItemNo&"|"&OptDesc&"|"&DispYn&"|"&SaleStatCd&"|"&StockQty
						strSql = ""
						strSql = strSql & " UPDATE oP "
					    strSql = strSql & " SET outmallOptName='"&html2DB(OptDesc)&"'"&VbCRLF
						strSql = strSql & " ,outmallOptCode='"&ItemNo&"'"&VbCRLF
					    strSql = strSql & " ,lastupdate=getdate()"&VbCRLF
					    strSql = strSql & " ,outMallSellyn='"&DispYn&"'"&VbCRLF
					    strSql = strSql & " ,outmalllimityn='Y'"&VbCRLF
					    strSql = strSql & " ,outMallLimitNo="&StockQty&VbCRLF
					    strSql = strSql & " ,checkdate=getdate()"&VbCRLF
					    strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
					    strSql = strSql & " WHERE itemid="&iitemid&VbCRLF
					    strSql = strSql & " and convert(int, outmallOptCode)='"&ItemNo&"'"&VbCRLF				'개편전 outmallOptCode는 001,002,003 이렇게 들어가있으나 개편 후엔 1,2,3이렇게 변함
					    strSql = strSql & " and mallid='lotteimall'"&VbCRLF
					    dbget.Execute strSql, AssignedRow
					    If (AssignedRow < 1) Then
					        strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption='')"
					        strSql = strSql & " BEGIN"
					        strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption=''"
					        strSql = strSql & " END"
					        dbget.Execute strSql

					        strSql = " Insert into db_item.dbo.tbl_OutMall_regedoption"
					        strSql = strSql & " (itemid,itemoption,mallid,outmallOptCode,outmallOptName,outMallSellyn,outmalllimityn,outMallLimitNo,checkdate)"
					        strSql = strSql & " values("&iitemid
					        strSql = strSql & " ,'"&ItemNo&"'" ''임시로 롯데 코드 넣음 //2013/04/01
					        strSql = strSql & " ,'lotteimall'"
					        strSql = strSql & " ,'"&ItemNo&"'"
					        strSql = strSql & " ,'"&html2DB(OptDesc)&"'"
					        strSql = strSql & " ,'"&DispYn&"'"
					        strSql = strSql & " ,'Y'"
					        strSql = strSql & " ,"&StockQty
					        strSql = strSql & " ,getdate()"
					        strSql = strSql & ")"
					        dbget.Execute strSql, AssignedRow

							If getRegOptCD = "" Then
								Dim newOptSQL
								newOptSQL = ""
								newOptSQL = newOptSQL & " SELECT TOP 1 itemoption FROM [db_item].[dbo].tbl_item_option WHERE itemid = '"&iitemid&"' and optionname = '"&html2DB(OptDesc)&"' " 
								rsget.Open newOptSQL, dbget
								If Not(rsget.EOF or rsget.BOF) Then
									getRegOptCD = rsget("itemoption")
								Else
									getRegOptCD = "0000"	
								End If
								rsget.Close
							End If

					        ''옵션 코드 매칭.
					        If (AssignedRow > 0) Then
					            strSql = " update oP"   &VbCRLF
					            strSql = strSql & " set itemoption='"&getRegOptCD&"'"&VbCRLF
					            strSql = strSql & " From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
'					            strSql = strSql & "     Join db_item.dbo.tbl_item_option o"&VbCRLF
'					            strSql = strSql & "     on oP.itemid=o.itemid"&VbCRLF
					            strSql = strSql & " where oP.mallid='lotteimall'"&VbCRLF
'					            strSql = strSql & " and o.itemid="&iitemid&VbCRLF
					            strSql = strSql & " and oP.itemid="&iitemid&VbCRLF
					            strSql = strSql & " and op.outmallOptCode='"&ItemNo&"'"&VbCRLF
'					            strSql = strSql & " and Replace(Replace(op.outmallOptName,' ',''),':','')=Replace(Replace(o.optionname,' ',''),':','')"&VbCRLF
					            dbget.Execute strSql, AssignedRow
					        End If
					        getRegOptCD = ""
					    Else
					    	''단일 상품일 때 tbl_OutMall_regedoption엔 데이터가 있으나 tbl_item_option엔 데이터가 없기에 하단 프로시저 호출
							Dim DanChkArr
							strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_ltimall '"&CMallName&"'," & iitemid
							rsget.CursorLocation = adUseClient
							rsget.CursorType = adOpenStatic
							rsget.LockType = adLockOptimistic
							rsget.Open strSql, dbget
							If Not(rsget.EOF or rsget.BOF) Then
							    DanChkArr = rsget.getRows
							End If
							rsget.close
							If UBound(DanChkArr,2) = 0 AND DanChkArr(0,1) = "0000"  Then

							Else
						        strSql = " update oP"   &VbCRLF
						        strSql = strSql & " set itemoption=o.itemoption"&VbCRLF
						        strSql = strSql & " From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
						        strSql = strSql & "     Join db_item.dbo.tbl_item_option o"&VbCRLF
						        strSql = strSql & "     on oP.itemid=o.itemid"&VbCRLF
						        strSql = strSql & " where oP.mallid='lotteimall'"&VbCRLF
						        strSql = strSql & " and o.itemid="&iitemid&VbCRLF
						        strSql = strSql & " and oP.itemid="&iitemid&VbCRLF
						        strSql = strSql & " and op.outmallOptCode='"&ItemNo&"'"&VbCRLF
						        strSql = strSql & " and Replace(Replace(op.outmallOptName,' ',''),':','')=Replace(Replace(o.optionname,' ',''),':','')"&VbCRLF
						        dbget.Execute strSql, AssignedRow
							End If
					    End If
					    actCnt = actCnt+AssignedRow
					Next

					If (actCnt > 0) Then
					    strSql = " update R"   &VbCRLF
			            strSql = strSql & " set regedOptCnt=isNULL(T.regedOptCnt,0)"   &VbCRLF
			            strSql = strSql & " from db_item.dbo.tbl_LTiMall_regItem R"   &VbCRLF
			            strSql = strSql & " 	Join ("   &VbCRLF
			            strSql = strSql & " 		select R.itemid,count(*) as CNT "
			            strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
			            strSql = strSql & "        from db_item.dbo.tbl_LTiMall_regItem R"   &VbCRLF
			            strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
			            strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
			            strSql = strSql & " 			and Ro.mallid='lotteimall'"   &VbCRLF
			            strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
			            strSql = strSql & " 		group by R.itemid"   &VbCRLF
			            strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF
			            dbget.Execute strSql
					End If
				End if
			LtiMallOneItemCheckStock =true
			Set xmlDOM = Nothing
		Else
		    iErrStr =  "롯데아이몰과 통신중에 오류가 발생 [" & iitemid & "]:"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit Function
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'전시카테고리 수정
Function LtiMallOneItemCateEdit(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strRst, iMessage, strSql
	On Error Resume Next
	LtiMallOneItemCateEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/updateGoodsCategoryOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				If Err <> 0 Then
				    iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.LtiMallOneItemCateEdit"
					dbget.Close: Exit Function
				End If

				'결과 코드
				on Error resume next
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
		        on Error Goto 0

				'// 오류 검출
				If Err <> 0 Then
		            iErrStr = "Error: " & xmlDOM.getElementsByTagName("Message").item(0).text
					dbget.Close: Exit Function
		        End If
			Set xmlDOM = Nothing
		Else
		    iErrStr ="롯데아이몰과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요"
			dbget.Close: Exit Function
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'전시상품 단품추가
Function LtiMallAddOpt(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strRst, iMessage, strSql
	Dim addOptCount, opt1Nm, opt2Nm, opt3Nm, opt4Nm, opt5Nm, opt1Tval, opt2Tval, opt3Tval, opt4Tval, opt5Tval
	On Error Resume Next
	LtiMallAddOpt = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/addGoodsItemInfo.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				If Err <> 0 Then
					iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.[ERR-NMEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If

    			addOptCount = xmlDOM.getElementsByTagName("itemCount").item(0).text
				If addOptCount > 0 Then
					opt1Nm 		= xmlDOM.getElementsByTagName("opt1Nm").item(0).text
					opt1Tval	= xmlDOM.getElementsByTagName("opt1Tval").item(0).text
					opt2Nm 		= xmlDOM.getElementsByTagName("opt2Nm").item(0).text
					opt2Tval	= xmlDOM.getElementsByTagName("opt2Tval").item(0).text
					opt3Nm 		= xmlDOM.getElementsByTagName("opt3Nm").item(0).text
					opt3Tval	= xmlDOM.getElementsByTagName("opt3Tval").item(0).text
					opt4Nm 		= xmlDOM.getElementsByTagName("opt4Nm").item(0).text
					opt4Tval	= xmlDOM.getElementsByTagName("opt4Tval").item(0).text
					opt5Nm 		= xmlDOM.getElementsByTagName("opt5Nm").item(0).text
					opt5Tval	= xmlDOM.getElementsByTagName("opt5Tval").item(0).text
				End If

			    If Err <> 0 Then
				    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
				End If

				'// 오류 검출
				If Err <> 0 Then
			        If (IsAutoScript) Then
			            rw "["&iitemid & "]:"&iMessage
			        Else
						iErrStr =  "["&iitemid & "]:"&iMessage
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    Else

			    End If
			Set xmlDOM = Nothing
			LtiMallAddOpt = True
		Else
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''롯데아이몰 상품 등록
Function LotteiMallOneItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iLotteSellYn)
	Dim objXML,xmlDOM,strRst, iMessage, lp
	Dim buf, LotteGoodNo, strSql, buf_item_list, pp, OptDesc, StockQty, AssignedRow

	On Error Resume Next
	LotteiMallOneItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/registApiGoodsInfo.lotte", false
'rw ltiMallAPIURL & "/openapi/registApiGoodsInfo.lotte?"&strparam
'response.end
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			LotteGoodNo = ""
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf ''BinaryToText(objXML.ResponseBody, "euc-kr")
				LotteGoodNo = xmlDOM.getElementsByTagName("goods_no").item(0).text

				If LotteGoodNo = "" Then
				    iMessage = xmlDOM.getElementsByTagName("Message").item(0).text
				    iErrStr =  "상품 등록중 오류 [" & iitemid & "]:"&iMessage
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
				    Exit Function
				End If

				If Err <> 0 Then
					iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit function
				End If

				'상품존재여부 확인
				strSql = "Select count(itemid) From db_item.dbo.tbl_LTiMall_regItem Where itemid='" & iitemid & "'"
				rsget.Open strSql,dbget,1

				If rsget(0) > 0 Then
					'// 존재 -> 수정
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					strSql = strSql & "	Set LTiMallLastUpdate = getdate() "  & VbCRLF
					strSql = strSql & "	, LTiMallTmpGoodNo = '" & LotteGoodNo & "'"  & VbCRLF
					strSql = strSql & "	, LTiMallPrice = " &iSellCash& VbCRLF
					strSql = strSql & "	, accFailCnt = 0"& VbCRLF
					strSql = strSql & "	, LTiMallRegdate = isNULL(LTiMallRegdate, getdate())" ''추가 2013/02/26
					If (LotteGoodNo <> "") Then
					    strSql = strSql & "	, LTiMallstatCD = '20'"& VbCRLF
					Else
						strSql = strSql & "	, LTiMallstatCD = '10'"& VbCRLF
					End If
					strSql = strSql & "	From db_item.dbo.tbl_LTiMall_regItem R"& VbCRLF
					strSql = strSql & " Where R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
				Else
					'// 없음 -> 신규등록
					strSql = ""
					strSql = strSql & " INSERT INTO db_item.dbo.tbl_LTiMall_regItem "
					strSql = strSql & " (itemid, reguserid, LTiMallRegdate, LTiMallLastUpdate, LTiMallTmpGoodNo, LTiMallPrice, LTiMallSellYn, LTiMallStatCd) VALUES " & VbCRLF
					strSql = strSql & " ('" & iitemid & "'" & VBCRLF
					strSql = strSql & " , '" & session("ssBctId") & "'" &_
					strSql = strSql & " , getdate(), getdate()" & VBCRLF
					strSql = strSql & " , '" & LotteGoodNo & "'" & VBCRLF
					strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
					strSql = strSql & " , '" & iLotteSellYn & "'" & VBCRLF
					If (LotteGoodNo <> "") Then
					    strSql = strSql & ",'20'"
					Else
					    strSql = strSql & ",'10'"
					End If
					strSql = strSql & ")"
					dbget.Execute(strSql)
					actCnt = actCnt + 1
				End If
				rsget.Close
			    IF (TRUE) THEN
			        pp = 1
			    ''옵션 리스트 저장(20120807 추가)
			        ''buf_item_list = xmlDOM.getElementsByTagName("item_list").item(0).text
			        ''buf_item_list = xmlDOM.getElementsByTagName("Arguments").item("item_list").text
			        If xmlDOM.getElementsByTagName("Argument").item(18).getAttribute("name")="item_list" Then
			            buf_item_list = xmlDOM.getElementsByTagName("Argument").item(18).getAttribute("value")
			        Else
			            buf_item_list = xmlDOM.getElementsByTagName("Argument").item(29).getAttribute("value")
			        End If
			        '''iitemid = oiMall.FItemList(i).Fitemid
			        If (buf_item_list <> "") Then
'			            buf_item_list = UniToHanbyChilkat(buf_item_list)
			            rw "["&iitemid&"]=="&LotteGoodNo&"=="&buf_item_list
			            buf_item_list = split(buf_item_list,":")
			            For lp = Lbound(buf_item_list) to Ubound(buf_item_list)
			                ''이중옵션인경우 이 구문 잘못 되었음.
			                OptDesc = split(buf_item_list(lp),",")(0)
			                StockQty = split(buf_item_list(lp),",")(1)
							strSql = ""
							strSql = strSql & " Insert into db_item.dbo.tbl_OutMall_regedoption"
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outMallSellyn, outmalllimityn, outMallLimitNo)"
							strSql = strSql & " values("&iitemid
							strSql = strSql & " ,''"
							strSql = strSql & " ,'lotteimall'"
							strSql = strSql & " ,'"&pp&"'"
							strSql = strSql & " ,'"&html2DB(OptDesc)&"'"
							strSql = strSql & " ,'Y'"
							strSql = strSql & " ,'Y'"
							strSql = strSql & " ,"&StockQty
							strSql = strSql & ")"
							dbget.Execute strSql, AssignedRow

							''옵션 코드 매칭.
							If (AssignedRow>0) Then
								strSql = ""
								strSql = strSql & " update oP"   & VBCRLF
								strSql = strSql & " SET itemoption = O.itemoption " & VBCRLF
								strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption oP " & VBCRLF
								strSql = strSql & " JOIN db_item.dbo.tbl_item_option o on oP.itemid=o.itemid " & VBCRLF
								strSql = strSql & " WHERE oP.mallid = 'lotteimall' " & VBCRLF
								strSql = strSql & " and o.itemid = "&iitemid & VBCRLF
								strSql = strSql & " and oP.itemid = "&iitemid & VBCRLF
								strSql = strSql & " and op.outmallOptCode = '"&pp&"'" & VBCRLF
								strSql = strSql & " and op.outmallOptName = o.optionname" & VBCRLF
								dbget.Execute strSql, AssignedRow
							End If
							pp = pp + 1
						Next
						strSql = ""
						strSql = strSql & " UPDATE R " & VBCRLF
						strSql = strSql & " SET regedOptCnt = isNULL(T.CNT,0) " & VBCRLF
						strSql = strSql & " FROM db_item.dbo.tbl_LTiMall_regItem R " & VBCRLF
						strSql = strSql & " Join ( " & VBCRLF
						strSql = strSql & " 	SELECT R.itemid, count(*) as CNT FROM db_item.dbo.tbl_LTiMall_regItem R " & VBCRLF
						strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'lotteimall' and Ro.itemid = " &iitemid & VBCRLF
						strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
						strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
						dbget.Execute strSql
					Else
						rw "["&iitemid&"]=="&LotteGoodNo
					End If
			    End If
			Set xmlDOM = Nothing
			LotteiMallOneItemReg= true
		Else
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-REG-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function LtiMallOneItemSellStatEdit(iitemid, iLotteGoodNo, ichgSellYn, byRef iErrStr)
    Dim strParam
    Dim objXML, xmlDOM
    Dim strRst, strSql, notitemId

    LtiMallOneItemSellStatEdit = False
	strParam = "?subscriptionId=" & ltiMallAuthNo										'롯데아이몰 인증번호	(*)
	strParam = strParam & "&goods_no=" & iLotteGoodNo                       			'롯데아이몰 상품번호	(*)
'	strParam = strParam & "&brnd_no=1099329"			                       			'롯데아이몰 브랜드코드

'	If ichgSellYn = "Y" Then															'판매여부(10:판매, 20:품절, 30:판매종료)
'		strParam = strParam & "&item_sale_stat_cd=10"
'	ElseIf ichgSellYn = "N" Then
'		strParam = strParam & "&item_sale_stat_cd=20"
'	ElseIf ichgSellYn = "X" Then                        '''X 기능 사용안함
'		strParam = strParam & "&item_sale_stat_cd=30"      ''판매종료되면 수정못함.
'		''strParam = strParam & "&sale_stat_cd=20"
'	End If

	If ichgSellYn = "Y" Then															'판매여부(10:판매, 20:품절, 30:판매종료)
		strParam = strParam & "&sale_stat_cd=10"
	ElseIf ichgSellYn = "N" Then
		strParam = strParam & "&sale_stat_cd=20"
	ElseIf ichgSellYn = "X" Then                        '''X 기능 사용안함
		''''''strParam = strParam & "&sale_stat_cd=30"      ''판매종료되면 수정못함.
		strParam = strParam & "&sale_stat_cd=20"
	End If

	strSql = ""
	strSql = "select count(*) as cnt from db_temp.dbo.tbl_jaehyumall_not_in_itemid where mallgubun = 'lotteimall' and itemid =" & iitemid
	rsget.Open strSql, dbget
	If Not(rsget.EOF or rsget.BOF) Then
		notitemId = rsget("cnt")
	End If
	rsget.close

	'2013-07-18 김진영 추가..등록제외된 상품이 자꾸 판매되어서...
	If notitemId > 0 Then
		strParam = strParam & "&sale_stat_cd=20"
	End If

'rw ltiMallAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte" & strParam
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", ltiMallAPIURL & "/openapi/updateGoodsSaleStat.lotte" & strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				If Err <> 0 Then
				    iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.LtiMallOneItemSellStatEdit"
					dbget.Close: Exit Function
				End If

				'결과 코드
				on Error resume next
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
		        on Error Goto 0

				'// 오류 검출
				If Err <> 0 Then
		            iErrStr = "Error: " & xmlDOM.getElementsByTagName("Message").item(0).text
					dbget.Close: Exit Function
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem " & VbCRLF
					strSql = strSql & " SET LtiMallLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,LtiMallSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
		        End If
			Set xmlDOM = Nothing
		Else
		    iErrStr ="롯데아이몰과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요"
			dbget.Close: Exit Function
		End If
	Set objXML = Nothing
	LtiMallOneItemSellStatEdit = true
End Function

''전시상품번호 매핑정보
Function CheckLtiMallTmpItemChk(iitemid,byRef iErrStr, byRef iLotteGoodNo, byref iLotteStatCd)
	Dim objXML,xmlDOM,strRst,iMessage
	Dim strParam, iLotteTmpID , SaleStatCd, GoodsViewCount
	Dim iRbody

	CheckLtiMallTmpItemChk = ""
	iLotteTmpID = getLtiMallTmpItemIdByTenItemID(iitemid)

	if iLotteTmpID="" then Exit function '' 2013/09/03 추가

	If iLotteTmpID = "전시상품" Then
		CheckLtiMallTmpItemChk = "전시상품"
	Else
		strParam = "subscriptionId=" & ltiMallAuthNo & "&goods_no="&iLotteTmpID
		On Error Resume Next
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "POST", ltiMallAPIURL & "/openapi/getRdToPrGoodsNoApi.lotte", false
'rw ltiMallAPIURL & "/openapi/getRdToPrGoodsNoApi.lotte"&strParam
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send(strParam)
			If objXML.Status = "200" Then
				Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
					xmlDOM.LoadXML iRbody
'rw iRbody
					If Err <> 0 Then
						iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다."
					    Set objXML = Nothing
					    Set xmlDOM = Nothing
					    Exit Function
					End If

					GoodsViewCount 		= Trim(xmlDOM.getElementsByTagName("GoodsCount").item(0).text)			'검색수

				    If Err <> 0 Then
					    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
					End If

					If (GoodsViewCount = "1") Then
						iLotteGoodNo		= Trim(xmlDOM.getElementsByTagName("GoodsNo").item(0).text)			'전시상품번호
						iLotteStatCd		= Trim(xmlDOM.getElementsByTagName("ConfStatCd").item(0).text)		'인증상태코드
				    End If

					If Err <> 0 Then
						iErrStr =  "["&iitemid & "]:"&iMessage
						Set objXML = Nothing
					    Set xmlDOM = Nothing
					    On Error Goto 0
					    Exit Function
				    Else

				    End If
				Set xmlDOM = Nothing

				If (GoodsViewCount <> "1") Then
				    iErrStr ="검색결과가 없습니다."&iMessage
				Else
				    CheckLtiMallTmpItemChk = iLotteStatCd
			    End If
			Else
				iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-ItemChk-002]"
			End If
		Set objXML = Nothing
		On Error Goto 0
	End If
End Function

''전시상품 조회
Function CheckLtiMallItemStat(iitemid,byRef iErrStr, byRef iSalePrc, byref iGoodsNm)
	Dim objXML,xmlDOM,strRst,iMessage
	Dim strParam, iLotteItemID , SaleStatCd, GoodsViewCount
	Dim iRbody

	CheckLtiMallItemStat = ""
	iLotteItemID = getLtiMallItemIdByTenItemID(iitemid)
	strParam = "subscriptionId=" & ltiMallAuthNo & "&goods_no="&iLotteItemID
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		'' rw "https://openapi.lotte.com/openapi/searchGoodsListOpenApiOther.lotte?"&strParam ''''전시상품조회aLL
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchGoodsListOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			'XML을 담을 DOM 객체 생성
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
			''rw iRbody
'			iRbody = replace(iRbody,"&","@@amp@@")   '' <![CDATA[]]> 로 안 묶여옴. 상품명에 < > 있음..
'			iRbody = replace(iRbody,"<GoodsNm>","<GoodsNm><![CDATA[")
'			iRbody = replace(iRbody,"</GoodsNm>","]]></GoodsNm>")
			xmlDOM.LoadXML iRbody

			if Err <> 0 Then
				iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.[ERR-ItemChk-001]"
			    Set objXML = Nothing
			    Set xmlDOM = Nothing
			    Exit Function
			End If
			GoodsViewCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text  ''결과수

		    If Err <> 0 Then
			    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
			End If
		''rw "GoodsViewCount="&GoodsViewCount
		''rw "Err="&Err
			If (GoodsViewCount = "1") Then
		    	SaleStatCd = xmlDOM.getElementsByTagName("SaleStatCd").item(0).text
		    	iSalePrc   = xmlDOM.getElementsByTagName("SalePrc").item(0).text
		    	iGoodsNm   = xmlDOM.getElementsByTagName("GoodsNm").item(0).text
		    	iGoodsNm   = replace(iGoodsNm,"@@amp@@","&")
		    End If

			'// 오류 검출
			If Err <> 0 Then
		        If (IsAutoScript) Then
		            rw "["&iitemid & "]:"&iMessage
		        Else
					iErrStr =  "["&iitemid & "]:"&iMessage
				End If
				Set objXML = Nothing
			    Set xmlDOM = Nothing
			    On Error Goto 0
			    Exit Function
		    Else

		    End If
			Set xmlDOM = Nothing

			If (GoodsViewCount <> "1") Then
			    iErrStr ="검색결과가 없습니다.."&iMessage
			ElseIf (SaleStatCd<>"0") then
			    CheckLtiMallItemStat = SaleStatCd
		    End If
		Else
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-ItemChk-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0

	'	rw "SaleStatCd="&SaleStatCd
	'	rw "GoodsViewCount="&GoodsViewCount
	'	rw "iMessage="&iMessage
	'	rw "iErrStr="&iErrStr
End Function


''전시상품 조회(날짜로)
Function CheckLtiMallDateItemStat(chkdate, byRef iErrStr, byRef iSalePrc, byref iGoodsNm)
	Dim objXML,xmlDOM,strRst,iMessage
	Dim strParam, iLotteItemID , SaleStatCd, GoodsViewCount
	Dim iRbody

	CheckLtiMallDateItemStat = ""
	strParam = "subscriptionId=" & ltiMallAuthNo & "&req_start_dtime="&replace(chkdate,"-","")&"&req_end_dtime="&replace(chkdate,"-","")
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchGoodsListOpenApi.lotte", false

		rw ltiMallAPIURL & "/openapi/searchGoodsListOpenApi.lotte?"&strParam
		rw "<input type='button' value='다음날짜' onclick=location.href='/admin/etc/LtiMall/actLotteiMallReq.asp?cmdparam=ChkDate&chkdate="&dateadd("d",1,chkdate)&"';>"

		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			'XML을 담을 DOM 객체 생성
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
			xmlDOM.LoadXML iRbody

			if Err <> 0 Then
				iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.[ERR-DateChk-001]"
			    Set objXML = Nothing
			    Set xmlDOM = Nothing
			    Exit Function
			End If

		    If Err <> 0 Then
			    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
			End If

		''rw "Err="&Err
			Dim Nodes, SubNodes, GoodsNo, iquery, GoodsRegDtime
			AssignedRow = 0

			GoodsViewCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text  ''결과수
			Set Nodes = xmlDOM.getElementsByTagName("GoodsInfoList")
			For each SubNodes in Nodes
			    GoodsNo    		= SubNodes.getElementsByTagName("GoodsNo").item(0).text
				GoodsRegDtime   = SubNodes.getElementsByTagName("GoodsRegDtime").Item(0).Text

				iquery = ""
				iquery = iquery & "SELECT count(*) as cnt FROM db_temp.dbo.tbl_tmp_ltimallGoodno where ltimallgoodno = '"&GoodsNo&"'"
				rsget.Open iquery, dbget
				If rsget("cnt") = 0 Then
					iquery = ""
					iquery = iquery & "INSERT INTO db_temp.dbo.tbl_tmp_ltimallGoodno (ltimallgoodno, GoodsRegDtime) VALUES ('"&GoodsNo&"', '"&GoodsRegDtime&"')"
					dbget.Execute iquery, AssignedRow
					rw GoodsNo&" 등록완료"
				End If
				rsget.Close
			Next

			'// 오류 검출
			If Err <> 0 Then
				iErrStr =  "["&iitemid & "]:"&iMessage
				Set objXML = Nothing
			    Set xmlDOM = Nothing
			    On Error Goto 0
			    Exit Function
		    Else

		    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-ItemChk-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function checkConfirmMatch(iitemid, iLotteItemStat, iLotteSalePrc, iLotteGoodsNm) ''롯데아이몰 (현)판매상태, (현)판매가 업데이트, 상품명 추가
	Dim sqlstr, LotteSellyn
	Dim assignedRow : assignedRow = 0
	Dim pbuf : pbuf = ""

	iLotteItemStat = Trim(iLotteItemStat)
	iLotteGoodsNm = Trim(iLotteGoodsNm)
	iLotteGoodsNm = Replace(iLotteGoodsNm,"&gt;",">")
	iLotteGoodsNm = Replace(iLotteGoodsNm,"&lt;","<")
	iLotteGoodsNm = Replace(iLotteGoodsNm,"&nbsp;"," ")
	iLotteGoodsNm = Replace(iLotteGoodsNm,"&amp;","&")

	If (iLotteItemStat="10") Then
	    LotteSellyn = "Y"
	ElseIf (iLotteItemStat="20") Then
	    LotteSellyn = "N"
	ElseIf (iLotteItemStat="30") Then
	    LotteSellyn = "X"
	End If

	sqlstr = ""
	sqlstr = sqlstr & " SELECT (convert(varchar(10),ltiMallStatCd)+','+LtiMallSellyn+','+convert(varchar(10),convert(int,ltiMallPrice)) )as pbuf"
	sqlstr = sqlstr & " FROM db_item.dbo.tbl_LTiMall_regItem R" & VbCRLF
	sqlstr = sqlstr & " WHERE R.itemid="&iitemid & VbCRLF
	rsget.Open sqlstr,dbget,1
	If Not rsget.EOF Then
	    pbuf = rsget("pbuf")
	End If
	rsget.close()

	sqlstr = "Update R" & VbCRLF
	sqlstr = sqlstr & " SET ltiMallPrice="&iLotteSalePrc & VbCRLF
	If (LotteSellyn <> "") Then
	    sqlstr = sqlstr & " ,ltiMallSellyn='"&LotteSellyn&"'"
	End If
	sqlstr = sqlstr & " ,regitemname='"&html2db(iLotteGoodsNm)&"'"
	sqlstr = sqlstr & " ,LtiMallStatCd=(CASE WHEN isNULL(LtiMallStatCd,-9)<7 THEN 7 ELSE LtiMallStatCd END)" ''2013/09/03 추가
	sqlstr = sqlstr & " ,lastStatCheckDate=getdate()"
	sqlstr = sqlstr & " FROM db_item.dbo.tbl_LTiMall_regItem R" & VbCRLF
	sqlstr = sqlstr & " WHERE R.itemid="&iitemid & VbCRLF
	sqlstr = sqlstr & " and (ltiMallPrice<>"&iLotteSalePrc&"" & VbCRLF
	sqlstr = sqlstr & "     or ltiMallSellyn<>'"&LotteSellyn&"'"& VbCRLF
	sqlstr = sqlstr & "     or isNULL(regitemname,'')<>'"&html2db(iLotteGoodsNm)&"'"& VbCRLF
	sqlstr = sqlstr & " )" & VbCRLF
	dbget.Execute sqlstr, assignedRow

	If (assignedRow < 1) Then
	    sqlstr = ""
	    sqlstr = sqlstr & "UPDATE R" & VbCRLF
	    sqlstr = sqlstr & " SET lastStatCheckDate = getdate()"
	    sqlstr = sqlstr & " ,LtiMallStatCd=(CASE WHEN isNULL(LtiMallStatCd,-9)<7 THEN 7 ELSE LtiMallStatCd END)"  ''
	    sqlstr = sqlstr & " From db_item.dbo.tbl_LTiMall_regItem R" & VbCRLF
	    sqlstr = sqlstr & " where R.itemid = "&iitemid & VbCRLF
	    dbget.Execute sqlstr
	Else
	    ''다른게 있으면 로그.
	    CALL Fn_AcctFailLog(CMALLNAME, iitemid, iLotteItemStat&","&LotteSellyn&", "&iLotteSalePrc&"::"&pbuf, "STAT_CHK")
	End If
End Function

''전시상품명수정
Function LtiMallOneItemNameEdit(iitemid, strParam, byRef iErrStr)
	Dim objXML,xmlDOM,strRst,iMessage
	On Error Resume Next
	LtiMallOneItemNameEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/updateGoodsNmOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				If Err <> 0 Then
					iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.[ERR-NMEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If

				strRst = xmlDOM.getElementsByTagName("Result").item(0).text

			    If Err <> 0 Then
				    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
				End If

				'// 오류 검출
				If Err <> 0 Then
			        If (IsAutoScript) Then
			            rw "["&iitemid & "]:"&iMessage
			        Else
						iErrStr =  "["&iitemid & "]:"&iMessage
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    Else

			    End If
			Set xmlDOM = Nothing
			LtiMallOneItemNameEdit = True
		Else
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'전시상품 판매가 수정
Function LtiMallOnItemPriceEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
	On Error Resume Next
	LtiMallOnItemPriceEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/updateGoodsSalePrcOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				If Err <> 0 Then
					iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.[ERR-PRCEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'결과 코드
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text

				If Err <> 0 Then
					iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
				End If

				If Err <> 0 Then
				    If (Trim(iMessage)="002. 세일진행(또는 진행예정)") Then
				        iErrStr = "["&iitemid & "]:"&Trim(iMessage)&"\n"
				        Set objXML = Nothing
					    Set xmlDOM = Nothing
					    On Error Goto 0
					    Exit Function
				    Else
						If (IsAutoScript) Then
							rw "["&iitemid & "]:"&iMessage
						Else
							iErrStr =  "["&iitemid & "]:"&iMessage
						End If
						Set objXML = Nothing
					    Set xmlDOM = Nothing
					    On Error Goto 0
					    Exit Function
					End If
			    Else

			    End If
			Set xmlDOM = Nothing
			LtiMallOnItemPriceEdit = True
		Else
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-PRCEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''롯데아이몰 상품정보 수정
Function LtiMallOneItemInfoEdit(iitemid, strParam, byRef iErrStr, isVer2)
	Dim objXML, xmlDOM, strRst, iMessage
	On Error Resume Next
	LtiMallOneItemInfoEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	If (isVer2) Then
	    objXML.Open "POST", ltiMallAPIURL & "/openapi/upateApiNewGoodsInfo.lotte", false          ''상품수정
	Else
	    objXML.Open "POST", ltiMallAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte", false      ''전시상품수정
	End If

'rw ltiMallAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte?"&strParam
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then

		'//전달받은 내용 확인
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				''rw BinaryToText(objXML.ResponseBody, "euc-kr")
				If Err<>0 then
				    iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'결과 코드

				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
			    If Err <> 0 Then
				    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
				End If

			    If (strRst <> "1") Then
			        ''rw "FitemDiv="&oiMall.FItemList(i).FitemDiv  ''주문제작 상품인경우 수정이 안되는듯. add_choc_tp_cd_20 에 값이 있는경우;;
			        ''rw "iMessage="&iMessage
			        iErrStr =  "상품 수정중 오류 [" & iitemid & "]:"&iMessage
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
			        On Error Goto 0
				    Exit Function
			    End If

			    If (strRst = "1") Then
					If Err <> 0 Then
					    If (IsAutoScript) Then
					        rw "["&iitemid & "]:"&iMessage
					    Else
							iErrStr = "["&iitemid & "]:"&iMessage
						End If
						Set objXML = Nothing
					    Set xmlDOM = Nothing
					    Exit Function
					Else

					End If
			    End If
			Set xmlDOM = Nothing
			LtiMallOneItemInfoEdit = True
		Else
		    iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-EDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function getItemOptionCount(iitemid)
	Dim ret, sqlstr
	ret = 0
	sqlstr = "SELECT count(*) as optCNT FROM db_item.dbo.tbl_item_option WHERE itemid = "&Left(iitemid,10)&" and isusing = 'Y'"
	rsget.Open sqlstr,dbget,1
	If Not rsget.EOF Then
	    ret = rsget("optCNT")
	End If
	rsget.close()
	getItemOptionCount = Trim(ret)
End Function
					''///////////////						위까지가 펑션, 아래부터 프로세스 				//////////////////
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim subcmd : subcmd = requestCheckVar(request("subcmd"),10)
Dim oiMall, i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim iLotteItemStat, iLotteSalePrc, ArrRows, iLotteGoodsNm, iLotteItemTmpChk
Dim retFlag
Dim iMessage
dim iLotteGoodNo, iItemName, pregitemname, iLotteStatCd
Dim delitemid
delitemid = requestCheckvar(request("delitemid"),10)

retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")
If (cmdparam = "RegSelectWait") Then						''선택상품 예정 등록.
	arrItemid = Trim(arrItemid)
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " INSERT into db_item.dbo.tbl_LTiMall_regItem "
	sqlStr = sqlStr & " (itemid, regdate, reguserid, LtiMallStatCD)"
	sqlStr = sqlStr & " SELECT i.itemid, getdate(), '"&session("SSBctID")&"', '0' "
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i"
	sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_LTiMall_regItem R on i.itemid = R.itemid "
	sqlStr = sqlStr & " WHERE i.itemid in ("&arrItemid&") "
	sqlStr = sqlStr & " and R.itemid is NULL"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 예정등록됨.');</script>"

	sqlStr = ""
	sqlStr = sqlStr & " update R "
	sqlStr = sqlStr & " set optAddPrcCnt= T.optAddPrcCnt "
	sqlStr = sqlStr & " from db_item.dbo.tbl_LTiMall_regItem R "
	sqlStr = sqlStr & " Join ( "
	sqlStr = sqlStr & " 	select ii.itemid,count(*) as optAddPrcCnt "
	sqlStr = sqlStr & " 	from db_item.dbo.tbl_item ii 	 "
	sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item_option o 	 "
	sqlStr = sqlStr & "		on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'	 "
	sqlStr = sqlStr & " 	group by ii.itemid "
	sqlStr = sqlStr & " ) T on R.itemid =T.itemid "
	sqlStr = sqlStr & " WHERE R.itemid in ("&arrItemid&") "
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>parent.location.reload();</script>"
ElseIf (cmdparam = "DelSelectWait") Then					''선택상품 예정 삭제.
	arrItemid = Trim(arrItemid)
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM db_item.dbo.tbl_LTiMall_regItem "
	sqlStr = sqlStr & " WHERE LtimallStatCD in ('0')"
	sqlStr = sqlStr & " and itemid in ("&arrItemid&")"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 예정 삭제됨.');parent.location.reload();</script>"
ElseIf (cmdparam="ChkDate") Then
	Dim chkdate, iLotteDateItemStat
	chkdate = request("chkdate")

    iErrStr = ""
	iLotteDateItemStat = CheckLtiMallDateItemStat(chkdate, iErrStr, iLotteSalePrc, iLotteGoodsNm)
	response.end
ElseIf (cmdparam = "EditSellYn") Then						''선택상품 판매여부 수정
	Set oiMall = new CLotteiMall
		oiMall.FPageSize	= 20
		oiMall.FRectItemID	= arrItemid
		''oiMall.FRectMatchCateNotCheck = "on"
		oiMall.getLTiMallEditedItemList

	If (chgSellYn="N") and (oiMall.FResultCount < 1) and (arrItemid = "") Then
	    oiMall.getLtiMallreqExpireItemList
	End If

	rw oiMall.FResultCount

	For i = 0 to (oiMall.FResultCount - 1)
	    iErrStr = ""
		If (LtiMallOneItemSellStatEdit(oiMall.FItemList(i).Fitemid, oiMall.FItemList(i).FLtiMallGoodNo, chgSellYn, iErrStr)) Then
			actCnt = actCnt+1
		Else
			rw "["&iitemid&"]"&iErrStr
		End If
		retErrStr = retErrStr & iErrStr
	Next
	Set oiMall = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam = "EditItemNm") Then							''선택상품명수정요청
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbget.Close: Response.End
	End If

	'## 수정된 상품 목록 접수
	Set oiMall = new CLotteiMall
		oiMall.FPageSize       = 30
		oiMall.FRectItemID	= arrItemid
		oiMall.getLTiMallEditedItemList

		For i = 0 to (oiMall.FResultCount - 1)
			on Error Resume Next
			strParam = oiMall.FItemList(i).GetLtiMallItemNameEditParameter()
			If (session("ssBctID") = "icommang") or (session("ssBctID") = "kjy8517") Then
				rw ltiMallAPIURL & "/openapi/updateGoodsNmOpenApi.lotte?" & strParam
			End If

			If Err <> 0 Then
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oiMall.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If
			on Error Goto 0

			iErrStr = ""
			ret1 = false
			ret1 = LtiMallOneItemNameEdit(oiMall.FItemList(i).Fitemid, strParam, iErrStr)
			If (ret1) Then
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_LTiMall_regItem  " & VBCRLF
				sqlStr = sqlStr & "	SET regitemname = '" & html2db(oiMall.FItemList(i).FItemName) &"'"& VBCRLF
				sqlStr = sqlStr & " WHERE itemid = '" & oiMall.FItemList(i).Fitemid & "'"& VBCRLF
				dbget.Execute(sqlStr)
				actCnt = actCnt+1
			Else
			    CALL Fn_AcctFailTouch("lotteimall", oiMall.FItemList(i).Fitemid, iErrStr)
			    rw "iErrStr="&iErrStr
			End If
			retErrStr = retErrStr & iErrStr
		Next
	Set oiMall = Nothing
    If (retErrStr <> "") Then
        Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
    End If
ElseIf (cmdparam = "CheckItemNmAuto") Then						''상품명수정(관리자)
	buf = ""
	CNT10 = 0
	sqlStr = ""
	sqlStr = sqlStr & " SELECT top 10 r.itemid, isNULL(r.ltiMallGoodNo,r.ltimallTmpGoodNO) as ltiMallGoodNo, i.ItemName "
	sqlStr = sqlStr & "	FROM db_item.dbo.tbl_LTiMall_regItem r "
	sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_item i on r.itemid = i.itemid "
	sqlStr = sqlStr & "	WHERE r.regitemname is Not NULL "
	sqlStr = sqlStr & "	and replace(r.regitemname,'"&CPREFIXITEMNAME&"','') <> i.itemname "
	sqlStr = sqlStr & "	and isNULL(r.ltiMallGoodNo,r.ltimallTmpGoodNO) is Not NULL"  ''수정 2013/07/12
	sqlStr = sqlStr & "	ORDER BY r.lastStatCheckDate DESC"
	rsget.Open sqlStr,dbget,1
	If not rsget.Eof then
		ArrRows = rsget.getRows()
	End If
	rsget.Close

	If isArray(ArrRows) Then
		For i =0 To UBound(ArrRows,2)
			iErrStr = ""
			iitemid = CStr(ArrRows(0,i))
			buf = buf&iitemid&","
			iLotteGoodNo = CStr(ArrRows(1,i))
			iItemName    = CStr(ArrRows(2,i))
			If (iitemid <> "") Then
				strParam = fnGetLtiMallItemNameEditParameter(iLotteGoodNo, iItemName)
				ret1 = false
				ret1 = LtiMallOneItemNameEdit(iitemid, strParam, iErrStr)
				If (ret1) Then
					'// 상품명 수정
					pregitemname = ""
					sqlStr = ""
					sqlStr = sqlStr & " SELECT isNULL(regitemname,'') as regitemname FROM db_item.dbo.tbl_LTiMall_regItem "& VBCRLF
					sqlStr = sqlStr & "	WHERE itemid = '" & iitemid & "'"& VBCRLF
					rsget.Open sqlStr,dbget,1
					If not rsget.EOF Then
					    pregitemname = rsget("regitemname")
					End If
					rsget.Close

					sqlStr = ""
					sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_LTiMall_regItem  " & VbCRLF
					sqlStr = sqlStr & "	SET regitemname='" & html2db(iItemName) &"'"& VbCRLF
					sqlStr = sqlStr & " WHERE itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(sqlStr)
					CNT10 = CNT10+1
					If (pregitemname <> iItemName) Then
					    buf2 = buf2 & pregitemname & "::" & iItemName &"<br>"
					End If
				Else
					CALL Fn_AcctFailTouch("lotteimall", iitemid, iErrStr)
					rw "iErrStr="&iErrStr
				End If
			End If
		Next
	End If
	rw buf
	rw buf2
	rw CNT10&"건 상품명수정 성공"
	response.end
ElseIf (cmdparam = "EditPriceSelect") Then						''선택상품가격 수정
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbget.Close: Response.End
	End If

	'## 수정된 상품 목록 접수
	Set oiMall = new CLotteiMall
		oiMall.FPageSize       = 30
		oiMall.FRectItemID	= arrItemid
		oiMall.getLTiMallEditedItemList
		For i = 0 to (oiMall.FResultCount - 1)
		    on Error Resume Next
		    strParam = oiMall.FItemList(i).getLtiMallItemPriceEditParameter()
			If (session("ssBctID") = "icommang") or (session("ssBctID") = "kjy8517") Then
				rw ltiMallAPIURL & "/openapi/updateGoodsSalePrcOpenApi.lotte?" & strParam
			End If

			If Err <> 0 Then
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oiMall.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If
		    on Error Goto 0

		    iErrStr = ""
			ret1 = false
		    ret1 = LtiMallOnItemPriceEdit(oiMall.FItemList(i).Fitemid, strParam, iErrStr)
			If (ret1) Then
			    '// 상품가격정보 수정
			    sqlStr = ""
				sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_LTiMall_regItem  " & VbCRLF
				sqlStr = sqlStr & "	SET LtiMallPrice = " & oiMall.FItemList(i).MustPrice & VbCRLF
				''sqlStr = sqlStr & "	, accFailCnt = 0"& VbCRLF                                           ''정보수정만
				''sqlStr = sqlStr & "	, LtiMallLastUpdate = getdate() " & VbCRLF                          ''정보수정만
				sqlStr = sqlStr & " WHERE itemid='" & oiMall.FItemList(i).Fitemid & "'"& VbCRLF
				dbget.Execute(sqlStr)
				actCnt = actCnt + 1
		    Else
		        CALL Fn_AcctFailTouch("lotteimall", oiMall.FItemList(i).Fitemid, iErrStr)
		        rw "iErrStr="&iErrStr
			End If
		    retErrStr = retErrStr & iErrStr
		Next
	Set oiMall = Nothing

	If (retErrStr <> "") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam = "getconfirmList") Then					''선택상품 등록확인
	Dim AnotherStat
	arrItemid = split(Trim(arrItemid), ",")
	If IsArray(arrItemid) Then
		For i = LBound(arrItemid) to UBound(arrItemid)
		    iErrStr = ""
		    iitemid = Trim(arrItemid(i))
			If (iitemid <> "") Then
				iLotteItemTmpChk = CheckLtiMallTmpItemChk(iitemid, iErrStr, iLotteGoodNo, iLotteStatCd)
				Select Case iLotteItemTmpChk
					Case "10"	AnotherStat = "임시등록"
					Case "20"	AnotherStat = "승인요청"
					Case "30"	AnotherStat = "승인완료"
					Case "40"	AnotherStat = "반려"
					Case "50"	AnotherStat = "승인불가"
					Case "51"	AnotherStat = "재승인요청"
					Case "52"	AnotherStat = "수정요청"
					Case "60"	AnotherStat = "취소"
				End Select

				If (iLotteItemTmpChk = "30") Then
					strSql =""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem "
					strSql = strSql & "	SET lastConfirmdate = getdate() "
					strSql = strSql & "	,LtiMallStatCd='7' "
					strSql = strSql & " ,LtiMallGoodNo='" & iLotteGoodNo & "' "
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute strSql
					rw "["&iitemid&"] :"&iLotteItemTmpChk&":"&AnotherStat
				ElseIf iLotteItemTmpChk = "전시상품" Then
					strSql =""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem "
					strSql = strSql & "	SET LtiMallStatCd='7' "
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute strSql
					rw "["&iitemid&"] :"&iLotteItemTmpChk&": 이미 전시 상품입니다"

				Else
				    sqlstr = "Update R" & VbCRLF
                	sqlstr = sqlstr & " SET lastStatCheckDate=getdate()"
                	sqlstr = sqlstr & " FROM db_item.dbo.tbl_LTiMall_regItem R" & VbCRLF
                	sqlstr = sqlstr & " WHERE R.itemid="&iitemid & VbCRLF
                	''rw sqlstr
                	dbget.Execute sqlstr
				    rw "["&iitemid&"] ::"&iLotteItemTmpChk&":"&AnotherStat
				End If

				If iLotteItemTmpChk = "30" Then
					CALL LtiMallOneItemCheckStock(iitemid, iErrStr)				''2013-07-24 김진영 추가
				End If
		    End If
		Next
	End If
	response.end
ElseIf (cmdparam = "CheckItemStat") Then					''선택상품 판매상태확인
	arrItemid = split(Trim(arrItemid), ",")
	If IsArray(arrItemid) Then
		For i = LBound(arrItemid) to UBound(arrItemid)
		    iErrStr = ""
		    iitemid = Trim(arrItemid(i))
			If (iitemid <> "") Then
				iLotteItemStat = CheckLtiMallItemStat(iitemid, iErrStr, iLotteSalePrc, iLotteGoodsNm)

				If (iLotteItemStat <> "") Then
					rw "["&iitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr
				    CALL checkConfirmMatch(iitemid,iLotteItemStat,iLotteSalePrc,iLotteGoodsNm)
				Else
				    sqlstr = "Update R" & VbCRLF
                	sqlstr = sqlstr & " SET lastStatCheckDate=getdate()"
                	sqlstr = sqlstr & " FROM db_item.dbo.tbl_LTiMall_regItem R" & VbCRLF
                	sqlstr = sqlstr & " WHERE R.itemid="&iitemid & VbCRLF
                	''rw sqlstr
                	dbget.Execute sqlstr
				    rw "["&iitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr
				End If
		    End If
		Next
	End If
	response.end
ElseIf (cmdparam = "CheckItemStatAuto") Then				''판매상태Check(관리자)
	buf = ""
	CNT10=0
	CNT20=0
	CNT30=0

	sqlStr = ""
	sqlStr = sqlStr & " SELECT TOP 20 r.itemid "
	sqlStr = sqlStr & "	FROM db_item.dbo.tbl_LTiMall_regItem r"
	sqlStr = sqlStr & "	WHERE LtiMallGoodno is Not NULL"
	sqlStr = sqlStr & "	ORDER BY r.lastStatCheckDate, (CASE WHEN r.LtiMallsellyn = 'X' THEN '0' ELSE r.LtiMallsellyn END), r.LtiMallLastUpdate, r.itemid DESC"
	rsget.Open sqlStr,dbget,1
	If not rsget.EOF Then
		ArrRows = rsget.getRows()
	End If
	rsget.Close

	If isArray(ArrRows) Then
	    For i = 0 To UBound(ArrRows,2)
			iErrStr = ""
			iitemid = CStr(ArrRows(0,i))
			If (iitemid<>"") then
				iLotteItemStat = CheckLtiMallItemStat(iitemid, iErrStr, iLotteSalePrc, iLotteGoodsNm)
				If (iLotteItemStat <> "") Then
			        buf=buf& "["&iitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr &"<br>"
					If (iLotteItemStat = "10") Then
					    CNT10 = CNT10 + 1
					ElseIf (iLotteItemStat = "20") Then
					    CNT20 = CNT20 + 1
					ElseIf (iLotteItemStat = "30") Then
					    CNT30 = CNT30 + 1
					End If

					CALL checkConfirmMatch(iitemid,iLotteItemStat,iLotteSalePrc,iLotteGoodsNm)
			    Else
			        sqlstr = "Update R" & VbCRLF
                	sqlstr = sqlstr & " SET lastStatCheckDate=getdate()"
                	sqlstr = sqlstr & " FROM db_item.dbo.tbl_LTiMall_regItem R" & VbCRLF
                	sqlstr = sqlstr & " WHERE R.itemid="&iitemid & VbCRLF
                	''rw sqlstr
                	dbget.Execute sqlstr

			        buf=buf& "["&iitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr &"<br>"
			    End if
			End If
		Next
	End iF
	rw "STAT10:"&CNT10&"<br>STAT20:"&CNT20&"<br>STAT30:"&CNT30&"<br><br>"&buf
	response.end

ElseIf (cmdparam = "ChkStockSelect") Then				''선택상품 재고조회
	arrItemid = Trim(arrItemid)
	If (Right(arrItemid, 1) = ",") Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	arrItemid = split(arrItemid,",")
	For i = Lbound(arrItemid) to UBound(arrItemid)
		iitemid = arrItemid(i)
		iErrStr = ""
		ret1 = false
		ret1 = LtiMallOneItemCheckStock(iitemid, iErrStr)
		If (ret1) Then
			actCnt = actCnt + 1
		Else
		    rw "iErrStr="&iErrStr
		End If
		retErrStr = retErrStr & iErrStr
	Next
	If (retFlag <> "") Then
		Response.Write "<script language=javascript>parent."&retFlag&";</script>"
		response.end
	End If
	response.end
ElseIf (cmdparam = "RegSelect") Then				''선택상품 실제 등록
    arrItemid = Trim(arrItemid)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbget.Close: Response.End
	End If

	'## 선택상품 목록 접수
	Set oiMall = new CLotteiMall
		oiMall.FPageSize	= 20
		oiMall.FRectItemID	= arrItemid
		oiMall.getLTiMallNotRegItemList
	    If (oiMall.FResultCount < 1) Then
	        arrItemid = split(arrItemid,",")
	        For i = LBound(arrItemid) to UBound(arrItemid)
	            CALL Fn_AcctFailTouch("lotteimall",arrItemid(i),"등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...")
	        Next

	        If (IsAutoScript) Then
	            rw "S_ERR|등록가능상품 없음 :등록조건 확인: 판매Y, 할인..."
	            dbget.Close: Response.End
	        Else
	            Response.Write "<script language=javascript>alert('등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...');</script>"
				dbget.Close: Response.End
			End If
		End If

		For i = 0 to (oiMall.FResultCount - 1)
			sqlStr = ""
			sqlStr = sqlStr & " IF Exists(SELECT * FROM db_item.dbo.tbl_LTiMall_regitem where itemid="&oiMall.FItemList(i).Fitemid&")"
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " UPDATE R" & VbCRLF
			sqlStr = sqlStr & "	SET LTiMallLastUpdate = getdate() "  & VbCRLF
			sqlStr = sqlStr & "	, LTiMallstatCD='1'"& VbCRLF                     ''2013-06-14 진영//통신전 등록시도로 일단 넣음(롯데닷컴은 10, 아이몰은 1로 수정)
			sqlStr = sqlStr & "	FROM db_item.dbo.tbl_LTiMall_regitem R"& VbCRLF
			sqlStr = sqlStr & " WHERE R.itemid='" & oiMall.FItemList(i).Fitemid & "'"
			sqlStr = sqlStr & " END ELSE "
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_LTiMall_regitem "
	        sqlStr = sqlStr & " (itemid, regdate, reguserid, LTiMallstatCD)"
	        sqlStr = sqlStr & " VALUES ("&oiMall.FItemList(i).Fitemid&", getdate(), '"&session("SSBctID")&"', '1')"
			sqlStr = sqlStr & " END "
		    dbget.Execute sqlStr

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oiMall.FItemList(i).checkTenItemOptionValid Then
'			    On Error Resume Next
				'//상품등록 파라메터
				strParam = oiMall.FItemList(i).getLotteiMallItemRegParameter(FALSE)
				If (session("ssBctID") = "icommang") or (session("ssBctID") = "kjy8517") Then
					rw ltiMallAPIURL & "" & strParam
				End If

				If Err <> 0 Then
				    rw Err.Description
					Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oiMall.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				End If
	            On Error Goto 0

	            iErrStr = ""
	            ret1 = LotteiMallOneItemReg(oiMall.FItemList(i).Fitemid, strParam, iErrStr, oiMall.FItemList(i).FSellCash, oiMall.FItemList(i).getLotteiMallSellYn)

	            If (ret1) Then
	                actCnt = actCnt+1
	            Else
	                CALL Fn_AcctFailTouch("lotteimall", oiMall.FItemList(i).Fitemid, iErrStr)
	                retErrStr = retErrStr & iErrStr
	            End If

			Else
				CALL Fn_AcctFailTouch("lotteimall", oiMall.FItemList(i).Fitemid, iErrStr)
				iErrStr = "["&oiMall.FItemList(i).Fitemid&"] 옵션검사 실패"
				retErrStr = retErrStr & iErrStr
			End If
		Next
	Set oiMall = Nothing

    If (retErrStr <> "") Then
        Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
    End If
ElseIf (cmdparam = "EditSelect") Then				''선택상품정보/가격 수정
    dim skipItem 
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbget.Close: Response.End
	End If

	'## 수정된 상품 목록 접수
	Set oiMall = new CLotteiMall
		oiMall.FPageSize    = 30
		oiMall.FRectItemID	= arrItemid
		''oiMall.FRectMatchCateNotCheck="on" ''2013/08/30추가
		oiMall.getLTiMallEditedItemList
rw oiMall.FResultCount&"건"
		For i = 0 to (oiMall.FResultCount - 1)
			'//상품등록 파라메터
			On Error Resume Next
			If (oiMall.FItemList(i).FmaySoldOut = "Y") Then
				iErrStr = ""
				chgSellYn = CHKIIF(oiMall.FItemList(i).FLtiMallSellYn = "X", "X", "N")
				If (LtiMallOneItemSellStatEdit(oiMall.FItemList(i).Fitemid, oiMall.FItemList(i).FLtiMallGoodNo, chgSellYn, iErrStr)) Then
					actCnt = actCnt+1
					rw "["&oiMall.FItemList(i).Fitemid&"]"&"품절처리"
				Else
					rw "["&oiMall.FItemList(i).Fitemid&"]"&iErrStr
				End if
			Else
				'2013-07-01 전시상품 단품추가 될 경우 추가
				Dim dp, aoptNm, aoptDc
				strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_ltimall '"&CMallName&"'," & oiMall.FItemList(i).Fitemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open strSql, dbget
				If Not(rsget.EOF or rsget.BOF) Then
				    arrRows = rsget.getRows
				End If
				rsget.close
                
                skipItem=false
                ''추가된 옵션 등록
				If isArray(arrRows) Then
				    if (UBound(ArrRows, 2)<100) then ''옵션수가 너무 많은케이스 재낌 - 444286 // 2014/09/02 옵션 많은상품에 대해 정리 필요
    					For dp = 0 To UBound(ArrRows, 2)
    						If (ArrRows(11,dp)=0) and ArrRows(12,dp) = "1" AND ArrRows(15,dp) = "" Then		'옵션명이 다르고 옵션코드값이 없을 때 ==> 단품추가 의미// preged 0
    							aoptNm = Replace(db2Html(ArrRows(2,dp)),":","")
    							If aoptNm = "" Then
    								aoptNm = "옵션"
    							End If
    							aoptDc = aoptDc & Replace(Replace(db2Html(ArrRows(3,dp)),":",""),"'","")&","
    						End If
    					Next
    
    					If aoptDc <> "" Then
    					    rw "단품추가:"&aoptDc
    						aoptDc = Left(aoptDc, Len(aoptDc) - 1)
    						strParam = oiMall.FItemList(i).getLotteiMallAddOptParameter(aoptNm, aoptDc)
    						CALL LtiMallAddOpt(oiMall.FItemList(i).Fitemid, strParam, iErrStr)
    					End If
    				else
    				    skipItem =true
    			    end if
				End If
				'2013-07-01 단품추가 될 경우 추가 끝

'				If (oiMall.FItemList(i).FoptionCnt > 0) and (oiMall.FItemList(i).FregedOptCnt < 1) Then
                    '' 이곳에서도 타임아웃
                    iErrStr = ""
					CALL LtiMallOneItemCheckStock(oiMall.FItemList(i).Fitemid,iErrStr)
					rw iErrStr
'				End If

'			2014-10-29 10:05 김진영 // 타임아웃이 아래 원인인가 싶어 당분간 주석처리 시작
'				strParam = ""
'				strParam = oiMall.FItemList(i).getLotteiMallCateParamToEdit()				'2013-07-02진영추가 // 전시카테고리 수정(상품수정API에서 수정 안 됨)
'				If Err <> 0 Then
'					response.write Err.description
'					Response.Write "<script language=javascript>alert('텐바이텐 getLotteiMallCateParamToEdit 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oiMall.FItemList(i).Fitemid & "]');</script>"
'					dbget.Close: Response.End
'				End If
'
'				CALL LtiMallOneItemCateEdit(oiMall.FItemList(i).Fitemid, strParam, iErrStr)
'				rw iErrStr
'			2014-10-29 10:05 김진영 // 타임아웃이 아래 원인인가 싶어 당분간 주석처리 끝

				strParam = ""
				strParam = oiMall.FItemList(i).getLotteiMallItemEditParameter()
			    If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
					'rw ltiMallAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte?" & strParam
			    End If

				If Err <> 0 Then
					response.write Err.description
					Response.Write "<script language=javascript>alert('텐바이텐 getLotteiMallItemEditParameter 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oiMall.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				End If
		        on Error Goto 0

				iErrStr = ""
				ret1 = LtiMallOneItemInfoEdit(oiMall.FItemList(i).Fitemid, strParam, iErrStr, FALSE)

				If (ret1) Then
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_LTiMall_regItem " & VBCRLF
					strSql = strSql & "	SET LtiMallLastUpdate = getdate() " & VBCRLF
'					strSql = strSql & "	,LtiMallSellYn = '" & oiMall.FItemList(i).getLTiMallSellYn & "'" & VBCRLF
					strSql = strSql & " WHERE itemid = '" & oiMall.FItemList(i).Fitemid & "'"
					dbget.Execute(strSql)
					actCnt = actCnt + 1
				Else
					CALL Fn_AcctFailTouch("lotteimall", oiMall.FItemList(i).Fitemid, iErrStr)
					rw "[정보수정오류::]"&iErrStr
				End If

		        If (ret1) and (oiMall.FItemList(i).FSellcash <> oiMall.FItemList(i).FLtiMallPrice) Then
		            strParam = oiMall.FItemList(i).getLtiMallItemPriceEditParameter()

					If Err <> 0 Then
						Response.Write "<script language=javascript>alert('텐바이텐 PriceEditParameter 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oiMall.FItemList(i).Fitemid & "]');</script>"
						dbget.Close: Response.End
					End If

		            ret1 = false
		            ret1 = LtiMallOnItemPriceEdit(oiMall.FItemList(i).Fitemid,strParam,iErrStr)

					If (ret1) Then
					    '// 상품가격정보 수정
					    strSql = ""
		    			strSql = strSql & " UPDATe db_item.dbo.tbl_LTiMall_regItem  " & VbCRLF
		    			strSql = strSql & "	SET LtiMallLastUpdate=getdate() " & VbCRLF
		    			strSql = strSql & "	, LtiMallPrice = " & oiMall.FItemList(i).MustPrice & VbCRLF
		    			strSql = strSql & "	, accFailCnt = 0"& VbCRLF
		    			strSql = strSql & " Where itemid='" & oiMall.FItemList(i).Fitemid & "'"& VbCRLF
		    			dbget.Execute(strSql)
		            Else
		                CALL Fn_AcctFailTouch("lotteimall",oiMall.FItemList(i).Fitemid,iErrStr)
		                rw "[가격수정오류]"&iErrStr
					End If
		        End If

				'2013-10-08 14:01 김진영 하단 추가(재고조회부분) 한번 더 조회
				iErrStr = ""
				ret1 = false
				ret1 = LtiMallOneItemCheckStock(oiMall.FItemList(i).Fitemid, iErrStr)

				If (oiMall.FItemList(i).Fitemid <> "") Then
					iLotteItemStat = CheckLtiMallItemStat(oiMall.FItemList(i).Fitemid, iErrStr, iLotteSalePrc, iLotteGoodsNm)
					If (iLotteItemStat <> "") Then
						rw "["&oiMall.FItemList(i).Fitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr
					    CALL checkConfirmMatch(oiMall.FItemList(i).Fitemid,iLotteItemStat,iLotteSalePrc,iLotteGoodsNm)
					Else
					    sqlstr = "Update R" & VbCRLF
			        	sqlstr = sqlstr & " SET lastStatCheckDate=getdate()"
			        	sqlstr = sqlstr & " FROM db_item.dbo.tbl_LTiMall_regItem R" & VbCRLF
			        	sqlstr = sqlstr & " WHERE R.itemid="&oiMall.FItemList(i).Fitemid & VbCRLF
			        	''rw sqlstr
			        	dbget.Execute sqlstr
					    rw "["&oiMall.FItemList(i).Fitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr
					End If
			    End If

		    End If
		    retErrStr = retErrStr & iErrStr
		Next
	Set oiMall = Nothing

	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
	''rw "Err.Number"&Err.Number

	If (retFlag <> "") Then
		Response.Write "<script language=javascript>parent."&retFlag&";</script>"
		response.end
	End If

	IF (session("ssBctID")="icommang") then
		response.end
	End If
ElseIf (cmdparam = "EditSelect2") Then				''선택상품정보 수정(등록대기)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbget.Close: Response.End
	End If
	'## 수정된 상품 목록 접수
	Set oiMall = new CLotteiMall
		oiMall.FPageSize       = 30
		oiMall.FRectItemID	= arrItemid
		oiMall.getLTiMallEditedItemList

		If (oiMall.FResultCount < 1) Then
		   rw "수정 가능상품 없음"& arrItemid
		End If

		For i = 0 to (oiMall.FResultCount - 1)
			on Error Resume Next
			strParam = oiMall.FItemList(i).getLotteiMallItemEditParameter2()

			If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
				rw ltiMallAPIURL & "/openapi/upateApiNewGoodsInfo.lotte?" & strParam
			End If

			If Err <> 0 Then
				response.write Err.description
				Response.Write "<script language=javascript>alert('텐바이텐 EditParameter2 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oiMall.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If
			on Error Goto 0

			iErrStr = ""
			ret1 = LtiMallOneItemInfoEdit(oiMall.FItemList(i).Fitemid,strParam,iErrStr,TRUE)

			If (ret1) Then
			    '// 상품정보 수정
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_LTiMall_regitem " & VBCRLF
				sqlStr = sqlStr & "	SET LtiMallLastUpdate = getdate() " & VBCRLF
				sqlStr = sqlStr & "	,LtiMallSellYn='" & oiMall.FItemList(i).getLTiMallSellYn & "'" & VBCRLF
				sqlStr = sqlStr & "	,accFailCnt=0"& VBCRLF
				sqlStr = sqlStr & " WHERE itemid='" & oiMall.FItemList(i).Fitemid & "'"
				dbget.Execute(sqlStr)
				actCnt = actCnt + 1
			Else
			    CALL Fn_AcctFailTouch("lotteimall",oiMall.FItemList(i).Fitemid,iErrStr)
			    rw "[정보수정오류]"&iErrStr
			End If
			retErrStr = retErrStr & iErrStr
		Next
	set oiMall = Nothing
	If (retErrStr <> "") Then
	    Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If

	If (session("ssBctID")="icommang") Then
		response.end
	End If
ElseIf (cmdparam = "CheckNDelReged") Then				'2014-02-21 김진영//상품종료상품 X버튼 클릭하여 삭제기능 추가
	iLotteItemStat = CheckLtiMallItemStat(delitemid, iErrStr, iLotteSalePrc, iLotteGoodsNm)
	If (iLotteItemStat = "30") Then
		strSql = ""
		strSql = strSql & " delete from db_item.dbo.tbl_LTiMall_regItem "
		strSql = strSql & " where LtiMallSellYn in ('X')"
		strSql = strSql & " and itemid in ("&delitemid&")"
		dbget.Execute strSql,AssignedRow
		actCnt = AssignedRow

		strSql = ""
		strSql = strSql & " delete from db_item.dbo.tbl_Outmall_regedoption "
		strSql = strSql & " where itemid in ("&delitemid&")"
		strSql = strSql & " and mallid = '"&CMALLNAME&"' "
		dbget.Execute strSql
		dbget.Close() : response.end
	Else
		If (iLotteItemStat = "") and (iErrStr = "검색결과가 없습니다.") Then
			strSql = "delete from db_item.dbo.tbl_LTiMall_regItem "
			strSql = strSql & " where LtiMallSellYn in ('X')"
			strSql = strSql & " and itemid in ("&delitemid&")"
			dbget.Execute strSql,AssignedRow
			actCnt = AssignedRow
			response.write "<script>alert('삭제 - ERR: 판매상태 : ["&iLotteItemStat&"]:"&iErrStr&"');</script>"
		Else
			response.write "<script>alert('ERR: 판매상태 : ["&iLotteItemStat&"]:"&iErrStr&"');</script>"
		End If
	End If
End If

If Err.Number = 0 Then
	If (IsAutoScript) then
		rw "OK|"& iMessage & "<br>"& actCnt & "건이 처리되었습니다."
	Else
		Response.Write "<script language=javascript>alert('" & iMessage & "\n"& actCnt & "건이 처리되었습니다.');</script>"
	End if
Else
	If (IsAutoScript) then
		rw "S_ERR|처리 중에 오류가 발생했습니다"
	Else
		Response.Write "<script language=javascript>alert('자료 저장 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
	End if
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->