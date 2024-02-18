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
''�Ե����̸� ��ǰ���� ��ȸ
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
	    iErrStr = "["&iitemid&"] �Ե� ���̸� �ڵ� ����."
	    Exit Function
	End If

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		strParam = "?subscriptionId=" & ltiMallAuthNo				'�Ե����̸� ������ȣ	(*)
		strParam = strParam & "&search_type=goods_no"
		strParam = strParam & "&search_value=" & ilottegoods_no		'�Ե����̸� ��ǰ��ȣ	(*)

		objXML.Open "GET", ltiMallAPIURL & "/openapi/searchStockList.lotte"&strParam, false
''rw ltiMallAPIURL & "/openapi/searchStockList.lotte"&strParam
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML replace(buf,"&","��")
				If Err <> 0 Then
					iErrStr =  "�Ե����̸� ��� �м� �߿� ������ �߻� LtiMallOneItemCheckStock[" & iitemid & "]:"
					Set objXML = Nothing
					Set xmlDOM = Nothing
					Exit function
				End If
				ProdCount   = Trim(xmlDOM.getElementsByTagName("GoodsCount").item(0).text)   '' ��ǰ ����

				'If (ProdCount = "1") Then                                                   ''2013/07/03 �ּ�ó��
				'	If (getItemOptionCount(iitemid) < 1) Then ProdCount = ""
				'End If
				If (ProdCount <> "") Then
			        Set oneProdInfo = xmlDOM.getElementsByTagName("GoodsInfoList")
			        strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption='')"
			        strSql = strSql & " BEGIN"
			        strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption=''"
			        strSql = strSql & " END"
			        dbget.Execute strSql

			        ''2013/05/30 �߰�
			        strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and Len(outmalloptCode)>6)"
			        strSql = strSql & " BEGIN"
			        strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and Len(outmalloptCode)>6"
			        strSql = strSql & " END"
			        dbget.Execute strSql
					For each SubNodes in oneProdInfo
						GoodNo	    = Trim(SubNodes.getElementsByTagName("GoodNo").item(0).text)
					    ItemNo	    = Trim(SubNodes.getElementsByTagName("ItemNo").item(0).text)        '' ��ǰ�ڵ� (���� 0,1,2,)
					    OptDesc	    = Trim(SubNodes.getElementsByTagName("OptDesc").item(0).text)
					    ''DispYn	    = Trim(SubNodes.getElementsByTagName("DispYn").item(0).text)         ''N:���� Y:����            ''2013/07/03 ����� �ȳѾ���°� ���� SaleStatCd �κ���
					    SaleStatCd	= Trim(SubNodes.getElementsByTagName("SaleStatCd").item(0).text) ''�Ǹ�����, �Ǹ�����, ǰ��'
					    StockQty	= Trim(SubNodes.getElementsByTagName("StockQty").item(0).text)
                        CorpItemNo  = Trim(SubNodes.getElementsByTagName("CorpItemNo").item(0).text)  '' ��ǰ�ڵ�_�ɼ��ڵ�

						getRegOptCD = Split(CorpItemNo,"_")(1)
					    OptDesc = replace(OptDesc, "��", "&")
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

						'################  2014-02-21 14:34 ������ ######################
						'OptDesc = replace(OptDesc, ",,", ",")���� �߰�..
						'���� : [db_item].[dbo].tbl_item_option_Multiple �� optionTypeName�� ,�� �� �ִ� ��찡 ����..
 						'ex)�̴ϼȰ���(5,000��)..�̷��� �Ǿ����� �� ,�� split��Ŵ�� ���� ,,�� ,�� ġȯ��
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
					    strSql = strSql & " and convert(int, outmallOptCode)='"&ItemNo&"'"&VbCRLF				'������ outmallOptCode�� 001,002,003 �̷��� �������� ���� �Ŀ� 1,2,3�̷��� ����
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
					        strSql = strSql & " ,'"&ItemNo&"'" ''�ӽ÷� �Ե� �ڵ� ���� //2013/04/01
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

					        ''�ɼ� �ڵ� ��Ī.
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
					    	''���� ��ǰ�� �� tbl_OutMall_regedoption�� �����Ͱ� ������ tbl_item_option�� �����Ͱ� ���⿡ �ϴ� ���ν��� ȣ��
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
		    iErrStr =  "�Ե����̸��� ����߿� ������ �߻� [" & iitemid & "]:"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit Function
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'����ī�װ� ����
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
				    iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.LtiMallOneItemCateEdit"
					dbget.Close: Exit Function
				End If

				'��� �ڵ�
				on Error resume next
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
		        on Error Goto 0

				'// ���� ����
				If Err <> 0 Then
		            iErrStr = "Error: " & xmlDOM.getElementsByTagName("Message").item(0).text
					dbget.Close: Exit Function
		        End If
			Set xmlDOM = Nothing
		Else
		    iErrStr ="�Ե����̸��� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����"
			dbget.Close: Exit Function
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���û�ǰ ��ǰ�߰�
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
					iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-NMEDIT-001]"
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

				'// ���� ����
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
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''�Ե����̸� ��ǰ ���
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
				    iErrStr =  "��ǰ ����� ���� [" & iitemid & "]:"&iMessage
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
				    Exit Function
				End If

				If Err <> 0 Then
					iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit function
				End If

				'��ǰ���翩�� Ȯ��
				strSql = "Select count(itemid) From db_item.dbo.tbl_LTiMall_regItem Where itemid='" & iitemid & "'"
				rsget.Open strSql,dbget,1

				If rsget(0) > 0 Then
					'// ���� -> ����
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					strSql = strSql & "	Set LTiMallLastUpdate = getdate() "  & VbCRLF
					strSql = strSql & "	, LTiMallTmpGoodNo = '" & LotteGoodNo & "'"  & VbCRLF
					strSql = strSql & "	, LTiMallPrice = " &iSellCash& VbCRLF
					strSql = strSql & "	, accFailCnt = 0"& VbCRLF
					strSql = strSql & "	, LTiMallRegdate = isNULL(LTiMallRegdate, getdate())" ''�߰� 2013/02/26
					If (LotteGoodNo <> "") Then
					    strSql = strSql & "	, LTiMallstatCD = '20'"& VbCRLF
					Else
						strSql = strSql & "	, LTiMallstatCD = '10'"& VbCRLF
					End If
					strSql = strSql & "	From db_item.dbo.tbl_LTiMall_regItem R"& VbCRLF
					strSql = strSql & " Where R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
				Else
					'// ���� -> �űԵ��
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
			    ''�ɼ� ����Ʈ ����(20120807 �߰�)
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
			                ''���߿ɼ��ΰ�� �� ���� �߸� �Ǿ���.
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

							''�ɼ� �ڵ� ��Ī.
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
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-REG-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function LtiMallOneItemSellStatEdit(iitemid, iLotteGoodNo, ichgSellYn, byRef iErrStr)
    Dim strParam
    Dim objXML, xmlDOM
    Dim strRst, strSql, notitemId

    LtiMallOneItemSellStatEdit = False
	strParam = "?subscriptionId=" & ltiMallAuthNo										'�Ե����̸� ������ȣ	(*)
	strParam = strParam & "&goods_no=" & iLotteGoodNo                       			'�Ե����̸� ��ǰ��ȣ	(*)
'	strParam = strParam & "&brnd_no=1099329"			                       			'�Ե����̸� �귣���ڵ�

'	If ichgSellYn = "Y" Then															'�Ǹſ���(10:�Ǹ�, 20:ǰ��, 30:�Ǹ�����)
'		strParam = strParam & "&item_sale_stat_cd=10"
'	ElseIf ichgSellYn = "N" Then
'		strParam = strParam & "&item_sale_stat_cd=20"
'	ElseIf ichgSellYn = "X" Then                        '''X ��� ������
'		strParam = strParam & "&item_sale_stat_cd=30"      ''�Ǹ�����Ǹ� ��������.
'		''strParam = strParam & "&sale_stat_cd=20"
'	End If

	If ichgSellYn = "Y" Then															'�Ǹſ���(10:�Ǹ�, 20:ǰ��, 30:�Ǹ�����)
		strParam = strParam & "&sale_stat_cd=10"
	ElseIf ichgSellYn = "N" Then
		strParam = strParam & "&sale_stat_cd=20"
	ElseIf ichgSellYn = "X" Then                        '''X ��� ������
		''''''strParam = strParam & "&sale_stat_cd=30"      ''�Ǹ�����Ǹ� ��������.
		strParam = strParam & "&sale_stat_cd=20"
	End If

	strSql = ""
	strSql = "select count(*) as cnt from db_temp.dbo.tbl_jaehyumall_not_in_itemid where mallgubun = 'lotteimall' and itemid =" & iitemid
	rsget.Open strSql, dbget
	If Not(rsget.EOF or rsget.BOF) Then
		notitemId = rsget("cnt")
	End If
	rsget.close

	'2013-07-18 ������ �߰�..������ܵ� ��ǰ�� �ڲ� �ǸŵǾ...
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
				    iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.LtiMallOneItemSellStatEdit"
					dbget.Close: Exit Function
				End If

				'��� �ڵ�
				on Error resume next
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
		        on Error Goto 0

				'// ���� ����
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
		    iErrStr ="�Ե����̸��� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����"
			dbget.Close: Exit Function
		End If
	Set objXML = Nothing
	LtiMallOneItemSellStatEdit = true
End Function

''���û�ǰ��ȣ ��������
Function CheckLtiMallTmpItemChk(iitemid,byRef iErrStr, byRef iLotteGoodNo, byref iLotteStatCd)
	Dim objXML,xmlDOM,strRst,iMessage
	Dim strParam, iLotteTmpID , SaleStatCd, GoodsViewCount
	Dim iRbody

	CheckLtiMallTmpItemChk = ""
	iLotteTmpID = getLtiMallTmpItemIdByTenItemID(iitemid)

	if iLotteTmpID="" then Exit function '' 2013/09/03 �߰�

	If iLotteTmpID = "���û�ǰ" Then
		CheckLtiMallTmpItemChk = "���û�ǰ"
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
						iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�."
					    Set objXML = Nothing
					    Set xmlDOM = Nothing
					    Exit Function
					End If

					GoodsViewCount 		= Trim(xmlDOM.getElementsByTagName("GoodsCount").item(0).text)			'�˻���

				    If Err <> 0 Then
					    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
					End If

					If (GoodsViewCount = "1") Then
						iLotteGoodNo		= Trim(xmlDOM.getElementsByTagName("GoodsNo").item(0).text)			'���û�ǰ��ȣ
						iLotteStatCd		= Trim(xmlDOM.getElementsByTagName("ConfStatCd").item(0).text)		'���������ڵ�
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
				    iErrStr ="�˻������ �����ϴ�."&iMessage
				Else
				    CheckLtiMallTmpItemChk = iLotteStatCd
			    End If
			Else
				iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-ItemChk-002]"
			End If
		Set objXML = Nothing
		On Error Goto 0
	End If
End Function

''���û�ǰ ��ȸ
Function CheckLtiMallItemStat(iitemid,byRef iErrStr, byRef iSalePrc, byref iGoodsNm)
	Dim objXML,xmlDOM,strRst,iMessage
	Dim strParam, iLotteItemID , SaleStatCd, GoodsViewCount
	Dim iRbody

	CheckLtiMallItemStat = ""
	iLotteItemID = getLtiMallItemIdByTenItemID(iitemid)
	strParam = "subscriptionId=" & ltiMallAuthNo & "&goods_no="&iLotteItemID
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		'' rw "https://openapi.lotte.com/openapi/searchGoodsListOpenApiOther.lotte?"&strParam ''''���û�ǰ��ȸaLL
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchGoodsListOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			'XML�� ���� DOM ��ü ����
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
			''rw iRbody
'			iRbody = replace(iRbody,"&","@@amp@@")   '' <![CDATA[]]> �� �� ������. ��ǰ�� < > ����..
'			iRbody = replace(iRbody,"<GoodsNm>","<GoodsNm><![CDATA[")
'			iRbody = replace(iRbody,"</GoodsNm>","]]></GoodsNm>")
			xmlDOM.LoadXML iRbody

			if Err <> 0 Then
				iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-ItemChk-001]"
			    Set objXML = Nothing
			    Set xmlDOM = Nothing
			    Exit Function
			End If
			GoodsViewCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text  ''�����

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

			'// ���� ����
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
			    iErrStr ="�˻������ �����ϴ�.."&iMessage
			ElseIf (SaleStatCd<>"0") then
			    CheckLtiMallItemStat = SaleStatCd
		    End If
		Else
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-ItemChk-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0

	'	rw "SaleStatCd="&SaleStatCd
	'	rw "GoodsViewCount="&GoodsViewCount
	'	rw "iMessage="&iMessage
	'	rw "iErrStr="&iErrStr
End Function


''���û�ǰ ��ȸ(��¥��)
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
		rw "<input type='button' value='������¥' onclick=location.href='/admin/etc/LtiMall/actLotteiMallReq.asp?cmdparam=ChkDate&chkdate="&dateadd("d",1,chkdate)&"';>"

		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			'XML�� ���� DOM ��ü ����
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
			xmlDOM.LoadXML iRbody

			if Err <> 0 Then
				iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-DateChk-001]"
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

			GoodsViewCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text  ''�����
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
					rw GoodsNo&" ��ϿϷ�"
				End If
				rsget.Close
			Next

			'// ���� ����
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
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-ItemChk-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function checkConfirmMatch(iitemid, iLotteItemStat, iLotteSalePrc, iLotteGoodsNm) ''�Ե����̸� (��)�ǸŻ���, (��)�ǸŰ� ������Ʈ, ��ǰ�� �߰�
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
	sqlstr = sqlstr & " ,LtiMallStatCd=(CASE WHEN isNULL(LtiMallStatCd,-9)<7 THEN 7 ELSE LtiMallStatCd END)" ''2013/09/03 �߰�
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
	    ''�ٸ��� ������ �α�.
	    CALL Fn_AcctFailLog(CMALLNAME, iitemid, iLotteItemStat&","&LotteSellyn&", "&iLotteSalePrc&"::"&pbuf, "STAT_CHK")
	End If
End Function

''���û�ǰ�����
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
					iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-NMEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If

				strRst = xmlDOM.getElementsByTagName("Result").item(0).text

			    If Err <> 0 Then
				    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
				End If

				'// ���� ����
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
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���û�ǰ �ǸŰ� ����
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
					iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRCEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text

				If Err <> 0 Then
					iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
				End If

				If Err <> 0 Then
				    If (Trim(iMessage)="002. ��������(�Ǵ� ���࿹��)") Then
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
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-PRCEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''�Ե����̸� ��ǰ���� ����
Function LtiMallOneItemInfoEdit(iitemid, strParam, byRef iErrStr, isVer2)
	Dim objXML, xmlDOM, strRst, iMessage
	On Error Resume Next
	LtiMallOneItemInfoEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	If (isVer2) Then
	    objXML.Open "POST", ltiMallAPIURL & "/openapi/upateApiNewGoodsInfo.lotte", false          ''��ǰ����
	Else
	    objXML.Open "POST", ltiMallAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte", false      ''���û�ǰ����
	End If

'rw ltiMallAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte?"&strParam
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then

		'//���޹��� ���� Ȯ��
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				''rw BinaryToText(objXML.ResponseBody, "euc-kr")
				If Err<>0 then
				    iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�

				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
			    If Err <> 0 Then
				    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
				End If

			    If (strRst <> "1") Then
			        ''rw "FitemDiv="&oiMall.FItemList(i).FitemDiv  ''�ֹ����� ��ǰ�ΰ�� ������ �ȵǴµ�. add_choc_tp_cd_20 �� ���� �ִ°��;;
			        ''rw "iMessage="&iMessage
			        iErrStr =  "��ǰ ������ ���� [" & iitemid & "]:"&iMessage
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
		    iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-EDIT-002]"
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
					''///////////////						�������� ���, �Ʒ����� ���μ��� 				//////////////////
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
If (cmdparam = "RegSelectWait") Then						''���û�ǰ ���� ���.
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
	response.write "<script>alert('"&AssignedRow&"�� ������ϵ�.');</script>"

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
ElseIf (cmdparam = "DelSelectWait") Then					''���û�ǰ ���� ����.
	arrItemid = Trim(arrItemid)
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM db_item.dbo.tbl_LTiMall_regItem "
	sqlStr = sqlStr & " WHERE LtimallStatCD in ('0')"
	sqlStr = sqlStr & " and itemid in ("&arrItemid&")"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� ���� ������.');parent.location.reload();</script>"
ElseIf (cmdparam="ChkDate") Then
	Dim chkdate, iLotteDateItemStat
	chkdate = request("chkdate")

    iErrStr = ""
	iLotteDateItemStat = CheckLtiMallDateItemStat(chkdate, iErrStr, iLotteSalePrc, iLotteGoodsNm)
	response.end
ElseIf (cmdparam = "EditSellYn") Then						''���û�ǰ �Ǹſ��� ����
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
ElseIf (cmdparam = "EditItemNm") Then							''���û�ǰ�������û
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ������ ��ǰ ��� ����
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
				Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oiMall.FItemList(i).Fitemid & "]');</script>"
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
ElseIf (cmdparam = "CheckItemNmAuto") Then						''��ǰ�����(������)
	buf = ""
	CNT10 = 0
	sqlStr = ""
	sqlStr = sqlStr & " SELECT top 10 r.itemid, isNULL(r.ltiMallGoodNo,r.ltimallTmpGoodNO) as ltiMallGoodNo, i.ItemName "
	sqlStr = sqlStr & "	FROM db_item.dbo.tbl_LTiMall_regItem r "
	sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_item i on r.itemid = i.itemid "
	sqlStr = sqlStr & "	WHERE r.regitemname is Not NULL "
	sqlStr = sqlStr & "	and replace(r.regitemname,'"&CPREFIXITEMNAME&"','') <> i.itemname "
	sqlStr = sqlStr & "	and isNULL(r.ltiMallGoodNo,r.ltimallTmpGoodNO) is Not NULL"  ''���� 2013/07/12
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
					'// ��ǰ�� ����
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
	rw CNT10&"�� ��ǰ����� ����"
	response.end
ElseIf (cmdparam = "EditPriceSelect") Then						''���û�ǰ���� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ������ ��ǰ ��� ����
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
				Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oiMall.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If
		    on Error Goto 0

		    iErrStr = ""
			ret1 = false
		    ret1 = LtiMallOnItemPriceEdit(oiMall.FItemList(i).Fitemid, strParam, iErrStr)
			If (ret1) Then
			    '// ��ǰ�������� ����
			    sqlStr = ""
				sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_LTiMall_regItem  " & VbCRLF
				sqlStr = sqlStr & "	SET LtiMallPrice = " & oiMall.FItemList(i).MustPrice & VbCRLF
				''sqlStr = sqlStr & "	, accFailCnt = 0"& VbCRLF                                           ''����������
				''sqlStr = sqlStr & "	, LtiMallLastUpdate = getdate() " & VbCRLF                          ''����������
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
ElseIf (cmdparam = "getconfirmList") Then					''���û�ǰ ���Ȯ��
	Dim AnotherStat
	arrItemid = split(Trim(arrItemid), ",")
	If IsArray(arrItemid) Then
		For i = LBound(arrItemid) to UBound(arrItemid)
		    iErrStr = ""
		    iitemid = Trim(arrItemid(i))
			If (iitemid <> "") Then
				iLotteItemTmpChk = CheckLtiMallTmpItemChk(iitemid, iErrStr, iLotteGoodNo, iLotteStatCd)
				Select Case iLotteItemTmpChk
					Case "10"	AnotherStat = "�ӽõ��"
					Case "20"	AnotherStat = "���ο�û"
					Case "30"	AnotherStat = "���οϷ�"
					Case "40"	AnotherStat = "�ݷ�"
					Case "50"	AnotherStat = "���κҰ�"
					Case "51"	AnotherStat = "����ο�û"
					Case "52"	AnotherStat = "������û"
					Case "60"	AnotherStat = "���"
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
				ElseIf iLotteItemTmpChk = "���û�ǰ" Then
					strSql =""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem "
					strSql = strSql & "	SET LtiMallStatCd='7' "
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute strSql
					rw "["&iitemid&"] :"&iLotteItemTmpChk&": �̹� ���� ��ǰ�Դϴ�"

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
					CALL LtiMallOneItemCheckStock(iitemid, iErrStr)				''2013-07-24 ������ �߰�
				End If
		    End If
		Next
	End If
	response.end
ElseIf (cmdparam = "CheckItemStat") Then					''���û�ǰ �ǸŻ���Ȯ��
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
ElseIf (cmdparam = "CheckItemStatAuto") Then				''�ǸŻ���Check(������)
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

ElseIf (cmdparam = "ChkStockSelect") Then				''���û�ǰ �����ȸ
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
ElseIf (cmdparam = "RegSelect") Then				''���û�ǰ ���� ���
    arrItemid = Trim(arrItemid)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ���û�ǰ ��� ����
	Set oiMall = new CLotteiMall
		oiMall.FPageSize	= 20
		oiMall.FRectItemID	= arrItemid
		oiMall.getLTiMallNotRegItemList
	    If (oiMall.FResultCount < 1) Then
	        arrItemid = split(arrItemid,",")
	        For i = LBound(arrItemid) to UBound(arrItemid)
	            CALL Fn_AcctFailTouch("lotteimall",arrItemid(i),"��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...")
	        Next

	        If (IsAutoScript) Then
	            rw "S_ERR|��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, ����..."
	            dbget.Close: Response.End
	        Else
	            Response.Write "<script language=javascript>alert('��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...');</script>"
				dbget.Close: Response.End
			End If
		End If

		For i = 0 to (oiMall.FResultCount - 1)
			sqlStr = ""
			sqlStr = sqlStr & " IF Exists(SELECT * FROM db_item.dbo.tbl_LTiMall_regitem where itemid="&oiMall.FItemList(i).Fitemid&")"
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " UPDATE R" & VbCRLF
			sqlStr = sqlStr & "	SET LTiMallLastUpdate = getdate() "  & VbCRLF
			sqlStr = sqlStr & "	, LTiMallstatCD='1'"& VbCRLF                     ''2013-06-14 ����//����� ��Ͻõ��� �ϴ� ����(�Ե������� 10, ���̸��� 1�� ����)
			sqlStr = sqlStr & "	FROM db_item.dbo.tbl_LTiMall_regitem R"& VbCRLF
			sqlStr = sqlStr & " WHERE R.itemid='" & oiMall.FItemList(i).Fitemid & "'"
			sqlStr = sqlStr & " END ELSE "
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_LTiMall_regitem "
	        sqlStr = sqlStr & " (itemid, regdate, reguserid, LTiMallstatCD)"
	        sqlStr = sqlStr & " VALUES ("&oiMall.FItemList(i).Fitemid&", getdate(), '"&session("SSBctID")&"', '1')"
			sqlStr = sqlStr & " END "
		    dbget.Execute sqlStr

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oiMall.FItemList(i).checkTenItemOptionValid Then
'			    On Error Resume Next
				'//��ǰ��� �Ķ����
				strParam = oiMall.FItemList(i).getLotteiMallItemRegParameter(FALSE)
				If (session("ssBctID") = "icommang") or (session("ssBctID") = "kjy8517") Then
					rw ltiMallAPIURL & "" & strParam
				End If

				If Err <> 0 Then
				    rw Err.Description
					Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oiMall.FItemList(i).Fitemid & "]');</script>"
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
				iErrStr = "["&oiMall.FItemList(i).Fitemid&"] �ɼǰ˻� ����"
				retErrStr = retErrStr & iErrStr
			End If
		Next
	Set oiMall = Nothing

    If (retErrStr <> "") Then
        Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
    End If
ElseIf (cmdparam = "EditSelect") Then				''���û�ǰ����/���� ����
    dim skipItem 
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ������ ��ǰ ��� ����
	Set oiMall = new CLotteiMall
		oiMall.FPageSize    = 30
		oiMall.FRectItemID	= arrItemid
		''oiMall.FRectMatchCateNotCheck="on" ''2013/08/30�߰�
		oiMall.getLTiMallEditedItemList
rw oiMall.FResultCount&"��"
		For i = 0 to (oiMall.FResultCount - 1)
			'//��ǰ��� �Ķ����
			On Error Resume Next
			If (oiMall.FItemList(i).FmaySoldOut = "Y") Then
				iErrStr = ""
				chgSellYn = CHKIIF(oiMall.FItemList(i).FLtiMallSellYn = "X", "X", "N")
				If (LtiMallOneItemSellStatEdit(oiMall.FItemList(i).Fitemid, oiMall.FItemList(i).FLtiMallGoodNo, chgSellYn, iErrStr)) Then
					actCnt = actCnt+1
					rw "["&oiMall.FItemList(i).Fitemid&"]"&"ǰ��ó��"
				Else
					rw "["&oiMall.FItemList(i).Fitemid&"]"&iErrStr
				End if
			Else
				'2013-07-01 ���û�ǰ ��ǰ�߰� �� ��� �߰�
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
                ''�߰��� �ɼ� ���
				If isArray(arrRows) Then
				    if (UBound(ArrRows, 2)<100) then ''�ɼǼ��� �ʹ� �������̽� �糦 - 444286 // 2014/09/02 �ɼ� ������ǰ�� ���� ���� �ʿ�
    					For dp = 0 To UBound(ArrRows, 2)
    						If (ArrRows(11,dp)=0) and ArrRows(12,dp) = "1" AND ArrRows(15,dp) = "" Then		'�ɼǸ��� �ٸ��� �ɼ��ڵ尪�� ���� �� ==> ��ǰ�߰� �ǹ�// preged 0
    							aoptNm = Replace(db2Html(ArrRows(2,dp)),":","")
    							If aoptNm = "" Then
    								aoptNm = "�ɼ�"
    							End If
    							aoptDc = aoptDc & Replace(Replace(db2Html(ArrRows(3,dp)),":",""),"'","")&","
    						End If
    					Next
    
    					If aoptDc <> "" Then
    					    rw "��ǰ�߰�:"&aoptDc
    						aoptDc = Left(aoptDc, Len(aoptDc) - 1)
    						strParam = oiMall.FItemList(i).getLotteiMallAddOptParameter(aoptNm, aoptDc)
    						CALL LtiMallAddOpt(oiMall.FItemList(i).Fitemid, strParam, iErrStr)
    					End If
    				else
    				    skipItem =true
    			    end if
				End If
				'2013-07-01 ��ǰ�߰� �� ��� �߰� ��

'				If (oiMall.FItemList(i).FoptionCnt > 0) and (oiMall.FItemList(i).FregedOptCnt < 1) Then
                    '' �̰������� Ÿ�Ӿƿ�
                    iErrStr = ""
					CALL LtiMallOneItemCheckStock(oiMall.FItemList(i).Fitemid,iErrStr)
					rw iErrStr
'				End If

'			2014-10-29 10:05 ������ // Ÿ�Ӿƿ��� �Ʒ� �����ΰ� �;� ��а� �ּ�ó�� ����
'				strParam = ""
'				strParam = oiMall.FItemList(i).getLotteiMallCateParamToEdit()				'2013-07-02�����߰� // ����ī�װ� ����(��ǰ����API���� ���� �� ��)
'				If Err <> 0 Then
'					response.write Err.description
'					Response.Write "<script language=javascript>alert('�ٹ����� getLotteiMallCateParamToEdit ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oiMall.FItemList(i).Fitemid & "]');</script>"
'					dbget.Close: Response.End
'				End If
'
'				CALL LtiMallOneItemCateEdit(oiMall.FItemList(i).Fitemid, strParam, iErrStr)
'				rw iErrStr
'			2014-10-29 10:05 ������ // Ÿ�Ӿƿ��� �Ʒ� �����ΰ� �;� ��а� �ּ�ó�� ��

				strParam = ""
				strParam = oiMall.FItemList(i).getLotteiMallItemEditParameter()
			    If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
					'rw ltiMallAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte?" & strParam
			    End If

				If Err <> 0 Then
					response.write Err.description
					Response.Write "<script language=javascript>alert('�ٹ����� getLotteiMallItemEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oiMall.FItemList(i).Fitemid & "]');</script>"
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
					rw "[������������::]"&iErrStr
				End If

		        If (ret1) and (oiMall.FItemList(i).FSellcash <> oiMall.FItemList(i).FLtiMallPrice) Then
		            strParam = oiMall.FItemList(i).getLtiMallItemPriceEditParameter()

					If Err <> 0 Then
						Response.Write "<script language=javascript>alert('�ٹ����� PriceEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oiMall.FItemList(i).Fitemid & "]');</script>"
						dbget.Close: Response.End
					End If

		            ret1 = false
		            ret1 = LtiMallOnItemPriceEdit(oiMall.FItemList(i).Fitemid,strParam,iErrStr)

					If (ret1) Then
					    '// ��ǰ�������� ����
					    strSql = ""
		    			strSql = strSql & " UPDATe db_item.dbo.tbl_LTiMall_regItem  " & VbCRLF
		    			strSql = strSql & "	SET LtiMallLastUpdate=getdate() " & VbCRLF
		    			strSql = strSql & "	, LtiMallPrice = " & oiMall.FItemList(i).MustPrice & VbCRLF
		    			strSql = strSql & "	, accFailCnt = 0"& VbCRLF
		    			strSql = strSql & " Where itemid='" & oiMall.FItemList(i).Fitemid & "'"& VbCRLF
		    			dbget.Execute(strSql)
		            Else
		                CALL Fn_AcctFailTouch("lotteimall",oiMall.FItemList(i).Fitemid,iErrStr)
		                rw "[���ݼ�������]"&iErrStr
					End If
		        End If

				'2013-10-08 14:01 ������ �ϴ� �߰�(�����ȸ�κ�) �ѹ� �� ��ȸ
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
ElseIf (cmdparam = "EditSelect2") Then				''���û�ǰ���� ����(��ϴ��)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
	Set oiMall = new CLotteiMall
		oiMall.FPageSize       = 30
		oiMall.FRectItemID	= arrItemid
		oiMall.getLTiMallEditedItemList

		If (oiMall.FResultCount < 1) Then
		   rw "���� ���ɻ�ǰ ����"& arrItemid
		End If

		For i = 0 to (oiMall.FResultCount - 1)
			on Error Resume Next
			strParam = oiMall.FItemList(i).getLotteiMallItemEditParameter2()

			If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
				rw ltiMallAPIURL & "/openapi/upateApiNewGoodsInfo.lotte?" & strParam
			End If

			If Err <> 0 Then
				response.write Err.description
				Response.Write "<script language=javascript>alert('�ٹ����� EditParameter2 ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oiMall.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If
			on Error Goto 0

			iErrStr = ""
			ret1 = LtiMallOneItemInfoEdit(oiMall.FItemList(i).Fitemid,strParam,iErrStr,TRUE)

			If (ret1) Then
			    '// ��ǰ���� ����
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
			    rw "[������������]"&iErrStr
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
ElseIf (cmdparam = "CheckNDelReged") Then				'2014-02-21 ������//��ǰ�����ǰ X��ư Ŭ���Ͽ� ������� �߰�
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
		If (iLotteItemStat = "") and (iErrStr = "�˻������ �����ϴ�.") Then
			strSql = "delete from db_item.dbo.tbl_LTiMall_regItem "
			strSql = strSql & " where LtiMallSellYn in ('X')"
			strSql = strSql & " and itemid in ("&delitemid&")"
			dbget.Execute strSql,AssignedRow
			actCnt = AssignedRow
			response.write "<script>alert('���� - ERR: �ǸŻ��� : ["&iLotteItemStat&"]:"&iErrStr&"');</script>"
		Else
			response.write "<script>alert('ERR: �ǸŻ��� : ["&iLotteItemStat&"]:"&iErrStr&"');</script>"
		End If
	End If
End If

If Err.Number = 0 Then
	If (IsAutoScript) then
		rw "OK|"& iMessage & "<br>"& actCnt & "���� ó���Ǿ����ϴ�."
	Else
		Response.Write "<script language=javascript>alert('" & iMessage & "\n"& actCnt & "���� ó���Ǿ����ϴ�.');</script>"
	End if
Else
	If (IsAutoScript) then
		rw "S_ERR|ó�� �߿� ������ �߻��߽��ϴ�"
	Else
		Response.Write "<script language=javascript>alert('�ڷ� ���� �߿� ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���.');</script>"
	End if
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->