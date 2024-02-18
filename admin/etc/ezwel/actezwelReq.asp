<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 600 %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ezwel/ezwelcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'#################################### 이지웰페어 기본 정보 Setting ####################################
Public ezwelAPIURL
Public postParam
Const dataHead = "<?xml version=""1.0"" encoding=""euc-kr"" standalone=""yes""?>"

IF application("Svr_Info") = "Dev" THEN
	ezwelAPIURL = "http://api.dev.ezwel.com/if/api/goodsInfoAPI.ez"
Else
	ezwelAPIURL = "http://api.ezwel.com/if/api/goodsInfoAPI.ez"
End if
postParam	= "cspCd="&cspCd&"&crtCd="&crtCd&"&dataSet="
'#####################################################################################################

'################################### 각종 Function Setting  ##########################################
Function EzwelOneItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iezwelSellYn, ilimityn, ilimitno, ilimitsold, iitemname, iimageNm)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql, tenOptCnt
	Dim retCode, goodsCd, iMessage, AssignedRow
	If (xmlStr = "") Then
		EzwelOneItemReg = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
		objXML.send(postParam&xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If

		goodsCd		= xmlDOM.getElementsByTagName("goodsCd").item(0).text
		retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
		iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text

		If retCode = "200" Then		'성공(200)
			strSql = "SELECT COUNT(itemid) FROM db_outmall.dbo.tbl_ezwel_regItem WHERE itemid='" & iitemid & "' and ezwelgoodno = '"&goodsCd&"'"
			rsCTget.Open strSql, dbCTget, 1
			If rsCTget(0) = 0 Then
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCRLF
				strSql = strSql & "	Set ezwelLastUpdate = getdate() "  & VbCRLF
				strSql = strSql & "	, ezwelGoodNo = '" & goodsCd & "'"  & VbCRLF
				strSql = strSql & "	, ezwelPrice = " &iSellCash& VbCRLF
				strSql = strSql & "	, accFailCnt = 0"& VbCRLF
				strSql = strSql & "	, ezwelRegdate = isNULL(ezwelRegdate, getdate())"
				If (goodsCd <> "") Then
				    strSql = strSql & "	, ezwelstatCD = '7'"& VbCRLF					'등록완료(임시)
				Else
					strSql = strSql & "	, ezwelstatCD = '1'"& VbCRLF					'전송시도
				End If
				strSql = strSql & "	From db_outmall.dbo.tbl_ezwel_regItem R"& VbCRLF
				strSql = strSql & " Where R.itemid = '" & iitemid & "'"
				dbCTget.Execute(strSql)
			Else
				'// 없음 -> 신규등록
				strSql = ""
				strSql = strSql & " INSERT INTO db_outmall.dbo.tbl_ezwel_regItem "
				strSql = strSql & " (itemid, regitemname, reguserid, ezwelRegdate, ezwelLastUpdate, ezwelGoodNo, ezwelPrice, ezwelSellYn, ezwelStatCd, regImageName) VALUES " & VbCRLF
				strSql = strSql & " ('" & iitemid & "'" & VBCRLF
				strSql = strSql & " , '" & iitemname & "'" &_
				strSql = strSql & " , '" & session("ssBctId") & "'" &_
				strSql = strSql & " , getdate(), getdate()" & VBCRLF
				strSql = strSql & " , '" & goodsCd & "'" & VBCRLF
				strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
				strSql = strSql & " , '" & iezwelSellYn & "'" & VBCRLF
				If (goodsCd <> "") Then
				    strSql = strSql & ",'7'"											'등록완료(임시)
				Else
				    strSql = strSql & ",'1'"											'전송시도
				End If
				strSql = strSql & " , '" & iimageNm & "'" & VBCRLF
				strSql = strSql & ")"
				dbCTget.Execute(strSql)
				actCnt = actCnt + 1
			End If
			rsCTget.Close

			strSql = ""
			strSql = strSql &  "SELECT count(*) as cnt "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
			strSql = strSql & " WHERE itemid=" & iitemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.Open strSql,dbget,1
				tenOptCnt = rsget("cnt")
			rsget.Close

			strSql = ""
			strSql = strSql & " UPDATE db_outmall.dbo.tbl_ezwel_regItem SET "
			strSql = strSql & " regedOptCnt = " & tenOptCnt
			strSql = strSql & " WHERE itemid = " & iitemid
			dbCTget.Execute strSql
			EzwelOneItemReg = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage
		Else						'실패(E)
		    iErrStr =  "상품 등록중 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
		On Error Goto 0
	End If
End Function

Function EzwelOneItemEdit(iitemid, iEzwelGoodNo, byRef iErrStr, strParam, imustprice, isellyn, optMust)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql, tenOptCnt
	Dim retCode, goodsCd, iMessage, AssignedRow, oMsg, ocount
	If (xmlStr = "") Then
		EzwelOneItemEdit = false
		Exit Function
    End If
'rw oEzwel.FItemList(i).isImageChanged
'rw oEzwel.FItemList(i).getBasicImage
'rw "=-=-="
'response.end
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
		objXML.send(postParam&xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
			'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If

		goodsCd		= xmlDOM.getElementsByTagName("goodsCd").item(0).text
		retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
		iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text

		If retCode = "200" Then		'성공(200)
			strSql = ""
			strSql = strSql & " UPDATE db_outmall.dbo.tbl_ezwel_regItem " & VbCRLF
			strSql = strSql & " SET accFailCnt = 0 " & VbCRLF
			strSql = strSql & " ,ezwelLastUpdate = getdate() " & VbCRLF
			strSql = strSql & " ,ezwelprice = '"&imustprice&"' " & VbCRLF
			strSql = strSql & " ,ezwelsellyn = '"&isellyn&"' " & VbCRLF
			If oEzwel.FItemList(i).isImageChanged Then
				strSql = strSql & " ,regImageName = '"&oEzwel.FItemList(i).getBasicImage&"' " & VbCRLF	
			End If
			strSql = strSql & " WHERE itemid='" & iitemid & "'"
			dbCTget.Execute(strSql)
			EzwelOneItemEdit = true
			If optMust = "all" Then
				strSql = ""
				strSql = strSql &  "SELECT count(*) as cnt "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & oEzwel.FItemList(i).Fitemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
				rsget.Open strSql,dbget,1
					ocount = rsget("cnt")
				rsget.Close
	
				strSql = ""
				strSql = strSql & " UPDATE db_outmall.dbo.tbl_ezwel_regItem SET "
				strSql = strSql & " regedOptCnt = " & ocount
				strSql = strSql & " WHERE itemid = " & oEzwel.FItemList(i).Fitemid
				dbCTget.Execute strSql
			End If

			Set objXML = Nothing
			Set xmlDOM = Nothing
			If isellyn = "N" Then
				oMsg = " | 품절처리"
			End If
			rw "[" & iitemid & "]:" & iMessage & oMsg
		Else						'실패(E)
			iErrStr =  "상품 수정 중 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
		On Error Goto 0
	End If
End Function
'#####################################################################################################
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),25)
Dim arrItemid : arrItemid = request("cksel")
Dim alertMsg, iMessage, actCnt, sqlStr, retErrStr
Dim oEzwel, i, strParam, iErrStr, ret1, chgSellYn, iitemid
Dim getMustprice, sellgubun, chkparam
chgSellYn = request("chgSellYn")
actCnt = 0

If (cmdparam = "RegSelect") Then				'선택상품 실제 등록
    arrItemid = Trim(arrItemid)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If

	'## 선택상품 목록 접수
	Set oEzwel = new CEzwel
		oEzwel.FPageSize	= 20
		oEzwel.FRectItemID	= arrItemid
		oEzwel.getEzwelNotRegItemList

	    If (oEzwel.FResultCount < 1) Then
	        arrItemid = split(arrItemid,",")
	        For i = LBound(arrItemid) to UBound(arrItemid)
	            CALL Fn_AcctFailTouch("ezwel",arrItemid(i),"등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...")
	        Next

	        If (IsAutoScript) Then
	            rw "S_ERR|등록가능상품 없음 :등록조건 확인: 판매Y, 할인..."
	            dbCTget.Close: Response.End
	        Else
	            Response.Write "<script language=javascript>alert('등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...');</script>"
				dbCTget.Close: Response.End
			End If
		End If

		For i = 0 to (oEzwel.FResultCount - 1)
			If (oEzwel.FItemList(i).FDepthCode = "") OR (oEzwel.FItemList(i).FdepthCode = "0") Then
				Response.Write "<script language=javascript>alert('카테고리 매칭을 하지 않은 상품번호: [" & oEzwel.FItemList(i).Fitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
			sqlStr = ""
			sqlStr = sqlStr & " IF NOT Exists(SELECT * FROM db_outmall.dbo.tbl_ezwel_regItem where itemid="&oEzwel.FItemList(i).Fitemid&")"
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_ezwel_regItem "
	        sqlStr = sqlStr & " (itemid, regdate, reguserid, ezwelstatCD, regitemname)"
	        sqlStr = sqlStr & " VALUES ("&oEzwel.FItemList(i).Fitemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oEzwel.FItemList(i).FItemName)&"')"
			sqlStr = sqlStr & " END "
			dbCTget.Execute sqlStr
			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oEzwel.FItemList(i).checkTenItemOptionValid Then
			    On Error Resume Next
				'//상품등록 파라메터
				strParam = oEzwel.FItemList(i).getEzwelItemRegXML("Reg")

				If Err <> 0 Then
				    rw Err.Description
					Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oEzwel.FItemList(i).Fitemid & "]');</script>"
					dbCTget.Close: Response.End
				End If

				On Error Goto 0
				iErrStr = ""
				ret1 = EzwelOneItemReg(oEzwel.FItemList(i).FItemid, strParam, iErrStr, oEzwel.FItemList(i).FSellCash, oEzwel.FItemList(i).getEzwelSellYn, oEzwel.FItemList(i).FLimityn, oEzwel.FItemList(i).FLimitNo, oEzwel.FItemList(i).FLimitSold, html2db(oEzwel.FItemList(i).FItemName), oEzwel.FItemList(i).FbasicimageNm)
				If (ret1) Then
					actCnt = actCnt+1
				Else
					CALL Fn_AcctFailTouch("ezwel", oEzwel.FItemList(i).Fitemid, iErrStr)
					retErrStr = retErrStr & iErrStr
					rw iErrStr
				End If
			Else
				CALL Fn_AcctFailTouch("ezwel", oEzwel.FItemList(i).Fitemid, iErrStr)
				iErrStr = "["&oEzwel.FItemList(i).Fitemid&"] : 옵션검사 실패 | 한정옵션의 개수가 5개 이하일 수 있음"
				retErrStr = retErrStr & iErrStr
			End If
		Next
	Set oEzwel = Nothing
    If (retErrStr <> "") Then
        Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
    End If
ElseIf (cmdparam = "EditSelect") Then				'정보 수정
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If
	'## 수정할 상품 목록 접수
	Set oEzwel = new CEzwel
		oEzwel.FPageSize	= 20
		oEzwel.FRectItemID	= arrItemid
		oEzwel.getEzwelEditedItemList
		For i = 0 to (oEzwel.FResultCount - 1)
			On Error Resume Next
			strParam = ""
			chkparam = ""
			iErrStr = ""
			If (oEzwel.FItemList(i).FmaySoldOut = "Y") OR (oEzwel.FItemList(i).IsSoldOutLimit5Sell) Then
				strParam = oEzwel.FItemList(i).getEzwelItemRegXML("SellN")
				chgSellYn = "N"
			Else
				strParam = oEzwel.FItemList(i).getEzwelItemRegXML("SellY")
				chgSellYn = "Y"
			End If
			getMustprice = ""
			getMustprice = oEzwel.FItemList(i).fngetMustPrice()
			'*********************************************************************************************************************************************************
			'2014-11-06 김진영 | dev_Comment
			'API가 전송되는 족족 상품옵션을 인식하지 않음 | 등록된 옵션카운트가 크다면 10x10에서 옵션 삭제한 것이 살아있음
			'결국 이지웰의 옵션사용안함으로 돌리면 옵션이 초기화 됨을 발견
			'옵션 없음으로 API전송 후 한번 더 옵션있는 정상상태로 전송하면 원하는 데이터가 확인 됨.
			'추가 : 두번 API전송시 많은 확률로 에러가 뜸 | 아마 이지웰페어 DB쪽 상품가격 수정하는 데 뭔가 걸려있는 듯 함..
			'		따라서 우선 이런 상품은 품절로 시켜두었다가 에러나는 항목들만 수기로 수정하던 하는 조치가 필요해 보임.
			'추가_김진영(2015-03-04) 1173474 4:2 이런형식이라서 If CInt(rsCTget("optioncnt")) <> CInt(rsCTget("regedoptcnt")) Then 이런 조건을 변경
			sqlStr = ""
			sqlStr = sqlStr &  "SELECT top 1 r.itemid, i.optioncnt, r.regedoptcnt "
			sqlStr = sqlStr & " FROM db_Appwish.dbo.tbl_item as i "
			sqlStr = sqlStr & " join db_outmall.dbo.tbl_ezwel_regitem as r on i.itemid=r.itemid "
			sqlStr = sqlStr & " WHERE i.itemid=" & oEzwel.FItemList(i).Fitemid
			rsCTget.Open sqlStr,dbCTget,1
			If not rsCTget.EOF Then
				If CInt(rsCTget("optioncnt")) > 0 Then
					If CInt(rsCTget("optioncnt")) <> CInt(rsCTget("regedoptcnt")) Then
						chkparam = oEzwel.FItemList(i).getEzwelItemOptZeroNotScheduleXML("SellN")
						Call EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, chkparam, getMustprice, "N", "optMustN")
					End If
				End If
			End If
			rsCTget.Close
			'*********************************************************************************************************************************************************
			iErrStr = ""
			ret1 = EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, "all")
			If (ret1) Then
				actCnt = actCnt+1
			Else
				CALL Fn_AcctFailTouch("ezwel", oEzwel.FItemList(i).Fitemid, iErrStr)
				retErrStr = retErrStr & iErrStr
				rw iErrStr
			End If

			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oEzwel.FItemList(i).Fitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
			On Error Goto 0
			retErrStr = retErrStr & iErrStr
		Next
	Set oEzwel = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam = "EditSelectNotSchedule") Then				'스케줄링 사용N(마케팅 수기 버튼) | 이미지, 상품설명은 안 보냄
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If
	'## 수정할 상품 목록 접수
	Set oEzwel = new CEzwel
		oEzwel.FPageSize	= 20
		oEzwel.FRectItemID	= arrItemid
		oEzwel.getEzwelEditedItemList
		For i = 0 to (oEzwel.FResultCount - 1)
			On Error Resume Next
			strParam = ""
			chkparam = ""
			iErrStr = ""
			If (oEzwel.FItemList(i).FmaySoldOut = "Y") OR (oEzwel.FItemList(i).IsSoldOutLimit5Sell) Then
				strParam = oEzwel.FItemList(i).getEzwelItemEditNotScheduleXML("SellN")
				chgSellYn = "N"
			Else
				strParam = oEzwel.FItemList(i).getEzwelItemEditNotScheduleXML("SellY")
				chgSellYn = "Y"
			End If

			getMustprice = ""
			getMustprice = oEzwel.FItemList(i).fngetMustPrice()
			'*********************************************************************************************************************************************************
			'2014-11-06 김진영 | dev_Comment
			'API가 전송되는 족족 상품옵션을 인식하지 않음 | 등록된 옵션카운트가 크다면 10x10에서 옵션 삭제한 것이 살아있음
			'결국 이지웰의 옵션사용안함으로 돌리면 옵션이 초기화 됨을 발견
			'옵션 없음으로 API전송 후 한번 더 옵션있는 정상상태로 전송하면 원하는 데이터가 확인 됨.
			'추가 : 두번 API전송시 많은 확률로 에러가 뜸 | 아마 이지웰페어 DB쪽 상품가격 수정하는 데 뭔가 걸려있는 듯 함..
			'		따라서 우선 이런 상품은 품절로 시켜두었다가 에러나는 항목들만 수기로 수정하던 하는 조치가 필요해 보임.
			'추가_김진영(2015-03-04) 1173474 4:2 이런형식이라서 If CInt(rsCTget("optioncnt")) <> CInt(rsCTget("regedoptcnt")) Then 이런 조건을 변경
			sqlStr = ""
			sqlStr = sqlStr &  "SELECT top 1 r.itemid, i.optioncnt, r.regedoptcnt "
			sqlStr = sqlStr & " FROM db_Appwish.dbo.tbl_item as i "
			sqlStr = sqlStr & " join db_outmall.dbo.tbl_ezwel_regitem as r on i.itemid=r.itemid "
			sqlStr = sqlStr & " WHERE i.itemid=" & oEzwel.FItemList(i).Fitemid
			rsCTget.Open sqlStr,dbCTget,1
			If not rsCTget.EOF Then
				If CInt(rsCTget("optioncnt")) > 0 Then
					If CInt(rsCTget("optioncnt")) <> CInt(rsCTget("regedoptcnt")) Then
						chkparam = oEzwel.FItemList(i).getEzwelItemOptZeroNotScheduleXML("SellN")
						Call EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, chkparam, getMustprice, "N", "optMustN")
					End If
				End If
			End If
			rsCTget.Close
			'*********************************************************************************************************************************************************
			iErrStr = ""
			ret1 = EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, "all")
			If (ret1) Then
				actCnt = actCnt+1
			Else
				CALL Fn_AcctFailTouch("ezwel", oEzwel.FItemList(i).Fitemid, iErrStr)
				retErrStr = retErrStr & iErrStr
				rw iErrStr
			End If

			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oEzwel.FItemList(i).Fitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
			On Error Goto 0
			retErrStr = retErrStr & iErrStr
		Next
	Set oEzwel = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam = "EditSellYn") Then				'판매상태 수정
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If
	'## 수정할 상품 목록 접수
	Set oEzwel = new CEzwel
		If (session("ssBctID")="kjy8517") Then
			oEzwel.FPageSize	= 100
		Else
			oEzwel.FPageSize	= 20
		End If
		oEzwel.FRectItemID	= arrItemid
		oEzwel.getEzwelEditedItemList

'		If (chgSellYn="N") and (oEzwel.FResultCount < 1) and (arrItemid = "") Then
'		    oEzwel.getEzwelreqExpireItemList
'		End If

		If chgSellYn = "N" Then
			sellgubun = "SellN"
		Else
			sellgubun = "SellY"
		End If

		For i = 0 to (oEzwel.FResultCount - 1)
			strParam = oEzwel.FItemList(i).getEzwelItemRegXML(sellgubun)
			getMustprice = ""
			getMustprice = oEzwel.FItemList(i).fngetMustPrice()
		    iErrStr = ""
			If EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, "all") Then
				actCnt = actCnt+1
			Else
				rw "["&iitemid&"]"&iErrStr
			End If
			retErrStr = retErrStr & iErrStr
		Next
	Set oEzwel = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
Else
	rw "미지정 ["&cmdparam&"]"
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
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->