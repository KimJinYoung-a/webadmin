<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/benepia/benepiaCls.asp"-->
<!-- #include virtual="/admin/etc/benepia/incbenepiaFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, failCnt, chgSellYn, goodsGrpCd, cateCode, mallInfoDiv
Dim resultMessage, strSql, AssignedRow, ccd, selectCateCode, arrRows, i
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
goodsGrpCd		= request("goodsGrpCd")
ccd				= request("ccd")
failCnt			= 0


If action = "REG" Then												'상품등록
	Call fnbenepiaItemReg(itemid, action, resultMessage)
ElseIf action = "REGITEM" Then										'상품등록
	Call fnbenepiaOnlyItemReg(itemid, action, resultMessage)
ElseIf action = "IMAGE" Then										'상품이미지등록
	Call fnbenepiaImageReg(itemid, action, resultMessage)
ElseIf action = "EditSellYn" Then									'상태변경
	Call fnbenepiaSellYN(itemid, action, chgSellYn, resultMessage)
ElseIf action = "EDIT" Then											'상품수정
	Call fnbenepiaItemEdit(itemid, action, resultMessage)
ElseIf action = "EDITCATE" Then										'상품카테고리수정
	Call fnbenepiaCategoryEdit(itemid, action, resultMessage)
ElseIf action = "PRICE" Then										'상품 가격 수정
	Call fnbenepiaPriceEdit(itemid, action, resultMessage)
ElseIf action = "QTY" Then											'상품 재고 수정
	Call fnbenepiaQuantityEdit(itemid, action, resultMessage)
ElseIf action = "EDITINFO" Then										'상품 정보 수정
	Call fnbenepiaItemInfoEdit(itemid, action, resultMessage)
ElseIf action = "EDITDELIVERY" Then									'배송 정보 수정
	Call fnbenepiaItemDeliveryEdit(itemid, action, resultMessage)
ElseIf action = "EDITIMAGEPC" OR action = "EDITIMAGEMOB" Then		'이미지 수정
	Call fnbenepiaItemImageEdit(itemid, action, resultMessage)
ElseIf action = "CONTENT" Then										'컨텐츠 수정
	Call fnbenepiaItemContentEdit(itemid, action, resultMessage)
ElseIf action = "SAFEINFO" Then										'안전인증 수정
	Call fnbenepiaItemSafeInfoEdit(itemid, action, resultMessage)
ElseIf action = "INFOCODE" Then										'정보고시 수정
	Call fnbenepiaItemInfoCodeEdit(itemid, action, resultMessage)
ElseIf action = "OPTEDIT" Then										'옵션 수정
	Call fnbenepiaItemEditOption(itemid, action, resultMessage)
ElseIf action = "CHKSTAT" Then										'상품조회 + 옵션조회
	Call fnwbenepiaChkstat(itemid, action, resultMessage)
ElseIf action = "CHKITEM" Then										'상품조회
	Call fnwbenepiaChkItem(itemid, action, resultMessage)
ElseIf action = "CHKOPT" Then										'옵션조회
	Call fnwbenepiaChkOpt(itemid, action, resultMessage)
'goodsGrpCd는 수기로 입력한다. 1~4까지 있는 듯
ElseIf action = "benepiaCommonCode" Then
	If ccd = "category" Then
		If goodsGrpCd <> "e" Then
			'1. 최초 depth 초기화
			rw "########## START ###########"
			strSql = ""
			strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_benepia_category] WHERE depth = '"& goodsGrpCd &"' "
			dbget.execute strSql

			If goodsGrpCd = "1" Then
				Call fnbenepiaCategory(goodsGrpCd, cateCode)
			ElseIf goodsGrpCd = "2" OR goodsGrpCd = "3" OR goodsGrpCd = "4" Then
				selectCateCode = Cint(goodsGrpCd) - 1
				strSql = ""
				strSql = strSql & " SELECT CateKey "
				strSql = strSql & " FROM db_etcmall.[dbo].[tbl_benepia_category] "
				strSql = strSql & " WHERE depth = "& selectCateCode &" "
				strSql = strSql & " AND lastLevel = '0' "
				strSql = strSql & " ORDER BY regdate ASC "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					arrRows = rsget.getRows
				End If
				rsget.close

				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						cateCode = ""
						cateCode = arrRows(0, i)
						Call fnbenepiaCategory(goodsGrpCd, cateCode)
						response.flush
						response.clear
					Next
				End If
			Else
				rw "1~4까지만 입력가능합니다."
				response.end
			End If
			rw "########## End ###########"
			response.end
		Else
			strSql = ""
			strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_Ten_API_benepia_Category_Make] "
			dbget.Execute strSql
			rw "OK"
		End If
	ElseIf ccd = "infocodedtl" Then
		strSql = ""
		strSql = strSql & " SELECT mallinfoDiv "
		strSql = strSql & " FROM db_item.dbo.tbl_outmall_infoDivMap "
		strSql = strSql & " WHERE mallid = 'benepia1010' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		If isArray(arrRows) Then
			strSql = ""
			strSql = strSql & " DELETE FROM db_item.dbo.tbl_benepia_infoCode "
			dbget.execute strSql
			For i = 0 To UBound(arrRows,2)
				mallInfoDiv = ""
				mallInfoDiv = arrRows(0, i)
				Call fnbenepiaCommonCode(ccd, mallInfoDiv)
				response.flush
				response.clear
			Next
		End If

	Else
		Call fnbenepiaCommonCode(ccd, "")
		response.end
	End If
ElseIf action = "updateSendState" Then						'주문상태변경 / benepia_SongjangProc.asp에서 넘어온다.
	AssignedRow = fnbenepiaSongjangUploadByManager(CMALLNAME, request("ord_no"), request("ord_dtl_sn"), request("updateSendState"))
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

response.write  "<script>" & vbCrLf &_
				"	var str, t; " & vbCrLf &_
				"	t = parent.document.getElementById('actStr') " & vbCrLf &_
				"	str = t.innerHTML; " & vbCrLf &_
				"	str += '"&resultMessage&"<br>' " & vbCrLf &_
				"	t.innerHTML = str; " & vbCrLf &_
				"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
				"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->