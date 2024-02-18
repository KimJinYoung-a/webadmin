<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/admin/etc/order/lib/xSiteOrderLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim act     : act = requestCheckVar(request("act"),32)
Dim param1  : param1 = requestCheckVar(request("param1"),32)
Dim param2  : param2 = requestCheckVar(request("param2"),32)
Dim param3  : param3 = requestCheckVar(request("param3"),32)
Dim param4  : param4 = requestCheckVar(request("param4"),32)
Dim param5  : param5 = requestCheckVar(request("param5"),32)
Dim sqlStr, i, paramData, retVal, ireturnJsonList, strObj, datalist
Dim cnt
Dim OutMallOrderSerialArr, TenOrderserial
Dim OrgDetailKeyArr, songjangDivArr
Dim songjangNoArr, sendReqCntArr, beasongdateArr, outmallGoodsIDArr

Select Case act
	Case "outmallSongJangIp"
		If (LCASE(param1)="ezwel") or (LCASE(param1)="kakaostore") or (LCASE(param1)="boribori1010") or (LCASE(param1)="wconcept1010") or (LCASE(param1)="benepia1010") Then
			Call GetInvoiceList(param1, ireturnJsonList)
			Set strObj = JSON.parse(ireturnJsonList)
				Set datalist = strObj.result
					If datalist.length > 0 Then
						rw "CNT="&datalist.length
						For i=0 to datalist.length - 1
							paramData = "redSsnKey=system&ord_no="&datalist.get(i).outMallOrderSerial&"&ord_dtl_sn="&datalist.get(i).originDetailKey&"&hdc_cd="&datalist.get(i).invoiceDivision&"&inv_no="& Trim(getNumeric(datalist.get(i).invoiceNumber))
							If (LCASE(param1)="ezwel") Then
								retVal = SendReq("https://webadmin.10x10.co.kr/admin/etc/ezwel/Ezwel_SongjangProc.asp",paramData)
							ElseIf (LCASE(param1)="kakaostore") Then
								retVal = SendReq("https://webadmin.10x10.co.kr/admin/etc/kakaostore/kakaostore_SongjangProc.asp",paramData)
							ElseIf (LCASE(param1)="boribori1010") Then
								retVal = SendReq("https://webadmin.10x10.co.kr/admin/etc/boribori/boribori_SongjangProc.asp",paramData)
							ElseIf (LCASE(param1)="wconcept1010") Then
								retVal = SendReq("https://webadmin.10x10.co.kr/admin/etc/wconcept/wconcept_SongjangProc.asp",paramData)
							ElseIf (LCASE(param1)="benepia1010") Then
								retVal = SendReq("https://webadmin.10x10.co.kr/admin/etc/benepia/benepia_SongjangProc.asp",paramData)
							End If
							rw retVal
						Next
					Else
						response.Write "S_NONE.."
						dbget.Close() : response.end
					End If
				Set datalist = nothing
			Set strObj = nothing
		End If
	' Case "outmallSongJangIp"
	' 	sqlStr = "select top 30 T.orderserial, T.OutMallOrderSerial"
	' 	sqlStr = sqlStr & " ,T.OrgDetailKey, IsNULL(T.sendState,0) as sendState"
	' 	sqlStr = sqlStr & " ,D.songjangDiv, D.songjangNo, D.itemNo, D.beasongdate, T.outMallGoodsNo"
	' 	sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
	' 	sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master M"
	' 	sqlStr = sqlStr & " 	on T.orderserial=M.orderserial"
	' 	sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D"
	' 	sqlStr = sqlStr & " 	on T.orderserial=D.orderserial"
	' 	sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid"					'// 기존 주문에 합쳐진 경우(빨강1개,파랑1개 -> 파랑2개)
	' 	sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption"
	' 	sqlStr = sqlStr & " 	and D.currstate=7"
	' 	sqlStr = sqlStr & " 	left join db_order.dbo.tbl_songjang_div V"
	' 	sqlStr = sqlStr & " 	on D.songjangDiv=V.divcd"
    '     sqlStr = sqlStr & " WHERE 1=1"
	' 	sqlStr = sqlStr & " and T.regdate > dateadd(month, -2, getdate()) "    ''7개월 -> 2개월로 변경..2021-11-18 김진영
	' 	sqlStr = sqlStr & " and T.sellsite='"&param1&"'"
	' 	sqlStr = sqlStr & " and T.OrgDetailKey is Not NULL"             ''디테일키 입력 주문건만..
	' 	sqlStr = sqlStr & " and IsNULL(T.sendState,0)=0"
	' 	sqlStr = sqlStr & " and T.sendReqCnt<3"                         ''여러번 시도 안되도록. 추가.
	' 	sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"      ''교환 취소 반품 제외.
	' 	sqlStr = sqlStr & " order by D.beasongdate desc"
	' 	rsget.CursorLocation = adUseClient
	' 	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	' 	cnt = rsget.RecordCount
	' 	ReDim TenOrderserial(cnt)
	' 	ReDim OutMallOrderSerialArr(cnt)
	' 	ReDim OrgDetailKeyArr(cnt)
	' 	ReDim songjangDivArr(cnt)
	' 	ReDim songjangNoArr(cnt)
	' 	Redim sendReqCntArr(cnt)
	' 	Redim beasongdateArr(cnt)
	' 	Redim outmallGoodsIDArr(cnt)
	' 	i = 0
	' 	If Not rsget.Eof Then
	' 		Do Until rsget.eof
	' 		TenOrderserial(i) = rsget("orderserial")
	' 		OutMallOrderSerialArr(i) = rsget("OutMallOrderSerial")
	' 		OrgDetailKeyArr(i) = rsget("OrgDetailKey")
	' 		songjangDivArr(i) = rsget("songjangDiv")
	' 		songjangNoArr(i) = rsget("songjangNo")
	' 		sendReqCntArr(i) = rsget("itemNo")
	' 		beasongdateArr(i) = rsget("beasongdate")
	' 		outmallGoodsIDArr(i) = rsget("outMallGoodsNo")
	' 		i=i+1
	' 		rsget.MoveNext
	' 		Loop
	' 	End If
	' 	rsget.close

	' 	If (cnt < 1) Then
	' 		response.Write "S_NONE.."
	' 		dbget.Close() : response.end
	' 	Else
	' 		rw "CNT="&CNT

	' 		For i=LBound(OutMallOrderSerialArr) to UBound(OutMallOrderSerialArr)
	' 			If (OutMallOrderSerialArr(i)<>"") Then
	' 			    If (LCASE(param1)="ezwel") Then
	' 			        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&songjangDivArr(i)&"&inv_no="&songjangNoArr(i)
	' 					retVal = SendReq("https://webadmin.10x10.co.kr/admin/etc/ezwel/Ezwel_SongjangProc.asp",paramData)
	' 					rw retVal
	' 			    End If
	' 			End If
	' 		Next
    '     End If
    Case Else
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->