<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  상품검색
' History : 2009.04.07 서동석 생성
'			2012.08.29 한용민 수정
'####################################################
%>
<% If request.cookies("commonpop")("islogics") <> "ok" Then %>
<%'<!-- #include virtual="/admin/incSessionAdmin.asp" -->%>
<% server.Execute("/admin/incSessionAdmin.asp") %>
<% End If %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itembarcode/totalitembarcodeCls.asp" -->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp" -->
<%
	Dim vItemCode, i, j, vAction, vQuery, vMessage
	Dim retVal, paramData
    Dim itemgubunArr, itemidArr, itemoptionArr
    vItemCode = requestCheckVar(request("itemcode"),32)
	If vItemCode = "" Then
		vItemCode = requestCheckVar(request("itembarcode"),32)
	End If
	vAction = requestCheckVar(request("action"),32)

	Dim siteSeq, itemgubun, itemid, itemoption

	if BF_IsMaybeTenBarcode(vItemCode) then
		siteSeq 	= BF_GetItemGubun(vItemCode)
		itemgubun 	= BF_GetItemGubun(vItemCode)
		itemid 		= BF_GetItemId(vItemCode)
		''itemoption 	= BF_GetItemOption(vItemCode)
	End If


	If vAction = "delete" Then
		vQuery = "EXECUTE [db_item].[dbo].[sp_Ten_itemBarCode_Reg] '" & itemgubun & "', '" & BF_GetFormattedItemId(itemid) & "', '" & Right(vItemCode,4) & "', '' "
		dbget.execute vQuery
		dbget.close()

        paramData = "mode=senditeminfo&ordertype=items&itemgubun=" & itemgubun & "&itemid=" & BF_GetFormattedItemId(itemid) & "&itemoption=" & Right(vItemCode,4) & ""
        retVal = SendReqGet("http://wapi.10x10.co.kr/agv/api.asp",paramData)

		Response.Write "<script type='text/javascript'>parent.document.location.href='popBarcodeManage.asp?itemcode="&vItemCode&"&isok=o';</script>"
		Response.End
	End If


	Dim vOptionCount, vTemp, vArrOpt, vArrBCode, vItemOption, vItemBarCode, vExitOption
	vOptionCount = Request("optioncount")
	vArrOpt		 = Replace(Request("itemoption")," ","")
	vArrBCode	 = Replace(Request("publicbar")," ","")

	If InStr(vArrOpt, "<") Or InStr(vArrOpt, "=") Or InStr(vArrOpt, "--") Or InStr(vArrOpt, "'") Then
		Response.Write "<script type='text/javascript'>alert('\n\n시스템팀 문의 : 허용되지 않는 특수문자입니다.\n\n');</script>"
		Response.Write "시스템팀 문의 : 허용되지 않는 특수문자입니다."
		Response.End
	End If

	If InStr(vArrBCode, "<") Or InStr(vArrBCode, "=") Or InStr(vArrBCode, "--") Or InStr(vArrBCode, "'") Then
		Response.Write "<script type='text/javascript'>alert('\n\n시스템팀 문의 : 허용되지 않는 특수문자입니다.\n\n');</script>"
		Response.Write "시스템팀 문의 : 허용되지 않는 특수문자입니다."
		Response.End
	End If

	For i = 0 To (vOptionCount-1)
		vTemp = Trim(Split(vArrBCode,",")(i))
		If vTemp <> "" Then
			For j = 0 To (vOptionCount-1)
				If i <> j AND vTemp = Trim(Split(vArrBCode,",")(j)) Then
					vExitOption = Trim(Split(vArrOpt,",")(j))
					Exit For
				End If
			Next
		End IF
		If vExitOption <> "" Then
			Exit For
		End If
	Next

	If vExitOption <> "" Then
		Response.Write "<script type='text/javascript'>parent.jsMessageReset();parent.document.getElementById('publicbarspan"&vExitOption&"').innerHTML = '<font color=red>* ["&vExitOption&"] 중복 입력된 바코드입니다. 중복 바코드 등록 불가.</font>';parent.FocusAndSelect(parent.document.frmbar, parent.document.frmbar.publicbar["&j&"]);</script>"
	Else
		'####### DB중복체크 #######

		''### tbl_item_option_stock 체크
		vQuery = "SELECT (Convert(varchar,itemgubun) + Convert(varchar,itemid) + Convert(varchar,itemoption)) as itemcode, barcode, itemgubun, itemid, itemoption FROM [db_item].[dbo].[tbl_item_option_stock] "
		vQuery = vQuery & "WHERE (Convert(varchar,itemgubun) + Convert(varchar,itemid)) <> '" & itemgubun + CStr(itemid) & "' AND barcode <> '' AND barcode IN ('" & Replace(vArrBCode,",","','") & "') "
		''rw vQuery
		rsget.Open vQuery,dbget,1
		If Not rsget.Eof Then
			Do Until rsget.Eof
				vMessage = vMessage & "[<a href=""javascript:barcodeManageRe(\'" & BF_MakeTenBarcode(rsget("itemgubun"), rsget("itemid"), rsget("itemoption")) & "\');"">바로가기</a>]" & BF_MakeTenBarcode(rsget("itemgubun"), rsget("itemid"), rsget("itemoption")) & " : " & rsget("barcode") & "<br>"
			rsget.MoveNext
			Loop

			Response.Write "<script type='text/javascript'>parent.jsMessageReset();parent.document.getElementById('notregmessage').innerHTML = '<font color=red>* 아래 상품에 이미 등록된 바코드 입니다.<br>"&vMessage&"</font>';</script>"

		End If
		rsget.close()

		If vMessage = "" Then
			vMessage = ""

			''### tbl_shop_item 체크
			vQuery = "SELECT itemgubun, shopitemid, itemoption, extbarcode FROM [db_shop].[dbo].[tbl_shop_item] "
			vQuery = vQuery & "WHERE (Convert(varchar,itemgubun) + Convert(varchar,shopitemid)) <> '" & itemgubun + CStr(itemid) & "' AND extbarcode <> '' AND extbarcode IN ('" & Replace(vArrBCode,",","','") & "') "

			rsget.Open vQuery,dbget,1
			If Not rsget.Eof Then
				Do Until rsget.Eof
					vMessage = vMessage & rsget("itemgubun") & " : " & rsget("shopitemid") & " : " & rsget("itemoption") & " - " & rsget("extbarcode") & "<br>"
				rsget.MoveNext
				Loop
				Response.Write "<script type='text/javascript'>parent.jsMessageReset();parent.document.getElementById('notregmessage').innerHTML = '<font color=red>* [db_shop]에 아래 상품이 이미 등록된 바코드 입니다.<br>"&vMessage&"</font>';</script>"
				rsget.close()
			Else
				rsget.close()
				vQuery = ""

                itemgubunArr = ""
                itemidArr = ""
                itemoptionArr = ""
				'### tbl_item_option_stock 와 tbl_shop_item UPDATE
				For i = 0 To (vOptionCount-1)
					vTemp = Trim(Split(vArrBCode,",")(i))
					If vTemp = "000000000100" Then
						vTemp = ""
					End IF

					vQuery = vQuery & " EXECUTE [db_item].[dbo].[sp_Ten_itemBarCode_Reg] '" & itemgubun & "', '" & BF_GetFormattedItemId(itemid) & "', '" & Trim(Split(vArrOpt,",")(i)) & "', '" & vTemp & "' "

                    itemgubunArr = itemgubunArr & "," & itemgubun
                    itemidArr = itemidArr & "," & BF_GetFormattedItemId(itemid)
                    itemoptionArr = itemoptionArr & "," & Trim(Split(vArrOpt,",")(i))
				Next

                if (vQuery <> "") then
				    dbget.execute vQuery

                    paramData = "mode=senditeminfo&ordertype=items&itemgubun=" & itemgubunArr & "&itemid=" & itemidArr & "&itemoption=" & itemoptionArr & ""
                    retVal = SendReqGet("http://wapi.10x10.co.kr/agv/api.asp",paramData)
                end if

                ''rw retVal
				Response.Write "<script type='text/javascript'>parent.document.location.href='popBarcodeManage.asp?itemcode="&vItemCode&"&isok=o';</script>"
			End If
		End If


	End If
'384950283945
'493827123897
'849302384756
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
