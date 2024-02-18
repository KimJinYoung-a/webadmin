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
<!-- #include virtual="/lib/BarcodeFunction.asp"-->

<%
	Dim vItemCode, i, j, vAction, vQuery, vMessage
	vItemCode = request("itemcode")
	If vItemCode = "" Then
		vItemCode = request("itembarcode")
	End If
	vAction = request("action")

	Dim siteSeq, itemgubun, itemid, itemoption

	if BF_IsMaybeTenBarcode(vItemCode) then
		siteSeq 	= BF_GetItemGubun(vItemCode)
		itemgubun 	= BF_GetItemGubun(vItemCode)
		itemid 		= BF_GetItemId(vItemCode)
		''itemoption 	= BF_GetItemOption(vItemCode)
	End If


	If vAction = "delete" Then
		vQuery = "EXECUTE [db_item].[dbo].[sp_Ten_UpcheManageCode_Reg] '" & itemgubun & "', '" & BF_GetFormattedItemId(itemid) & "', '" & Right(vItemCode,4) & "', '' "
		dbget.execute vQuery
		dbget.close()
		Response.Write "<script type='text/javascript'>parent.document.location.href='popUpcheManageCode.asp?itemcode="&vItemCode&"&isok=o';</script>"
		Response.End
	End If


	Dim vOptionCount, vTemp, vArrOpt, vArrUMCode, vItemOption, vItemBarCode, vExitOption
	vOptionCount = Request("optioncount")
	vArrOpt		 = Replace(Request("itemoption")," ","")
	vArrUMCode	 = Replace(Request("upchemanagecode")," ","")


	For i = 0 To (vOptionCount-1)
		vTemp = Trim(Split(vArrUMCode,",")(i))
		If vTemp <> "" Then
			For j = 0 To (vOptionCount-1)
				If i <> j AND vTemp = Trim(Split(vArrUMCode,",")(j)) Then
					vExitOption = Trim(Split(vArrOpt,",")(j))
					Exit For
				End If
			Next
		End IF
		If vExitOption <> "" Then
			Exit For
		End If
	Next

	'// 업체코드는 업체 마음대로 등록한다.(중복입력 허용), skyer9, 2017-04-04
	If vExitOption <> "" And False Then
		Response.Write "<script type='text/javascript'>parent.jsMessageReset();parent.document.getElementById('publicbarspan"&vExitOption&"').innerHTML = '<font color=red>* ["&vExitOption&"] 중복 입력된 업체코드입니다. 중복 업체코드 등록 불가.</font>';parent.FocusAndSelect(parent.document.frmbar, parent.document.frmbar.upchemanagecode["&j&"]);</script>"
	Else
		For i = 0 To (vOptionCount-1)
			vTemp = Trim(Split(vArrUMCode,",")(i))
			If vTemp = "000000000100" Then
				vTemp = ""
			End IF

			vQuery = vQuery & " EXECUTE [db_item].[dbo].[sp_Ten_UpcheManageCode_Reg] '" & itemgubun & "', '" & BF_GetFormattedItemId(itemid) & "', '" & Trim(Split(vArrOpt,",")(i)) & "', '" & vTemp & "' "
		Next
		dbget.execute vQuery
		Response.Write "<script type='text/javascript'>parent.document.location.href='popUpcheManageCode.asp?itemcode="&vItemCode&"&isok=o';</script>"
	End If
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
