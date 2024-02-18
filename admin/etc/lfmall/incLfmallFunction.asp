<%
'############################################## ���� �����ϴ� API �Լ� ���� ���� ############################################
'��ǰ ���
Public Function fnLfmallItemReg(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "?itemid="&iitemid&"&scmid="&session("ssBctID")
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://110.93.128.100:8090/lfmall/product/productreg" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ���] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				iMessage		= strObj.message
				If isSuccess = true Then
					iErrStr = "OK||"&iitemid&"||����[��ǰ���]"
				Else
					iErrStr = "ERR||"&iitemid&"||����[��ǰ���] " & iMessage
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||����[��ǰ���] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||����[��ǰ���_NO] ��ſ���"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'��ǰ ����
Public Function fnLfmallItemEdit(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "?itemid="&iitemid&"&scmid="&session("ssBctID")
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://110.93.128.100:8090/lfmall/product/update" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ����] " & html2db(Err.Description)
			Exit Function
		End If
		'rw objXML.Status
		'rw BinaryToText(objXML.ResponseBody,"utf-8")
		'response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				iMessage		= strObj.message
				If isSuccess = true Then
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||����[��ǰ����]"
				Else
					iErrStr = "ERR||"&iitemid&"||����[��ǰ����] " & iMessage
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||����[��ǰ����] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||����[��ǰ����] "& html2db(replace(iRbody, """", ""))
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'��ǰ ���� ����
Public Function fnLfmallSellYN(iitemid, ichgSellYn, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, itemstat, isSuccess

	Select Case ichgSellYn
		Case "Y"	itemstat = "7"		'��ǰ�ǸŽ��� MD ���� �ʿ�
		Case "X"	itemstat = "5"		'��ǰ����(��������)
		Case Else	itemstat = "6"		'�Ͻ�ǰ��
	End Select

	istrParam = "?itemid="&iitemid&"&stat="&itemstat&"&scmid="&session("ssBctID")
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://110.93.128.100:8090/lfmall/product/manage" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[���¼���] " & html2db(Err.Description)
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				iMessage		= strObj.message
				If isSuccess = true Then
					If ichgSellyn = "Y" Then
						iErrStr =  "OK||"&iitemid&"||�ǸŽ��δ��[���¼���]"
					ElseIf ichgSellyn = "N" Then
						iErrStr =  "OK||"&iitemid&"||ǰ��ó��[���¼���]"
					End If
				Else
					If Instr(iMessage, "��ǰ �ڵ��� �Ǹ� ���� �� ���� ���� �ڵ� Ȯ�� �ٶ��ϴ�") > 0 Then
						If ichgSellyn = "Y" Then
							iErrStr =  "OK||"&iitemid&"||�ǸŽ��δ��[���¼���]"
						ElseIf ichgSellyn = "N" Then
							iErrStr =  "OK||"&iitemid&"||ǰ��ó��[���¼���]"
						End If
					Else
						iErrStr = "ERR||"&iitemid&"||����[���¼���] "& html2db(iMessage)
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||����[���¼���] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||����[���¼���] ��ſ���"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'��ǰ ��ȸ
Public Function fnLfmallStatChk(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, isSuccess
	Dim slitmAprvlGbcd, productStatusName, itemSellGbcd, ihmallSellYn, ihmallPrice, ostkYn, itemAthzGbcd, itemAthzGbcdNm
	istrParam = "?itemid="&iitemid&"&scmid="&session("ssBctID")

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://110.93.128.100:8090/lfmall/product/info" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ȸ] " & html2db(Err.Description)
			Exit Function
		End If
'		rw objXML.Status
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				iMessage		= strObj.message
				If isSuccess = true Then
					productStatusName = strObj.outValue.body.product.productStatusName
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lfmall_regitem " & VbCRLF
					strSql = strSql & " SET lastConfirmdate = getdate() " & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||����[��ȸ("&html2db(iMessage)&")]"
				Else
					iErrStr = "ERR||"&iitemid&"||����[��ȸ] "& html2db(iMessage)
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||����[��ȸ] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||����[��ȸ] ��ſ���"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function
'############################################## ���� �����ϴ� API �Լ� ���� �� ############################################
%>