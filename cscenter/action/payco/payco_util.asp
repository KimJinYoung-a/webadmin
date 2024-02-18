<%
	'-----------------------------------------------------------------------------
	' payco_util.asp version 1.0
	' 2015-03-25	PAYCO������� <dl_payco_ts@nhnent.com>
	'-----------------------------------------------------------------------------

	'---------------------------------------------------------------------------------
	' �ֹ� ���� API ȣ�� �Լ�
	' ��� ��� : Call payco_reserve(mData)
	' mData - JSON ������
	'---------------------------------------------------------------------------------
	Function payco_reserve(mData)
		Dim Result, resultValue, tmpJSON
		Set Result = payco_Call_API(payco_URL_reserve, "json", mData)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "�ֹ� ���� API ȣ�� ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		payco_reserve = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' ���� ���� API ȣ�� �Լ�
	' ��� ��� : Call payco_cancelmileage(mData)
	' mData - JSON ������
	'---------------------------------------------------------------------------------
	Function payco_approval(mData)
		Dim Result, resultValue, tmpJSON
		Set Result = payco_Call_API(payco_URL_approval, "json", mData)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "���� ���� API ȣ�� ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'��� ����
		payco_approval = resultValue
	End Function

	'-----------------------------------------------------------------------------
	' PAYCO �ֹ� ��� ���� ���� API ȣ�� �Լ�
	' ��� ��� : Call payco_cancel_check(mData)
	' mData - JSON ������
	'-----------------------------------------------------------------------------
	Function payco_cancel_check(mData)
		Dim Result, resultValue, tmpJSON
		Set Result = payco_Call_API(payco_URL_cancel_check, "json", mData)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "�ֹ� ���� ��� ���� ���� ��ȸ ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'��� ����
		payco_cancel_check = resultValue
	End Function

	'-----------------------------------------------------------------------------
	' PAYCO �ֹ� ��� API ȣ�� �Լ�
	' ��� ��� : Call payco_cancel(mData)
	' mData - JSON ������
	'-----------------------------------------------------------------------------
	Function payco_cancel(mData)
		Dim Result, resultValue, tmpJSON
		Set Result = payco_Call_API(payco_URL_cancel, "json", mData)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "�ֹ� ���� ��� ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'��� ����
		payco_cancel = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' �ֹ� ���� ���� API ȣ�� �Լ�
	' ��� ��� : Call payco_upstatus(mData)
	' mData - JSON ������
	'---------------------------------------------------------------------------------
	Function payco_upstatus(mData)
		Dim Result, resultValue, tmpJSON
		Set Result = payco_Call_API(payco_URL_upstatus, "json", mData)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "�ֹ� ���� ���� ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'��� ����
		payco_upstatus = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' ���ϸ��� ���� ��� API ȣ�� �Լ�
	' ��� ��� : Call payco_cancelmileage(mData)
	' mData - JSON ������
	'---------------------------------------------------------------------------------
	Function payco_cancelmileage(mData)
		Dim Result, resultValue, tmpJSON
		Set Result = payco_Call_API(payco_URL_cancelMileage, "json", mData)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "���ϸ��� ���� ��� ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'��� ����
		payco_cancelmileage = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' �������� ����Ű ��ȿ�� üũ API ȣ�� �Լ�
	' ��� ��� : Call payco_keycheck(mData)
	' mData - JSON ������
	'---------------------------------------------------------------------------------
	Function payco_keycheck(mData)
		Dim Result, resultValue, tmpJSON
		Set Result = payco_Call_API(payco_URL_checkUsability, "json", mData)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "�������� ����Ű ��ȿ�� üũ ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		payco_keycheck = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' ���� �� ��ȸ API ȣ�� �Լ�
	' ��� ��� : Call payco_verifypayment(mData)
	' mData - JSON ������
	'---------------------------------------------------------------------------------
	Function payco_verifypayment(mData)
		Dim Result, resultValue, tmpJSON
		Set Result = payco_Call_API(payco_URL_verifyPayment, "json", mData)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "���� �� ���� ȣ�� ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		payco_verifypayment = resultValue
	End Function


	'-----------------------------------------------------------------------------
	' payco_URLDecode
	' ��� ��� : Call payco_URLDecode(Encoding URL)
	' Encoding URL : ���ڵ��� URL
	'-----------------------------------------------------------------------------
	Function payco_URLDecode(sStr)
		Dim sTemp, sChar, nLen
		Dim nPos, sResult, sHex

		'On Error Resume Next

		nLen = Len(sStr)

		sTemp = Replace(sStr, "+", " ")
		For nPos = 1 To nLen
			sChar = Mid(sTemp, nPos, 1)
			If sChar = "%" Then
				If nPos + 2 <= nLen Then
					sHex = Mid(sTemp, nPos+1, 2)
					If IsHexaString(sHex) Then
						sResult = sResult & Chr(CLng("&H" & sHex))
						nPos = nPos + 2
					Else
						sResult = sResult & sChar
					End If
				Else
					sResult = sResult & sChar
				End If
			Else
				sResult = sResult & sChar
			End If
		Next

		If Err Then
			Call payco_Write_Log("payco_URLDecode(" & sStr & "). " & Err.description)
		End If

		'On Error GoTo 0

		payco_URLDecode = sResult
	End Function

	'-----------------------------------------------------------------------------
	' �α� ��� �Լ� ( ����׿� )
	' ��� ��� : Call payco_Write_Log(Log_String)
	' Log_String : �α� ���Ͽ� ����� ����
	'-----------------------------------------------------------------------------
	Const fsoForReading = 1		'- Open a file for reading. You cannot write to this file.
	Const fsoForWriting = 2		'- Open a file for writing.
	Const fsoForAppend = 8		'- Open a file and write to the end of the file.
	Sub payco_Write_Log(Log_String)
		If Not Payco_LogUse Then Exit Sub
		'On Error Resume Next
		Dim oFSO
		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		Dim oTextStream
		Set oTextStream = oFSO.OpenTextFile(payco_Write_LogFile, fsoForAppend, True, 0)
		'-----------------------------------------------------------------------------
		' ���� ���
		'-----------------------------------------------------------------------------
		oTextStream.WriteLine  CStr(FormatDateTime(Now,0)) + " " + Replace(CStr(Log_String),Chr(0),"'")
		'-----------------------------------------------------------------------------
		' ���ҽ� ����
		'-----------------------------------------------------------------------------
		oTextStream.Close
		Set oTextStream = Nothing
		Set oFSO = Nothing
	End Sub

	'-----------------------------------------------------------------------------
	' API ȣ�� �Լ�( POST ���� - PAYCO ������ ��� API ȣ�⿡ POST���� ����մϴ�. )
	' ��� ��� : payco_Call_API(SiteURL, App_Mode, Param)
	' SiteURL : ȣ���� API �ּ�
	' App_Mode : ������ ���� ���� ( ��: json, x-www-form-urlencoded �� )
	' Param : ������ POST ������
	'-----------------------------------------------------------------------------
	Function payco_Call_API(SiteURL, App_Mode, Param)
		Dim HTTP_Object

		'-----------------------------------------------------------------------------
		' WinHttpRequest ����
		'-----------------------------------------------------------------------------
		Set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		With HTTP_Object
			'API ��� Timeout �� 30�ʷ� ����
			.SetTimeouts 30000, 30000, 30000, 30000
			.Open "POST", SiteURL, False
			.SetRequestHeader "Content-Type", "application/"+CStr(App_Mode)+"; charset=UTF-8"
			'-----------------------------------------------------------------------------
			' API ���� ������ �α� ���Ͽ� ����
			'-----------------------------------------------------------------------------
			Call payco_Write_Log("Call API   "+CStr(SiteURL)+" Mode : "  + CStr(App_Mode))
			Call payco_Write_Log("Call API   "+CStr(SiteURL)+" Data : "  + CStr(Param))
			.Send Param
			.WaitForResponse
			'-----------------------------------------------------------------------------
			' ���� ����� �����ϱ� ���� ���� ���� �� �� ����
			'-----------------------------------------------------------------------------
			Dim Result
			Set Result = New clsHTTP_Object
			Result.Status = CStr(.Status)
			Result.ResponseText = CStr(.ResponseText)
			'-----------------------------------------------------------------------------
			' API ���� ����� �α� ���Ͽ� ����
			'-----------------------------------------------------------------------------
			Call payco_Write_Log("API Result "+CStr(SiteURL) + " Status : " + CStr(.Status))
			Call payco_Write_Log("API Result "+CStr(SiteURL) + " ResponseText : " + CStr(.ResponseText))
		End With
		Set payco_Call_API = Result
	End Function

	'-----------------------------------------------------------------------------
	' API ��� ���ۿ� ������ ���� ����
	' Status �� ResponseText ���� �����Ѵ�.
	'-----------------------------------------------------------------------------
	Class clsHTTP_Object
		private m_Status
		private m_ResponseText

		public property get Status()
			Status = m_Status
		end property

		public property get ResponseText()
			ResponseText = m_ResponseText
		end property

		public property let Status(p_Status)
			m_Status = p_Status
		end property

		public property let ResponseText(p_ResponseText)
			m_ResponseText = p_ResponseText
		end property

		Private Sub Class_Initialize
			m_Status = ""
			m_ResponseText = ""
		End Sub
	End Class

%>
