<%
	'-----------------------------------------------------------------------------
	' payco_util.asp version 1.0
	' 2015-03-25	PAYCO기술지원 <dl_payco_ts@nhnent.com>
	'-----------------------------------------------------------------------------

	'---------------------------------------------------------------------------------
	' 주문 예약 API 호출 함수
	' 사용 방법 : Call payco_reserve(mData)
	' mData - JSON 데이터
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
						.Add "result", "주문 예약 API 호출 도중 오류가 발생하였습니다."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		payco_reserve = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 결제 승인 API 호출 함수
	' 사용 방법 : Call payco_cancelmileage(mData)
	' mData - JSON 데이터
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
						.Add "result", "결제 승인 API 호출 도중 오류가 발생하였습니다."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'결과 전달
		payco_approval = resultValue
	End Function

	'-----------------------------------------------------------------------------
	' PAYCO 주문 취소 가능 여부 API 호출 함수
	' 사용 방법 : Call payco_cancel_check(mData)
	' mData - JSON 데이터
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
						.Add "result", "주문 결제 취소 가능 여부 조회 도중 오류가 발생하였습니다."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'결과 전달
		payco_cancel_check = resultValue
	End Function

	'-----------------------------------------------------------------------------
	' PAYCO 주문 취소 API 호출 함수
	' 사용 방법 : Call payco_cancel(mData)
	' mData - JSON 데이터
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
						.Add "result", "주문 결제 취소 도중 오류가 발생하였습니다."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'결과 전달
		payco_cancel = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 주문 상태 변경 API 호출 함수
	' 사용 방법 : Call payco_upstatus(mData)
	' mData - JSON 데이터
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
						.Add "result", "주문 상태 변경 도중 오류가 발생하였습니다."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'결과 전달
		payco_upstatus = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 마일리지 적립 취소 API 호출 함수
	' 사용 방법 : Call payco_cancelmileage(mData)
	' mData - JSON 데이터
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
						.Add "result", "마일리지 적립 취소 도중 오류가 발생하였습니다."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		'결과 전달
		payco_cancelmileage = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 가맹점별 연동키 유효성 체크 API 호출 함수
	' 사용 방법 : Call payco_keycheck(mData)
	' mData - JSON 데이터
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
						.Add "result", "가맹점별 연동키 유효성 체크 도중 오류가 발생하였습니다."
						.Add "message", .ResponseText
						.Add "code", .Status
					End With
					resultValue = tmpJSON.JSONoutput()
			End Select
		End With
		payco_keycheck = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 결제 상세 조회 API 호출 함수
	' 사용 방법 : Call payco_verifypayment(mData)
	' mData - JSON 데이터
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
						.Add "result", "결제 상세 내역 호출 도중 오류가 발생하였습니다."
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
	' 사용 방법 : Call payco_URLDecode(Encoding URL)
	' Encoding URL : 인코딩된 URL
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
	' 로그 기록 함수 ( 디버그용 )
	' 사용 방법 : Call payco_Write_Log(Log_String)
	' Log_String : 로그 파일에 기록할 내용
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
		' 내용 기록
		'-----------------------------------------------------------------------------
		oTextStream.WriteLine  CStr(FormatDateTime(Now,0)) + " " + Replace(CStr(Log_String),Chr(0),"'")
		'-----------------------------------------------------------------------------
		' 리소스 해제
		'-----------------------------------------------------------------------------
		oTextStream.Close
		Set oTextStream = Nothing
		Set oFSO = Nothing
	End Sub

	'-----------------------------------------------------------------------------
	' API 호출 함수( POST 전용 - PAYCO 연동은 모든 API 호출에 POST만을 사용합니다. )
	' 사용 방법 : payco_Call_API(SiteURL, App_Mode, Param)
	' SiteURL : 호출할 API 주소
	' App_Mode : 데이터 전송 형태 ( 예: json, x-www-form-urlencoded 등 )
	' Param : 전송할 POST 데이터
	'-----------------------------------------------------------------------------
	Function payco_Call_API(SiteURL, App_Mode, Param)
		Dim HTTP_Object

		'-----------------------------------------------------------------------------
		' WinHttpRequest 선언
		'-----------------------------------------------------------------------------
		Set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		With HTTP_Object
			'API 통신 Timeout 을 30초로 지정
			.SetTimeouts 30000, 30000, 30000, 30000
			.Open "POST", SiteURL, False
			.SetRequestHeader "Content-Type", "application/"+CStr(App_Mode)+"; charset=UTF-8"
			'-----------------------------------------------------------------------------
			' API 전송 정보를 로그 파일에 저장
			'-----------------------------------------------------------------------------
			Call payco_Write_Log("Call API   "+CStr(SiteURL)+" Mode : "  + CStr(App_Mode))
			Call payco_Write_Log("Call API   "+CStr(SiteURL)+" Data : "  + CStr(Param))
			.Send Param
			.WaitForResponse
			'-----------------------------------------------------------------------------
			' 전송 결과를 리턴하기 위해 변수 선언 및 값 대입
			'-----------------------------------------------------------------------------
			Dim Result
			Set Result = New clsHTTP_Object
			Result.Status = CStr(.Status)
			Result.ResponseText = CStr(.ResponseText)
			'-----------------------------------------------------------------------------
			' API 전송 결과를 로그 파일에 저장
			'-----------------------------------------------------------------------------
			Call payco_Write_Log("API Result "+CStr(SiteURL) + " Status : " + CStr(.Status))
			Call payco_Write_Log("API Result "+CStr(SiteURL) + " ResponseText : " + CStr(.ResponseText))
		End With
		Set payco_Call_API = Result
	End Function

	'-----------------------------------------------------------------------------
	' API 결과 전송용 데이터 구조 선언
	' Status 와 ResponseText 만을 전송한다.
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
