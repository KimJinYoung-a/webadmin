
<%
	'-----------------------------------------------------------------------------
	' �佺 ���� ȯ�漳�� ������ ( ASP )
	' incKtosspayCommon.asp
	' 2019.10.11 ������ ����
	'-----------------------------------------------------------------------------
%>
<%
    ''Response.charset = "UTF-8"
	'---------------------------------------------------------------------------------
	' ȯ�溯�� ���� (apiKey ���� sk_test_M1Neq1wmNjM1NeG3Klkj �̸� �׽�Ʈ��)
	' ����
	'   - userinfo���� ����ϴ� paymethod���� 980
	'   - DB�� ���� �α��� Ű���� TS
	'-----------------------------------------------------------------------------
    Dim TossPay_Ready_Url '//����������û URL
	Dim TossPay_Payment_Approve_Url '//�������� URL
	Dim TossPay_Payment_Cancel '//������� URL(���� ��� ���� ������ ���)
	Dim TossPay_Payment_Refunds '// ���� ȯ�� URL(���� �Ϸ� �� �����ݾ� �Ϻ� �Ǵ� ���� ȯ��)
	Dim TossPay_Payment_Status '// ���� ���� Ȯ��(������ ������ ���� ���¸� ��ȸ)

    Dim TossPay_RestApi_Key '//RestAPI Ű
    Dim TossPay_OrderSuccess_Url '//�����������̵��� URL
    Dim TossPay_OrderFail_Url '//�������н��̵��� URL
    Dim TossPay_OrderCancel_Url '//������ҽ��̵��� URL
    Dim TossPay_ApiKey '//�������ڵ�
    Dim TossPay_Custom_Json '//����ȭ�鿡 �����ְ� ���� custom metadata(key, value����){"size":"XL","color":"Red"}
	Dim TossPay_Payment_Method_Type '// �������� ���к���(TOSS_MONEY, CARD �� ���� ������ ��ü)
	Dim TossPay_CashReceipt	'// ���ݿ����� �߱� ���� ����(true, false �⺻�� true�̹Ƿ� ���� ��Ű�� �������� ����)
	Dim TossPay_CashReceiptOption '// ���ݿ����� �߱�Ÿ��(CULTURE - ��ȭ��, GENERAL - �Ϲ�, PUBLIC_TP - �����) �ϴ� ������
	'// 2019-10-16�� ���� �佺 ���� ī�� ���
	'// ���� - 1, ���� - 2, �Ｚ - 3, ����(������) - 4, �Ե� - 5, �ϳ� - 6, �츮(������) - 7, ����(������) - 8, ��Ƽ(������) - 9, �� - 10
	Dim TossPay_Available_Cards '// ī������Ѹ��(���������ü){'options':[{'cardCompanyCode':3}]}
    Dim TossPay_LogUse '//�α׻�뿩��



    '// �佺 ���� ���� RestApi Url �� Key��
    TossPay_Ready_Url                                  	= "https://pay.toss.im/api/v1/payments" '//����������û URL
    TossPay_Payment_Approve_Url                        	= "https://pay.toss.im/api/v2/execute" '//�������� URL
	TossPay_Payment_Cancel								= "https://pay.toss.im/api/v1/cancel" '// ������� URL
	TossPay_Payment_Refunds 							= "https://pay.toss.im/api/v2/refunds" '// ����ȯ�� URL
	TossPay_Payment_Status 								= "https://pay.toss.im/api/v1/status" '// ���� ���� Ȯ�� URL

    if (application("Svr_Info")="Dev") then
		'���߼���
        TossPay_OrderSuccess_Url       = "http://2015www.10x10.co.kr/inipay/tosspay/ordertemp_tossresult.asp" '//�����������̵���url
        TossPay_OrderFail_Url          = "http://2015www.10x10.co.kr/inipay/tosspay/ordertemp_tossfail.asp" '//�������н��̵���url
        TossPay_OrderCancel_Url        = "http://2015www.10x10.co.kr/inipay/tosspay/ordertemp_tossfail.asp" '//������ҽ��̵���url
        TossPay_RestApi_Key            = "sk_test_M1Neq1wmNjM1NeG3Klkj"'//�������ڵ�(�׽�Ʈ��)
		TossPay_Payment_Method_Type	   = ""
		TossPay_CashReceipt			   = ""
		TossPay_CashReceiptOption	   = ""
		TossPay_Available_Cards        = ""
        TossPay_LogUse                 = False

    Else
		'�Ǽ���
        TossPay_OrderSuccess_Url       = "https://www.10x10.co.kr/inipay/tosspay/ordertemp_tossresult.asp" '//�����������̵���url
        TossPay_OrderFail_Url          = "https://www.10x10.co.kr/inipay/tosspay/ordertemp_tossfail.asp" '//�������н��̵���url
        TossPay_OrderCancel_Url        = "https://www.10x10.co.kr/inipay/tosspay/ordertemp_tossfail.asp" '//������ҽ��̵���url
        ''TossPay_RestApi_Key            = "sk_test_M1Neq1wmNjM1NeG3Klkj"'//�������ڵ�(�׽�Ʈ��)
        TossPay_RestApi_Key            = "sk_live_3AkvOVG7263AkvlMPLN6"'//�������ڵ�(�ǰ�����)
		TossPay_Payment_Method_Type	   = ""
		TossPay_CashReceipt			   = ""
		TossPay_CashReceiptOption	   = ""
		TossPay_Available_Cards        = ""
        TossPay_LogUse                 = False
    End If

	'---------------------------------------------------------------------------------
	' �α� ���� ���� ( ��Ʈ��κ��� \tosspay\asp\log �������� ������ �� �����ϴ�. )
	'---------------------------------------------------------------------------------
	Dim Write_LogFile
	Write_LogFile = Server.MapPath(".") + "\log\Tosspay_Log_"+Replace(FormatDateTime(Now,2),"-","")+"_asp.txt"


	'-----------------------------------------------------------------------------
	' �α� ��� �Լ� ( ����׿� )
	' ��� ��� : Call Write_Log(Log_String)
	' Log_String : �α� ���Ͽ� ����� ����
	'-----------------------------------------------------------------------------
	''Const fsoForReading = 1		'- Open a file for reading. You cannot write to this file.
	''Const fsoForWriting = 2		'- Open a file for writing.
	''Const fsoForAppend = 8		'- Open a file and write to the end of the file.
	Sub Toss_Write_Log(Log_String)
		If Not Tosspay_Log_ Then Exit Sub
		'On Error Resume Next
		Dim oFSO
		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		Dim oTextStream
		Set oTextStream = oFSO.OpenTextFile(Write_LogFile, 8, True, 0)
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
	' API ȣ�� �Լ�( POST ���� - TOSSPAY ������ ��� API ȣ�⿡ POST���� ����մϴ�. )
	' ��� ��� : Call_API(SiteURL, App_Mode, Param)
	' SiteURL : ȣ���� API �ּ�
	' App_Mode : ������ ���� ���� ( ��: json, x-www-form-urlencoded �� )
	' Param : ������ POST ������
	'-----------------------------------------------------------------------------
	Function Call_API(SiteURL, App_Mode, Param)
		Dim HTTP_Object

		'-----------------------------------------------------------------------------
		' WinHttpRequest ����
		'-----------------------------------------------------------------------------
		If (application("Svr_Info")	= "Dev") Then
			'set HTTP_Object = Server.CreateObject("Msxml2.ServerXMLHTTP")	'xmlHTTP���۳�Ʈ ����
			set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")	'xmlHTTP���۳�Ʈ ����
		Else
			Set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		End If
		With HTTP_Object
			'API ��� Timeout �� 30�ʷ� ����
			.SetTimeouts 30000, 30000, 30000, 30000
			.Open "POST", SiteURL, False
			.SetRequestHeader "Content-Type", "application/"+CStr(App_Mode)+"; charset=UTF-8"
			'-----------------------------------------------------------------------------
			' API ���� ������ �α� ���Ͽ� ����
			'-----------------------------------------------------------------------------
			'Call Toss_Write_Log("Call API   "+CStr(SiteURL)+" Mode : "  + CStr(App_Mode))
			'Call Toss_Write_Log("Call API   "+CStr(SiteURL)+" Data : "  + CStr(Param))
			.Send Param
			.WaitForResponse 60
			'-----------------------------------------------------------------------------
			' ���� ����� �����ϱ� ���� ���� ���� �� �� ����
			'-----------------------------------------------------------------------------
			Dim Result
			Set Result = New clsHTTP_Toss_Object
			Result.Status = CStr(.Status)
			Result.ResponseText = CStr(.ResponseText)
			'-----------------------------------------------------------------------------
			' API ���� ����� �α� ���Ͽ� ����
			'-----------------------------------------------------------------------------
			'Call Toss_Write_Log("API Result "+CStr(SiteURL) + " Status : " + CStr(.Status))
			'Call Toss_Write_Log("API Result "+CStr(SiteURL) + " ResponseText : " + CStr(.ResponseText))
		End With
		Set Call_API = Result
	End Function

	'---------------------------------------------------------------------------------
	' �ֹ� ���� API ȣ�� �Լ�
	' ��� ��� : Call toss_reserve(mData)
	' mData - parameter ������
	'---------------------------------------------------------------------------------
	Function tossapi_reserve(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Ready_Url, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "���� ���� ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", resultJson.data("message")
						'.Add "message_code", resultJson.data("errorCode")
						'.Add "code", resultJson.data("code")
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_reserve = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' ���� ���� API ȣ�� �Լ�
	' ��� ��� : Call tossapi_order_confirm(mData)
	' �佺 ���������� x-www-form-urlencoded�� �ƴ� json���� �����ߵ�-_-
	' mData - parameter ������
	'---------------------------------------------------------------------------------
	Function tossapi_order_confirm(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Payment_Approve_Url, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "���� ���� ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", resultJson.data("message")
						'If resultJson.data("msg") <> "" Then
						'	.Add "message", resultJson.data("msg")
						'	.Add "message_code", resultJson.data("errorCode")
						'Else
						'	.Add "message", ""
						'	.Add "message_code", ""
						'End If
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_order_confirm = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' �ֹ� ���� Ȯ�� API ȣ�� �Լ�
	' ��� ��� : Call tossapi_ordercheck(mData)
	' mData - parameter ������
	'---------------------------------------------------------------------------------
	Function tossapi_ordercheck(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Payment_Status, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "���� �� ���� ȣ�� ���� ������ �߻��Ͽ����ϴ�."
						.Add "message", resultJson.data("message")
						'.Add "message_code", resultJson.data("errorCode")
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_ordercheck = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' ���� ���� �� ��� API ȣ�� �Լ�(�̰� ���� ȯ���� �ƴ� ���� ��� ������ �����)
	' ��� ��� : Call tossapi_ordercancel(mData)
	' mData - parameter ������
	'---------------------------------------------------------------------------------
	Function tossapi_ordercancel(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Payment_Cancel, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "���� ��� �� ������ �߻��Ͽ����ϴ�."
						.Add "message", resultJson.data("message")
						'.Add "message_code", resultJson.data("errorCode")
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_ordercancel = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' ȯ�� API ȣ�� �Լ�
	' ��� ��� : Call tossapi_refund(mData)
	' mData - parameter ������
	'---------------------------------------------------------------------------------
	Function tossapi_refund(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Payment_Refunds, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
        ''Response.write Result.ResponseText & "<br />"
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "ERR"
						.Add "message", resultJson.data("message")
						'.Add "message_code", resultJson.data("errorCode")
						'.Add "code", resultJson.data("code")
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_refund = resultValue
	End Function

	Function tossapi_return_order_status_value(v)
		Select Case trim(v)
			Case "PAY_STANDBY"
				'// ���� ���� �����Ǿ���, �������� ���� ������ ��� ���� ����.
				'// �� ���¿��� �����ڳ� �������� ������ ����� �� �ֽ��ϴ�.
				'// ���� ������ '���� �Ⱓ'�� �����ϸ� �ڵ����� ��� �˴ϴ�.
				tossapi_return_order_status_value = "���� ��� ��"

			Case "PAY_APPROVED"
				'// ������ ���� ������ ���� �Ϸ�ǰ�, �������� ���� ������ ��ٸ��� ����.
				'// (���� ���� �� 'autoExecute'�� false�� ������ ��쿡�� �� �ܰ踦 ��Ĩ�ϴ�)
				'// �ٹ������� ���� ���� Ȯ���� �츮�ʿ��� �ϹǷ� autoExecute�� false �̾����.
				tossapi_return_order_status_value = "������ ���� �Ϸ�"

			Case "PAY_CANCEL"
				'// ������ �Ϸ�Ǳ� ���� �����ڳ� �������� ������ ����� �����Դϴ�.
				'// (��������� �ݾ��� �̵� ���� ����� ��)
				tossapi_return_order_status_value = "���� ���"

			Case "PAY_PROGRESS"
				'// �����ڰ� ������ �����Ͽ� �������� ���¿��� ���� �ݾ��� ��� ó�� ���� �����Դϴ�.
				tossapi_return_order_status_value = "���� ���� ��"

			Case "PAY_COMPLETE"
				'// ������ �� �������� ���� ���� �� ����� ���������� �Ϸ�� �����Դϴ�.
				tossapi_return_order_status_value = "���� �Ϸ�"

			Case "REFUND_PROGRESS"
				'// ���� �Ǵ� �κ� ȯ���� ���� ���� ���·�, �Ϸ�Ǳ� �� ���� �ٸ� ȯ���� ������ �� �����ϴ�.
				tossapi_return_order_status_value = "ȯ�� ���� ��"

			Case "REFUND_SUCCESS"
				'// ���� �Ǵ� �κ� ȯ���� �Ϸ�Ǿ�, ȯ�� ó���� �ݾ��� �������� ���·� �Ա� �Ϸ�� �����Դϴ�.
				tossapi_return_order_status_value = "ȯ�� ����"

			Case "SETTLEMENT_COMPLETE"
				'// ���� �Ϸ�� �ݾ׿� ���� ������ �Ϸ�Ǿ� �� �̻� ȯ���� �Ұ��� �����Դϴ�.
				'// (������ �Ǵ� ���� Ȯ���Ϸκ��� 1�� ���)
				tossapi_return_order_status_value = "���� �Ϸ�"

			Case "SETTLEMENT_REFUND_COMPLETE"
				'// ���� �Ǵ� �κ� ȯ�ҿ� ���� ������ �Ϸ�Ǿ� �� �̻� ȯ���� �Ұ��� �����Դϴ�.
				'// (������ �Ǵ� ���� Ȯ���Ϸκ��� 1�� ����߰ų� ���� ȯ�ҿ� ���� ���� �Ϸ�� ���)
				tossapi_return_order_status_value = "ȯ�� ���� �Ϸ�"

			Case Else
				tossapi_return_order_status_value = ""
		End Select
	End Function

	'-----------------------------------------------------------------------------
	' API ��� ���ۿ� ������ ���� ����
	' Status �� ResponseText ���� �����Ѵ�.
	'-----------------------------------------------------------------------------
	Class clsHTTP_Toss_Object
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
