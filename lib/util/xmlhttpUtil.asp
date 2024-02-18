<%
Function SendReq(call_url, sedata)
    dim objHttp, ret_txt, status
    Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")

    on error resume next
    objHttp.Open "POST", call_url, False
    objHttp.setRequestHeader "Connection", "close"
    objHttp.setRequestHeader "Content-Length", Len(sedata)
    objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHttp.setTimeouts 5000,90000,90000,90000
    objHttp.Send  sedata
    '������ ����� �������°��� ������ �ɴϴ�.
    status = objHttp.status

	'������ �ְų� (������ ������� err.number�� 0 ���� ����) status ���� 200 (�ε� ����) �� �ƴҰ��
	if err.number <> 0 or status <> 200 then
	      if status = 404 then
	            ret_txt = "[404]�������� �ʴ� ������ �Դϴ�."
	      elseif status >= 401 and status < 402 then
	            ret_txt = "[401]������ ������ ������ �Դϴ�."
	      elseif status >= 500 and status <= 600 then
	            ret_txt = "[500]���� ���� ���� �Դϴ�."
	      else
	            ret_txt = "[err]������ �ٿ�Ǿ��ų� �ùٸ� ��ΰ� �ƴմϴ�."
	      end if
	'������ ���� (������ ���������� �ε���)
	else
	      ret_txt = objHttp.ResponseBody
	end if
	on Error Goto 0
    set objHttp = Nothing
    SendReq = Trim(BinToText(ret_txt,8192))
end function


Function SendReqGet(call_url, sedata)
    dim objHttp, ret_txt, status
    Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")

    on error resume next
    objHttp.Open "GET", call_url & "?" & sedata , False
'    objHttp.setRequestHeader "Connection", "close"
'    objHttp.setRequestHeader "Content-Length", Len(sedata)
'    objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHttp.setTimeouts 5000,90000,90000,90000
    objHttp.Send 

    '������ ����� �������°��� ������ �ɴϴ�.
    status = objHttp.status

	'������ �ְų� (������ ������� err.number�� 0 ���� ����) status ���� 200 (�ε� ����) �� �ƴҰ��
	if err.number <> 0 or status <> 200 then
	      if status = 404 then
	            ret_txt = "[404]�������� �ʴ� ������ �Դϴ�."
	      elseif status >= 401 and status < 402 then
	            ret_txt = "[401]������ ������ ������ �Դϴ�."
	      elseif status >= 500 and status <= 600 then
	            ret_txt = "[500]���� ���� ���� �Դϴ�."
	      else
	            ret_txt = "[err]������ �ٿ�Ǿ��ų� �ùٸ� ��ΰ� �ƴմϴ�."
	      end if
	'������ ���� (������ ���������� �ε���)
	else
	      ret_txt = objHttp.ResponseBody
	end if
	on Error Goto 0
    set objHttp = Nothing
    
    SendReqGet = Trim(BinToText(ret_txt,8192))
end function

Function BinToText(varBinData, intDataSizeBytes)
	Const adFldLong = &H00000080
	Const adVarChar = 200

	dim objRS, strV, tmpMsg,isError

	Set objRS = CreateObject("ADODB.Recordset")
	objRS.Fields.Append "txt", adVarChar, intDataSizeBytes, adFldLong
	objRS.Open
	objRS.AddNew
	objRS.Fields("txt").AppendChunk varBinData
	strV=objRS("txt").Value
	BinToText = strV
	objRS.Close
	Set objRS=Nothing
End Function

Function StripTags(htmlDoc)
	Dim rex
	Set rex = new Regexp
	rex.Pattern= "<[^>]+>"
	rex.Global=True
	StripTags =rex.Replace(htmlDoc,"")
	Set rex = Nothing
End Function
%>
