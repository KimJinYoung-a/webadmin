<%

function CheckJenkinsServerIP(ref)
    CheckJenkinsServerIP = false

    dim VaildIP
    if (application("Svr_Info") = "Dev") then
        VaildIP = Array("::1", "192.168.1.67","114.31.63.82")
    else
        VaildIP = Array("61.252.133.67","114.31.63.82", "172.16.0.225", "121.78.103.60")		'// �׽�Ʈ������ : 121.78.103.60
    end if
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckJenkinsServerIP = true
            exit function
        end if
    next
end function

'// ============================================================================
'//
'// status
'//
'// 0000 : ����
'//
'// 1001 : ����Ұ�
'// 2000 : ���뿡��
'//
'// ============================================================================
function WriteJenkinsJsonResponse(status, content)
	dim jsonStr

	jsonStr = "{"
	jsonStr = jsonStr & """status"" : """ & Replace(status, """", "") & ""","
	jsonStr = jsonStr & """content"" : """ & Replace(content, """", "") & """"
	jsonStr = jsonStr & "}"
	Response.ContentType = "application/json"
	Response.Write(jsonStr)
end function

%>
