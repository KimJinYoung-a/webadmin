<%

'// SCM,로직스에 있는 "/lib/NoUSBAllowIpList.asp" 은 동일해야 합니다.

function fnIsNoUsbAllowIp()
    dim icheck_UserIP : icheck_UserIP = request.ServerVariables("REMOTE_ADDR")
    dim i,buf1
    dim C_ALLOW_NOUSB_IPLIST
    C_ALLOW_NOUSB_IPLIST = Array("115.94.163.42","115.94.163.43","115.94.163.44","115.94.163.45","115.94.163.46","211.206.236.117")

    fnIsNoUsbAllowIp = false
    for i=0 to UBound(C_ALLOW_NOUSB_IPLIST)
        buf1 = C_ALLOW_NOUSB_IPLIST(i)
        if (icheck_UserIP=buf1) then
            fnIsNoUsbAllowIp = true
            Exit For
        end if
    next

    if (application("Svr_Info")="Dev") then
        fnIsNoUsbAllowIp = (Left(icheck_UserIP,7)="192.168") or (G_IsLocalDev)
    end if

    ''61.252.133.1~126
    buf1 = replace(icheck_UserIP,".","")
    if (LEft(buf1,8)="61252133") then
        if (Mid(buf1,9,3)>0 and Mid(buf1,9,3)<127) then
            fnIsNoUsbAllowIp = true
        end if
    end if
end function

%>
