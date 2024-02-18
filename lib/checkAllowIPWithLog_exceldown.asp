<%
dim downFilemenupos,downPersonalInformation_rowcnt

''IP Check :: '' [db_log].[dbo].tbl_partner_login_log 에서 IP 로그 확인가능.
function fnChkAllowIpLog_exceldown(bltype, refip, downFileGubun, downFilemenupos, downPersonalInformation_rowcnt)
    dim sqlStr
    dim scrname : scrname = Request.ServerVariables("SCRIPT_NAME")
    dim strMethod : strMethod = Request.ServerVariables("REQUEST_METHOD")
    dim qryStr
    If strMethod = "POST" Then
        qryStr = (Request.Form)  ''Server.HTMLEncode
    else
        qryStr = Request.QueryString
    end if

    if downFilemenupos="" or downFilemenupos="0" then downFilemenupos="NULL"
    if downPersonalInformation_rowcnt="" or isnull(downPersonalInformation_rowcnt) then downPersonalInformation_rowcnt=0

    sqlStr = "exec db_log.dbo.sp_TEN_ChkAllowIpLog '"&bltype&"','"&session("ssBctID")&"','"&refip&"','"&scrname&"','"&replace(qryStr,"'","")&"','"&LEFT(strMethod,1)&"','"& downFileGubun &"',"& downFilemenupos &","& downPersonalInformation_rowcnt &""
    dbget.Execute sqlStr
end function

dim TMP_check_UserIP
TMP_check_UserIP = request.ServerVariables("REMOTE_ADDR")

''61.252.133.2   - 61.252.133.126 사내 고정.
''115.94.163.43  - 115.94.163.45  사내 유동(Xpeed office)
''203.84.251.209 - '203.84.251.214 물류 고정.
'' 203.84.253.58 - 물류 유동 2015/03/24 3/25 삭제  ,"203.84.253.58" _

''222.111.61.89, 59.5.170.49        물류 유동.
''115.93.29.58 - 115.93.29.60 물류 엑스피드
''115.91.233.123 물류 (3PL)
'' 175.123.149.94   ' 대학로매장
'' "61.252.133.92" ''모의해킹 (2015/06/08)

dim C_ALLOWIPLIST
C_ALLOWIPLIST = Array(  "203.84.251.209","203.84.251.210","203.84.251.211","203.84.251.212","203.84.251.213","203.84.251.214" _
                        ,"115.94.163.42","115.94.163.43","115.94.163.44","115.94.163.45","115.94.163.46" _
                        ,"115.93.29.58","115.93.29.59","115.93.29.60","115.93.29.61","115.93.29.62" _
                        ,"115.91.233.126","115.91.233.124" _
                        ,"61.252.133.2","61.252.133.3","61.252.133.4","61.252.133.5","61.252.133.6" _
                        ,"61.252.133.7","61.252.133.8","61.252.133.9","61.252.133.10","61.252.133.11" _
                        ,"61.252.133.12","61.252.133.13","61.252.133.14","61.252.133.15","61.252.133.16" _
                        ,"61.252.133.17","61.252.133.18","61.252.133.19","61.252.133.20","61.252.133.21" _
                        ,"61.252.133.22","61.252.133.23","61.252.133.24","61.252.133.25","61.252.133.26" _
                        ,"61.252.133.27","61.252.133.28","61.252.133.29","61.252.133.30","61.252.133.31" _
                        ,"61.252.133.32","61.252.133.33","61.252.133.34","61.252.133.35","61.252.133.36" _
                        ,"61.252.133.37","61.252.133.38","61.252.133.39","61.252.133.40","61.252.133.41" _
                        ,"61.252.133.67","61.252.133.68","61.252.133.69","61.252.133.70" _
                        ,"61.252.133.71","61.252.133.72","61.252.133.73","61.252.133.74","61.252.133.75" _
                        ,"61.252.133.76","61.252.133.77","61.252.133.78","61.252.133.79","61.252.133.80" _
                        ,"61.252.133.81","61.252.133.82","61.252.133.83","61.252.133.84","61.252.133.85","61.252.133.86","61.252.133.87","61.252.133.88","61.252.133.91","61.252.133.96" _
                        ,"61.252.133.100","61.252.133.103","61.252.133.104","61.252.133.105","61.252.133.106","61.252.133.107" _
                        ,"61.252.133.113","61.252.133.114","61.252.133.115","61.252.133.116","61.252.133.117","61.252.133.118" _
                        ,"61.252.133.121","61.252.133.122","61.252.133.123","61.252.133.124","61.252.133.125", "61.252.133.92","61.252.133.102" _
                        ,"115.91.233.123","211.46.21.145" _
                        ,"175.123.149.94","118.221.70.31" _
						,"112.218.65.240","112.218.65.241","112.218.65.242","112.218.65.243","112.218.65.244","112.218.65.245" _
						,"112.218.65.246","112.218.65.247","112.218.65.248","112.218.65.249","112.218.65.250","112.218.65.251" _
						,"112.218.65.252","112.218.65.253","112.218.65.254" _
                      )

dim IPCheckOK
dim tmp_ip_i, tmp_ip_buf1
IPCheckOK = false
for tmp_ip_i=0 to UBound(C_ALLOWIPLIST)
    tmp_ip_buf1 = C_ALLOWIPLIST(tmp_ip_i)
    if (TMP_check_UserIP=tmp_ip_buf1) then
        IPCheckOK = true
        Exit For
    end if
next

if (Not IPCheckOK) then
    ' 로그인 권한 IP 관리 DB화   ' 2019.09.20 한용민
	IPCheckOK = fncheckAllowIPWithByDB("Y", "", "")
end if

if (application("Svr_Info")="Dev") then
    Call fnChkAllowIpLog_exceldown("A",TMP_check_UserIP, "EXCEL", downFilemenupos, downPersonalInformation_rowcnt)
else
    if (Not IPCheckOK) then
        Call fnChkAllowIpLog_exceldown("B",TMP_check_UserIP, "EXCEL", downFilemenupos, downPersonalInformation_rowcnt)
    	response.write "승인된 페이지가 아닙니다. 관리자 문의요망 [" & TMP_check_UserIP & "]"
    	response.end
    else
        Call fnChkAllowIpLog_exceldown("A",TMP_check_UserIP, "EXCEL", downFilemenupos, downPersonalInformation_rowcnt)
    end if
end if

%>
