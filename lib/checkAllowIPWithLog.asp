<%
''IP Check :: '' [db_log].[dbo].tbl_partner_login_log ���� IP �α� Ȯ�ΰ���.
function fnChkAllowIpLog(bltype, refip)
    dim sqlStr
    dim scrname : scrname = Request.ServerVariables("SCRIPT_NAME")
    dim strMethod : strMethod = Request.ServerVariables("REQUEST_METHOD")
    dim qryStr
    If strMethod = "POST" Then
        qryStr = (Request.Form)  ''Server.HTMLEncode
    else
        qryStr = Request.QueryString
    end if

    sqlStr = "exec db_log.dbo.sp_TEN_ChkAllowIpLog '"&bltype&"','"&session("ssBctID")&"','"&refip&"','"&scrname&"','"&replace(qryStr,"'","")&"','"&LEFT(strMethod,1)&"'"
    dbget.Execute sqlStr
end function

dim TMP_check_UserIP
TMP_check_UserIP = request.ServerVariables("REMOTE_ADDR")

''210.92.223.194 - 210.92.223.253 �系 ����(��) - �������.
''125.128.177.13                    ���� - �������.
''221.138.98.95, 58.77.2.17         ������ - �������.
''59.7.192.159                      ���� - �������.
''61.252.133.2   - 61.252.133.126 �系(������) ����. => ����
''115.94.163.43  - 115.94.163.45  �系 ����(Xpeed office) => ����

''121.78.103.2   - 121.78.103.126 �系(���ǵ�) ����.
''112.218.65.240 - 112.218.65.254  �系 ����(Xpeed office)
''110.11.187.233 - �系 ����(WiFi)
''203.84.251.209 - '203.84.251.214 ���� ����.
'' 203.84.253.58 - ���� ���� 2015/03/24 3/25 ����  ,"203.84.253.58" _

''222.111.61.89, 59.5.170.49        ���� ����.
''115.93.29.58 - 115.93.29.60 ���� �����ǵ�
''112.221.93.178 - 180 ���Ʒ���  - �������.
''115.91.233.123 ���� (3PL)
''223.62.169.24 , 61.72.241.14 ��Ÿ
'' 59.5.52.207=> 59.5.54.57 =>59.5.52.25=> 59.5.52.31=> 59.5.52.47 => 59.5.52.109���з� =>175.123.149.94
'' 118.221.62.51 �����Ե�	'/2017.06.21
'' 119.206.18.23 ���� �˾������
'' 175.252.136.132 ����

'' 119.192.0.192 �̹��� �ӽ� 2014/08/02
'' 175.209.38.216 ''GS ���� �ӽ� 2014/12/09
'' 210.223.20.106 ����AK��
'' "61.252.133.92" ''������ŷ (2015/06/08) => ����
'' 123.109.32.109 �̿��� �ӽ� (2015/06/12) => ����
'' 175.253.18.49 �̿��� �ӽ� (2015/06/12) => ����
' 112.171.18.242 ������ ���� => ����
'' 119.192.0.105 �̹��� ���� �ӽ�. (2015/12/29)
'' 116.42.227.138 <= 124.58.8.230 <= 58.127.43.37 ����� ���� �ӽ� (2016/02/29)
'' 222.110.25.49 �̴� ��ť���� 2017/01/12
'' 61.73.50.7 �̴� ��ť���� 2017/03/09
'' 61.73.86.179 �̴� ��ť���� 2017/03/15
'/ 221.155.223.196 �ϻ�	'/2018.04.11
' 14.40.95.157 �Ǵ�
'218.146.160.217 �λ��� '/2017.06.08
'211.53.55.89 DDP��	'/2017.12.19
'61.37.16.3 => 210.216.195.7 ���罺Ÿ�ʵ���	'/2019.09.16



dim C_ALLOWIPLIST
C_ALLOWIPLIST = Array(  "203.84.251.209","203.84.251.210","203.84.251.211","203.84.251.212","203.84.251.213" _
                        ,"115.93.29.58","115.93.29.59","115.93.29.60","115.93.29.61","115.93.29.62" _
                        ,"115.91.233.126","211.53.55.89","115.91.233.124" _
                        ,"115.91.233.123","61.73.86.179","14.40.95.157","211.46.21.145","210.216.195.7" _
                        ,"223.62.169.24","61.72.241.14","175.123.149.94","118.221.70.31","119.206.18.23","175.252.136.132" _
                        ,"119.192.0.105","119.192.0.192","210.223.20.106", "116.42.227.138", "118.221.62.51","221.155.223.196","218.146.160.217" _
						,"121.78.103.1","121.78.103.2","121.78.103.3","121.78.103.4","121.78.103.5","121.78.103.6","121.78.103.7","121.78.103.8" _
						,"121.78.103.9","121.78.103.10","121.78.103.11","121.78.103.12","121.78.103.13","121.78.103.14","121.78.103.15","121.78.103.16" _
						,"121.78.103.17","121.78.103.18","121.78.103.19","121.78.103.20","121.78.103.21","121.78.103.22","121.78.103.23","121.78.103.24" _
						,"121.78.103.25","121.78.103.26","121.78.103.27","121.78.103.28","121.78.103.29","121.78.103.30","121.78.103.41","121.78.103.42" _
						,"121.78.103.43","121.78.103.44","121.78.103.45","121.78.103.46","121.78.103.47","121.78.103.48","121.78.103.49","121.78.103.50" _
						,"121.78.103.51","121.78.103.52","121.78.103.53","121.78.103.54","121.78.103.55","121.78.103.56","121.78.103.57","121.78.103.58" _
						,"121.78.103.59","121.78.103.60","121.78.103.61","121.78.103.62","121.78.103.63","121.78.103.64" _
						,"112.218.65.240","112.218.65.241","112.218.65.242","112.218.65.243","112.218.65.244","112.218.65.245" _
						,"112.218.65.246","112.218.65.247","112.218.65.248","112.218.65.249","112.218.65.250","112.218.65.251" _
						,"112.218.65.252","112.218.65.253","112.218.65.254","110.11.187.233" _
						,"192.168." _
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
	if UBound(Split(tmp_ip_buf1, ".")) < 3 then
		if Left(TMP_check_UserIP, Len(tmp_ip_buf1)) = tmp_ip_buf1 then
			IPCheckOK = true
			Exit For
		end if
	end if
next

if (Not IPCheckOK) then
    ' �α��� ���� IP ���� DBȭ   ' 2019.09.20 �ѿ��
	IPCheckOK = fncheckAllowIPWithByDB("Y", "", "")
end if

if (application("Svr_Info")="Dev") then
    Call fnChkAllowIpLog("A",TMP_check_UserIP)
else
    if (Not IPCheckOK) then
        Call fnChkAllowIpLog("B",TMP_check_UserIP)
    	response.write "���ε� �������� �ƴմϴ�. ������ ���ǿ�� [" & TMP_check_UserIP & "]"
    	response.end
    else
        Call fnChkAllowIpLog("A",TMP_check_UserIP)
    end if
end if

%>