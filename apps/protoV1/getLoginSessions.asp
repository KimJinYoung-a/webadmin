<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
'###############################################
' PageName : /apps/protoV1/getLoginSession.asp
' Discription : SCM 로그인 세션정보 처리
' Request : none
' Response : response > 결과
' History : 2018.05.29 허진원 : 신규 생성
'###############################################

'//헤더 출력
Response.ContentType = "text/json"

dim sToken, oJson, clientIp

'// 접근 IP 확인
dim C_ALLOWIPLIST
C_ALLOWIPLIST = Array(  "192.168.50.4" _
						,"110.93.128.99","172.16.0.99" _
                        ,"61.252.133.71","192.168.1.71" _
                        ,"61.252.133.84","192.168.1.84" _
						,"61.252.133.74","192.168.1.74","61.252.133.4", "192.168.1.99", "110.93.128.94" _
                      )
dim IPCheckOK
dim tmp_ip_i, tmp_ip_buf1

clientIp = request.ServerVariables("REMOTE_ADDR")

IPCheckOK = false
for tmp_ip_i=0 to UBound(C_ALLOWIPLIST)
    tmp_ip_buf1 = C_ALLOWIPLIST(tmp_ip_i)
    if (clientIp=tmp_ip_buf1) then
        IPCheckOK = true
        Exit For
    end if
next

if Not(IPCheckOK) then
  response.Status="403 Forbidden"
  response.Write(response.Status)
  dbget.Close(): response.End
end if


'// HEADER 데이터 접수
sToken = request.ServerVariables("HTTP_Authorization")
if instr(lcase(sToken),"bearer ")>0 then
	sToken = right(sToken,len(sToken)-instr(sToken," "))
else
	response.Status="401 Unauthorized"
	response.Write(response.Status)
	dbget.Close(): response.End
end if
'키값 확인
if sToken<>"1L6O9L>N8CAM@CEFH:D<G:N?O:L6NO6e8O>[7F?^?FGO>=FHF=NTN4K9M4U\7R6P6I>N>IFT" then
	response.Status="401 Unauthorized"
	response.Write(response.Status)
	dbget.Close(): response.End
end if


'// 로그인 데이터 화인
if session("ssBctId")="" then
	response.Status="403.2 Forbidden"
	response.Write(response.Status)
	dbget.Close(): response.End
end if

'// 전송결과 파징
on Error Resume Next

'// json객체 선언
Set oJson = jsObject()

oJson("ssBctId") = session("ssBctId")			'로그인 아이디
oJson("ssBctDiv") = session("ssBctDiv")			'회원구분
oJson("ssBctBigo") = session("ssBctBigo")		'매장 추가 정보
oJson("ssBctSn") = session("ssBctSn")			'직원번호
oJson("ssBctCname") = session("ssBctCname")		'직원 이름
oJson("ssBctEmail") = session("ssBctEmail")		'직원 이메일

oJson("ssGroupid") = session("ssGroupid")		'그룹 코드
oJson("ssAdminPsn") = session("ssAdminPsn")		'부서 번호
oJson("ssAdminLsn") = session("ssAdminLsn")		'등급 번호
oJson("ssAdminPOsn") = session("ssAdminPOsn")		'직책 번호
oJson("ssAdminPOSITsn") = session("ssAdminPOSITsn")	'직급 번호
oJson("ssAdminCLsn") = session("ssAdminCLsn")		'개인정보 취급권한

if ERR then Call OnErrNoti()
On Error Goto 0
'Json 출력(JSON)
oJson.flush

Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->