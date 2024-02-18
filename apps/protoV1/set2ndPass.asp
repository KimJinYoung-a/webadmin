<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/apps/common/appFunction.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/apps/protoV1/protoFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' PageName : /apps/protoV1/set2ndPass.asp
' Discription : 2단계 비밀번호 셋팅
' Request : json > type, pushid, OS, versioncode, versionname, verserion
' Response : response > 결과
' History : 2016.10.18 허진원 : 신규 생성
'###############################################

'//헤더 출력
Response.ContentType = "text/html"

'---------------------------
'@@ 서버 점검시 아래 주석을 풀어주세요
Set oJson = jsObject()
oJson("response") = getErrMsg("9999",sFDesc)
oJson("faildesc") = "핑거스 서비스가 텐바이텐으로 이전되고 종료 되었습니다. 텐바이텐으로 문의 바랍니다. 감사합니다."
oJson.flush
Set oJson = Nothing
Response.End
'---------------------------

'// 웹어드민 접속 로그 저장 함수
Sub AddLoginLog(param1,param2,param3)
    dim sqlStr, reFAddr
    reFAddr = request.ServerVariables("REMOTE_ADDR")

    ''최종 로그인 일자 저장 //2014/07/14 '' tbl_user_tenbyten 사번로그인 제외
    sqlStr = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLog '"&param1&"','"&Left(reFAddr,16)&"','"&param2&"','"&param3&"',0"
    dbget.Execute sqlStr

end Sub

Dim sFDesc
Dim sType
Dim sOS, sVerCd, sUid, sSaved_ID
Dim sData : sData = Request("json")
Dim oJson
Dim sUserPassSec, tokenSn, lgnMethod, AuthNo, sDeviceId
Dim isdbpassword_sec, dbpassword, dbpassword2, Enc_2userpass64
AuthNo=""
tokenSn=""
'// 전송결과 파징
on Error Resume Next

dim oResult
set oResult = JSON.parse(sData)
	sType = oResult.type
	sUid = requestCheckVar(oResult.id,32)
	sOS = requestCheckVar(oResult.OS,10)
	sVerCd = requestCheckVar(oResult.versionname,6)     ''' versioncode => versionname ''2016/12/19
	sSaved_ID = requestCheckVar(oResult.saved_id,1)
	sUserPassSec = requestCheckVar(oResult.pass2,32)
	'sDeviceId = requestCheckVar(oResult.pushid,256)
set oResult = Nothing

'// json객체 선언
Set oJson = jsObject()

dim lockTerm, failNo, loginlock, errmsg, sql
failNo = 5			'// 로그인 실패 허용수
lockTerm = 15		'// 계정 장금 시간 설정(분)

IF (Err) then
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다."
ElseIf sType<>"set2ndpass" then
	'// 2차 비번 설정 타입 아님
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "잘못된 접근입니다."
ElseIf sUid="" then
	'// 잘못된 접근
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "파라메터 정보가 없습니다."
Else

	Dim lastlogindt, lastpwchgdt, lastInfoChgDT, isFirstConnect, isRequirePwdUp, isRequireInfoUp
	'### 유저정보 접수 '//2011-03-9 한용민(정윤정) 수정 - 오푸샵 아이디 추가
	sql = "select top 1 A.id, A.company_name, A.tel, A.fax, A.url, A.email, A.userdiv, A.Enc_password, A.Enc_password64, A.groupid " & vbCrlf
	sql = sql & "	, B.part_sn, A.level_sn, B.job_sn, B.username,  B.direct070, B.usermail, B.posit_sn, IsNull(B.empno, '') as empno " & vbCrlf
	sql = sql & "	, A.lastlogindt, isNULL(A.lastpwchgdt,'2001-01-01') as lastpwchgdt " & vbCrlf
	sql = sql & "	, isNULL(A.lastInfoChgDT,A.regdate) as lastInfoChgDT, IsNull(B.criticinfouser,0) as criticinfouser " & vbCrlf
	sql = sql & " ,(select top 1 shopid" & vbCrlf
	sql = sql & " 	from db_partner.dbo.tbl_partner_shopuser" & vbCrlf
	sql = sql & " 	where b.empno=empno and firstisusing='Y') as firstshopid" & vbCrlf 
	sql = sql & " 	, A.Enc_2password64" 
	sql = sql & " from [db_partner].[dbo].tbl_partner as A " & vbCrlf
	sql = sql & " left join db_partner.dbo.tbl_user_tenbyten as B ON A.id = B.userid AND B.isUsing = 1" & vbCrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sql = sql & " 		and (b.statediv ='Y' or (b.statediv ='N' and datediff(dd,b.retireday,getdate())<=0))" & vbcrlf
	sql = sql & " where A.id = '" & sUid & "'" & vbCrlf
	sql = sql & " and A.isusing='Y'"

	'response.write sql & "<br>"
	rsget.Open sql,dbget,1

	If Not rsget.EOF Then '1
		If sUserPassSec ="" Then'3
			rsget.close
			oJson("response") = getErrMsg("9999",sFDesc)
			oJson("faildesc") = "2차 비밀번호 값이 없습니다.확인해주세요."
		Else'3
			Enc_2userpass64 = SHA256(md5(sUserPassSec))
		
			Dim objCmd,returnValue 
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
				With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdText
					.CommandText = "{?= call db_partner.[dbo].[sp_Ten_partner_SetSecondPassWord]('"&sUid&"',  '"&Enc_2userpass64&"' )}"
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With
					returnValue = objCmd(0).Value
			Set objCmd = Nothing

			If returnValue =-1 Then'2
				rsget.close
				oJson("response") = getErrMsg("9999",sFDesc)
				oJson("faildesc") = "2차 비밀번호가 이미 등록되어있습니다."
			ElseIf returnValue =0 Then'2
				rsget.close
				oJson("response") = getErrMsg("9999",sFDesc)
				oJson("faildesc") = "2차 비밀번호 등록에 실패했습니다. 확인 후 다시 등록해주세요."
			Else'등록 성공'2
				oJson("response") = getErrMsg("1000",sFDesc)
				oJson("resultmsg") = "2차 비밀번호가 등록되었습니다."

				dbpassword  = rsget("Enc_password64")
				dbpassword2  = rsget("Enc_2password64")
				lastlogindt = rsget("lastlogindt")  ''최종 접속 성공일
				lastpwchgdt = rsget("lastpwchgdt")  ''최종 패스워드 변경일
				lastInfoChgDT = rsget("lastInfoChgDT")  ''최종 담당자정보  변경일
		
				isFirstConnect = isNULL(lastlogindt)

		
				If (isFirstConnect) Then
					isRequirePwdUp = True
					isRequireInfoUp = True
				Else
					'' isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
					isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91) And (datediff("d",lastlogindt,now())>0) '' 패스워드 최종변경일을 2014/07/15 부터 넣었으므로.. 우선 lastlogindt 조건넣음.
					If (CLNG(rsget("userdiv"))<10) Then ''일단 직원
						isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
					End If
					isRequireInfoUp=  (datediff("d",lastInfoChgDT,now())>91)
				End If
				
				response.Cookies("partner").domain = "10x10.co.kr"
				If sSaved_ID = "Y" Then
				response.Cookies("partner").Expires = Date + 365	'1년간 쿠키 저장
				Else
				response.Cookies("partner").Expires = Date + 1	'1일간 쿠키 저장
				End If

				response.Cookies("partner")("userid") = rsget("id")
				response.Cookies("partner")("userdiv") = rsget("userdiv")

				'response.Cookies("partner")("ssBctBigo") = rsget("firstshopid")


				response.Cookies("partner")("ssBctSn") = rsget("empno")
				IF rsget("userdiv") <= 9 THEN
				response.Cookies("partner")("ssBctCname") = rsget("username")
				response.Cookies("partner")("ssBctEmail") = db2html(rsget("usermail"))
				Else
					If isnull(rsget("company_name")) Then
						response.Cookies("partner")("ssBctCname") = rsget("username")
						oJson("username") = rsget("username")
					Else
						response.Cookies("partner")("ssBctCname") = db2html(rsget("company_name"))
						oJson("username") = db2html(rsget("company_name"))
					End If
					response.Cookies("partner")("ssBctEmail") = db2html(rsget("email"))
				End If
				
				response.Cookies("partner")("ssGroupid") = rsget("groupid")
				'response.Cookies("partner")("ssAdminPsn") = rsget("part_sn")		'부서 번호

				response.Cookies("partner")("ssAdminLsn") = rsget("level_sn")		'등급 번호
				'response.Cookies("partner")("ssAdminPOsn") = rsget("job_sn")		'직책 번호

				'response.Cookies("partner")("ssAdminPOSITsn") = rsget("posit_sn")		'직급 번호

				response.Cookies("partner")("ssAdminCLsn") = rsget("criticinfouser")	'개인정보 취급권한

				'아이디저장
				response.Cookies("PASave").domain = "10x10.co.kr"
				response.cookies("PASave").Expires = Date + 30	'1개월간 쿠키 저장

				If sSaved_ID = "Y" Then
					response.cookies("PASave")("SAVED_ID") = tenEnc(CStr(rsget("id")))
				Else
					response.cookies("PASave")("SAVED_ID") = ""
				End If

				'3PL SSO 용 쿠키생성(유저아이디 + 접속아이피 + 접속일자) 및 암호화
				'로그인 후 아이피가 변경되면(스마트폰 접속 등) 로그인이 실패한다.
				'코딩 간소화를 위해 비밀번호는 쿠키로 생성하지 않는다. 차후 변경필요.(비번 단방향 암호화 및 쿠키저장)
				'Response.Cookies("ThreePL").Domain = "10x10.co.kr"

				'Response.Cookies("ThreePL")("UserID") = TBTEncrypt(CStr(rsget("id")) & "," & Request.ServerVariables("REMOTE_HOST") & "," & Left(now(), 10))

				'2014-12-17 김진영 // API서버용 쿠키생성
				'Response.Cookies("wapi").Domain = "10x10.co.kr"
				'Response.Cookies("wapi")("UserID") = TBTEncrypt(CStr(rsget("id")) & "," & Request.ServerVariables("REMOTE_HOST") & "," & Left(now(), 10))

				If isnull(rsget("part_sn")) OR rsget("part_sn") = "" Then
				Else
					'Response.Cookies("wapi")("PartSN") = TBTEncrypt(rsget("part_sn") & "," & Request.ServerVariables("REMOTE_HOST") & "," & Left(now(), 10))
				End If
				''로그저장(성공)
				rsget.close
				'######## 강사 이미지 불러오기
				sql = "select top 1 image_400x400 from db_academy.dbo.tbl_corner_good_item" & vbCrlf
				sql = sql & " where lecturer_id = '" & sUid & "'" & vbCrlf
				sql = sql & " and isusing='Y'"
				rsACADEMYget.Open sql,dbACADEMYget,1
				If Not rsACADEMYget.EOF Then
					oJson("userimg") = rsACADEMYget("image_400x400")
				Else
					oJson("userimg") ="none"
				End If
				rsACADEMYget.close
				'######## 푸시 알림 설정 불러오기
				sql = "select top 1 isnull(pushyn,'') from [db_academy].[dbo].[tbl_app_regInfo]" & vbCrlf
				sql = sql & " where deviceid = '" & sDeviceId & "'" & vbCrlf
				sql = sql & " and isusing='Y'"
				rsACADEMYget.Open sql,dbACADEMYget,1
				If Not rsACADEMYget.EOF Then
					If rsACADEMYget("pushyn")="" Then
						oJson("notiyn") = ""
					Else
						oJson("notiyn") = rsACADEMYget("pushyn")
					End If
				Else
					oJson("notiyn") =""
				End If
				rsACADEMYget.close
				If (isFirstConnect) Then
					''최초접속인경우 성공으로 안봄 비번 변경으로 // 비번 변경후 최초로그인 일자 Update , 3개월단위 강제로 할경우 isRequirePwdUp 조건추가.
				Else
					If AuthNo<>"" Then
						Call AddLoginLog (sUid,"Y",AuthNo)
					Else
						Call AddLoginLog (sUid,"Y",tokenSn)
					End If
				End If
			End If'2
		End If'3
	Else'1
		rsget.close
		'// 계정없음
		oJson("response") = getErrMsg("9999",sFDesc)
		oJson("faildesc") = "계정이 활성화 되지 않았거나, 아이디 또는 비밀번호가 틀렸습니다."
	End If'1
End If

If ERR Then Call OnErrNoti()
On Error Goto 0
'Json 출력(JSON)
oJson.flush

Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->