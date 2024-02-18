<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/apps/common/appFunction.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/apps/protoV1/protoFunction.asp"-->
<!-- #include virtual="/apps/academy/lib/tenSessionLib.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' PageName : /apps/protoV1/loginProc.asp
' Discription : 로그인 처리
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
Dim sType, sDeviceId, sAppKey, sUUID, snID
Dim sOS, sVerCd, sUid, sUPw, sSaved_ID
Dim sData : sData = Request("json")
Dim oJson, ssnlogindt
Dim Enc_userpass, Enc_userpass64, sUserPassSec, tokenSn, lgnMethod, AuthNo
Dim isdbpassword_sec, dbpassword, dbpassword2, Enc_2userpass64
Dim ilecturer_name, sVerCheck
AuthNo=""
tokenSn="fApp" ''로그용 //eastone

'// 전송결과 파징
on Error Resume Next

dim oResult
set oResult = JSON.parse(sData)
	sType = oResult.type
	sUid = requestCheckVar(oResult.id,32)
	sUPw = requestCheckVar(oResult.pass,32)
	sDeviceId = requestCheckVar(oResult.pushid,256)
	sUUID = requestCheckVar(oResult.uuid,40)                ''2016/11/29 추가
	snID = requestCheckVar(oResult.idfa,40)                 ''2016/11/29 추가
	sOS = requestCheckVar(oResult.OS,10)
	sAppKey = getWishAppKey(sOS)
	sVerCd = requestCheckVar(oResult.versionname,6)         ''' versioncode => versionname ''2016/12/19
	sVerCheck = requestCheckVar(oResult.version,6)         ''' 2016/12/19 추가 api 버전 확인
	sSaved_ID = requestCheckVar(oResult.saved_id,1)
	sUserPassSec = requestCheckVar(oResult.pass2,32)
set oResult = Nothing

Enc_userpass = md5(sUPw)
Enc_userpass64 = SHA256(md5(sUPw))


if Len(sUPw)<32 then '' 0을 재끼는듯
    sUPw=format00(32,sUPw)
end if

'// json객체 선언
Set oJson = jsObject()

dim lockTerm, failNo, loginlock, errmsg, sql
failNo = 5			'// 로그인 실패 허용수
lockTerm = 15		'// 계정 장금 시간 설정(분)

IF (Err) then
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다."
	oJson("type") = 0

ElseIf sType<>"login" then
	'// 로그인 타입 아님
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "잘못된 접근입니다."
	oJson("type") = 0

ElseIf sUid="" then
	'// 잘못된 접근
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "파라메터 정보가 없습니다."
	oJson("type") = 0
Else
	'### 계정 로그 확인
	sql = "select  isNull(max(regdate),getdate()) as regdate " &VbCRLF
	sql = sql & "	,isNull(sum(Case loginSuccess When 'N' Then 1 end),0) as FailCnt " &VbCRLF
	sql = sql & "from (select top " & failNo & " regdate, loginSuccess " &VbCRLF
	sql = sql & "	from [db_log].[dbo].tbl_partner_login_log " &VbCRLF
	sql = sql & "	where userid='" & sUid & "' " &VbCRLF
	sql = sql & "	order by idx desc) as pLog " &VbCRLF
    
    rsget.CursorLocation = adUseClient
	rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
		'// 연속 로그인 실패 후 지정시간 동안 계정 잠금
		If (datediff("n",rsget("regdate"),now)<lockTerm) And (rsget("FailCnt")>=failNo) Then
			loginlock="Y"
			errmsg = "비밀번호를 연속으로 " & failNo & "번 틀려 아이디가 잠겼습니다." & (lockTerm-datediff("n",rsget("regdate"),now)) & "분 후 다시 로그인을 해주세요."
		Else
			loginlock="N"
		End If
	rsget.Close

    ''check 계약서 confirm check  2016/12/09
    if (loginlock = "N") then
        sql = "db_partner.[dbo].[sp_Ten_partner_ACA_agree_Check] '"&sUid&"'"
        
        rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
        If Not rsget.EOF Then
            if (rsget("retVal")>0) then
                if (rsget("retVal")=9) then
                    loginlock="Y"
                    errmsg = "FINGERS 작품/강사 아이디만 접근가능합니다."
                else
                    loginlock="Y"
                    errmsg = "업체 계약 승인후 사용 가능합니다. "&vbCRLF&vbCRLF&"PC버전 SCM 로그인 후 업체계약관리 메뉴에서 동의 후 사용가능합니다."
                end if
            end if
        end if
        rsget.close
    end if
    
    
	If sUserPassSec = "" Then
		'### 1차 로그인 확인
		If loginlock = "Y" Then
			oJson("response") = getErrMsg("9999",sFDesc)
			oJson("faildesc") = errmsg
			oJson("type") = 0
		Else
		    '' tbl_user_c JOin ''강사만 접속 가능.
			sql = "select top 1 A.id,   A.Enc_password, A.Enc_password64, A.groupid,  A.Enc_2password64 " & vbCrlf 
			sql = sql & " from [db_partner].[dbo].tbl_partner as A " & vbCrlf 
			sql = sql & "   JOin db_user.dbo.tbl_user_c c" & vbCrlf 
			sql = sql & " on A.id=c.userid" & vbCrlf 
			sql = sql & " where A.id = '" & sUid & "'" & vbCrlf
			sql = sql & " and A.isusing='Y'" & vbCrlf
			sql = sql & " and A.userdiv='9999'"
			sql = sql & " and c.userdiv='14'"
			
			
            rsget.CursorLocation = adUseClient
			rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

			If Not rsget.EOF Then
				'// 로그인 정보 확인
				If rtrim(UCase(rsget("Enc_password64")))=trim(UCase(Enc_userpass64)) Then
					If isnull(rtrim(UCase(rsget("Enc_2password64")))) Or rtrim(UCase(rsget("Enc_2password64")))="" Then
						oJson("response") = getErrMsg("1000",sFDesc)
						oJson("resultmsg") = "등록된 2차 비밀번호가 없습니다. 새로 설정해주세요.."
						oJson("type") = 3
					Else
						dbpassword  = rsget("Enc_password64")
					 
						If isNull(rsget("Enc_2password64"))   Or rsget("Enc_2password64")  ="" Then
							isdbpassword_sec= "N"
						Else
							isdbpassword_sec= "Y"
						End If
						'// 로그인 OK
						oJson("response") = getErrMsg("1000",sFDesc)
						oJson("hidsec") = cStr(isdbpassword_sec)
						oJson("saved_id") = cStr(sSaved_ID)
						oJson("type") = 1
					End If
				Else
					''로그저장(실패)
					If AuthNo<>"" Then
						Call AddLoginLog (sUid,"N",AuthNo)
					Else
						Call AddLoginLog (sUid,"N",tokenSn)
					End If
					oJson("response") = getErrMsg("9999",sFDesc)
					oJson("faildesc") = "아이디 또는 비밀번호가 틀렸습니다. 비밀번호 대소문자를 확인해주세요."
					oJson("type") = 0
				End If
			Else
				'// 계정없음
				oJson("response") = getErrMsg("9999",sFDesc)
				oJson("faildesc") = "계정이 활성화 되지 않았거나, 아이디 또는 비밀번호가 틀렸습니다."
				oJson("type") = 0
			End If
			rsget.close  ''위치변경 eastone
		End If

	Else
		'### 2차 로그인 확인
		If loginlock = "Y" Then '-2
			oJson("response") = getErrMsg("9999",sFDesc)
			oJson("faildesc") = errmsg
			oJson("type") = 0
		Else
			Dim lastlogindt, lastpwchgdt, lastInfoChgDT, isFirstConnect, isRequirePwdUp, isRequireInfoUp
			'### 유저정보 접수 
			sql = "select top 1 A.id, isNULL(A.company_name,'') as company_name, isNULL(A.email,'') as email, A.userdiv, A.Enc_password, A.Enc_password64, A.groupid " & vbCrlf
			sql = sql & "	, A.level_sn " & vbCrlf
			sql = sql & "	, A.lastlogindt, isNULL(A.lastpwchgdt,'2001-01-01') as lastpwchgdt " & vbCrlf
			sql = sql & "	, isNULL(A.lastInfoChgDT,A.regdate) as lastInfoChgDT, '0' as criticinfouser " & vbCrlf
			sql = sql & " 	, A.Enc_2password64" 
			sql = sql & " from [db_partner].[dbo].tbl_partner as A " & vbCrlf
			sql = sql & " where A.id = '" & sUid & "'" & vbCrlf
			sql = sql & " and A.isusing='Y'"
			
			rsget.CursorLocation = adUseClient
			rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

			If Not rsget.EOF Then '1
				'// 로그인 정보 확인
				If rtrim(UCase(rsget("Enc_password64")))=trim(UCase(Enc_userpass64)) Then'2
					If sUserPassSec ="" Then'3
						rsget.close
						oJson("response") = getErrMsg("9999",sFDesc)
						oJson("faildesc") = "2차 비밀번호 값이 없습니다.확인해주세요."
						oJson("type") = 0
					Else'3
						Enc_2userpass64 = SHA256(md5(sUserPassSec))
						If isnull(rtrim(UCase(rsget("Enc_2password64")))) Or rtrim(UCase(rsget("Enc_2password64")))="" Then
							rsget.close
							oJson("response") = getErrMsg("1000",sFDesc)
							oJson("resultmsg") = "등록된 2차 비밀번호가 없습니다. 새로 설정해주세요."
							oJson("type") = 3
						Else
							If rtrim(UCase(rsget("Enc_2password64")))<>trim(UCase(Enc_2userpass64)) Then'4
								rsget.close
								If AuthNo<>"" Then
									Call AddLoginLog (sUid,"N",AuthNo)
								Else
									Call AddLoginLog (sUid,"N",tokenSn)
								End If
								oJson("response") = getErrMsg("9999",sFDesc)
								oJson("faildesc") = "2차 비밀번호가 틀렸습니다.확인후 다시 시도해주세요."
								oJson("type") = 0
							Else'4
								'로그인 성공시 성공 처리
								'// 로그인 OK
								oJson("response") = getErrMsg("1000",sFDesc)
								oJson("type") = 2
								dbpassword  = rsget("Enc_password64")
								dbpassword2  = rsget("Enc_2password64")
								lastlogindt = rsget("lastlogindt")  ''최종 접속 성공일
								lastpwchgdt = rsget("lastpwchgdt")  ''최종 패스워드 변경일
								lastInfoChgDT = rsget("lastInfoChgDT")  ''최종 담당자정보  변경일
						
								isFirstConnect = isNULL(lastlogindt)

						
								if (isFirstConnect) then
									isRequirePwdUp = true
									isRequireInfoUp = true
								else
									'' isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
									isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91) and (datediff("d",lastlogindt,now())>0) '' 패스워드 최종변경일을 2014/07/15 부터 넣었으므로.. 우선 lastlogindt 조건넣음.
									isRequireInfoUp=  (datediff("d",lastInfoChgDT,now())>91)
								end If
								
								ssnlogindt = fnDateTimeToLongTime(now())
								
								response.Cookies("partner").domain = "10x10.co.kr"
								if (sUid="fingertest01") then
								    response.Cookies("partner").Expires = Date + 3	'일정기간  쿠키 저장 => X
								else
							        response.Cookies("partner").Expires = Date + 3	'1일정기간  쿠키 저장 => X => 페이지 이동시 문제가 있음.. ios location.href 로 이동시 ///admin2009scm/apps/academy/ordermaster|orderList.asp
                                end if
                            
								response.Cookies("partner")("ssndt") = ssnlogindt
								''## 보안강화 세션 처리 2017/05/19=================================
								session("ssnuserid") = LCase(rsget("id"))	'' 로그인 세션값
								session("ssnlogindt") = ssnlogindt                      '' 로그인시각 세션값
								session("ssnlastcheckdt") = ssnlogindt                  '' 최종 세션체크시각
								Call fnDBSessionCreate("A")								'' 디비에 세션검증값 저장.


								response.Cookies("partner")("userid") = rsget("id")
								response.Cookies("partner")("userdiv") = rsget("userdiv")
								response.Cookies("partner")("ssBctCname") = db2html(rsget("company_name"))
								response.Cookies("partner")("ssBctEmail") = db2html(rsget("email"))
							
								
								response.Cookies("partner")("ssGroupid") = rsget("groupid")
								response.Cookies("partner")("ssAdminLsn") = rsget("level_sn")		'등급 번호
								''response.Cookies("partner")("ssAdminCLsn") = "0" ''rsget("criticinfouser")	'개인정보 취급권한

								'아이디저장
								response.Cookies("PASave").domain = "10x10.co.kr"
								response.cookies("PASave").Expires = Date + 365	'1개월간 쿠키 저장

								If sSaved_ID = "Y" Then
									response.cookies("PASave")("SAVED_ID") = tenEnc(CStr(rsget("id")))
								Else
									response.cookies("PASave")("SAVED_ID") = ""
								End If
                                
                                ilecturer_name = db2html(rsget("company_name"))
								''로그저장(성공)
								rsget.close
								
								'######## 강사 이미지 불러오기
								sql = "select top 1 isnull(newImage_profile,'') as newImage_profile, isNULL(lecturer_name,'') as lecturer_name from db_academy.dbo.tbl_corner_good" & vbCrlf
								sql = sql & " where lecturer_id = '" & sUid & "'" & vbCrlf
								sql = sql & " and isusing='Y'"
								rsACADEMYget.CursorLocation = adUseClient
								rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
								If Not rsACADEMYget.EOF Then
									if rsACADEMYget("newImage_profile")<>"" then
										oJson("userimg") = "http://image.thefingers.co.kr/corner/newImage_profile/" & rsACADEMYget("newImage_profile")
									Else
										oJson("userimg") = "none"
									end if
									
									if rsACADEMYget("lecturer_name")<>"" then
									    oJson("username") = db2html(rsACADEMYget("lecturer_name"))
										oJson("userid") = sUid
									else
									    oJson("username") = ilecturer_name
										oJson("userid") = sUid
									end if
								Else
									oJson("userimg") ="none"
								End If
								rsACADEMYget.close
								
								'######## 푸시 알림 설정 불러오기
								sql = "select top 1 isnull(pushyn,'A') as pushyn, isAlarm01, isAlarm02, isAlarm03, isAlarm04 from [db_academy].[dbo].[tbl_app_regInfo]" & vbCrlf
								sql = sql & " where deviceid = '" & sDeviceId & "'" & vbCrlf
								sql = sql & " and isusing='Y'"
								rsACADEMYget.CursorLocation = adUseClient
								rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
								If Not rsACADEMYget.EOF Then
									If rsACADEMYget("pushyn")="" Then
										oJson("notiyn") = ""
										oJson("notipush") = "Y"
										oJson("itempush") = "Y"
										oJson("orderpush") = "Y"
										oJson("qnapush") = "Y"
									Else
										oJson("notiyn") = rsACADEMYget("pushyn")
										oJson("notipush") = rsACADEMYget("isAlarm01")
										oJson("itempush") = rsACADEMYget("isAlarm02")
										oJson("orderpush") = rsACADEMYget("isAlarm03")
										oJson("qnapush") = rsACADEMYget("isAlarm04")
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
								
								if (sVerCheck=2) Then
									Dim totalordercount, noticount
									totalordercount = 0
									'푸시알림 뱃지 카운트
									'아직 정해진 것이 없어 임시로 카운트 숫자만
									sql = "exec db_AppNoti.dbo.sp_ACA_getAppHisRecentNotiList_Academy_Artist_ViewCnt '"&sDeviceId&"','"&sUid&"'," & sAppKey
									If sDeviceId<>"" Then
										rsAppNotiget.CursorLocation = adUseClient
										rsAppNotiget.Open sql,dbAppNotiget,adOpenForwardOnly, adLockReadOnly
										If Not rsAppNotiget.Eof Then
											oJson("noticount")=rsAppNotiget("cnt")
											noticount=rsAppNotiget("cnt")
										Else
											oJson("noticount")=0
											noticount=0
										End If
										rsAppNotiget.Close
									Else
										oJson("noticount")=0
										noticount=0
									End If

									'//뱃지 카운트 설정 및 확인
									sql = "exec [db_academy].[dbo].[sp_Academy_App_IconBadgeCountSet] '" + Cstr(sUid) + "', " + Cstr(noticount)
									rsACADEMYget.CursorLocation = adUseClient
									rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
									if not rsACADEMYget.EOF Then
										oJson("ordercount")=rsACADEMYget("mibaljucnt") + rsACADEMYget("ordercnt") + rsACADEMYget("cscnt")
										oJson("qnacount")=rsACADEMYget("qnacnt")
									end if
									rsACADEMYget.Close
								End If
								'// 아티스트 앱 회원정보 생성  2016/11/29 eastone
                        		sql = "IF Not EXISTS(Select userid from db_academy.dbo.tbl_App_artist_userInfo where userid='" & sUid & "') " & vbCrLf
                        		sql = sql & " begin " & vbCrLf
                        		sql = sql & "	Insert Into db_academy.dbo.tbl_App_artist_userInfo (userid,nid,uuid) values ('" & sUid & "','"&snID&"','"&sUUID&"') " & vbCrLf
                        		sql = sql & " end" & vbCrLf
                        		sql = sql & " ELSE" & vbCrLf
                        		sql = sql & " begin " & vbCrLf
                        		sql = sql & "	update db_academy.dbo.tbl_App_artist_userInfo " & vbCrLf
                        		sql = sql & "	SET nid='"&snID&"'" & vbCrLf
                        		sql = sql & "	,uuid='"&sUUID&"'" & vbCrLf
                        		sql = sql & "	,lastLogin=getdate()" & vbCrLf
                        		sql = sql & "	where userid='"&sUid&"'" & vbCrLf
                        		sql = sql & " end"
                        		dbACADEMYget.Execute(sql)
                        		
                        		'// tbl_app_regInfo 테이블 업데이트.
                        		sql = " update db_academy.dbo.tbl_app_regInfo " & vbCrLf
                        		sql = sql & " Set userid='" & sUid & "' " & vbCrLf
                        		sql = sql & "	,appVer='"&sVerCd&"'" & vbCrLf
                        		sql = sql & "	,lastUpdate=getdate() " & vbCrLf
                        		sql = sql & "	,lastact='lgn'" & vbCrLf
                        		sql = sql & " Where appKey='" & sAppKey & "' " & vbCrLf
                        		sql = sql & "	and deviceid='" & sDeviceId & "'"
                        		dbACADEMYget.Execute(sql)
								
								''tbl_app_NidInfo lastUserid 
                                if (snID<>"") then
                                    sql = "update db_academy.dbo.tbl_app_NidInfo " & vbCrLf
                        		    sql = sql & " Set lastUserid='" & sUid & "' " & vbCrLf       
                        		    sql = sql & " where Nid='"&snID&"'" 
                        		    dbACADEMYget.Execute(sql)
                                end if
        
							End If'4
						End If
						
					End If'3
				Else'2
					''로그저장(실패)
					rsget.close
					If AuthNo<>"" Then
						Call AddLoginLog (sUid,"N",AuthNo)
					Else
						Call AddLoginLog (sUid,"N",tokenSn)
					End If
					oJson("response") = getErrMsg("9999",sFDesc)
					oJson("faildesc") = "아이디 또는 비밀번호가 틀렸습니다. 비밀번호 대소문자를 확인해주세요."
					oJson("type") = 0
				End If'2
			Else'1
				rsget.close
				'// 계정없음
				oJson("response") = getErrMsg("9999",sFDesc)
				oJson("faildesc") = "계정이 활성화 되지 않았거나, 아이디 또는 비밀번호가 틀렸습니다."
				oJson("type") = 0
			End If'1
		End If '-2
	End If ''2차로그인 끝

End If

if ERR then Call OnErrNoti()
On Error Goto 0
'Json 출력(JSON)
oJson.flush

Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->