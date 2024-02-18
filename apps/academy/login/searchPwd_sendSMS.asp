<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
Session.CodePage = 65001
'###########################################################
' Description :  비밀번호찾기 인증문자 발송
' History : 2016.10.15 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim sql
dim sCharge,susercell , userid,refip ,authNo, sendMsg,sKey
dim manager_hp,jungsan_hp,deliver_hp
userid 		= requestCheckVar(Request.Form("uid"),32)
sCharge 	= requestCheckVar(Request.Form("shp"),1)
refip			= request.ServerVariables("REMOTE_ADDR")
skey  		=	requestCheckVar(Request.Form("sKey"),32)

if userid ="" or sCharge ="" then
	 response.write("<script>alert('인증번호 전송 파라미터에 문제가 발생했습니다. 관리자에게 문의해주세요');</script>") 
response.end
end if
 
if skey <> md5(userid&"TPUSMS") then
	response.write("<script>alert('아이디 암호화에 문제가 발생했습니다. 확인 후 다시 시도해주세요');</script>") 
response.end
end if

sql =" select id, manager_name, manager_hp ,jungsan_name, jungsan_hp, deliver_name, deliver_hp from db_partner.dbo.tbl_partner where id ='"&userid&"'"
	rsget.Open sql,dbget,1
  if  not rsget.EOF  then 
   		manager_hp =rsget("manager_hp")  
   		jungsan_hp =rsget("jungsan_hp")  
   		deliver_hp =rsget("deliver_hp")  
   end if
  rsget.close
 
 if sCharge = "J" then 
 	susercell = jungsan_hp
elseif  sCharge = "D" then 
 	susercell = deliver_hp
 else
 	susercell = manager_hp
end if	

if susercell = "" then
	 response.write("<script>alert('핸드폰 번호가 존재하지 않습니다. 확인 후 재시도 해주세요');</script>") 
response.end
end if
 
'// 인증번호 발생 및 DB저장 후 SMS발송
Randomize()
authNo = int(Rnd()*1000000)		'6자리 난수
authNo = Num2Str(authNo,6,"0","R")

'#문자 전송
''Call SendNormalSMS(susercell,"","[텐바이텐SCM] " & susername & "님 인증번호는 ["&authNo&"]입니다.")
Call SendNormalSMS_LINK(susercell,"","[핑거스Artist] 비밀번호찾기 인증번호는 ["&authNo&"]입니다.")

'#디비저장
sql = "insert into db_partner.dbo.tbl_partner_searchPWD_authno( userid, refip,authno)"&vbcrlf
sql = sql & " values ('"&userid&"','"&refip&"','"&authNo&"')"
dbget.Execute sql

'//발송 안내 및 카운터 시작  
	sendMsg= "<script language=javascript>" 
	sendMsg = sendMsg &		" parent.document.all.dvAuth.style.display ='';" 
	sendMsg = sendMsg &		"	parent.startLimitCounter('new');"  
	sendMsg = sendMsg &		"	alert('인증번호가 발송되었습니다..\n수신 받은 인증번호를 입력해주세요.');"  
			IF application("Svr_Info")="Dev" THEN 		'// TEST서버이면 그냥 Alert처리 
	sendMsg = sendMsg &	"	alert(' [텐바이텐Partner] 비밀번호찾기 인증번호는 [" & authNo & "]입니다.');"  
			end if
 sendMsg = sendMsg & "</script>" 
 
response.write sendMsg
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->