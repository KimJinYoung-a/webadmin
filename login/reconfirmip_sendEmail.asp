<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  업체 계정 인증 메일 발송
' History : 2019.05.20 허진원 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim sql
dim sCharge,suserMail , userid,refip ,authNo, sendMsg,sKey, mailCont
dim manager_email ,jungsan_email, deliver_email
userid 		= requestCheckVar(Request.Form("uid"),32)
sCharge 	= requestCheckVar(Request.Form("shp"),1)
refip			= request.ServerVariables("REMOTE_ADDR")
skey  		=	requestCheckVar(Request.Form("sKey"),32)

if userid ="" or sCharge ="" then
	 response.write("<script>alert('인증번호 전송 파라미터에 문제가 발생했습니다. 관리자에게 문의해주세요');</script>") 
response.end
end if
 
if (LCASE(userid)<>LCASE(session("reauthUID"))) then
    response.write("<script>alert(' 암호화에 문제가 발생했습니다. 확인 후 다시 시도해주세요');</script>") 
response.end
end if

if skey <> md5(userid&"TPUSMS") then
	response.write("<script>alert('아이디 암호화에 문제가 발생했습니다. 확인 후 다시 시도해주세요');</script>") 
response.end
end if

sql =" select id, email as manager_email ,jungsan_email, deliver_email from db_partner.dbo.tbl_partner where id ='"&userid&"'"
	rsget.Open sql,dbget,1
  if  not rsget.EOF  then 
   		manager_email =rsget("manager_email")  
   		jungsan_email =rsget("jungsan_email")  
   		deliver_email =rsget("deliver_email")  
   end if
  rsget.close
 
 if sCharge = "J" then 
 	suserMail = jungsan_email
elseif  sCharge = "D" then 
 	suserMail = deliver_email
 else
 	suserMail = manager_email
end if	

if suserMail = "" then
	 response.write("<script>alert('담당자 이메일이 존재하지 않습니다. 확인 후 재시도 해주세요');</script>") 
response.end
end if
 
'// 인증번호 발생 및 DB저장 후 SMS발송
Randomize()
authNo = int(Rnd()*1000000)		'6자리 난수
authNo = Num2Str(authNo,6,"0","R")

'#인증메일 전송

mailCont = "<h1><b>요청하신 <span style=""color:#f22727"">인증번호</span>를 알려드립니다.</b></h1><p></p>"
mailCont = mailCont & "<p>아래 인증번호 6자리를 입력창에 입력해주세요.<br />인증번호는 메일이 발송되는 시점부터 <b>10분간만 유효</b>합니다.</p><br />"
mailCont = mailCont & "<table style=""border-top:2px solid #A0A0A0;border-bottom:2px solid #A0A0A0;text-align:center;"">"
mailCont = mailCont & "<tr style=""border-bottom:1px solid #E0E0E0;"">"
mailCont = mailCont & "<td style=""background-color:#F0F0F0;padding:5px;"">인증번호</td>"
mailCont = mailCont & "<td><b>" & authNo & "</b></td></tr>"
mailCont = mailCont & "<tr>"
mailCont = mailCont & "<td style=""background-color:#F0F0F0;padding:5px;"">발송시간</td>"
mailCont = mailCont & "<td style=""padding:5px;"">" & now() & "</td>"
mailCont = mailCont & "</tr>"
mailCont = mailCont & "</table>"

'#디비저장
sql = "insert into db_partner.dbo.tbl_partner_searchPWD_authno( userid, refip,authno)"&vbcrlf
sql = sql & " values ('"&userid&"','"&refip&"','"&authNo&"')"
dbget.Execute sql

'# 이메일 발송
call sendmailCS(suserMail, "요청하신 인증번호를 알려드립니다.", mailCont)
'response.Write mailCont

'//발송 안내 및 카운터 시작  
sendMsg= "<script type=""text/javascript"">" 
sendMsg = sendMsg &		" parent.document.all.dvAuth.style.display ='';" 
sendMsg = sendMsg &		"	parent.startLimitCounter('newMail');"  
sendMsg = sendMsg &		"	alert('인증 이메일을 발송했습니다.\n이메일을 확인 후 로그인해주세요.');"  
		IF application("Svr_Info")="Dev" THEN 		'// TEST서버이면 그냥 Alert처리 
sendMsg = sendMsg &	"	alert(' [텐바이텐Partner] 접속환경IP인증 인증번호는 [" & authNo & "]입니다.');"  
		end if
sendMsg = sendMsg & "</script>" 
 
response.write sendMsg
%>