<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 고객파일전송관리
' History : 2019.11.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/customer_file_cls.asp" -->
<%
dim menupos, authidx, senduserhp, senduserid, sendorderserial, filecertsendgubun, i, RndNo, KakaoTalkYN, smsyn, adminid
dim certNo, authidxtmp, smstitlestr, smsmsgstr, btnJson, kakaomsgstr, mode, sql, sendasmasteridx
    menupos = requestcheckvar(getNumeric(request("menupos")),10)
    authidx = requestcheckvar(getNumeric(request("authidx")),10)
	senduserhp = requestcheckvar(request("senduserhp"),16)
	senduserid = requestcheckvar(request("senduserid"),32)
	sendorderserial = requestcheckvar(request("sendorderserial"),16)
	filecertsendgubun = requestcheckvar(request("filecertsendgubun"),32)
    mode = requestcheckvar(request("mode"),32)
	sendasmasteridx = requestcheckvar(getNumeric(request("sendasmasteridx")),10)

adminid = session("ssBctId")
smsyn="N"
KakaoTalkYN="N"

'// 인증번호
Randomize()
RndNo = int(Rnd()*1000000)		'6자리 난수
RndNo = Num2Str(RndNo,6,"0","R")

if mode = "fileusersend" then
	if senduserhp = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('휴대폰 번호가 없습니다');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	if filecertsendgubun = "KAKAOTALK" or filecertsendgubun = "SMS" then
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('인증 받으실 구분(카카오톡,SMS)이 없습니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

    senduserhp = replace(senduserhp,"'","")
	if filecertsendgubun = "KAKAOTALK" then
		KakaoTalkYN="Y"

	elseif filecertsendgubun = "SMS" then
		smsyn="Y"
	end if

	if trim(sendasmasteridx)="" or isnull(sendasmasteridx) then sendasmasteridx="NULL"

	'/ 인증정보 등록
	sql = "insert into db_cs.dbo.tbl_customer_filelist(" & vbcrlf
    sql = sql & " userhp,userid,orderserial,smsyn,kakaotalkyn,status,certno,isusing,regdate, adminid, customer_file_regdate, asmasteridx)" & vbcrlf
	sql = sql & " 	select '"& trim(senduserhp)&"', '"& trim(senduserid)&"', '"&trim(sendorderserial)&"', '"&smsyn&"', '"&KakaoTalkYN&"'" & vbcrlf
	sql = sql & " 	,0, '"& trim(RndNo) &"', 'Y', getdate(), '"& adminid &"', NULL, "& sendasmasteridx &"" & vbcrlf

	'response.write sql &"<br>"
	'response.end
	dbget.execute sql

	sql = "select IDENT_CURRENT('db_cs.dbo.tbl_customer_filelist') as authidx"

	'response.write sql &"<br>"
	rsget.open sql ,dbget ,1

	if not(rsget.eof) then
		authidxtmp = rsget("authidx")
	end if

	rsget.close()

    certNo = md5(trim(authidxtmp) & trim(RndNo) & replace(trim(senduserhp),"-",""))

    smstitlestr = "[텐바이텐]문의내용과 파일을 입력해 주세요."
    smsmsgstr = "안녕하세요. 텐바이텐 고객센터입니다." & vbCrLf
    smsmsgstr = smsmsgstr & "문의하실 내용과 파일을 아래 링크에서 입력해 주시기 바랍니다." & vbCrLf
    smsmsgstr = smsmsgstr & "감사합니다. 즐거운 하루 되세요. :D" & vbCrLf
    smsmsgstr = smsmsgstr & "https://m.10x10.co.kr/cscenter/cs_file_send.asp?nb="& trim(authidxtmp) &"&certNo="& trim(certNo) &""

    btnJson = "{""button"":[{""name"":""문의파일전송"",""type"":""WL"", ""url_mobile"":""https://m.10x10.co.kr/cscenter/cs_file_send.asp?nb="& trim(authidxtmp) &"&certNo="& trim(certNo) &"""}]}"
    kakaomsgstr = "안녕하세요. 텐바이텐 고객센터입니다." & vbCrLf & vbCrLf
    kakaomsgstr = kakaomsgstr & "문의하실 내용과 파일을 아래 링크에서 입력해 주시기 바랍니다." & vbCrLf & vbCrLf
    kakaomsgstr = kakaomsgstr & "감사합니다. 즐거운 하루 되세요. :D"

    ' 카카오톡 발송. 같은 내용을 또 재발송 하면 안됨. IP막힘. 테섭에서도 하지 말것. 실제 발송됨.
    if filecertsendgubun = "KAKAOTALK" then
        'Call SendKakaoMsg_LINK(trim(senduserhp),"1644-6030","A-0008",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson)
		Call SendKakaoCSMsg_LINK("",trim(senduserhp),"1644-6030","KC-0008",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson,"",trim(senduserid))

    ' SMS 발송
    elseif filecertsendgubun = "SMS" then
		call SendNormalLMS(trim(senduserhp), smstitlestr, "1644-6030", smsmsgstr)
        'sql = "INSERT INTO [SMSDB].db_LgSMS.dbo.MMS_MSG (SUBJECT,PHONE,CALLBACK,STATUS,REQDATE,MSG,FILE_CNT, EXPIRETIME)" & vbcrlf
        'sql = sql & " 	select '"& smstitlestr &"', '"& trim(senduserhp) &"', '1644-6030','0',getdate(),'"& smsmsgstr &"','0','43200'" & vbcrlf

        'response.write sql &"<br>"
        'dbget.execute sql
    end if

    response.write "<script type='text/javascript'>"
    response.write "	alert('파일첨부용 링크가 고객님께 발송 되었습니다.');"
    response.write "	location.replace('/cscenter/action/pop_cs_file_send.asp?authidx="& trim(authidxtmp) &"&userhp="& trim(senduserhp) &"&menupos="&menupos&"')"
    response.write "</script>"
    dbget.close()	:	response.End

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('잘못된 경로를 지정 하셨습니다.');"
	response.write "	history.back();"
	response.write "</script>"
	dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
