<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : ���������۰���
' History : 2019.11.25 �ѿ�� ����
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

'// ������ȣ
Randomize()
RndNo = int(Rnd()*1000000)		'6�ڸ� ����
RndNo = Num2Str(RndNo,6,"0","R")

if mode = "fileusersend" then
	if senduserhp = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�޴��� ��ȣ�� �����ϴ�');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	if filecertsendgubun = "KAKAOTALK" or filecertsendgubun = "SMS" then
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('���� ������ ����(īī����,SMS)�� �����ϴ�.');"
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

	'/ �������� ���
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

    smstitlestr = "[�ٹ�����]���ǳ���� ������ �Է��� �ּ���."
    smsmsgstr = "�ȳ��ϼ���. �ٹ����� �������Դϴ�." & vbCrLf
    smsmsgstr = smsmsgstr & "�����Ͻ� ����� ������ �Ʒ� ��ũ���� �Է��� �ֽñ� �ٶ��ϴ�." & vbCrLf
    smsmsgstr = smsmsgstr & "�����մϴ�. ��ſ� �Ϸ� �Ǽ���. :D" & vbCrLf
    smsmsgstr = smsmsgstr & "https://m.10x10.co.kr/cscenter/cs_file_send.asp?nb="& trim(authidxtmp) &"&certNo="& trim(certNo) &""

    btnJson = "{""button"":[{""name"":""������������"",""type"":""WL"", ""url_mobile"":""https://m.10x10.co.kr/cscenter/cs_file_send.asp?nb="& trim(authidxtmp) &"&certNo="& trim(certNo) &"""}]}"
    kakaomsgstr = "�ȳ��ϼ���. �ٹ����� �������Դϴ�." & vbCrLf & vbCrLf
    kakaomsgstr = kakaomsgstr & "�����Ͻ� ����� ������ �Ʒ� ��ũ���� �Է��� �ֽñ� �ٶ��ϴ�." & vbCrLf & vbCrLf
    kakaomsgstr = kakaomsgstr & "�����մϴ�. ��ſ� �Ϸ� �Ǽ���. :D"

    ' īī���� �߼�. ���� ������ �� ��߼� �ϸ� �ȵ�. IP����. �׼������� ���� ����. ���� �߼۵�.
    if filecertsendgubun = "KAKAOTALK" then
        'Call SendKakaoMsg_LINK(trim(senduserhp),"1644-6030","A-0008",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson)
		Call SendKakaoCSMsg_LINK("",trim(senduserhp),"1644-6030","KC-0008",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson,"",trim(senduserid))

    ' SMS �߼�
    elseif filecertsendgubun = "SMS" then
		call SendNormalLMS(trim(senduserhp), smstitlestr, "1644-6030", smsmsgstr)
        'sql = "INSERT INTO [SMSDB].db_LgSMS.dbo.MMS_MSG (SUBJECT,PHONE,CALLBACK,STATUS,REQDATE,MSG,FILE_CNT, EXPIRETIME)" & vbcrlf
        'sql = sql & " 	select '"& smstitlestr &"', '"& trim(senduserhp) &"', '1644-6030','0',getdate(),'"& smsmsgstr &"','0','43200'" & vbcrlf

        'response.write sql &"<br>"
        'dbget.execute sql
    end if

    response.write "<script type='text/javascript'>"
    response.write "	alert('����÷�ο� ��ũ�� ���Բ� �߼� �Ǿ����ϴ�.');"
    response.write "	location.replace('/cscenter/action/pop_cs_file_send.asp?authidx="& trim(authidxtmp) &"&userhp="& trim(senduserhp) &"&menupos="&menupos&"')"
    response.write "</script>"
    dbget.close()	:	response.End

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('�߸��� ��θ� ���� �ϼ̽��ϴ�.');"
	response.write "	history.back();"
	response.write "</script>"
	dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
