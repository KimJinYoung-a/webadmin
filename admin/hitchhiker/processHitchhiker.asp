<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	PageName 	: /admin/hitchhiker/downHitchhiker.asp
'	Description : ��ġ����Ŀ ��ûȸ������Ʈ �߼�Ȯ��,
'					�߼۽�û, ��߼۽�û
'	History		: 2006.11.30 ������ ����
'                 2008.02.27 ������ ���� : Max(SendVol) null�� �⺻�� ����
'				  2016.07.19 �ѿ�� ���� SSL ����
'				  2018.09.17 ������ ���� : �߼�Ȯ�� ó���� sms���� ��� �߰�
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%	Dim pMode, smsyn
	Dim strSql, strSqlAdd, strQuery
	Dim iHVol, iAVol, blnSend, sUserID, receviename
	Dim zipcode, addr1, addr2, userphone, usercell
	Dim chkSType
	Dim iMaxV : iMaxV = 1
	dim tmpSql

	pMode = requestCheckVar(Request.Form("pMode"), 32)
	iHVol = requestCheckVar(Request.Form("iHV"), 32)
	iAVol = requestCheckVar(Request.Form("iAV"), 32)
	blnSend	= requestCheckVar(Request.Form("blnS"), 32)
	smsyn	= requestCheckVar(Request.Form("smsyn"), 1)

SELECT CASE pMode
CASE "C"	'//�߼�Ȯ��
	IF iAVol = "" and blnSend <> "" THEN
		strSqlAdd = " and (isnull(ApplyVol,0) > isnull(SendVol,0)) "
	ELSE
		strSqlAdd = " and ApplyVol = "&iAVol
	END IF

	'// LMS�߼�
	' IF application("Svr_Info") = "Dev" THEN
    ' 	tmpSql = " insert into [ACADEMYDB].db_LgSMS.dbo.mms_msg( "
    ' else
    ' 	tmpSql = " insert into [LOGISTICSDB].db_LgSMS.dbo.mms_msg( "
    ' end if

	' tmpSql = tmpSql + " 	subject "
	' tmpSql = tmpSql + " 	, phone "
	' tmpSql = tmpSql + " 	, callback "
	' tmpSql = tmpSql + " 	, status "
	' tmpSql = tmpSql + " 	, reqdate "
	' tmpSql = tmpSql + " 	, msg "
	' tmpSql = tmpSql + " 	, file_cnt "
	' tmpSql = tmpSql + " 	, file_path1 "
	' tmpSql = tmpSql + " 	, expiretime) "
	' tmpSql = tmpSql + " SELECT "
	' tmpSql = tmpSql + " 	'" + html2db("[�ٹ�����] ��ġ����Ŀ �߼۾ȳ�") + "' "
	' tmpSql = tmpSql + " 	, usercell "
	' tmpSql = tmpSql + " 	, '1644-6030' "
	' tmpSql = tmpSql + " 	, '0' "
	' tmpSql = tmpSql + " 	, getdate() "
	' tmpSql = tmpSql + " 	, convert(varchar(4000),'" + ("��ġ����Ŀ " & iHVol & "ȸ�� ����߼۵Ǿ����ϴ�." & vbCrLf & vbCrLf & "7���̳� �����Գ� Ȯ�� �����ϸ�, ��Ÿ ���ǻ����� ������ : 1644-6030 ���� ���� ��Ź �帳�ϴ�." & vbCrLf & vbCrLf & "�ູ ������ �Ϸ� �����ñ� �ٶ��ϴ� :)") + "') "
	' tmpSql = tmpSql + " 	, 0 "
	' tmpSql = tmpSql + " 	, null "
	' tmpSql = tmpSql + " 	, '43200' "
	tmpSql = "INSERT INTO smsdb.[db_kakaoSMS].[dbo].MMS_MSG ( REQDATE, STATUS, TYPE, PHONE, CALLBACK, SUBJECT, MSG, FILE_CNT )"
	tmpSql = tmpSql & " 	select"
	tmpSql = tmpSql & " 	getdate() as REQDATE, '1' as STATUS, '0' as TYPE, usercell as PHONE, '1644-6030' as CALLBACK"
	tmpSql = tmpSql & " 	, N'[�ٹ�����]��ġ����Ŀ �߼۾ȳ�' as SUBJECT"
	tmpSql = tmpSql & " 	, N'��ġ����Ŀ " & iHVol & "ȸ�� ����߼۵Ǿ����ϴ�." & vbCrLf & vbCrLf & "7���̳� �����Գ� Ȯ�� �����ϸ�, ��Ÿ ���ǻ����� �ٹ����� 1:1 ����� �̿����ּ���." & vbCrLf & "�ູ ������ �Ϸ� �����ñ� �ٶ��ϴ� :)' as MSG"
	tmpSql = tmpSql & " 	, '1' as FILE_CNT"
	tmpSql = tmpSql + " 	FROM [db_user].[dbo].[tbl_user_hitchhiker] with (nolock) WHERE"
	tmpSql = tmpSql + " 	SendDate is NULL and HVol = " & iHVol & strSqlAdd
	If smsyn = "Y" Then
		dbget.execute tmpSql
	End If

	strSql = "UPDATE [db_user].[dbo].[tbl_user_hitchhiker]  "&_
			" SET SendVol = ApplyVol , SendDate =getdate() "&_
			" WHERE HVol = "&iHVol &strSqlAdd
	dbget.execute strSql

	IF Err.Number <> 0 THEN
%>
	<script language="javascript">
	alert("������ ó���� ������ �߻��Ͽ����ϴ�. �����ڿ��� ������ �ֽʽÿ�.");
	history.back(-1);
	</script>
<%		dbget.close()	:	response.End
	ELSE
%>
	<script language="javascript">
	alert("�߼�Ȯ�� ó�� �Ǿ����ϴ�.");
	location.href= "<%= getSCMSSLURL %>/admin/hitchhiker/index.asp?iHV=<%=iHVol%>&iAV=<%=iAVol%>&blnS=<%=blnSend%>&chkList=view";
	</script>
<%		dbget.close()	:	response.End
	END IF
CASE "A"	'// �߼۽�û
	sUserID = Request.Form("sUID")
	receviename = Request.Form("receviename")
	zipcode = request.Form("zipcode")
	addr1 = html2db(request.Form("addr1"))
	addr2 = html2db(request.Form("addr2"))
	userphone = request.Form("userphone1") + "-" + request.Form("userphone2") + "-" + request.Form("userphone3")
	usercell = request.Form("usercell1")+ "-" + request.Form("usercell2") + "-" +request.Form("usercell3")

	'��߼����� Ȯ��
	strQuery = " SELECT userid FROM  [db_user].[dbo].[tbl_user_hitchhiker] "&_
			" WHERE HVol = "&iHVol& " and userid = '"&sUserID&"'"
	rsget.Open strQuery, dbget,1
	IF not (rsget.EOF or rsget.BOF) THEN
		chkSType = TRUE
	ELSE
		chkSType = FALSE
	END IF
	rsget.Close

	dbget.beginTrans

	strSql = "SELECT isNull(max(SendVol)+1,1) FROM [db_user].[dbo].[tbl_user_hitchhiker] WHERE HVol = "&iHVol
	rsget.Open strSql, dbget,1
	IF not rsget.eof THEN
		iMaxV = rsget(0)
	END IF
	rsget.Close

	IF chkSType THEN '//��߼�ó��
		strSql = " UPDATE [db_user].[dbo].[tbl_user_hitchhiker] "&_
				" SET ApplyVol = "&iMaxV&",  SendDate = NULL, AdminID = '"&session("ssBctId")&"'"&_
				" WHERE HVol = "&iHVol& " and userid = '"&sUserID&"'"
		dbget.execute strSql

	'��߼� ��û�� �α����̺� ���
	strSql = "INSERT INTO [db_log].[dbo].[tbl_user_hitchhikerLog]  "	&_
			" (iHvol, iAvol, iAvol2, userid, regdate,AdminID)"&_
			" VALUES "&_
			" ("&iHVol&",'',"&iMaxV&",'"&sUserID&"',getdate(),'"&session("ssBctId")&"')"
	dbget.execute strSql
	ELSE
		strSql = "INSERT INTO [db_user].[dbo].[tbl_user_hitchhiker] "	&_
				" (HVol, userid, ApplyVol,recevieName, zipcode, zipaddr, useraddr, userphone, usercell, AdminID)"&_
				" VALUES "&_
				" ("&iHVol&",'"&sUserID&"',"&iMaxV&", '"&recevieName&"','"&zipcode&"', '"&addr1&"', '"&addr2&"', '"&userphone&"', '"&usercell&"', '"&session("ssBctId")&"')"
		dbget.execute strSql
	END IF

	strSql = " UPDATE [db_user].[dbo].tbl_user_n"&_
	 		" SET "&_
	 		" zipcode='" + zipcode + "'"&_
			" ,zipaddr='" + addr1 + "'"&_
			" ,useraddr='" + addr2 + "'"&_
	 		" ,userphone='" + userphone + "'"&_
	 		" ,usercell='" + usercell + "'"  &_
	 		" where userid='" + sUserID + "'"
	dbget.execute strSql

	IF Err.Number = 0 THEN
		dbget.CommitTrans
%>
	<script language="javascript">
	alert("�߼� ��û�� �ּ�Ȯ�� �Ǿ����ϴ�.");
	//location.href= "<%'= getSCMSSLURL %>/admin/hitchhiker/index.asp?iHV=<%'=iHVol%>&iAV=<%'=iAVol%>&blnS=<%'=blnSend%>";
	self.close();
	</script>
<%		dbget.close()	:	response.End
	Else
	   	dbget.RollBackTrans
%>
	<script language="javascript">
	alert("������ ó���� ������ �߻��Ͽ����ϴ�. �����ڿ��� ������ �ֽʽÿ�.");
	history.back(-1);
	</script>
<%		dbget.close()	:	response.End
	End IF
CASE "R"	'// ��߼۽�û
	sUserID = Request.Form("sUID")
	zipcode = request.Form("zipcode")
	addr1 = html2db(request.Form("addr1"))
	addr2 = html2db(request.Form("addr2"))
	userphone = request.Form("userphone1") + "-" + request.Form("userphone2") + "-" + request.Form("userphone3")
	usercell = request.Form("usercell1")+ "-" + request.Form("usercell2") + "-" +request.Form("usercell3")

	dbget.beginTrans

	strSql = "SELECT isNull(max(SendVol)+1,1) FROM [db_user].[dbo].[tbl_user_hitchhiker] WHERE HVol = "&iHVol
	rsget.Open strSql, dbget,1
	IF not rsget.eof THEN
		iMaxV = rsget(0)
	END IF
	rsget.Close

	strSql = " UPDATE [db_user].[dbo].[tbl_user_hitchhiker] "&_
				" SET ApplyVol = "&iMaxV&", SendDate = NULL, AdminID = '"&session("ssBctId")&"'"&_
				" WHERE HVol = "&iHVol& " and userid = '"&sUserID&"'"
	'response.Write strSql
	'dbget.close()	:	response.End
	dbget.execute strSql

	strSql = " UPDATE [db_user].[dbo].tbl_user_n" & vbcrlf
	strSql = strSql & " set zipcode='" & trim(zipcode) & "'" & vbcrlf
	strSql = strSql & " , zipaddr='" & trim(addr1) & "'" & vbcrlf
	strSql = strSql & " , useraddr='" & trim(addr2) & "'" & vbcrlf
	strSql = strSql & " , userphone='" & trim(userphone) & "'" & vbcrlf
	strSql = strSql & " , usercell='" & trim(usercell) & "' where" & vbcrlf
	strSql = strSql & " userid='" & trim(sUserID) & "'" & vbcrlf

	'response.write strSql & "<br>"
	dbget.execute strSql

	'��߼� ��û�� �α����̺� ���
	strSql = "INSERT INTO [db_log].[dbo].[tbl_user_hitchhikerLog]  "	&_
			" (iHvol, iAvol, iAvol2, userid, regdate,AdminID)"&_
			" VALUES "&_
			" ("&iHVol&","&iAVol&","&iMaxV&",'"&sUserID&"',getdate(),'"&session("ssBctId")&"')"
	dbget.execute strSql

	IF Err.Number = 0 THEN
		dbget.CommitTrans
%>
	<script language="javascript">
	alert("��߼� ��û�� �ּ�Ȯ�� �Ǿ����ϴ�.");
	//location.href= "<%'= getSCMSSLURL %>/admin/hitchhiker/index.asp?iHV=<%'=iHVol%>&iAV=<%'=iAVol%>&blnS=<%'=blnSend%>&chkList=view";
	self.close();
	</script>
<%		dbget.close()	:	response.End
	Else
	   	dbget.RollBackTrans
%>
	<script language="javascript">
	alert("������ ó���� ������ �߻��Ͽ����ϴ�. �����ڿ��� ������ �ֽʽÿ�.");
	history.back(-1);
	</script>
<%		dbget.close()	:	response.End
	End IF
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
