<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : ���ر� ����
'			 2022.07.12 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->

<%
	Dim strSql, vWorkerTemp, vWorkerViewTemp, vReferTemp, vReferViewTemp, vFileTemp, vRFileTemp, i, sDoc_workername, sDoc_Refer, sDoc_ReferName, sDoc_R_SMS
	Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status, sDoc_Start, sDoc_End, sDoc_Type, sDoc_Import, sDoc_Diffi, sDoc_Subj, sDoc_Content, sDoc_Worker, sDoc_File, sDoc_RealFile, sDoc_WorkerView, sDoc_UseYN, sDoc_SMS
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	sDoc_Id			= session("ssBctId")
	sDoc_UseYN		= NullFillWith(requestCheckVar(Request("doc_useyn"),1),"Y")
	sDoc_Status		= NullFillWith(requestCheckVar(Request("doc_status"),2),1)
	sDoc_Start		= NullFillWith(requestCheckVar(Request("doc_start"),50),"")
	sDoc_End		= NullFillWith(requestCheckVar(Request("doc_end"),50),"")
	sDoc_Type		= NullFillWith(requestCheckVar(Request("doc_type"),2),0)
	sDoc_Import		= NullFillWith(requestCheckVar(Request("doc_important"),2),0)
	sDoc_Diffi		= NullFillWith(requestCheckVar(Request("doc_difficult"),2),0)
	sDoc_Worker		= NullFillWith(requestCheckVar(Request("doc_worker"),1000),"")
	'sDoc_WorkerView	= Replace(Request("doc_workerview"),"x","")
	sDoc_Subj		= html2db(Request("doc_subject"))
	sDoc_Content	= html2db(Request("doc_content"))
	sDoc_File		= NullFillWith(Request("doc_file"),"")
	sDoc_RealFile	= NullFillWith(Request("doc_realfile"),"")
	sDoc_SMS		= NullFillWith(Request("sms_send"),"")
	sDoc_R_SMS		= NullFillWith(Request("sms_r_send"),"")
	sDoc_workername	= NullFillWith(Request("doc_workername"),"")
	sDoc_Refer		= NullFillWith(Request("doc_refer"),"")
	sDoc_ReferName	= NullFillWith(Request("doc_refername"),"")	'####### 2011-06-30 ������ �۾�����.

	If (checkNotValidHTML(sDoc_Subj) = true) Then
		response.write "<script type='text/javascript'>alert('���񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
		response.End
	End If	

	If (checkNotValidHTML(sDoc_File) = true) Then
		response.write "<script type='text/javascript'>alert('���ϸ��� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
		response.End
	End If

	If (checkNotValidHTML(sDoc_RealFile) = true) Then
		response.write "<script type='text/javascript'>alert('���ϸ��� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
		response.End
	End If			
	
'On Error Resume Next
'dbget.beginTrans
	
	If iDoc_Idx = "" Then
		strSql = " INSERT INTO [db_partner].[dbo].tbl_cooperate_document " & _
				 "		(id, doc_startdate, doc_enddate, doc_type, doc_important, doc_difficult, doc_subject, doc_content, doc_status, doc_workername, doc_refername) " & _
				 "	VALUES " & _
				 "		('" & sDoc_Id & "', '" & sDoc_Start & "', '" & sDoc_End & "', '" & sDoc_Type & "', '" & sDoc_Import & "', '" & sDoc_Diffi & "', " & _
				 "		'" & sDoc_Subj & "', '" & sDoc_Content & "', '" & sDoc_Status & "', '" & sDoc_workername & "', '" & sDoc_ReferName & "') "
		dbget.execute strSql
		
		strSql = " SELECT SCOPE_IDENTITY() "
		rsget.Open strSql,dbget
 		IF Not rsget.EOF THEN
 			iDoc_Idx = rsget(0)
 		ELSE	
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "") 	
 		END IF
 		rsget.close
		
		'####### �α� ���� (insert:1, ������ �ۼ�:1) #######
		Call LogInsert(iDoc_Idx,"1","1")
		'####### �α� ���� #######


	Else
		Dim strSub
		If Request("read") = "o" Then
			'�б� ���(������ ���� ���� ����)
			strSql = " UPDATE [db_partner].[dbo].tbl_cooperate_document SET " & _
					"		doc_status = '" & sDoc_Status & "', " & _
					"		doc_workername = '" & sDoc_workername & "', " & _
					"		doc_refername = '" & sDoc_ReferName & "' " & _
					"	WHERE " & _
					"		doc_idx = '" & iDoc_Idx & "'"
		Else
			'�������(�ۼ��ڸ� ����)
			strSql = " UPDATE [db_partner].[dbo].tbl_cooperate_document SET " & _
					"		doc_startdate = '" & sDoc_Start & "', " & _
					"		doc_enddate = '" & sDoc_End & "', " & _
					"		doc_type = '" & sDoc_Type & "', " & _
					"		doc_important = '" & sDoc_Import & "', " & _
					"		doc_difficult = '" & sDoc_Diffi & "', " & _
					"		doc_subject = '" & sDoc_Subj & "', doc_content = '" & sDoc_Content & "', " & _
					"		doc_status = '" & sDoc_Status & "', " & _
					"		doc_workername = '" & sDoc_workername & "', " & _
					"		doc_refername = '" & sDoc_ReferName & "', " & _
					"		doc_useyn = '" & sDoc_UseYN & "' " & _
					"	WHERE " & _
					"		doc_idx = '" & iDoc_Idx & "' and id='"& trim(sDoc_Id) &"'"
		End If

		'response.write strSql & "<br>"
		dbget.execute strSql

		'####### �α� ���� (update:2, ������ ����:2) (delete:3, ������ ����:3) #######
		If sDoc_UseYN = "Y" Then
			Call LogInsert(iDoc_Idx,"2","2")
		ElseIf  sDoc_UseYN = "N" Then
			Call LogInsert(iDoc_Idx,"3","3")
		End IF
		'####### �α� ���� #######


	End If


	'####### �۾��� ���� #######
	Dim vWTempRs, vWTemp, j
	strSql = ""
	If iDoc_Idx <> "" Then
		'####### ���� �۾��� viewdate ���ؿ�. #######
		strSql = "SELECT worker_id, Convert(varchar(20),worker_viewdate,120) AS worker_viewdate From [db_partner].[dbo].tbl_cooperate_worker WHERE doc_idx = '" & iDoc_Idx & "' "
		rsget.Open strSql,dbget,1
		If Not rsget.Eof Then
			Do Until rsget.Eof
				vWTempRs = vWTempRs & rsget("worker_id") & "=" & rsget("worker_viewdate") & ","
			rsget.Movenext
			Loop
			vWTempRs = Left(vWTempRs,Len(vWTempRs)-1)
			'<!-- //-->
		End If
		rsget.close()
		strSql = " DELETE [db_partner].[dbo].tbl_cooperate_worker WHERE doc_idx = '" & iDoc_Idx & "' "
	End If
	
	
	vWorkerTemp = Split(sDoc_Worker, ",")
	vWTemp = Split(vWTempRs, ",")
	'response.write vWTempRs & "<p>"
	
	For i = 0 To UBOUND(vWorkerTemp)
		For j=0 To UBOUND(vWTemp)
		'response.write Split(vWorkerTemp(i),"|")(0) & "<br>"
		'response.write Split(vWTemp(j),"=")(0) & "<p>"
			If Split(vWorkerTemp(i),"|")(0) = Split(vWTemp(j),"=")(0) Then
				If Split(vWTemp(j), "=")(1) <> "" Then
					vWorkerViewTemp = ", '" & Split(vWTemp(j), "=")(1) & "' "
				Else
					vWorkerViewTemp = ""
				End If
				'<!-- //-->
			End If
		Next
		
		If vWorkerViewTemp = "" Then
			vWorkerViewTemp = ", null"
		End If

		strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_cooperate_worker " & _
						  "		(doc_idx, worker_id, part_sn, worker_viewdate) " & _
						  "	VALUES " & _
						  "		('" & iDoc_Idx & "', '" & Split(vWorkerTemp(i),"|")(0) & "', '" & Split(vWorkerTemp(i),"|")(1) & "' " & vWorkerViewTemp & ") " & vbCrLf
						  
		vWorkerViewTemp = ""
		
		'####### SMS ���� ######
		If sDoc_SMS = "o" Then
			dim StrSMS
			'// SMS ���� �Ӹ��� �ۼ�
			Select Case sDoc_Import
				Case "1"
					StrSMS = "[���]"
				Case "2"
					StrSMS = "[����]"
				Case "3"
					StrSMS = "[����]"
				Case Else
					StrSMS = ""
			End Select
			'StrSMS = StrSMS & session("ssBctCname") & "���� ���������� ���½��ϴ�.(No." & iDoc_Idx & ")"	'####### �ʱ� ���.

			StrSMS = StrSMS & session("ssBctCname") & "���� ��������-" & Trim(Replace(sDoc_Subj,"'",""))
			''StrSMS = chrbyte(Trim(StrSMS),75,"Y")	'####### ... 3byte�� ���� 2byte
			''Call SendNormalSMS_LINK(fnGetMemberHp(Split(vWorkerTemp(i),"|")(0)),"",StrSMS)

			StrSMS = chrbyte(Trim(StrSMS),1000,"Y")
			Call SendRadioWebHookMessage(fnGetMemberEmail(Split(vWorkerTemp(i),"|")(0)),"admin","���� �˸�","��������",StrSMS,"")

			'####### �α� ���� (insert:1, ������ �۾��ڿ��� SMS ����:8) #######
			Call LogInsert(iDoc_Idx,"1","8")
			'####### �α� ���� #######
		End If
	Next
	'response.write strSql & "<br>"
	dbget.execute strSql
	
	
	'####### ������ ���� #######
	Dim vRTempRs, vRTemp
	strSql = ""
	If iDoc_Idx <> "" Then
		'####### ���� ������ ���ؿ�. (������ ��� ����ؼ� viewdate �ִ°� ó�� �۾� �� ����. ���� �߰��ش޶�� �ϸ� �ʵ常 �߰��ؼ� �ణ ������ �ϸ� ��.) #######
		strSql = "SELECT refer_id, Convert(varchar(20),refer_viewdate,120) AS refer_viewdate From [db_partner].[dbo].tbl_cooperate_refer WHERE doc_idx = '" & iDoc_Idx & "' "
		rsget.Open strSql,dbget,1
		If Not rsget.Eof Then
			Do Until rsget.Eof
				vRTempRs = vRTempRs & rsget("refer_id") & "=" & rsget("refer_viewdate") & ","
			rsget.Movenext
			Loop
			vRTempRs = Left(vRTempRs,Len(vRTempRs)-1)
			'<!-- //-->
		End If
		rsget.close()
		strSql = " DELETE [db_partner].[dbo].tbl_cooperate_refer WHERE doc_idx = '" & iDoc_Idx & "' "
	End If
	
	
	vReferTemp = Split(sDoc_Refer, ",")
	vRTemp = Split(vRTempRs, ",")
	'response.write vRTempRs & "<p>"
	
	For i = 0 To UBOUND(vReferTemp)
		For j=0 To UBOUND(vRTemp)
		'response.write Split(vReferTemp(i),"|")(0) & "<br>"
		'response.write Split(vRTemp(j),"=")(0) & "<p>"
			If Split(vReferTemp(i),"|")(0) = Split(vRTemp(j),"=")(0) Then
				If Split(vRTemp(j), "=")(1) <> "" Then
					vReferViewTemp = ", '" & Split(vRTemp(j), "=")(1) & "' "
				Else
					vReferViewTemp = ""
				End If
				'<!-- //-->
			End If
		Next
		
		If vReferViewTemp = "" Then
			vReferViewTemp = ", null"
		End If

		strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_cooperate_refer " & _
						  "		(doc_idx, refer_id, part_sn, refer_viewdate) " & _
						  "	VALUES " & _
						  "		('" & iDoc_Idx & "', '" & Split(vReferTemp(i),"|")(0) & "', '" & Split(vReferTemp(i),"|")(1) & "' " & vReferViewTemp & ") " & vbCrLf
						  
		vReferViewTemp = ""
		
		'####### SMS ���� ######
		If sDoc_R_SMS = "o" Then
			dim StrRSMS
			'// SMS ���� �Ӹ��� �ۼ�
			Select Case sDoc_Import
				Case "1"
					StrRSMS = "[���]"
				Case "2"
					StrRSMS = "[����]"
				Case "3"
					StrRSMS = "[����]"
				Case Else
					StrRSMS = ""
			End Select

			StrRSMS = StrRSMS & session("ssBctCname") & "���� ��������(����)-" & Trim(Replace(sDoc_Subj,"'",""))
			'StrRSMS = chrbyte(Trim(StrRSMS),75,"Y")		'####### ... 3byte�� ���� 2byte
			'Call SendNormalSMS_LINK(fnGetMemberHp(Split(vReferTemp(i),"|")(0)),"",StrRSMS)

			StrRSMS = chrbyte(Trim(StrRSMS),1000,"Y")
			Call SendRadioWebHookMessage(fnGetMemberEmail(Split(vReferTemp(i),"|")(0)),"admin","���� �˸�","��������(����)",StrRSMS,"")
			
			'####### �α� ���� (insert:1, ������ �����ڿ��� SMS ����:9) #######
			Call LogInsert(iDoc_Idx,"1","9")
			'####### �α� ���� #######
		End If
	Next
	'response.write strSql & "<br>"
	dbget.execute strSql
	
	
	'####### ÷������ ���� #######
	If sDoc_File <> "" Then
		strSql = ""
		If iDoc_Idx <> "" Then
			strSql = " DELETE [db_partner].[dbo].tbl_cooperate_file WHERE doc_idx = '" & iDoc_Idx & "' "
		End If
		vFileTemp 	= Split(sDoc_File, ",")
		vRFileTemp	= Split(sDoc_RealFile, ",")
		For i = 0 To UBOUND(vFileTemp)
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_cooperate_file " & _
							  "		(file_name, real_name, doc_idx) " & _
							  "	VALUES " & _
							  "		('" & Trim(vFileTemp(i)) & "', '" & Trim(vRFileTemp(i)) & "', '" & iDoc_Idx & "') " & vbCrLf
		Next
		dbget.execute strSql
	Else
		If requestCheckVar(Request("isfile"),1) = "o" Then
			dbget.execute " DELETE [db_partner].[dbo].tbl_cooperate_file WHERE doc_idx = '" & iDoc_Idx & "' "
		End If
	End If

'0dbget.RollBackTrans
'dbget.CommitTrans
'Response.End
'on error Goto 0

	If Request("gubun") = "write" Then
		Response.Write "<script type='text/javascript'>alert('����Ǿ����ϴ�.');location.href='/admin/notice/cooperate/my_cooperate.asp?menupos="&g_MenuPos&"';</script>"
	Else
		Response.Write "<script type='text/javascript'>alert('����Ǿ����ϴ�.');location.href='/admin/notice/cooperate/?menupos="&g_MenuPos&"';</script>"
	End If
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
