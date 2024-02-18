<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü�������/����
' History : 2015.05.27 ���ر� ����
'			2021.12.06 �ѿ�� ����(���Ѽ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->

<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

Dim tidx, reguserid, groupid, company_name, ceoname, company_no, jungsan_gubun, company_zipcode, company_address, company_address2, company_upjong, company_uptae
Dim company_tel, company_fax, return_zipcode, return_address, return_address2, jungsan_bank, jungsan_acctno, jungsan_acctname, jungsan_date, jungsan_date_off
Dim jungsan_name, jungsan_phone, jungsan_email, jungsan_hp, groupid_old
Dim manager_name, manager_phone, manager_hp, manager_email, comment, gubun, confirmuserid, uid, old_uid, vQuery, vTempTIdx, vStatus, vExtRegUserID, i, vIsSMS

tidx						= request("tidx")
reguserid					= session("ssBctId")
groupid						= request("groupid")
groupid_old					= request("groupid_old")
company_name 				= html2db(request("company_name"))
ceoname						= html2db(request("ceoname"))
company_no  				= request("company_no")
jungsan_gubun 				= request("jungsan_gubun")
company_zipcode 			= request("company_zipcode")
company_address 			= request("company_address")
company_address2 			= request("company_address2")
company_upjong  			= Left(html2db(request("company_upjong")),32)
company_uptae   			= Left(html2db(request("company_uptae")),25)
company_tel 				= request("company_tel")
company_fax 				= request("company_fax")
return_zipcode 				= request("return_zipcode")
return_address 				= request("return_address")
return_address2 			= request("return_address2")
jungsan_bank 				= html2db(request("jungsan_bank"))
jungsan_acctno 				= request("jungsan_acctno")
jungsan_acctname 			= html2db(request("jungsan_acctname"))
jungsan_date 				= request("jungsan_date")
jungsan_date_off			= request("jungsan_date_off")
manager_name 				= html2db(request("manager_name"))
manager_phone 				= request("manager_phone")
manager_hp 					= request("manager_hp")
manager_email 				= request("manager_email")
jungsan_name				= html2db(request("jungsan_name"))
jungsan_phone				= request("jungsan_phone")
jungsan_email				= request("jungsan_email")
jungsan_hp					= request("jungsan_hp")
comment		 				= html2db(request("comment"))
gubun						= request("gubun")
confirmuserid				= session("ssBctId")
uid							= Trim(request("uid"))
old_uid						= Trim(request("old_uid"))
vStatus						= request("status")
vIsSMS = "x"

if (checkNotValidHTML(company_name) = true) Then
	response.write "<script>alert('ȸ���(��ȣ)���� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	dbget.Close
	response.End
End If

'// �� text, input �׸� script ���� Ȯ��(script�� �ԷµǸ� ƨ�ܳ� 2016.07.04 ������ �߰�)
if (checkNotValidHTMLcritical(company_name) = true) Then
	response.write "<script>alert('�ش��׸񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	response.End
End If
if (checkNotValidHTMLcritical(ceoname) = true) Then
	response.write "<script>alert('�ش��׸񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	response.End
End If
if (checkNotValidHTMLcritical(company_upjong) = true) Then
	response.write "<script>alert('�ش��׸񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	response.End
End If
if (checkNotValidHTMLcritical(company_uptae) = true) Then
	response.write "<script>alert('�ش��׸񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	response.End
End If
if (checkNotValidHTMLcritical(jungsan_bank) = true) Then
	response.write "<script>alert('�ش��׸񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	response.End
End If
if (checkNotValidHTMLcritical(jungsan_acctname) = true) Then
	response.write "<script>alert('�ش��׸񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	response.End
End If
if (checkNotValidHTMLcritical(manager_name) = true) Then
	response.write "<script>alert('�ش��׸񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	response.End
End If
if (checkNotValidHTMLcritical(jungsan_name) = true) Then
	response.write "<script>alert('�ش��׸񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	response.End
End If
if (checkNotValidHTMLcritical(comment) = true) Then
	response.write "<script>alert('�ش��׸񿡴� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	response.End
End If

On Error Resume Next
dbget.beginTrans

If tidx = "" Then
	If groupid <> "" Then
		vQuery = "SELECT TOP 1 (SELECT username FROM [db_partner].[dbo].[tbl_user_tenbyten] WHERE userid = A.reguserid) FROM [db_partner].[dbo].[tbl_partner_temp_info] AS A WHERE groupid = '" & groupid & "' AND status NOT IN ('0','3') "
		rsget.Open vQuery,dbget
		IF Not rsget.EOF THEN
			Response.Write "<script>alert('" & rsget(0) & " ���� ����ü�� ��û�� ������� �ֽ��ϴ�.\n�� ���� �Ϸ�� �� ��û�� �� �ֽ��ϴ�.');history.back();</script>"
			rsget.close()
			dbget.RollBackTrans
			dbget.close()
			Response.End
		Else
			rsget.close()
		END IF
	END IF
	
	vQuery = "INSERT INTO [db_partner].[dbo].[tbl_partner_temp_info]" & VbCRLF
	vQuery = vQuery & "(" & VbCRLF
	vQuery = vQuery & "		reguserid, groupid, company_name, ceoname, company_no, jungsan_gubun, company_zipcode, company_address, " & VbCRLF
	vQuery = vQuery & "		company_address2, company_uptae, company_upjong, company_tel, company_fax, return_zipcode, return_address, " & VbCRLF
	vQuery = vQuery & "		return_address2, jungsan_bank, jungsan_acctno, jungsan_acctname, jungsan_date, jungsan_date_off, " & VbCRLF
	vQuery = vQuery & "		manager_name, manager_phone, manager_hp, manager_email, comment, gubun, " & VbCRLF
	vQuery = vQuery & "		jungsan_name, jungsan_phone, jungsan_hp, jungsan_email, groupid_old " & VbCRLF
	vQuery = vQuery & "		,encCompNo" & VbCRLF
	vQuery = vQuery & ") VALUES(" & VbCRLF
	vQuery = vQuery & "		'" & reguserid & "', '" & groupid & "', '" & company_name & "', '" & ceoname & "', '" & socialnoReplace(company_no) & "', '" & jungsan_gubun & "', '" & company_zipcode & "', '" & company_address & "'," & VbCRLF
	vQuery = vQuery & "		'" & company_address2 & "', '" & company_uptae & "', '" & company_upjong & "', '" & company_tel & "', '" & company_fax & "', '" & return_zipcode & "', '" & return_address & "'," & VbCRLF
	vQuery = vQuery & "		'" & return_address2 & "', '" & jungsan_bank & "', '" & jungsan_acctno & "', '" & jungsan_acctname & "', '" & jungsan_date & "', '" & jungsan_date_off & "'," & VbCRLF
	vQuery = vQuery & "		'" & manager_name & "', '" & manager_phone & "', '" & manager_hp & "','" & manager_email & "','" & comment & "', '" & gubun & "'," & VbCRLF
	vQuery = vQuery & "		'" & jungsan_name & "', '" & jungsan_phone & "', '" & jungsan_hp & "', '" & jungsan_email & "', '" & groupid_old & "'" & VbCRLF
	vQuery = vQuery & "		,[db_partner].[dbo].[uf_EncSOCNoPH1]('"&company_no&"')" & VbCRLF    ''2016/08/04 �߰�
	vQuery = vQuery & ")"
	
	dbget.Execute vQuery
	
	vQuery = " SELECT SCOPE_IDENTITY() "
	rsget.Open vQuery,dbget
	IF Not rsget.EOF THEN
		vTempTIdx = rsget(0)
	ELSE
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
	END IF
	rsget.close
	
	if (LEN(Trim(replace(company_no,"-","")))=13) then
		vQuery = "exec [db_cs].[dbo].[usp_Ten_partner_temp_info_Enc_companyno] "&vTempTIdx&",'"&company_no&"'"
		dbget.Execute vQuery
	end if
	
	If groupid = "" Then
		vQuery = "UPDATE [db_partner].[dbo].[tbl_partner_temp_info] SET groupid_old = '" & groupid_old & "' WHERE tidx = '" & vTempTIdx & "'"
		dbget.Execute vQuery
	End If
	
	vQuery = ""
	For i = LBound(Split(uid,",")) To UBound(Split(uid,","))
		vQuery = vQuery & " INSERT INTO [db_partner].[dbo].[tbl_partner_temp_makerid](tidx, makerid) VALUES('" & vTempTIdx & "','" & Trim(Split(uid,",")(i)) & "') " & vbCrLf
	Next
	IF vQuery <> "" Then
		dbget.Execute vQuery
	End IF
	
	vIsSMS = "o"
	tidx = vTempTIdx
Else
	vQuery = "UPDATE [db_partner].[dbo].[tbl_partner_temp_info] SET " & VbCRLF
	vQuery = vQuery & "		groupid = '" & groupid & "', " & VbCRLF
	vQuery = vQuery & "		company_name = '" & company_name & "', " & VbCRLF
	vQuery = vQuery & "		ceoname = '" & ceoname & "', " & VbCRLF
	vQuery = vQuery & "		company_no = '" & socialnoReplace(company_no) & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_gubun = '" & jungsan_gubun & "', " & VbCRLF
	vQuery = vQuery & "		company_zipcode = '" & company_zipcode & "', " & VbCRLF
	vQuery = vQuery & "		company_address = '" & company_address & "', " & VbCRLF
	vQuery = vQuery & "		company_address2 = '" & company_address2 & "', " & VbCRLF
	vQuery = vQuery & "		company_uptae = '" & company_uptae & "', " & VbCRLF
	vQuery = vQuery & "		company_upjong = '" & company_upjong & "', " & VbCRLF
	vQuery = vQuery & "		company_tel = '" & company_tel & "', " & VbCRLF
	vQuery = vQuery & "		company_fax = '" & company_fax & "', " & VbCRLF
	vQuery = vQuery & "		return_zipcode = '" & return_zipcode & "', " & VbCRLF
	vQuery = vQuery & "		return_address = '" & return_address & "', " & VbCRLF
	vQuery = vQuery & "		return_address2 = '" & return_address2 & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_bank = '" & jungsan_bank & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_acctno = '" & jungsan_acctno & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_acctname = '" & jungsan_acctname & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_date = '" & jungsan_date & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_date_off = '" & jungsan_date_off & "', " & VbCRLF
	vQuery = vQuery & "		manager_name = '" & manager_name & "', " & VbCRLF
	vQuery = vQuery & "		manager_phone = '" & manager_phone & "', " & VbCRLF
	vQuery = vQuery & "		manager_hp = '" & manager_hp & "', " & VbCRLF
	vQuery = vQuery & "		manager_email = '" & manager_email & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_name = '" & jungsan_name & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_phone = '" & jungsan_phone & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_hp = '" & jungsan_hp & "', " & VbCRLF
	vQuery = vQuery & "		jungsan_email = '" & jungsan_email & "', " & VbCRLF
	vQuery = vQuery & "		comment = '" & comment & "', " & VbCRLF
	vQuery = vQuery & "		status = '" & vStatus & "', " & VbCRLF
	
	if C_MngPart or C_ADMIN_AUTH then
		vQuery = vQuery & "		confirmuserid = '" & confirmuserid & "', " & VbCRLF
	End If
	vQuery = vQuery & "		lastupdate = getdate() " & VbCRLF
	vQuery = vQuery & "		,encCompNo=[db_partner].[dbo].[uf_EncSOCNoPH1]('"&company_no&"')"& VbCRLF       ''2016/08/04 �߰�
	vQuery = vQuery & "	WHERE " & VbCRLF
	vQuery = vQuery & "		tidx = '" & tidx & "'"
	dbget.Execute vQuery
	
	if (LEN(Trim(replace(company_no,"-","")))=13) then
		vQuery = "exec [db_cs].[dbo].[usp_Ten_partner_temp_info_Enc_companyno] "&tidx&",'"&company_no&"'"
		dbget.Execute vQuery
	end if

	If uid <> old_uid Then
		vQuery = " DELETE [db_partner].[dbo].[tbl_partner_temp_makerid] WHERE tidx = '" & tidx & "' " & vbCrLf
		For i = LBound(Split(uid,",")) To UBound(Split(uid,","))
			vQuery = vQuery & " INSERT INTO [db_partner].[dbo].[tbl_partner_temp_makerid](tidx, makerid) VALUES('" & tidx & "','" & Trim(Split(uid,",")(i)) & "') " & vbCrLf
		Next
		dbget.Execute vQuery
	End If
End IF


'####### ÷������ ���� #######
Dim vFileTemp, vRFileTemp, vInfo_File, vInfo_RealFile

vInfo_RealFile	= NullFillWith(Request("info_realfile"),"")
'vInfo_File = NullFillWith(Request("info_file"),"")  '2015.02.04 ������ �ѱ����� ������ �������� ���� ���ϸ� ���� ���� 2015.02.04
vInfo_File = vInfo_RealFile

If vInfo_File <> "" Then
	vQuery = ""
	If tidx <> "" Then
		vQuery = " DELETE [db_partner].[dbo].tbl_partner_temp_file WHERE tidx = '" & tidx & "' "
	End If
	vFileTemp 	= Split(vInfo_File, ",")
	vRFileTemp	= Split(vInfo_RealFile, ",")
	For i = 0 To UBOUND(vFileTemp)
		vQuery = vQuery & " INSERT INTO [db_partner].[dbo].tbl_partner_temp_file " & vbCrLf
		vQuery = vQuery & "		(file_name, real_name, tidx) " & vbCrLf
		vQuery = vQuery & "	VALUES " & vbCrLf
		vQuery = vQuery & "		('" & Trim(vFileTemp(i)) & "', '" & Trim(vRFileTemp(i)) & "', '" & tidx & "') " & vbCrLf
	Next
	dbget.execute vQuery
Else
	If requestCheckVar(Request("isfile"),1) = "o" Then
		dbget.execute " DELETE [db_partner].[dbo].tbl_partner_temp_file WHERE tidx = '" & tidx & "' "
	End If
End If


If Err.Number = 0 Then
	dbget.CommitTrans
Else
	dbget.RollBackTrans
	Response.Write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\n�Է��� ������ �ʹ� ���� �ʴ��� Ȯ�ιٶ��ϴ�.\n�ַ� ���¿� �������� ������ ���� ��Ÿ���ϴ�.')</script>"
	dbget.close()
	Response.End
End If

On Error Goto 0

If gubun = "newcompreg" Then
	If vIsSMS = "o" Then
		'Call SendNormalSMS("010-8460-0212","","[����ڵ��(�ű�)]" & chrbyte(Trim(company_name),20,"Y") & " ��û�� �����Ǿ����ϴ�.")
		'Call SendNormalSMS_LINK("010-8460-0212","","[����ڵ��(�ű�)]" & chrbyte(Trim(company_name),20,"Y") & " ��û�� �����Ǿ����ϴ�.")	'' �����
		'Call SendRadioWebHookMessage("hrkang97@10x10.co.kr","admin","SCM �˸�","����ڵ��(�ű�)",Trim(company_name),"")	' 2022.12.07 �������� ����ڿ��� ������ �־���.
	End IF
	Response.Write "<script language='javascript'>alert('����Ǿ����ϴ�.');top.opener.location.reload();top.document.location.href='/admin/member/partner/upcheinfo_new.asp?groupid="&groupid&"&gb="&gubun&"&tidx="&tidx&"';</script>"
Else
	Response.Write "<script language='javascript'>alert('����Ǿ����ϴ�.');top.opener.location.reload();top.document.location.href='/admin/member/partner/upcheinfo_edit_parent.asp?groupid="&groupid&"&gb="&gubun&"&tidx="&tidx&"';</script>"
End IF
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->