<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/admin/member/partner/partnerCls.asp"-->

<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

Dim tidx, reguserid, groupid, company_name, ceoname, company_no, jungsan_gubun, company_zipcode, company_address, company_address2, company_upjong, company_uptae
Dim company_tel, company_fax, return_zipcode, return_address, return_address2, jungsan_bank, jungsan_acctno, jungsan_acctname, jungsan_date, jungsan_date_off
Dim jungsan_name, jungsan_phone, jungsan_email, jungsan_hp, psocno, sqlStr
Dim manager_name, manager_phone, manager_hp, manager_email, comment, gubun, confirmuserid, uid, old_uid, vQuery, vTempTIdx, vStatus, vExtRegUserID, i
Dim modType

tidx						= request("tidx")
reguserid					= session("ssBctId")
groupid						= request("groupid")
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
psocno 						= request("psocno")



If tidx = "" Then
	Response.Write "<script>alert('잘못된 경로입니다.');history.back();</script>"
	dbget.close()
	Response.End
End IF


On Error Resume Next
dbget.beginTrans


Dim alreadySocNoExists
If (Replace(psocno,"-","")<>Replace(company_no,"-","")) Then
    sqlStr = "select count(*) as cnt from [db_partner].[dbo].tbl_partner_group"
    sqlStr = sqlStr &" where Replace(company_no,'-','')='"&Replace(company_no,"-","")&"'"
    if (LEN(TRIM(replace(company_no,"-","")))=13) then '' 주민번호인경우 2016/08/04
        sqlStr = sqlStr &" or (replace([db_partner].[dbo].[uf_DecSOCNoPH1]([encCompNo]),'-','')='"&Replace(company_no,"-","")&"')"
    end if
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        alreadySocNoExists = rsget("cnt")>0
    rsget.CLose

	''주민번호암호화64 //2018/09/28
	if (NOT alreadySocNoExists) and (LEN(TRIM(replace(company_no,"-","")))=13) then
		sqlStr = "exec db_cs.[dbo].[usp_Ten_partner_Enc_companyno_ExistsCNT] '"&company_no&"'"
		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			alreadySocNoExists = rsget("cnt")>0
		rsget.CLose
	end if

    IF (alreadySocNoExists) then
        response.write "<script>alert('사업자 번호 변경 불가.("&company_no&") - 이미 존재하는 사업자 번호.');history.back();</script>"
        dbget.RollBackTrans
        dbget.Close() : response.end
    end if
End If

''주민번호 타입인지. 2016/08/08 수정---------------------------------------------------------------
dim bufComno
if (LEN(TRIM(replace(company_no,"-","")))=13) and (right(company_no,2)="**") then
    If groupid = "" Then
        response.write "<script>alert('사업자(주민) 번호 오류 .("&bufComno&") - 관리자 문의 요망.');history.back();</script>"
        dbget.Close() : response.end
    else
        ' sqlStr = "select isNULL([db_partner].[dbo].[uf_DecSOCNoPH1](encCompNo),'') as DecCompNo"
        ' sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_group where groupid='"&groupid&"'"
        ' rsget.Open sqlStr,dbget,1
        ' if NOT rsget.Eof then
        '     bufComno = rsget("DecCompNo")
        ' end if
        ' rsget.CLose

		''암호화방식변경.
		sqlStr = "select isNULL(db_cs.[dbo].[uf_DecCompanyNoAES256](encCompNo64),'') as DecCompNo64"
        sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_group_adddata where groupid='"&groupid&"'"
        rsget.Open sqlStr,dbget,1
        if NOT rsget.Eof then
            bufComno = rsget("DecCompNo64")
        end if
        rsget.CLose
        
        if ((bufComno="") or (LEN(TRIM(replace(bufComno,"-","")))<>13) or (right(bufComno,2)="**")) then
            response.write "<script>alert('사업자(주민) 번호 오류 .("&bufComno&") - 관리자 문의 요망.');history.back();</script>"
            dbget.Close() : response.end
        end if
        
        company_no = bufComno
    end if
end if
'' ------------------------------------------------------------------------------------------------
    
If groupid = "" Then	'####### 사업자번호 변경인 경우 #######
		modType = "Y"
		sqlStr = "select top 1 groupid from [db_partner].[dbo].tbl_partner_group"
		sqlStr = sqlStr + " order by groupid desc"
		rsget.Open sqlStr,dbget,1
			if rsget.Eof then
				groupid = 1
			else
				groupid = rsget("groupid")
				groupid = Right(groupid,5)
				groupid = CLng(groupid) +1
			end if
		rsget.Close
		groupid = "G" + Format00(5,groupid)
		
		
		sqlStr = "insert into [db_partner].[dbo].tbl_partner_group"
		sqlStr = sqlStr & " (groupid, company_name, company_no, ceoname, company_uptae, "
		sqlStr = sqlStr & " company_upjong, company_zipcode, company_address, company_address2, "
		sqlStr = sqlStr & " company_tel, company_fax, return_zipcode, return_address, return_address2, "
		sqlStr = sqlStr & " jungsan_gubun, jungsan_bank, jungsan_date, jungsan_date_off, jungsan_acctname, jungsan_acctno, "
		sqlStr = sqlStr & " manager_name, manager_phone, manager_hp, manager_email, "
		sqlStr = sqlStr & " jungsan_name, jungsan_phone, jungsan_hp, jungsan_email, "
		sqlStr = sqlStr & " encCompNo, " ''2016/08/04 추가
		sqlStr = sqlStr & " lastupdate)"
		sqlStr = sqlStr & " values('" & groupid & "'"
		sqlStr = sqlStr & " ,'" & company_name & "'"
		sqlStr = sqlStr & " ,'" & socialnoReplace(company_no) & "'"
		sqlStr = sqlStr & " ,'" & ceoname & "'"
		sqlStr = sqlStr & " ,'" & company_uptae & "'"
		sqlStr = sqlStr & " ,'" & company_upjong & "'"
		sqlStr = sqlStr & " ,'" & company_zipcode & "'"
		sqlStr = sqlStr & " ,'" & company_address & "'"
		sqlStr = sqlStr & " ,'" & company_address2 & "'"
		sqlStr = sqlStr & " ,'" & company_tel & "'"
		sqlStr = sqlStr & " ,'" & company_fax & "'"
		sqlStr = sqlStr & " ,'" & return_zipcode & "'"
		sqlStr = sqlStr & " ,'" & return_address & "'"
		sqlStr = sqlStr & " ,'" & return_address2 & "'"
		sqlStr = sqlStr & " ,'" & jungsan_gubun & "'"
		sqlStr = sqlStr & " ,'" & jungsan_bank & "'"
		sqlStr = sqlStr & " ,'" & jungsan_date & "'"
		sqlStr = sqlStr & " ,'" & jungsan_date_off & "'"
		sqlStr = sqlStr & " ,'" & jungsan_acctname & "'"
		sqlStr = sqlStr & " ,'" & jungsan_acctno & "'"
		sqlStr = sqlStr & " ,'" & manager_name & "'"
		sqlStr = sqlStr & " ,'" & manager_phone & "'"
		sqlStr = sqlStr & " ,'" & manager_hp & "'"
		sqlStr = sqlStr & " ,'" & manager_email & "'"
		sqlStr = sqlStr & " ,'" & jungsan_name & "'"
		sqlStr = sqlStr & " ,'" & jungsan_phone & "'"
		sqlStr = sqlStr & " ,'" & jungsan_hp & "'"
		sqlStr = sqlStr & " ,'" & jungsan_email & "'"
		sqlStr = sqlStr & " ,[db_partner].[dbo].[uf_EncSOCNoPH1]('" & company_no & "')" ''2016/08/04 추가
		sqlStr = sqlStr & " ,getdate()"
		sqlStr = sqlStr & " )"
		dbget.Execute sqlStr

		if (LEN(Trim(replace(company_no,"-","")))=13) then
			sqlStr = "exec [db_cs].[dbo].[usp_Ten_partner_Enc_companyno] '"&groupid&"','"&company_no&"'"
			dbget.Execute sqlStr
		end if

		For i = LBound(Split(uid,",")) To UBound(Split(uid,","))
			sqlStr = "update [db_partner].[dbo].tbl_partner set groupid = '" & groupid & "' where id = '" & Trim(Split(uid,",")(i)) & "'" & vbCrLf
			dbget.Execute sqlStr
		Next


		sqlStr = "update [db_partner].[dbo].tbl_partner" & vbCrLf
		sqlStr = sqlStr & " set company_name = '" & company_name & "'" & vbCrLf
		sqlStr = sqlStr & " ,ceoname = '" & ceoname & "'" & vbCrLf
		sqlStr = sqlStr & " ,company_no = '" & socialnoReplace(company_no) & "'" & vbCrLf
		sqlStr = sqlStr & " ,company_upjong = '" & company_upjong & "'" & vbCrLf
		sqlStr = sqlStr & " ,company_uptae = '" & company_uptae & "'" & vbCrLf
		sqlStr = sqlStr & " ,zipcode = '" & company_zipcode& "'" & vbCrLf
		sqlStr = sqlStr & " ,address = '" & company_address & "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_address = '" & company_address2 & "'" & vbCrLf
		sqlStr = sqlStr & " ,tel = '" & company_tel & "'" & vbCrLf
		sqlStr = sqlStr & " ,fax = '" & company_fax & "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_name = '" & manager_name & "'" & vbCrLf
		sqlStr = sqlStr & " ,email = '" & manager_email & "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_phone = '" & manager_phone & "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_hp = '" & manager_hp & "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_name = '" & jungsan_name& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_phone = '" & jungsan_phone& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_hp = '" & jungsan_hp& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_email = '" & jungsan_email& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_gubun = '" & jungsan_gubun& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_bank = '" & jungsan_bank& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_acctname = '" & jungsan_acctname& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_acctno = '" & jungsan_acctno& "'" & vbCrLf

		if (jungsan_date<>"") then
		    sqlStr = sqlStr & " ,jungsan_date = '" & jungsan_date& "'" & vbCrLf
	    end if

	    if (jungsan_date_off<>"") then
		    sqlStr = sqlStr & " ,jungsan_date_off = '" & jungsan_date_off& "'" & vbCrLf
		    sqlStr = sqlStr & " ,jungsan_date_frn = '" & jungsan_date_off& "'" & vbCrLf
		end if

		sqlStr = sqlStr & " where groupid = '" & groupid & "'"
		dbget.Execute sqlStr

		
Else	'####### 사업자번호 변경이 아닌 경우 #######
	modType = "N"
		sqlStr = "update [db_partner].[dbo].tbl_partner_group" & vbCrLf
		sqlStr = sqlStr & " set company_name = '" & company_name & "'" & vbCrLf
		sqlStr = sqlStr & " ,company_no = '" & socialnoReplace(company_no) & "'" & vbCrLf   
		sqlStr = sqlStr & " ,ceoname = '" & ceoname & "'" & vbCrLf
		sqlStr = sqlStr & " ,company_uptae = '" & company_uptae& "'" & vbCrLf
		sqlStr = sqlStr & " ,company_upjong = '" & company_upjong& "'" & vbCrLf
		sqlStr = sqlStr & " ,company_zipcode = '" & company_zipcode& "'" & vbCrLf
		sqlStr = sqlStr & " ,company_address = '" & company_address& "'" & vbCrLf
		sqlStr = sqlStr & " ,company_address2 = '" & company_address2& "'" & vbCrLf
		sqlStr = sqlStr & " ,company_tel = '" & company_tel& "'" & vbCrLf
		sqlStr = sqlStr & " ,company_fax = '" & company_fax& "'" & vbCrLf
		sqlStr = sqlStr & " ,return_zipcode = '" & return_zipcode & "'" & vbCrLf
		sqlStr = sqlStr & " ,return_address = '" & return_address& "'" & vbCrLf
		sqlStr = sqlStr & " ,return_address2 = '" & return_address2& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_gubun = '" & jungsan_gubun& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_bank = '" & jungsan_bank& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_acctname = '" & jungsan_acctname& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_acctno = '" & jungsan_acctno& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_date = '" & jungsan_date& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_date_off = '" & jungsan_date_off& "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_name = '" & manager_name& "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_phone = '" & manager_phone& "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_hp = '" & manager_hp& "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_email = '" & manager_email& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_name = '" & jungsan_name& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_phone = '" & jungsan_phone& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_hp = '" & jungsan_hp& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_email = '" & jungsan_email& "'" & vbCrLf
		sqlStr = sqlStr & " ,lastupdate = getdate()" & vbCrLf
		sqlStr = sqlStr & " ,encCompNo=[db_partner].[dbo].[uf_EncSOCNoPH1]('" & company_no & "')" ''2016/08/04 추가
		sqlStr = sqlStr & " where groupid = '" & groupid & "'"
		dbget.Execute sqlStr

		if (LEN(Trim(replace(company_no,"-","")))=13) then
			sqlStr = "exec [db_cs].[dbo].[usp_Ten_partner_Enc_companyno] '"&groupid&"','"&company_no&"'"
			dbget.Execute sqlStr
		end if

		For i = LBound(Split(uid,",")) To UBound(Split(uid,","))
			sqlStr = "update [db_partner].[dbo].tbl_partner set groupid = '" & groupid & "' where id = '" & Trim(Split(uid,",")(i)) & "'" & vbCrLf
			dbget.Execute sqlStr
		Next


		sqlStr = "update [db_partner].[dbo].tbl_partner" & vbCrLf
		sqlStr = sqlStr & " set company_name = '" & company_name & "'" & vbCrLf
		sqlStr = sqlStr & " ,ceoname = '" & ceoname & "'" & vbCrLf
		sqlStr = sqlStr & " ,company_no = '" & socialnoReplace(company_no) & "'" & vbCrLf
		sqlStr = sqlStr & " ,company_upjong = '" & company_upjong & "'" & vbCrLf
		sqlStr = sqlStr & " ,company_uptae = '" & company_uptae & "'" & vbCrLf
		sqlStr = sqlStr & " ,zipcode = '" & company_zipcode& "'" & vbCrLf
		sqlStr = sqlStr & " ,address = '" & company_address & "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_address = '" & company_address2 & "'" & vbCrLf
		sqlStr = sqlStr & " ,tel = '" & company_tel & "'" & vbCrLf
		sqlStr = sqlStr & " ,fax = '" & company_fax & "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_name = '" & manager_name & "'" & vbCrLf
		sqlStr = sqlStr & " ,email = '" & manager_email & "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_phone = '" & manager_phone & "'" & vbCrLf
		sqlStr = sqlStr & " ,manager_hp = '" & manager_hp & "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_name = '" & jungsan_name& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_phone = '" & jungsan_phone& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_hp = '" & jungsan_hp& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_email = '" & jungsan_email& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_gubun = '" & jungsan_gubun& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_bank = '" & jungsan_bank& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_acctname = '" & jungsan_acctname& "'" & vbCrLf
		sqlStr = sqlStr & " ,jungsan_acctno = '" & jungsan_acctno& "'" & vbCrLf

		if (jungsan_date<>"") then
		    sqlStr = sqlStr & " ,jungsan_date = '" & jungsan_date& "'" & vbCrLf
	    end if

	    if (jungsan_date_off<>"") then
		    sqlStr = sqlStr & " ,jungsan_date_off = '" & jungsan_date_off& "'" & vbCrLf
		    sqlStr = sqlStr & " ,jungsan_date_frn = '" & jungsan_date_off& "'" & vbCrLf
		end if

		sqlStr = sqlStr & " where groupid = '" & groupid & "'"
		dbget.Execute sqlStr
		
		''2016/12/14 추가 ==========================================================
		sqlStr = "update P" + VbCrlf
        sqlStr = sqlStr + " set jungsan_bank=A.jungsan_bank" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_acctname=A.jungsan_acctname" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_acctno=A.jungsan_acctno" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_date=A.jungsan_date" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_date_off=A.jungsan_date_off" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_date_frn=A.jungsan_date_off" + VbCrlf
        sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner P"
        sqlStr = sqlStr + "     Join db_partner.dbo.tbl_partner_addJungsanInfo A"
        sqlStr = sqlStr + "     on P.id=A.partnerid"
        sqlStr = sqlStr + " where P.groupid='" + groupid + "'"

        dbget.Execute sqlStr
        ''===========================================================================
        
		sqlStr = "UPDATE [db_partner].[dbo].[tbl_partner_temp_info] SET " & vbCrLf
		sqlStr = sqlStr & " status = '3', " & vbCrLf
		sqlStr = sqlStr & " confirmuserid = '" & confirmuserid & "', " & vbCrLf
		sqlStr = sqlStr & " lastupdate = getdate() " & vbCrLf
		sqlStr = sqlStr & " WHERE tidx = '" & tidx & "'"
		dbget.Execute sqlStr
End IF


vQuery = "UPDATE [db_partner].[dbo].[tbl_partner_temp_info] SET " & vbCrLf
vQuery = vQuery & "		groupid = '" & groupid & "', " & vbCrLf
vQuery = vQuery & "		company_name = '" & company_name & "', " & vbCrLf
vQuery = vQuery & "		ceoname = '" & ceoname & "', " & vbCrLf
vQuery = vQuery & "		company_no = '" & socialnoReplace(company_no) & "', " & vbCrLf
vQuery = vQuery & "		jungsan_gubun = '" & jungsan_gubun & "', " & vbCrLf
vQuery = vQuery & "		company_zipcode = '" & company_zipcode & "', " & vbCrLf
vQuery = vQuery & "		company_address = '" & company_address & "', " & vbCrLf
vQuery = vQuery & "		company_address2 = '" & company_address2 & "', " & vbCrLf
vQuery = vQuery & "		company_uptae = '" & company_uptae & "', " & vbCrLf
vQuery = vQuery & "		company_upjong = '" & company_upjong & "', " & vbCrLf
vQuery = vQuery & "		company_tel = '" & company_tel & "', " & vbCrLf
vQuery = vQuery & "		company_fax = '" & company_fax & "', " & vbCrLf
vQuery = vQuery & "		return_zipcode = '" & return_zipcode & "', " & vbCrLf
vQuery = vQuery & "		return_address = '" & return_address & "', " & vbCrLf
vQuery = vQuery & "		return_address2 = '" & return_address2 & "', " & vbCrLf
vQuery = vQuery & "		jungsan_bank = '" & jungsan_bank & "', " & vbCrLf
vQuery = vQuery & "		jungsan_acctno = '" & jungsan_acctno & "', " & vbCrLf
vQuery = vQuery & "		jungsan_acctname = '" & jungsan_acctname & "', " & vbCrLf
vQuery = vQuery & "		jungsan_date = '" & jungsan_date & "', " & vbCrLf
vQuery = vQuery & "		jungsan_date_off = '" & jungsan_date_off & "', " & vbCrLf
vQuery = vQuery & "		manager_name = '" & manager_name & "', " & vbCrLf
vQuery = vQuery & "		manager_phone = '" & manager_phone & "', " & vbCrLf
vQuery = vQuery & "		manager_hp = '" & manager_hp & "', " & vbCrLf
vQuery = vQuery & "		manager_email = '" & manager_email & "', " & vbCrLf
vQuery = vQuery & "		jungsan_name = '" & jungsan_name & "', " & vbCrLf
vQuery = vQuery & "		jungsan_phone = '" & jungsan_phone & "', " & vbCrLf
vQuery = vQuery & "		jungsan_hp = '" & jungsan_hp & "', " & vbCrLf
vQuery = vQuery & "		jungsan_email = '" & jungsan_email & "', " & vbCrLf
vQuery = vQuery & "		comment = '" & comment & "', " & vbCrLf
vQuery = vQuery & "		status = '3', " & vbCrLf
vQuery = vQuery & "		confirmuserid = '" & confirmuserid & "', " & vbCrLf
vQuery = vQuery & "		encCompNo=[db_partner].[dbo].[uf_EncSOCNoPH1]('"&company_no&"'), " & vbCrLf  ''2016/08/04 추가
vQuery = vQuery & "		lastupdate = getdate() " & vbCrLf
vQuery = vQuery & "	WHERE " & vbCrLf
vQuery = vQuery & "		tidx = '" & tidx & "'"
dbget.Execute vQuery

if (LEN(Trim(replace(company_no,"-","")))=13) then
	vQuery = "exec [db_cs].[dbo].[usp_Ten_partner_temp_info_Enc_companyno] "&tidx&",'"&company_no&"'"
	dbget.Execute vQuery
end if

If Err.Number = 0 Then
	dbget.CommitTrans
	
	'####### 신청자에게 변경완료 SMS 발송 #######
	Dim StrRSMS, vPhoneNo
	vQuery = "SELECT isNull(G.company_name,''), isNull(T.usercell,'0') FROM [db_partner].[dbo].[tbl_partner_temp_info] AS I " & vbCrLf
	vQuery = vQuery & "		LEFT JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS T ON I.reguserid = T.userid " & vbCrLf
	vQuery = vQuery & "		LEFT JOIN [db_partner].[dbo].[tbl_partner_group] AS G ON I.groupid_old = G.groupid " & vbCrLf
	vQuery = vQuery & "	WHERE I.tidx = '" & tidx & "'"
	rsget.Open vQuery,dbget,1
	If Not rsget.Eof Then
		If rsget(0) <> "" Then
			company_name = rsget(0)
		End IF
		vPhoneNo = rsget(1)
	End IF
	rsget.close()
	StrRSMS = "[업체정보변경]" & chrbyte(Trim(company_name),23,"Y") & "의 " & RequestDocumentName(gubun) & "변경이 완료되었습니다."
	if modType = "Y" then '사업자번호 변경일 경우에만 계약서 생성
		StrRSMS = StrRSMS & "새로운 계약서를 생성해서 발송해주세요"
	end if
	Call SendMultiRowsSMS(vPhoneNo,"",StrRSMS,"") 
Else
	dbget.RollBackTrans
	Response.Write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\n입력한 값들이 너무 길지 않는지 확인바랍니다.\n주로 업태와 업종에서 에러가 자주 나타납니다.')</script>"
	dbget.close()
	Response.End
End If

On Error Goto 0

Response.Write "<script language='javascript'>alert('저장되었습니다.');top.opener.location.reload();top.window.close();</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->