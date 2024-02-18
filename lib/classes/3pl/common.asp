<%

Sub SelectBoxCompanyID(selectedname, selectedId, useyn)
	dim tmp_str,query1

	response.write "<select class='select' name='" & selectedname & "'>"
	response.write "<option value=''>선택</option>"
	query1 = " select * from [db_threepl].[dbo].[tbl_company] c"
	query1 = query1 & " where 1 = 1 "
	if (useyn <> "") then
		query1 = query1 & " and c.useyn='" + Cstr(useyn) + "' "
	end if
	query1 = query1 & " order by c.indt desc"
	rsget_TPL.Open query1,dbget_TPL,1

	if  not rsget_TPL.EOF  then
		rsget_TPL.Movefirst

		do until rsget_TPL.EOF
			tmp_str = ""
			if Lcase(selectedId) = Lcase(rsget_TPL("companyid")) then
				tmp_str = " selected"
			end if
			response.write("<option value='"&rsget_TPL("companyid")&"' "&tmp_str&">"&db2html(rsget_TPL("company_name"))&"</option>")

			rsget_TPL.MoveNext
		loop
	end if
	rsget_TPL.close

	response.write "</select>"
End Sub

Sub SelectBoxPartnerCompanyID(selectedname, selectedId, useyn)
	dim tmp_str,query1

	response.write "<select class='select' name='" & selectedname & "'>"
	response.write "<option value=''>선택</option>"
	query1 = " select * from [db_threepl].[dbo].[tbl_partnerinfo] c"
	query1 = query1 & " where 1 = 1 "
	if (useyn <> "") then
		query1 = query1 & " and c.useyn='" + Cstr(useyn) + "' "
	end if
	query1 = query1 & " order by c.regdate desc"
	rsget_TPL.Open query1,dbget_TPL,1

	if  not rsget_TPL.EOF  then
		rsget_TPL.Movefirst

		do until rsget_TPL.EOF
			tmp_str = ""
			if Lcase(selectedId) = Lcase(rsget_TPL("partnercompanyid")) then
				tmp_str = " selected"
			end if
			response.write("<option value='"&rsget_TPL("partnercompanyid")&"' "&tmp_str&">"&db2html(rsget_TPL("partnercompanyname"))&"</option>")

			rsget_TPL.MoveNext
		loop
	end if
	rsget_TPL.close

	response.write "</select>"
End Sub

function CreateBrandOne(companyid)
	dim query1, brandSeq, Cmd, ScopeID

	brandSeq = ""

	query1 = " exec [db_threepl].[dbo].[usp_create_brand_one] '" & companyid & "' "
	rsget_TPL.Open query1,dbget_TPL,1
	if  not rsget_TPL.EOF  then
		brandSeq = rsget_TPL("brandSeq")
	end if
	rsget_TPL.close

	CreateBrandOne = brandSeq
end function

Sub SelectBoxBrandID(companyid, selectedname, selectedId, useyn)
	dim tmp_str,query1

	response.write "<select class='select' name='" & selectedname & "'>"
	response.write "<option value=''>선택</option>"
	query1 = " select * from [db_threepl].[dbo].[tbl_brand] b"
	query1 = query1 & " where 1 = 1 "
	if (companyid <> "") then
		query1 = query1 & " and b.companyid='" + Cstr(companyid) + "' "
	end if
	if (useyn <> "") then
		query1 = query1 & " and b.useyn='" + Cstr(useyn) + "' "
	end if
	query1 = query1 & " order by b.regdate desc"
	rsget_TPL.Open query1,dbget_TPL,1

	if  not rsget_TPL.EOF  then
		rsget_TPL.Movefirst

		do until rsget_TPL.EOF
			tmp_str = ""
			if Lcase(selectedId) = Lcase(rsget_TPL("brandid")) then
				tmp_str = " selected"
			end if
			response.write("<option value='"&rsget_TPL("brandid")&"' "&tmp_str&">"&db2html(rsget_TPL("brand_name"))&"</option>")

			rsget_TPL.MoveNext
		loop
	end if
	rsget_TPL.close

	response.write "</select>"
End Sub

Sub SelectBoxEngBrandID(companyid, selectedname, selectedId, useyn)
	dim tmp_str,query1

	response.write "<select class='select' name='" & selectedname & "'>"
	response.write "<option value=''>선택</option>"
	query1 = " select * from [db_threepl].[dbo].[tbl_brand] b"
	query1 = query1 & " where 1 = 1 "
	if (companyid <> "") then
		query1 = query1 & " and b.companyid='" + Cstr(companyid) + "' "
	end if
	if (useyn <> "") then
		query1 = query1 & " and b.useyn='" + Cstr(useyn) + "' "
	end if
	query1 = query1 & " order by b.regdate desc"
	rsget_TPL.Open query1,dbget_TPL,1

	if  not rsget_TPL.EOF  then
		rsget_TPL.Movefirst

		do until rsget_TPL.EOF
			tmp_str = ""
			if Lcase(selectedId) = Lcase(rsget_TPL("brandid")) then
				tmp_str = " selected"
			end if
			response.write("<option value='"&rsget_TPL("brandid")&"' "&tmp_str&">"&db2html(rsget_TPL("brand_name_eng"))&"</option>")

			rsget_TPL.MoveNext
		loop
	end if
	rsget_TPL.close

	response.write "</select>"
End Sub

function FormatPrdCode(prdcode)
	if (Len(prdcode) = 12) then
		FormatPrdCode = Left(prdcode, 2) & "-" & Mid(prdcode, 3, 6) & "-" & Right(prdcode, 4)
	else
		FormatPrdCode = prdcode
	end if
end function

function CheckCompanyItemidExists(companyid, itemgubun, itemid, itemoption, itemoptionname)
	dim query1

	query1 = " exec [db_threepl].[dbo].[usp_check_itemid_exist] '" & companyid & "', '" & itemgubun & "', '" & itemid & "', '" & itemoption & "', '" & itemoptionname & "' "
	rsget_TPL.Open query1,dbget_TPL,1
	if  not rsget_TPL.EOF  then
		CheckCompanyItemidExists = (rsget_TPL("result") = "Y")
	end if
	rsget_TPL.close
end function

function CreatePrdcodeOne(companyid, itemgubun, itemid, itemoption, itemoptionname)
	dim query1

	query1 = " exec [db_threepl].[dbo].[usp_create_prdcode_one] '" & companyid & "', '" & itemgubun & "', '" & itemid & "', '" & itemoption & "', '" & itemoptionname & "' "
	rsget_TPL.Open query1,dbget_TPL,1
	if  not rsget_TPL.EOF  then
		CreatePrdcodeOne = rsget_TPL("prdcode")
	end if
	rsget_TPL.close
end function

Sub SelectBoxApiInput(companyid, selectedname, selectedId, useyn)
	dim tmp_str,query1

	response.write "<select class='select' name='" & selectedname & "'>"
	response.write "<option value=''>선택</option>"
	query1 = " select * from [db_threepl].[dbo].[tbl_partnercompany] p"
	query1 = query1 & " where 1 = 1 "
	query1 = query1 & " and apiAvail = 'Y' "
	if (companyid <> "") then
		query1 = query1 & " and p.companyid='" + Cstr(companyid) + "' "
	end if
	if (useyn <> "") then
		query1 = query1 & " and p.useyn='" + Cstr(useyn) + "' "
	end if
	query1 = query1 & " order by p.regdate desc"
	rsget_TPL.Open query1,dbget_TPL,1

	if  not rsget_TPL.EOF  then
		rsget_TPL.Movefirst

		do until rsget_TPL.EOF
			tmp_str = ""
			if Lcase(selectedId) = Lcase(rsget_TPL("companyid")&","&rsget_TPL("partnercompanyid")) then
				tmp_str = " selected"
			end if
			response.write("<option value='"& (rsget_TPL("companyid")&","&rsget_TPL("partnercompanyid")) &"' "&tmp_str&">"& (db2html(rsget_TPL("companyid")) & " - " & db2html(rsget_TPL("partnercompanyname"))) &"</option>")

			rsget_TPL.MoveNext
		loop
	end if
	rsget_TPL.close

	response.write "</select>"
End Sub

%>
