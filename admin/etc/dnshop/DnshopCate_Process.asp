<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->

<%
dim ten_code, tecdl, tecdm, tecdn
dim mngcate, dispcate, storcate, sqlStr, ecate, rcate, spkey, secate
ten_code = request("ten_code")
mngcate = request("mngcate")
dispcate = request("dispcate")
storcate = request("storcate")
ecate = request("ecate")
rcate = request("rcate")
spkey = request("spkey")
secate = request("secate")


If ten_code <> "" Then
	Dim i, vTmpCode, vTmpL, vTmpM, vTmpN, vTmpMng, vTmpDisp, vTmpStor, vTmpEcate, vTmpRcate, vTmpSpkey, vTmpSeCate
	vTmpCode 	= Split(ten_code, ",")
	vTmpMng		= Split(mngcate, ",")
	vTmpDisp	= Split(dispcate, ",")
	vTmpStor	= Split(storcate, ",")
	vTmpEcate	= Split(ecate, ",")
	vTmpRcate	= Split(rcate, ",")
	vTmpSpkey	= Split(spkey, ",")
	vTmpSeCate	= Split(secate, ",")
	
	For i = 0 To UBOUND(vTmpCode)
		sqlStr = ""
		tecdl = ""
		tecdm = ""
		tecdn = ""
		
		tecdl = Trim(Split(Trim(vTmpCode(i)),"|")(0))
		tecdm = Trim(Split(Trim(vTmpCode(i)),"|")(1))
		tecdn = Trim(Split(Trim(vTmpCode(i)),"|")(2))
		
		sqlStr = " IF EXISTS (SELECT tencdl FROM [db_item].[dbo].tbl_dnshop_dspcategory_mapping WHERE tencdl = '" & tecdl & "' AND tencdm = '" & tecdm & "' AND tencdn = '" & tecdn & "') " & _
				 "		BEGIN " & _
				 " 			UPDATE [db_item].[dbo].tbl_dnshop_dspcategory_mapping " & _
				 "				SET dnshopdispcategory = '" & Trim(vTmpDisp(i)) & "', dnshopstorecategory = '" & Trim(vTmpStor(i)) & "' " & _
				 "					, dnshopEcategory = '" & Trim(vTmpEcate(i)) & "', dnshopRcategory = '" & Trim(vTmpRcate(i)) & "', dnshopSpkey = '" & Trim(vTmpSpkey(i)) & "' " & _
				 "					, dnshopSeCategory = '" & Trim(vTmpSeCate(i)) & "' " & _
				 "			WHERE " & _
				 "				tencdl = '" & tecdl & "' AND tencdm = '" & tecdm & "' AND tencdn = '" & tecdn & "' " & _
				 "		END	" & _
				 " ELSE	" & _
				 "		BEGIN " & _
				 "			INSERT INTO [db_item].[dbo].tbl_dnshop_dspcategory_mapping(tencdl, tencdm, tencdn, dnshopdispcategory, dnshopstorecategory, dnshopEcategory, dnshopRcategory, dnshopSpkey, dnshopSeCategory) " & _
				 "			VALUES('" & tecdl & "', '" & tecdm & "', '" & tecdn & "', '" & Trim(vTmpDisp(i)) & "', '" & Trim(vTmpStor(i)) & "', '" & Trim(vTmpEcate(i)) & "', '" & Trim(vTmpRcate(i)) & "', '" & Trim(vTmpSpkey(i)) & "', '" & Trim(vTmpSeCate(i)) & "') " & _
				 "		END	" & _
				 "		  " & _
				 " IF EXISTS (SELECT tencdl FROM [db_item].[dbo].tbl_dnshop_mngcategory_mapping WHERE tencdl = '" & tecdl & "' AND tencdm = '" & tecdm & "') " & _
				 "		BEGIN " & _
				 " 			UPDATE [db_item].[dbo].tbl_dnshop_mngcategory_mapping " & _
				 "				SET dnshopmngcategory = '" & Trim(vTmpMng(i)) & "' " & _
				 "			WHERE " & _
				 "				tencdl = '" & tecdl & "' AND tencdm = '" & tecdm & "' " & _
				 "		END	" & _
				 " ELSE	" & _
				 "		BEGIN " & _
				 "			INSERT INTO [db_item].[dbo].tbl_dnshop_mngcategory_mapping(tencdl, tencdm, dnshopmngcategory) " & _
				 "			VALUES('" & tecdl & "', '" & tecdm & "', '" & Trim(vTmpMng(i)) & "') " & _
				 "		END	"
		dbget.Execute sqlStr
	Next
End If

%>
<script language='javascript'>
alert('저장되었습니다.');
location.href = "DnshopCategory.asp";
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->