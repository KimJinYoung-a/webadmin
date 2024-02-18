<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sqlStr, cateCDArr, menupos, cnt, i, catecode, viewgubun
Dim yyyy, mm, tarMoneyarr, proMoneyarr, mmadd0
cateCDArr	= request("cateCDArr")
tarMoneyarr	= request("tarMoneyarr")
proMoneyarr	= request("proMoneyarr")
catecode	= request("catecode")
yyyy		= requestCheckvar(request("yyyy"),4)
mm			= requestCheckvar(request("mm"),2)
viewgubun	= requestCheckvar(request("gubun"),3)

If Len(mm) <> 2 Then
	mmadd0 = "0"&mm
Else
	mmadd0 = mm
End If

cateCDArr = split(cateCDArr,",")
cnt = ubound(cateCDArr)

If tarMoneyarr <> "" and proMoneyarr <> "" THEN
	tarMoneyarr = split(tarMoneyarr,",")
	proMoneyarr = split(proMoneyarr,",")

	If cnt <> Ubound(tarMoneyarr) OR cnt <> Ubound(proMoneyarr) Then
		response.write	"<script language='javascript'>" &_
						"	alert('입력값에 콤마(,)가 있습니다.\n확인 후 재입력 하세요');" &_
						"	history.back(-1);" &_
						"</script>"	
	End If

	For i=0 to cnt	
		If NOT isnumeric(tarMoneyarr(i)) OR NOT isnumeric(proMoneyarr(i)) Then
			response.write	"<script language='javascript'>" &_
							"	alert('입력값에 문자가 있습니다.\숫자만 입력 하세요');" &_
							"	history.back(-1);" &_
							"</script>"	
		End If

		sqlStr = ""
		sqlStr = sqlStr & " IF Exists(select * from db_partner.dbo.tbl_mdmenu_purpose where catecode='"&cateCDArr(i)&"' and yyyy = '"&yyyy&"' and mm = '"&mmadd0&"' and gubun = '"&viewgubun&"' )"
		sqlStr = sqlStr & " BEGIN"& VbCRLF
		sqlStr = sqlStr & " UPDATE R SET" & VbCRLF
		sqlStr = sqlStr & "	targetMoney='"&tarMoneyarr(i)&"', "  & VbCRLF
		sqlStr = sqlStr & "	profitMoney='"&proMoneyarr(i)&"' "  & VbCRLF
		sqlStr = sqlStr & "	FROM db_partner.dbo.tbl_mdmenu_purpose R"& VbCRLF
		sqlStr = sqlStr & " WHERE R.catecode='" & cateCDArr(i) & "' and yyyy = '"&yyyy&"' and mm = '"&mmadd0&"' and gubun = '"&viewgubun&"' "
		sqlStr = sqlStr & " END ELSE "
		sqlStr = sqlStr & " BEGIN"& VbCRLF
		sqlStr = sqlStr & " INSERT INTO db_partner.dbo.tbl_mdmenu_purpose "
	    sqlStr = sqlStr & " (catecode, yyyy, mm, gubun, targetMoney, profitMoney) "
	    sqlStr = sqlStr & " VALUES ("&cateCDArr(i)&", '" & yyyy & "', '"&mmadd0&"', '"&viewgubun&"', '"&tarMoneyarr(i)&"', '"&proMoneyarr(i)&"'  )"
		sqlStr = sqlStr & " END "
		dbget.execute sqlStr
	Next
END IF
response.write	"<script language='javascript'>" &_
				"	alert('저장되었습니다');" &_
				"	top.opener.location.reload();" &_
				"	location.replace('/admin/mdMenu/purpose/popRegPrice.asp?catecode="&catecode&"&yyyy="&yyyy&"&mm="&mm&"&gubun="&viewgubun&"');" &_
				"</script>"	
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->