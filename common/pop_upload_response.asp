<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 이미지 등록
' History : 2010.06.16 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%

'==============================================================================
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim pk, filename, realfilename, imagekind

'==============================================================================
pk				= requestCheckVar(request("pk"),32)
imagekind		= requestCheckVar(request("imagekind"),32)
filename		= requestCheckVar(request("filename"),64)
realfilename	= requestCheckVar(html2db(request("realfilename")),64)

'==============================================================================
dim sqlStr, i, iid

sqlStr = ""

if (imagekind = "mobileshopimage") then
	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_user "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	mobileshopimage = '" + CStr(filename) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	userid = '" + CStr(pk) + "'	 "
elseif (imagekind = "mobilemapimage") then
	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_user "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	mobilemapimage = '" + CStr(filename) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	userid = '" + CStr(pk) + "'	 "
elseif (imagekind = "mobilegiftitemimage") then
	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_shop].[dbo].[tbl_gift_off] "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	gift_img = '" + CStr(filename) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	gift_code = " + CStr(pk) + "	 "
end if

if (sqlStr <> "") then
	dbget.Execute sqlStr
end if

%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	opener.focus();
	window.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->