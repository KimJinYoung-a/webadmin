<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 가맹점 정산관리
' History : 2009.04.07 서동석 생성
'			2010.05.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim mode
dim idx, shopid, cardPrice, startDate, endDate, rateGubun, rateAmmount, isusing, regdate
dim sqlStr

mode    	= requestCheckVar(request("mode"),32)
idx    		= requestCheckVar(request("idx"),32)
shopid    	= requestCheckVar(request("shopid"),32)
cardPrice   = requestCheckVar(request("cardPrice"),32)
startDate   = requestCheckVar(request("startDate"),32)
endDate    	= requestCheckVar(request("endDate"),32)
rateGubun   = requestCheckVar(request("rateGubun"),32)
rateAmmount = requestCheckVar(request("rateAmmount"),32)
isusing    	= requestCheckVar(request("isusing"),32)


select case mode
	case "ins"
		sqlStr = " insert into [db_shop].[dbo].[tbl_shop_card_promotion](shopid, cardPrice, startDate, endDate, rateGubun, rateAmmount, isusing, regdate)" & vbCrLf
		sqlStr = sqlStr + " values('" & shopid & "', " & cardPrice & ", '" & startDate & "', '" & endDate & "', " & rateGubun & ", " & rateAmmount & ", '" & isusing & "', getdate()) " & vbCrLf
		dbget.execute sqlStr
	case "modi"
		response.write "modi"
		sqlStr = " update [db_shop].[dbo].[tbl_shop_card_promotion] " & vbCrLf
		sqlStr = sqlStr + " set " & vbCrLf
		sqlStr = sqlStr + " 	shopid = '" & shopid & "' " & vbCrLf
		sqlStr = sqlStr + " 	, cardPrice = " & cardPrice & " " & vbCrLf
		sqlStr = sqlStr + " 	, startDate = '" & startDate & "' " & vbCrLf
		sqlStr = sqlStr + " 	, endDate = '" & endDate & "' " & vbCrLf
		sqlStr = sqlStr + " 	, rateGubun = " & rateGubun & " " & vbCrLf
		sqlStr = sqlStr + " 	, rateAmmount = " & rateAmmount & " " & vbCrLf
		sqlStr = sqlStr + " 	, isusing = '" & isusing & "' " & vbCrLf
		sqlStr = sqlStr + " where " & vbCrLf
		sqlStr = sqlStr + " 	idx = " & idx
		dbget.execute sqlStr
	case default
		response.write "aaaaaaaaaaa"
end select

%>
<script type='text/javascript'>
<% if (mode = "ins") then %>
alert('저장되었습니다.');
opener.document.frm.page.value = "1";
opener.document.frm.submit();
opener.focus();
window.close();
<% elseif (mode = "modi") then %>
alert('저장되었습니다.');
opener.location.reload();
opener.focus();
window.close();
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
