<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  인트라넷 개인정보 
' History : 2007.07.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
dim userid,txName,txBirthday,txEmail1,txEmail2,txPhone ,birth_isSolar
dim txCell,txZip,txAddr1,txAddr2,txpass1,part1,part2,position,txintro

dim refer
refer = request.ServerVariables("HTTP_REFERER")	


	userid = session("ssBctId")
	txName = html2db(request("txName"))
	txBirthday = CStr(DateSerial(request("txBirthday1"),request("txBirthday2"),request("txBirthday3")))
	txEmail1 = html2db(request("txEmail1"))
	txEmail2 = html2db(request("txEmail2"))
	txPhone = html2db(request("txPhone1")) + "-" + html2db(request("txPhone2")) + "-" + html2db(request("txPhone3")) + "-" + html2db(request("txPhone4"))
	txCell = html2db(request("txCell1")) + "-" + html2db(request("txCell2")) + "-" + html2db(request("txCell3"))
	txZip = html2db(request("txZip1")) + "-" + html2db(request("txZip2"))
	txAddr1 = html2db(request("txAddr1"))
	txAddr2 = html2db(request("txAddr2"))
	txpass1 = html2db(request("txpass1"))
	part1 = html2db(request("part1"))
	part2 = html2db(request("part2"))
	position = html2db(request("position"))
	txintro = html2db(request("txintro"))
	birth_isSolar = request("birth_isSolar")
	
			
		dim sqlStr
		sqlStr = "update [db_partner].[dbo].tbl_partner"			& VbCrlf
		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), company_name='" + txName + "',"		& VbCrlf
		sqlStr = sqlStr + " email='" + txEmail1 + "',"				& VbCrlf
		sqlStr = sqlStr + " msn='" + txEmail2 + "',"				& VbCrlf
		sqlStr = sqlStr + " birthday='" + txBirthday + "',"			& VbCrlf
		sqlStr = sqlStr + " tel='" + txPhone + "',"					& VbCrlf
		sqlStr = sqlStr + " manager_hp='" + txCell + "',"			& VbCrlf
		sqlStr = sqlStr + " zipcode='" + txZip + "',"				& VbCrlf
		sqlStr = sqlStr + " address='" + txAddr1 + "',"				& VbCrlf
		sqlStr = sqlStr + " manager_address='" + txAddr2 + "',"		& VbCrlf
		sqlStr = sqlStr + " buseo='" + part1 + "',"					& VbCrlf
		sqlStr = sqlStr + " part='" + part2 + "',"					& VbCrlf
		sqlStr = sqlStr + " cposition='" + position + "',"			& VbCrlf
		sqlStr = sqlStr + " intro='" + txintro + "',"				& VbCrlf
		sqlStr = sqlStr + " birth_isSolar='" + birth_isSolar + "'"				& VbCrlf
		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"
		
		dbget.execute sqlStr
	%>
	
		<script language="javascript">
		alert('저장 되었습니다.');
		location.replace('<%= refer %>');
		</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->