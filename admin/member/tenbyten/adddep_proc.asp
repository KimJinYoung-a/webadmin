<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 추가부서등록
' Hieditor : 2017.08.23 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->

<%
dim empno ,department_id , mode , sql ,tmpvalue , tmpcnt
	empno =  requestcheckvar(request("sEPN"),32)
	department_id =requestcheckvar(request("department_id"),10)
	mode = requestcheckvar(request("mode")	,10)

dim ref
	ref = request.ServerVariables("HTTP_REFERER")

tmpcnt = 0
	
 
'//부서추가
if mode = "A" then
	 
	sql = "select sum(cnt) as sumcnt"
	sql = sql & " from ("
	sql = sql & " select count(*) as cnt "
	sql = sql & " from db_partner.[dbo].tbl_user_tenbyten "
	sql = sql & " where department_id = '"&department_id&"' and empno = '"&empno&"'" 
	sql = sql & " union all "
	sql = sql & " select count(*) as cnt "
	sql = sql & " from db_partner.[dbo].tbl_partner_addDepartment "
	sql = sql & " where departmentid = '"&department_id&"' and empno = '"&empno&"' and isusing=1 " 
	sql = sql & ") as t "
	 
	rsget.Open sql,dbget,1
	
	if not rsget.EOF  then        
		tmpvalue = rsget("sumcnt") > 0
	end if
	
	rsget.close
	
	if tmpvalue then		
		response.write "<script language='javascript'>"
		response.write " 	alert('동일한 부서에 이미 권한이 있습니다. 다른 부서를 선택해주세요');"
		response.write "	location.href='adddep_reg.asp?sEPN="&empno&"';"
		response.write "</script>"
		response.end
	end if

		sql = "IF(EXISTS(Select depidx from db_partner.dbo.tbl_partner_addDepartment where empno='"&empno&"' and departmentid='"&department_id&"')) " + vbcrlf
		sql = sql & "BEGIN " + vbcrlf
		sql = sql & "	UPDATE db_partner.dbo.tbl_partner_addDepartment SET isusing=1, regdate=getdate() where empno='"&empno&"' and departmentid='"&department_id&"' " + vbcrlf
		sql = sql & "END " + vbcrlf
		sql = sql & "ELSE " + vbcrlf
		sql = sql & "BEGIN " + vbcrlf
		sql = sql & "	INSERT INTO db_partner.dbo.tbl_partner_addDepartment (empno,userid,departmentid) " + vbcrlf
		sql = sql & " 	SELECT empno, userid, '"&department_id&"'"
		sql = sql & "   FROM db_partner.dbo.tbl_user_tenbyten where empno ='"&empno&"' "
		sql = sql & "END " + vbcrlf
		dbget.execute sql
		
		response.write "<script language='javascript'>"
		response.write " 	alert('OK');"
		response.write " 	opener.location.reload();"
		response.write "	location.href='adddep_reg.asp?sEPN="&empno&"';"
		response.write "</script>"
		response.end
		
	 

'//삭제
elseif mode = "D" then 
		dim depid
		depid = requestcheckvar(request("depid"),10)
		 
		sql = " update  db_partner.[dbo].[tbl_partner_addDepartment]  set  isusing =0 " & vbcrlf
		sql = sql & " where empno ='"&empno&"' and departmentid ='"&depid&"'"  
		dbget.execute sql
		
		response.write "<script language='javascript'>"
		response.write " 	alert('OK');"
		response.write " 	opener.location.reload();"
		response.write "	location.href='adddep_reg.asp?sEPN="&empno&"';"
		response.write "</script>"
		response.end
			
 
end if
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->