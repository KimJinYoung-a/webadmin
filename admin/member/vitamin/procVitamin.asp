<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sMode,adminid
Dim strSql 
dim arrempno,intLoop,arrmidx, idx, totvm
dim menupos

 
sMode =requestCheckvar(request("hidM"),1)
menupos =requestCheckvar(request("menupos"),10)
adminid	= session("ssBctId")
if sMode ="A" THEN '전체 등록(년단위)
	strSql = " exec [db_partner].[dbo].[usp_Ten_user_tenbyten_InsertAllYearVitamin] '"& adminid& "' " 
	dbget.Execute(strSql)
	response.write "<script>location.href='/admin/member/vitamin/?menupos="&menupos&"';alert('등록되었습니다.');</script>"
response.end
elseif sMode="I" THEN '미등록자 등록(월단위)
		arrempno = requestCheckvar(request("chki"),500)   
		arrempno = split(arrempno,",")
		 
		For intLoop = 0 To ubound(arrempno)
			strSql = " exec [db_partner].[dbo].[usp_ten_user_tenbyten_InsertMonthVitamin] '"&Trim(arrempno(intLoop))&"' ,'"& adminid& "' " 
		  dbget.Execute(strSql)
		Next			
		response.write "<script>self.close();opener.location.href='/admin/member/vitamin/?menupos="&menupos&"';alert('등록되었습니다.');</script>"
	response.end
elseif sMode ="P"	 THEN '인사지급처리
	arrmidx = requestCheckvar(request("chki"),500)   
		arrmidx = split(arrmidx,",")
		 
		For intLoop = 0 To ubound(arrmidx)
			strSql = " exec [db_partner].[dbo].[sp_ten_Vitamin_UpdateState] '"&Trim(arrmidx(intLoop))&"' ,'"& adminid& "' " 
		  dbget.Execute(strSql)
		Next			
		response.write "<script>location.href='/admin/member/vitamin/detailVMList.asp?menupos="&menupos&"';alert('등록되었습니다.');</script>"
	response.end
elseif sMode ="U"	 THEN '비타민수정
	idx = requestCheckvar(request("idx"),8)
	totvm = requestCheckvar(request("totvm"),12)
	if isNumeric(idx) and isNumeric(totvm) then
		strSql = "UPDATE db_partner.dbo.tbl_vitamin_master " 
		strSql = strSql & " SET totvm= " & totvm
		strSql = strSql & " 	,adminid = '" & adminid & "'"
		strSql = strSql & " 	,lastupdate = getdate()"
		strSql = strSql & " WHERE midx=" & idx
		dbget.Execute(strSql)
	end if
	response.write "<script>self.close();opener.location.reload();alert('수정되었습니다.');</script>"
	response.end
elseif sMode ="D"	 THEN '선택삭제
	arrmidx = requestCheckvar(request("chki"),500)    
		arrmidx = split(arrmidx,",")
		
		For intLoop = 0 To ubound(arrmidx)
			strSql = " exec [db_partner].[dbo].[sp_ten_Vitamin_Delete] '"&Trim(arrmidx(intLoop))&"' ,'"& adminid& "' " 
		  dbget.Execute(strSql)
		Next			
		response.write "<script>location.href='/admin/member/vitamin/detailVMList.asp?menupos="&menupos&"';alert('등록되었습니다.');</script>"
	response.end	
END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->