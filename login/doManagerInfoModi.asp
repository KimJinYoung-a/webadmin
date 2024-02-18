<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<%  
	'로그인 확인
	if session("ssBctId")="" or isNull(session("ssBctId")) then
		Call Alert_Return("잘못된 접속입니다.")
		dbget.close()	:	response.End
	end if
	
	if session("ssGroupid")="" or isNull(session("ssGroupid")) then
		Call Alert_Return("잘못된 접속입니다.2")
		dbget.close()	:	response.End
	end if

	'// 변수 선언 및 전송값 접수
	dim manager_name, manager_phone, manager_email, manager_hp,jungsan_name,jungsan_phone,jungsan_email,jungsan_hp
	dim sqlStr
	
	manager_name 	= requestCheckVar(request("manager_name"),32)
	manager_phone 	= requestCheckVar(request("manager_phone"),16)
	manager_email 	= requestCheckVar(request("manager_email"),64)
	manager_hp 		= requestCheckVar(request("manager_hp"),16)
	jungsan_name 	= requestCheckVar(request("jungsan_name"),32) 
	jungsan_phone 	= requestCheckVar(request("jungsan_phone"),16)
	jungsan_email 	= requestCheckVar(request("jungsan_email"),64)
	jungsan_hp 		= requestCheckVar(request("jungsan_hp"),16)
 
 dbget.beginTrans
 	sqlStr = "update [db_partner].[dbo].tbl_partner_group" + VbCrlf	
 	sqlStr = sqlStr + " set manager_name='" + manager_name+ "'" + VbCrlf
	sqlStr = sqlStr + " ,manager_phone='" + manager_phone+ "'" + VbCrlf
	sqlStr = sqlStr + " ,manager_hp='" + manager_hp+ "'" + VbCrlf
	sqlStr = sqlStr + " ,manager_email='" + manager_email+ "'" + VbCrlf
	sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
	sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
	sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
	sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
	sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
	sqlStr = sqlStr + " where groupid='" + session("ssGroupid") + "'"
	dbget.Execute sqlStr
	
	IF Err.Number <> 0 THEN
	dbget.RollBackTrans
		Call Alert_return ("데이터처리에 문제가 발생했습니다. 담당자에게 문의주세요") 
			dbget.close()	:	response.End
	END IF
	
	sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
	sqlStr = sqlStr + " set  manager_name='" + manager_name + "'" + VbCrlf
	sqlStr = sqlStr + " ,email='" + manager_email + "'" + VbCrlf
	sqlStr = sqlStr + " ,manager_phone='" + manager_phone + "'" + VbCrlf
	sqlStr = sqlStr + " ,manager_hp='" + manager_hp + "'" + VbCrlf
	sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
	sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
	sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
	sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
	sqlStr = sqlStr + " ,lastInfoChgDT = getdate()" + VbCrlf
	sqlStr = sqlStr + " where groupid='" +  session("ssGroupid") + "'"
 	dbget.Execute sqlStr
 	IF Err.Number <> 0 THEN
	dbget.RollBackTrans
			Call Alert_return ("데이터처리에 문제가 발생했습니다. 담당자에게 문의주세요") 
			dbget.close()	:	response.End
	END IF
	dbget.CommitTrans
	
	dim cuseridv
    sqlStr = "select top 1 IsNULL(userdiv,'') as userdiv "
    sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c"
    sqlStr = sqlStr + " where userid = '" + session("ssBctId") + "'" + vbCrlf
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
    	cuseridv = rsget("userdiv")
    end if
    rsget.close
    
    ''강사 임시
	if (cuseridv="14") then
	    session("ssUserCDiv")=cuseridv  ''2016/08/11 추가
		response.write "<script language='javascript'>alert('저장되었습니다.');top.location.replace('" & manageUrl & "/lectureadmin/index.asp')</script>"
    	dbget.close()	:	response.End
	end if
	
	response.write "<script language='javascript'>alert('저장되었습니다.');top.location.replace('" & manageUrl & "/partner/index.asp')</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->