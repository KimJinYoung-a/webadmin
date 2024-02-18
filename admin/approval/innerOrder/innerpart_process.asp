<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%


dim mode
dim idx, divcd, BIZSECTION_CD, scmid

mode = requestCheckvar(Request("mode"),32)

idx = requestCheckvar(Request("idx"),32)
divcd = requestCheckvar(Request("divcd"),32)
BIZSECTION_CD = requestCheckvar(Request("BIZSECTION_CD"),32)
scmid = requestCheckvar(Request("scmid"),32)

dim sqlStr
dim result

function CheckERPIdOK(erpid)
	dim sqlStr

	sqlStr = " select top 1 BIZSECTION_CD "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_partner.dbo.tbl_TMS_BA_BIZSECTION "
	sqlStr = sqlStr + " where BIZSECTION_CD = '" + CStr(erpid) + "' "
	rsget.Open sqlStr, dbget, 1
	'response.write sqlStr

	CheckERPIdOK = False
	if  not rsget.EOF  then
		CheckERPIdOK = True
	end if

	rsget.Close
end function

function CheckSCMIdOK(scmid)
	dim sqlStr

	sqlStr = " select top 1 id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_partner.dbo.tbl_partner "
	sqlStr = sqlStr + " where (id = '" + CStr(scmid) + "') or (groupid = '" + CStr(scmid) + "') "
	rsget.Open sqlStr, dbget, 1
	'response.write sqlStr

	CheckSCMIdOK = False
	if  not rsget.EOF  then
		CheckSCMIdOK = True
	end if

	rsget.Close
end function

function CheckDuplicateERP(divcd, BIZSECTION_CD)
	dim sqlStr

	sqlStr = " select top 1 BIZSECTION_CD "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_partner.dbo.tbl_InternalPart "
	sqlStr = sqlStr + " where divcd = '" + CStr(divcd) + "' and BIZSECTION_CD = '" + CStr(BIZSECTION_CD) + "' and useyn = 'Y' "
	rsget.Open sqlStr, dbget, 1
	'response.write sqlStr

	CheckDuplicateERP = False
	if  not rsget.EOF  then
		CheckDuplicateERP = True
	end if

	rsget.Close
end function

function CheckDuplicateSCM(divcd, scmid)
	dim sqlStr

	sqlStr = " select top 1 scmid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_partner.dbo.tbl_InternalPart "
	sqlStr = sqlStr + " where divcd = '" + CStr(divcd) + "' and scmid = '" + CStr(scmid) + "' and useyn = 'Y' "
	rsget.Open sqlStr, dbget, 1
	'response.write sqlStr

	CheckDuplicateSCM = False
	if  not rsget.EOF  then
		CheckDuplicateSCM = True
	end if

	rsget.Close
end function

if (mode = "regnewpart") then

	result = CheckERPIdOK(BIZSECTION_CD)
	if (result = False) then
	    response.write "<script>alert('잘못된 ERP부서코드입니다.(" + CStr(BIZSECTION_CD) + ")'); history.back();</script>"
	    dbget.close()	:	response.End
	end if

	result = CheckSCMIdOK(scmid)
	if (result = False) then
	    response.write "<script>alert('잘못된 SCM부서코드입니다.(" + CStr(scmid) + ")'); history.back();</script>"
	    dbget.close()	:	response.End
	end if

	result = CheckDuplicateERP(divcd, BIZSECTION_CD)
	if (result = True) then
	    response.write "<script>alert('중복된 ERP부서코드입니다.(" + CStr(BIZSECTION_CD) + ")'); history.back();</script>"
	    dbget.close()	:	response.End
	end if

	result = CheckDuplicateSCM(divcd, scmid)
	if (result = True) then
	    response.write "<script>alert('중복된 SCM부서코드입니다.(" + CStr(scmid) + ")'); history.back();</script>"
	    dbget.close()	:	response.End
	end if



	sqlStr = " insert into db_partner.dbo.tbl_InternalPart(divcd, BIZSECTION_CD, scmid) values('" + CStr(divcd) + "', '" + CStr(BIZSECTION_CD) + "', '" + CStr(scmid) + "') "
	'response.write sqlStr
	dbget.Execute sqlStr

    response.write "<script>alert('생성 되었습니다.');</script>"
    response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
    dbget.close()	:	response.End

elseif (mode = "delpart") then

	sqlStr = " update db_partner.dbo.tbl_InternalPart set useyn = 'N' where idx = " + CStr(idx) + " "
	'response.write sqlStr
	dbget.Execute sqlStr

    response.write "<script>alert('삭제 되었습니다.');</script>"
    response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
    dbget.close()	:	response.End

else
	'// 에러
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
