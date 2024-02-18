<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim makerid,id, mode, igroupid
makerid  = session("ssBctID")
igroupid = session("ssGroupid")
id       = requestCheckvar(request("id"),10)
mode     = requestCheckvar(request("mode"),20)

dim sqlStr, AssignedRow
if mode="confirm" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master" & vbCRLF
	sqlStr = sqlStr + " set finishflag='2'" & vbCRLF
	sqlStr = sqlStr + " where id=" + CStr(id) & vbCRLF
	sqlStr = sqlStr + " and finishflag='1'" & vbCRLF               ''업체확인대기
	sqlStr = sqlStr + " and groupid='"&igroupid&"'" & vbCRLF        ''같은GrouopCode만
	
	''sqlStr = sqlStr + " and designerid='"&makerid&"'" & vbCRLF

	dbget.execute sqlStr, AssignedRow
	
	if (AssignedRow>0) then
	    response.write "<script language='javascript'>"
        response.write " alert('확인 완료 처리 되었습니다.');"
        response.write " opener.location.reload();"
        response.write " window.close();"
        response.write "</script>"

	else
	    response.write "<script language='javascript'>"
        response.write " alert('처리중 오류가 발생하였습니다.');"
        response.write " opener.location.reload();"
        response.write " window.close();"
        response.write "</script>"
	end if
end if


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->