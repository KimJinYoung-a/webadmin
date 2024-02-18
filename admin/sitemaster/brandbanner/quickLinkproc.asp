<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/brand_banner_manageCls.asp"-->
<%
	Dim i, vAction, vQuery, vIdx, vName, vURL_PC, vURL_M, vRegUserName
	Dim vSDate, vEDate, vRegdate, vLastUserName, vLastdate, vUseYN
	Dim vQImgPC, vQImgM, vUserID, vExist
	Dim vShhmmss, vEhhmmss
	
	'### 퀵링크 기본 정보
	vIdx 		= requestCheckVar(Request("idx"),15)
	vAction	= requestCheckVar(Request("action"),10)
	vUserID	= session("ssBctId")
	vName 		= html2db(requestCheckVar(Request("quickname"),20))
	vURL_PC	= requestCheckVar(Request("url_pc"),200)
	vURL_M		= requestCheckVar(Request("url_m"),200)
	vSDate = requestCheckVar(Request("sdate"),10)
	vEDate = requestCheckVar(Request("edate"),10)
	vShhmmss = requestCheckVar(Request("shhmmss"),8)
	vEhhmmss = requestCheckVar(Request("ehhmmss"),8)
	vSDate = vSDate & " " & vShhmmss
	vEDate = vEDate & " " & vEhhmmss
	vUseYN		= requestCheckVar(Request("useyn"),1)
	vQImgPC = requestCheckVar(Request("qimgurlpc"),100)
	vQImgM = requestCheckVar(Request("qimgurlm"),100)
	
	If vAction = "" Then

		If vIdx = "" Then

			vQuery = "INSERT INTO [db_sitemaster].[dbo].[tbl_brand_link_banner]"
			vQuery = vQuery & "(name, url_pc, url_m, "
			vQuery = vQuery & " sdate, edate, qimgpc, qimgm, isusing, reguserid, lastupdateid) "
			vQuery = vQuery & " VALUES "
			vQuery = vQuery & "('" & vName & "', '" & vURL_PC & "', '" & vURL_M & "', "
			vQuery = vQuery & "'" & vSDate & "', '" & vEDate & "', '" & vQImgPC & "', '" & vQImgM & "', '" & vUseYN & "', '" & vUserID & "', '" & vUserID & "')"
			dbget.Execute vQuery
			
			vQuery = "select IDENT_CURRENT('[db_sitemaster].[dbo].[tbl_brand_link_banner]') as idx"
			rsget.CursorLocation = adUseClient
			rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
			If Not Rsget.Eof then
				vIdx = rsget("idx")
			end if
			rsget.close

		Else

			vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_brand_link_banner] SET "
			vQuery = vQuery & "name = '" & vName & "' "
			vQuery = vQuery & ", url_pc = '" & vURL_PC & "' "
			vQuery = vQuery & ", url_m = '" & vURL_M & "' "
			vQuery = vQuery & ", sdate = '" & vSDate & "' "
			vQuery = vQuery & ", edate = '" & vEDate & "' "
			vQuery = vQuery & ", qimgpc = '" & vQImgPC & "' "
			vQuery = vQuery & ", qimgm = '" & vQImgM & "' "
			vQuery = vQuery & ", isusing = '" & vUseYN & "' "
			vQuery = vQuery & ", lastupdateid = '" & vUserID & "' "
			vQuery = vQuery & ", lastupdatedate = getdate() "
			vQuery = vQuery & "where idx = '" & vIdx & "' "
			dbget.Execute vQuery

		End If
		
		Response.Write "<script>alert('처리되었습니다.');opener.location.reload();window.close();</script>"
		
	ElseIf vAction = "delete" Then
		
		vQuery = "DELETE [db_sitemaster].[dbo].[tbl_brand_link_banner] WHERE idx = '" & vIdx & "'; "
		dbget.Execute vQuery
		
    	vQuery = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    	vQuery = vQuery & "VALUES('" & session("ssBctId") & "', 'brand_quicklink', '" & vIdx & "', '0', "
    	vQuery = vQuery & "'퀵링크 idx="&vIdx&" 삭제', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vQuery)
		
		Response.Write "<script>alert('삭제되었습니다.');parent.location.reload();</script>"
		
	End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->