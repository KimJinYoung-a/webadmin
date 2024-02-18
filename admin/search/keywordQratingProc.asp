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
<!-- #include virtual="/lib/classes/search/search_manageCls.asp"-->
<%
	Dim i, vAction, vQuery, vIdx, vTitle, vViewGubun, vRegUserName, vSDate, vEDate, vRegdate, vLastUserName, vLastdate, vMemo, vUseYN, vKwArr
	Dim vKeyword, vKeywordInDB, vUserID, vExist, vUnit, vUnitCount, vUnitInDB, vTmpGB, vTmpIDX, vTmpSN, vUnitGubun, vUnitContentsIdx
	Dim vNext, vShhmmss, vEhhmmss, vBGClass
	
	'### 퀵링크 기본 정보
	vIdx 		= requestCheckVar(Request("idx"),15)
	vAction	= requestCheckVar(Request("action"),13)
	vUserID	= session("ssBctId")
	vTitle		= html2db(requestCheckVar(Request("title"),50))
	vViewGubun	= requestCheckVar(Request("viewgubun"),6)
	If vViewGubun = "always" Then	'### 상시노출
		vSDate = ""
		vEDate = ""
	Else	'### 기간설정
		vSDate = requestCheckVar(Request("sdate"),10)
		vEDate = requestCheckVar(Request("edate"),10)
		vShhmmss = requestCheckVar(Request("shhmmss"),8)
		vEhhmmss = requestCheckVar(Request("ehhmmss"),8)
		
		vSDate = vSDate & " " & vShhmmss
		vEDate = vEDate & " " & vEhhmmss
	End If
	vUseYN		= requestCheckVar(Request("useyn"),1)
	vBGClass	= requestCheckVar(Request("bgclass"),10)
	vMemo 		= html2db(Request("memo"))
	vUnit		= requestCheckVar(Request("unit"),100)
	vUnitCount	= requestCheckVar(Request("unitount"),2)
	vUnitInDB	= requestCheckVar(Request("unit_in_db"),100)

	
	'### 검색키워드
	vKeyword = requestCheckVar(Request("keyword"),200)
	vKeywordInDB = requestCheckVar(Request("keyword_in_db"),200)


	vUnitGubun = requestCheckVar(Request("unitgubun"),10)
	vUnitContentsIdx = requestCheckVar(Request("unitcontentsidx"),10)


	If vAction = "" Then
		
		vExist = fnKeywordExistCheck(vKeyword,"c",vIdx)
		If vExist = "1" Then
			Response.Write "<script>alert('동일 글 내 검색키워드에 같은 키워드가 2개 이상 있습니다.\n확인해보시기 바랍니다.');</script>"
			dbget.close
			Response.End
		ElseIf vExist = "2" Then
			Response.Write "<script>alert('전체 검색키워드에 같은 키워드가 있습니다.\n확인해보시기 바랍니다.');</script>"
			dbget.close
			Response.End
		End If
		
		
		If vIdx = "" Then
			
			vNext = "unitreg"

			vQuery = "INSERT INTO [db_sitemaster].[dbo].[tbl_search_curator]"
			vQuery = vQuery & "(title, viewgubun, sdate, edate, useyn, bgclass, memo, "
			vQuery = vQuery & "reguserid, lastupdateid) "
			vQuery = vQuery & " VALUES "
			vQuery = vQuery & "('" & vTitle & "', '" & vViewGubun & "', '" & vSDate & "', '" & vEDate & "', '" & vUseYN & "', '" & vBGClass & "', '" & vMemo & "', "
			vQuery = vQuery & "'" & vUserID & "', '" & vUserID & "')"
			dbget.Execute vQuery
			
			vQuery = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_search_curator') as idx"
			rsget.CursorLocation = adUseClient
			rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
			If Not Rsget.Eof then
				vIdx = rsget("idx")
			end if
			rsget.close

		Else

			vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_search_curator] SET "
			vQuery = vQuery & "title = '" & vTitle & "' "
			vQuery = vQuery & ", viewgubun = '" & vViewGubun & "' "
			vQuery = vQuery & ", sdate = '" & vSDate & "' "
			vQuery = vQuery & ", edate = '" & vEDate & "' "
			vQuery = vQuery & ", useyn = '" & vUseYN & "' "
			vQuery = vQuery & ", bgclass = '" & vBGClass & "' "
			vQuery = vQuery & ", memo = '" & vMemo & "' "
			vQuery = vQuery & ", lastupdateid = '" & vUserID & "' "
			vQuery = vQuery & ", lastupdatedate = getdate() "
			vQuery = vQuery & "where idx = '" & vIdx & "' "
			dbget.Execute vQuery

		End If

		
		If vKeyword <> vKeywordInDB Then	'### 기존과 다를때만 업데이트.
			vQuery = "DELETE [db_sitemaster].[dbo].[tbl_search_keyword] WHERE topidx = '" & vIdx & "' and topgubun = 'c'; "
			For i = LBound(Split(vKeyword,",")) To UBound(Split(vKeyword,","))
				vQuery = vQuery & "INSERT INTO [db_sitemaster].[dbo].[tbl_search_keyword](topidx, topgubun, keyword) "
				vQuery = vQuery & "VALUES('" & vIdx & "', 'c', '" & Trim(Split(vKeyword,",")(i)) & "');"
			Next
			dbget.Execute vQuery
		End If
		
		
		If vUnit <> vUnitInDB Then	'### 기존과 다를때만 업데이트.
			vQuery = "DELETE [db_sitemaster].[dbo].[tbl_search_curator_unit] WHERE topidx = '" & vIdx & "'; "
			For i = LBound(Split(vUnit,",")) To UBound(Split(vUnit,","))
			
				'Split(Split(vUnit,",")(i),"$")() 0:gubun, 1: contentsidx   ex) event$67890
				vTmpGB = Trim(Split(Split(vUnit,",")(i),"$")(0))
				vTmpIDX = Trim(Split(Split(vUnit,",")(i),"$")(1))
			
				vQuery = vQuery & "INSERT INTO [db_sitemaster].[dbo].[tbl_search_curator_unit](topidx, gubun, contentsidx, sortno) "
				vQuery = vQuery & "VALUES('" & vIdx & "', '" & vTmpGB & "', '" & vTmpIDX & "', '" & (i+1) & "');"
			Next
			dbget.Execute vQuery
		End If
		

		If vNext = "unitreg" Then
			Response.Write "<script>parent.location.href='keywordQratingManage.asp?idx="&vIdx&"';</script>"
		Else
			Response.Write "<script>alert('처리되었습니다.');parent.location.href='keywordQratingManageList.asp';</script>"
		End IF
		
	ElseIf vAction = "unitdelete" Then
		vQuery = "DELETE [db_sitemaster].[dbo].[tbl_search_curator_unit] "
		vQuery = vQuery & "WHERE topidx = '" & vIdx & "' and gubun = '" & vUnitGubun & "' and contentsidx = '" & vUnitContentsIdx & "'"
		dbget.Execute vQuery
		
		Response.Write "<script>alert('처리되었습니다.');parent.location.reload();</script>"
		
	ElseIf vAction = "unitdeletepop" Then
		vQuery = "DELETE [db_sitemaster].[dbo].[tbl_search_curator_unit] "
		vQuery = vQuery & "WHERE topidx = '" & vIdx & "' and gubun = '" & vUnitGubun & "' and contentsidx = '" & vUnitContentsIdx & "'"
		dbget.Execute vQuery
		
		Response.Write "<script>alert('처리되었습니다.');parent.jsAllReload();</script>"
		
	ElseIf vAction = "delete" Then
		
		vQuery = "DELETE [db_sitemaster].[dbo].[tbl_search_curator] WHERE idx = '" & vIdx & "'; "
		vQuery = vQuery & "DELETE [db_sitemaster].[dbo].[tbl_search_curator_unit] WHERE topidx = '" & vIdx & "'; "
		vQuery = vQuery & "DELETE [db_sitemaster].[dbo].[tbl_search_keyword] WHERE topidx = '" & vIdx & "' and topgubun = 'c'; "
		dbget.Execute vQuery
		
    	vQuery = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    	vQuery = vQuery & "VALUES('" & session("ssBctId") & "', 'curator', '" & vIdx & "', '0', "
    	vQuery = vQuery & "'큐레이터 idx="&vIdx&" 삭제', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vQuery)
		
		Response.Write "<script>alert('삭제되었습니다.');parent.location.reload();</script>"
		
	End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->