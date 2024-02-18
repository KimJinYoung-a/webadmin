<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 플레이모바일
' Hieditor : 이종화 생성
'			 2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play_moCls.asp" -->
<%
	Dim vQuery, vAction, vIdx, vPlayType, vViewNO, vTitle, vSubCopy, vStartDate, vPartmdid, vPartwdid, vPartpbid, vState, vListImg, vIsUsing, vColorCD, vContents, vStyleCD, vWorkComm, vViewNoTxt, vContentsIdx, vSortNo
	vAction = requestCheckVar(Request("action"),10)
	vIdx = requestCheckVar(Request("idx"),10)
	vPlayType = requestCheckVar(Request("playtype"),3)
	vViewNO = requestCheckVar(Request("viewno"),50)
	vViewNoTxt = requestCheckVar(Request("viewnotxt"),50)
	vTitle = requestCheckVar(Request("title"),200)
	vSubCopy = requestCheckVar(Request("subcopy"),200)
	vStartDate = requestCheckVar(Request("startdate"),10)
	vPartmdid = requestCheckVar(Request("partmdid"),32)
	vPartwdid = requestCheckVar(Request("partwdid"),32)
	vPartpbid = requestCheckVar(Request("partpbid"),32)
	vState = requestCheckVar(Request("state"),10)
	vListImg = requestCheckVar(Request("listimg"),300)
	vIsUsing = requestCheckVar(Request("isusing"),10)
	vColorCD = requestCheckVar(Request("colorcd"),10)
	vStyleCD = requestCheckVar(Request("playstyle"),10)
	vContents = Request("contents")
	vWorkComm = Request("workcomment")
	vContentsIdx = requestCheckVar(Request("contentsidx"),10)
	vSortNo = requestCheckVar(Request("sortno"),10)
	
	
	If vAction = "insert" Then
		if vViewNO <> "" and not(isnull(vViewNO)) then
			vViewNO = ReplaceBracket(vViewNO)
		end If
		if vViewNoTxt <> "" and not(isnull(vViewNoTxt)) then
			vViewNoTxt = ReplaceBracket(vViewNoTxt)
		end If
		if vTitle <> "" and not(isnull(vTitle)) then
			vTitle = ReplaceBracket(vTitle)
		end If
		if vSubCopy <> "" and not(isnull(vSubCopy)) then
			vSubCopy = ReplaceBracket(vSubCopy)
		end If
		if vWorkComm <> "" and not(isnull(vWorkComm)) then
			vWorkComm = ReplaceBracket(vWorkComm)
		end If
		if vContents <> "" and not(isnull(vContents)) then
			vContents = ReplaceBracket(vContents)
		end If

		vQuery = "INSERT INTO [db_sitemaster].[dbo].[tbl_play_mo](viewno, viewnotxt, type, title, subcopy, startdate, state, isusing, partwdid, contents_idx, "
		vQuery = vQuery & " partmdid, partpbid, listimg, contents, colorcd, stylecd, lastadminid, workcomment, sortno) VALUES"
		vQuery = vQuery & "('" & vViewNO & "', '" & html2db(vViewNoTxt) & "', '" & vPlayType & "', '" & html2db(vTitle) & "', '" & html2db(vSubCopy) & "', '" & vStartDate & "', "
		vQuery = vQuery & "'" & vState & "', '" & vIsUsing & "', '" & vPartwdid & "', '" & vContentsIdx & "', "
		vQuery = vQuery & "'" & vPartmdid & "', '" & vPartpbid & "', '" & vListImg & "', '" & html2db(vContents) & "', '" & vColorCD & "', '" & vStyleCD & "', '" & session("ssBctId") & "', '" & html2db(vWorkComm) & "', '" & vSortNo & "')"
		dbget.Execute vQuery
		
		vQuery = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_play_mo') as idx"
		rsget.Open vQuery, dbget, 1
		If Not Rsget.Eof then
			vIdx = rsget("idx")
		end if
		rsget.close
		
		Response.Write "<script type='text/javascript'>alert('처리되었습니다.');opener.location.reload();location.href='/admin/sitemaster/play_mo/write.asp?idx="&vIdx&"';</script>"
	ElseIf vAction = "update" Then
		if vViewNO <> "" and not(isnull(vViewNO)) then
			vViewNO = ReplaceBracket(vViewNO)
		end If
		if vViewNoTxt <> "" and not(isnull(vViewNoTxt)) then
			vViewNoTxt = ReplaceBracket(vViewNoTxt)
		end If
		if vTitle <> "" and not(isnull(vTitle)) then
			vTitle = ReplaceBracket(vTitle)
		end If
		if vSubCopy <> "" and not(isnull(vSubCopy)) then
			vSubCopy = ReplaceBracket(vSubCopy)
		end If
		if vWorkComm <> "" and not(isnull(vWorkComm)) then
			vWorkComm = ReplaceBracket(vWorkComm)
		end If
		if vContents <> "" and not(isnull(vContents)) then
			vContents = ReplaceBracket(vContents)
		end If

		vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_play_mo] SET "
		vQuery = vQuery & " viewno = '" & vViewNO & "', "
		vQuery = vQuery & " viewnotxt = '" & html2db(vViewNoTxt) & "', "
		vQuery = vQuery & " type = '" & vPlayType & "', "
		vQuery = vQuery & " title = '" & html2db(vTitle) & "', "
		vQuery = vQuery & " subcopy = '" & html2db(vSubCopy) & "', "
		vQuery = vQuery & " startdate = '" & vStartDate & "', "
		vQuery = vQuery & " state = '" & vState & "', "
		vQuery = vQuery & " isusing = '" & vIsUsing & "', "
		vQuery = vQuery & " partwdid = '" & vPartwdid & "', "
		vQuery = vQuery & " partmdid = '" & vPartmdid & "', "
		vQuery = vQuery & " partpbid = '" & vPartpbid & "', "
		vQuery = vQuery & " listimg = '" & vListImg & "', "
		vQuery = vQuery & " contents = '" & html2db(vContents) & "', "
		vQuery = vQuery & " colorcd = '" & vColorCD & "', "
		vQuery = vQuery & " stylecd = '" & vStyleCD & "', "
		vQuery = vQuery & " lastadminid = '" & session("ssBctId") & "', "
		vQuery = vQuery & " workcomment = '" & html2db(vWorkComm) & "', "
		vQuery = vQuery & " contents_idx = '" & vContentsIdx & "', "
		vQuery = vQuery & " sortno = '" & vSortNo & "', "
		vQuery = vQuery & " lastupdate = getdate() "
		vQuery = vQuery & " WHERE idx = '" & vIdx & "'"
		dbget.Execute vQuery
		
		Response.Write "<script type='text/javascript'>alert('처리되었습니다.');opener.location.reload();location.href='/admin/sitemaster/play_mo/write.asp?idx="&vIdx&"';</script>"
	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->