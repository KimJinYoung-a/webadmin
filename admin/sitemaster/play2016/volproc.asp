<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : PLAYing
' Hieditor : 이종화 생성
'			 2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim vQuery, vAction, vMIdx, vVolNum, vTitle, vOpenDate, vState, vMoBGColor, vWorkText, vPartWDID, vPartMKID, vPartPBID
	vAction = requestCheckVar(Request("action"),10)
	vMIdx = requestCheckVar(Request("midx"),10)
	vVolNum = requestCheckVar(Request("volnum"),3)
	vTitle = requestCheckVar(Request("title"),100)
	vOpenDate = requestCheckVar(Request("opendate"),10)
	vState = requestCheckVar(Request("state"),2)
	vMoBGColor = requestCheckVar(Request("mo_bgcolor"),6)
	vWorkText = Request("worktext")
	vPartMKID = requestCheckVar(Request("partmkid"),32)
	vPartWDID = requestCheckVar(Request("partwdid"),32)
	vPartPBID = requestCheckVar(Request("partpbid"),32)
	
	If vAction = "insert" Then
		if vTitle <> "" and not(isnull(vTitle)) then
			vTitle = ReplaceBracket(vTitle)
		end If
		if vWorkText <> "" and not(isnull(vWorkText)) then
			vWorkText = ReplaceBracket(vWorkText)
		end If

		vQuery = "INSERT INTO [db_giftplus].[dbo].[tbl_play_master](volnum, title, startdate, state, mo_bgcolor, worktext, partwdid, partmkid, partpbid, lastupdate, lastupdateID) VALUES "
		vQuery = vQuery & "('" & vVolNum & "', '" & html2db(vTitle) & "', '" & vOpenDate & "', '" & vState & "', '" & vMoBGColor & "', '" & html2db(vWorkText) & "', "
		vQuery = vQuery & "'" & vPartWDID & "', '" & vPartMKID & "', '" & vPartPBID & "', getdate(), '" & session("ssBctId") & "')"
		dbget.Execute vQuery
		
		vQuery = "select IDENT_CURRENT('db_giftplus.dbo.tbl_play_master') as midx"
		rsget.Open vQuery, dbget, 1
		If Not Rsget.Eof then
			vMIdx = rsget("midx")
		end if
		rsget.close
		
		Response.Write "<script type='text/javascript'>alert('처리되었습니다.');location.href='/admin/sitemaster/play2016/volwrite.asp?midx="&vMIdx&"';</script>"
	ElseIf vAction = "update" Then
		if vTitle <> "" and not(isnull(vTitle)) then
			vTitle = ReplaceBracket(vTitle)
		end If
		if vWorkText <> "" and not(isnull(vWorkText)) then
			vWorkText = ReplaceBracket(vWorkText)
		end If

		vQuery = "UPDATE [db_giftplus].[dbo].[tbl_play_master] SET "
		vQuery = vQuery & " volnum = '" & vVolNum & "', "
		vQuery = vQuery & " title = '" & html2db(vTitle) & "', "
		vQuery = vQuery & " startdate = '" & vOpenDate & "', "
		vQuery = vQuery & " state = '" & vState & "', "
		vQuery = vQuery & " mo_bgcolor = '" & vMoBGColor & "', "
		vQuery = vQuery & " worktext = '" & html2db(vWorkText) & "', "
		vQuery = vQuery & " partwdid = '" & vPartWDID & "', "
		vQuery = vQuery & " partmkid = '" & vPartMKID & "', "
		vQuery = vQuery & " partpbid = '" & vPartPBID & "', "
		vQuery = vQuery & " lastupdate = getdate(), "
		vQuery = vQuery & " lastupdateID = '" & session("ssBctId") & "' "
		vQuery = vQuery & " WHERE midx = '" & vMIdx & "'"
		dbget.Execute vQuery
		
		Response.Write "<script type='text/javascript'>alert('처리되었습니다.');location.href='/admin/sitemaster/play2016/volwrite.asp?midx="&vMIdx&"';</script>"
	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->