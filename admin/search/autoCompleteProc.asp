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
	'### 기본정보 ###
	Dim i, vAction, vQuery, vIdx, vAutoType, vTitle, vURL_PC, vURL_M, vIcon, vRegUserID, vMemo, vUseYN, vSortNo, vUserID
	Dim vIdxArr, vTitleArr, vAutoTypeArr, vIsExist
	vIdx 		= requestCheckVar(Request("idx"),15)
	vAction	= requestCheckVar(Request("action"),10)
	vUserID	= session("ssBctId")
	vAutoType	= requestCheckVar(Request("autotype"),2)
	vTitle 	= html2db(requestCheckVar(Request("title"),20))
	vURL_PC	= requestCheckVar(Request("url_pc"),200)
	vURL_M		= requestCheckVar(Request("url_m"),200)
	vIcon		= requestCheckVar(Request("icon"),4)
	vMemo 		= html2db(Request("memo"))
	vUseYN		= requestCheckVar(Request("useyn"),1)
	vSortNo	= requestCheckVar(Request("sortno"),4)
	
	vIdxArr	= Replace(requestCheckVar(Request("idxarr"),200), " ", "")
	vTitleArr	= requestCheckVar(Request("titlearr"),300)
	vAutoTypeArr	= Replace(requestCheckVar(Request("autotypearr"),200), " ", "")


	If vAction = "" Then
		
		If Not vIsExist Then
			If vIdx = "" Then
				vIsExist = fnIsExistValue(0,vAutoType,vTitle)
				
				vQuery = "INSERT INTO [db_sitemaster].[dbo].[tbl_search_autocomplete](autotype, title, url_pc, url_m, icon, memo, useyn, reguserid, lastupdateid) "
				vQuery = vQuery & " VALUES "
				vQuery = vQuery & "('" & vAutoType & "', '" & vTitle & "', '" & vURL_PC & "', '" & vURL_M & "', '" & vIcon & "', '" & vMemo & "', "
				vQuery = vQuery & "'" & vUseYN & "','" & vUserID & "', '" & vUserID & "')"
				dbget.Execute vQuery

			Else
				vIsExist = fnIsExistValue(vIdx,vAutoType,vTitle)
				
				vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_search_autocomplete] SET "
				vQuery = vQuery & "autotype = '" & vAutoType & "' "
				vQuery = vQuery & ", title = '" & vTitle & "' "
				vQuery = vQuery & ", url_pc = '" & vURL_PC & "' "
				vQuery = vQuery & ", url_m = '" & vURL_M & "' "
				vQuery = vQuery & ", icon = '" & vIcon & "' "
				vQuery = vQuery & ", memo = '" & vMemo & "' "
				vQuery = vQuery & ", useyn = '" & vUseYN & "' "
				vQuery = vQuery & ", lastupdateid = '" & vUserID & "' "
				vQuery = vQuery & ", lastupdatedate = getdate() "
				'vQuery = vQuery & ", sortno = '" & vSortNo & "' "	'### 추후 추가 될 경우 사용
				vQuery = vQuery & "where idx = '" & vIdx & "' "
				dbget.Execute vQuery

			End If
		Else
			Response.Write "<script>alert('"&vTitle&" 와 같은 제목이 있습니다.');</script>"
			
		End If
		
		Response.Write "<script>alert('처리되었습니다.');opener.location.reload();window.close();</script>"
	ElseIf vAction = "update_arr" Then
		
		vQuery = ""
		For i = LBound(Split(vIdxArr,",")) To UBound(Split(vIdxArr,","))
		
			vIsExist = fnIsExistValue(Split(vIdxArr,",")(i),Split(vAutoTypeArr,",")(i),Trim(Split(vTitleArr,",")(i)))
			
			If Not vIsExist Then
				vQuery = vQuery & "UPDATE [db_sitemaster].[dbo].[tbl_search_autocomplete] SET "
				vQuery = vQuery & "title = '" & Trim(Split(vTitleArr,",")(i)) & "' "
				vQuery = vQuery & ", lastupdateid = '" & vUserID & "' "
				vQuery = vQuery & ", lastupdatedate = getdate() "
				vQuery = vQuery & "WHERE idx = '" & Split(vIdxArr,",")(i) & "'; "
			Else
				Response.Write "<script>alert('"&Trim(Split(vTitleArr,",")(i))&" 와 같은 제목이 있습니다.');</script>"
				dbget.close
				Response.End
				Exit For
			End IF
		Next
		
		If vQuery <> "" Then
			dbget.Execute vQuery
		End If
		
		Response.Write "<script>alert('처리되었습니다.');parent.location.reload();</script>"
	ElseIf vAction = "delete_arr" Then
		vQuery = ""
		vQuery = vQuery & "DELETE From [db_sitemaster].[dbo].[tbl_search_autocomplete] WHERE idx in(" & vIdxArr & ")"
		dbget.Execute vQuery
		
		Response.Write "<script>alert('처리되었습니다.');parent.location.reload();</script>"
	End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->