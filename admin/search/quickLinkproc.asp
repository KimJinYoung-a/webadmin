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
	Dim i, vAction, vQuery, vIdx, vQuickType, vName, vSubCopy, vURL_PC, vURL_M, vViewGubun, vRegUserName, vHtmlCont, vBtnName, vBtnPCLink, vBtnMLink
	Dim vSDate, vEDate, vRegdate, vLastUserName, vLastdate, vMemo, vUseYN, vKwArr, vBgGubun, vBgColor, vBgImgPC, vBgImgM, vQuickBrID
	Dim vQImgUseYN, vQImgPC, vQImgM, vBtnColor, vKeyword, vKeywordInDB, vUserID, vExist
	Dim vShhmmss, vEhhmmss
	
	'### ����ũ �⺻ ����
	vIdx 		= requestCheckVar(Request("idx"),15)
	vAction	= requestCheckVar(Request("action"),10)
	vUserID	= session("ssBctId")
	vQuickType	= requestCheckVar(Request("quicktype"),3)
	vName 		= html2db(requestCheckVar(Request("quickname"),20))
	vQuickBrID	= requestCheckVar(Request("quickbrid"),50)
	vSubCopy	= html2db(requestCheckVar(Request("subcopy"),70))
	vURL_PC	= requestCheckVar(Request("url_pc"),200)
	vURL_M		= requestCheckVar(Request("url_m"),200)
	vViewGubun	= requestCheckVar(Request("viewgubun"),6)
	If vViewGubun = "always" Then	'### ��ó���
		vSDate = ""
		vEDate = ""
	Else	'### �Ⱓ����
		vSDate = requestCheckVar(Request("sdate"),10)
		vEDate = requestCheckVar(Request("edate"),10)
		vShhmmss = requestCheckVar(Request("shhmmss"),8)
		vEhhmmss = requestCheckVar(Request("ehhmmss"),8)
		
		vSDate = vSDate & " " & vShhmmss
		vEDate = vEDate & " " & vEhhmmss
	End If
	vUseYN		= requestCheckVar(Request("useyn"),1)
	vMemo 		= html2db(Request("memo"))
	
	
	'### ����ũ �Ӽ� ����
	vBtnName	= html2db(requestCheckVar(Request("btnname"),25))
	vBtnPCLink	= requestCheckVar(Request("btnlinkpc"),200)
	vBtnMLink	= requestCheckVar(Request("btnlinkm"),200)
	vBtnColor	= requestCheckVar(Request("btn_color"),6)
	
	'vBgGubun	= requestCheckVar(Request("bggubun"),1)
	'If vBgGubun = "c" Then	'### ����ũ��� �ܻ�
	'	vBgImgPC = ""
	'	vBgImgM = ""
	'	vBgColor	= requestCheckVar(Request("bgcolor"),6)
	'ElseIf vBgGubun = "i" Then	'### ����ũ��� �̹������
	'	vBgImgPC	= requestCheckVar(Request("qbgimgurlpc"),100)
	'	vBgImgM	= requestCheckVar(Request("qbgimgurlm"),100)
	'	vBgColor	= ""
	'End If
    vBgColor	= requestCheckVar(Request("bgcolor"),6)

	vQImgUseYN = requestCheckVar(Request("qimg_useyn"),1)
	If vQImgUseYN = "y" Then	'### ����ũ�̹��� ����� ���
		vQImgPC = requestCheckVar(Request("qimgurlpc"),100)
		vQImgM = requestCheckVar(Request("qimgurlm"),100)
	End If
	
	'### ����ũ �Ӽ� ���� Ŀ������ html
	vHtmlCont = Request("htmlcont")
	
	'### �˻�Ű����
	vKeyword = requestCheckVar(Request("keyword"),200)
	vKeywordInDB = requestCheckVar(Request("keyword_in_db"),200)



	If vAction = "" Then

		vExist = fnKeywordExistCheck(vKeyword,"q",vIdx)
		If vExist = "1" Then
			Response.Write "<script>alert('���� �� �� �˻�Ű���忡 ���� Ű���尡 2�� �̻� �ֽ��ϴ�.\nȮ���غ��ñ� �ٶ��ϴ�.');history.back();</script>"
			dbget.close
			Response.End
		ElseIf vExist = "2" Then
			Response.Write "<script>alert('��ü �˻�Ű���忡 ���� Ű���尡 �ֽ��ϴ�.\nȮ���غ��ñ� �ٶ��ϴ�.');history.back();</script>"
			dbget.close
			Response.End
		End If

		If vIdx = "" Then

			vQuery = "INSERT INTO [db_sitemaster].[dbo].[tbl_search_quicklink]"
			vQuery = vQuery & "(type, name, brandid, subcopy, url_pc, url_m, "
			vQuery = vQuery & "viewgubun, sdate, edate, btnname, btn_pclink, btn_mlink, "
			vQuery = vQuery & "btn_color, bggubun, bgcolor, bgimgpc, bgimgm, qimg_useyn, "
			vQuery = vQuery & "qimgpc, qimgm, htmlcont, useyn, memo, reguserid, lastupdateid) "
			vQuery = vQuery & " VALUES "
			vQuery = vQuery & "('" & vQuickType & "', '" & vName & "', '" & vQuickBrID & "', '" & vSubCopy & "', '" & vURL_PC & "', '" & vURL_M & "', "
			vQuery = vQuery & "'" & vViewGubun & "', '" & vSDate & "', '" & vEDate & "', '" & vBtnName & "', '" & vBtnPCLink & "', '" & vBtnMLink & "', "
			vQuery = vQuery & "'" & vBtnColor & "', '" & vBgGubun & "', '" & vBgColor & "', '" & vQImgPC & "', '" & vQImgM & "', '" & vQImgUseYN & "', "
			vQuery = vQuery & "'" & vQImgPC & "', '" & vQImgM & "', '" & vHtmlCont & "', '" & vUseYN & "', '" & vMemo & "', '" & vUserID & "', '" & vUserID & "')"
			dbget.Execute vQuery
			
			vQuery = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_search_quicklink') as idx"
			rsget.CursorLocation = adUseClient
			rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
			If Not Rsget.Eof then
				vIdx = rsget("idx")
			end if
			rsget.close

		Else

			vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_search_quicklink] SET "
			vQuery = vQuery & "name = '" & vName & "' "
			vQuery = vQuery & ", brandid = '" & vQuickBrID & "' "
			vQuery = vQuery & ", subcopy = '" & vSubCopy & "' "
			vQuery = vQuery & ", url_pc = '" & vURL_PC & "' "
			vQuery = vQuery & ", url_m = '" & vURL_M & "' "
			vQuery = vQuery & ", viewgubun = '" & vViewGubun & "' "
			vQuery = vQuery & ", sdate = '" & vSDate & "' "
			vQuery = vQuery & ", edate = '" & vEDate & "' "
			vQuery = vQuery & ", btnname = '" & vBtnName & "' "
			vQuery = vQuery & ", btn_pclink = '" & vBtnPCLink & "' "
			vQuery = vQuery & ", btn_mlink = '" & vBtnMLink & "' "
			vQuery = vQuery & ", btn_color = '" & vBtnColor & "' "
			vQuery = vQuery & ", bggubun = '" & vBgGubun & "' "
			vQuery = vQuery & ", bgcolor = '" & vBgColor & "' "
			vQuery = vQuery & ", bgimgpc = '" & vQImgPC & "' "
			vQuery = vQuery & ", bgimgm = '" & vQImgM & "' "
			vQuery = vQuery & ", qimg_useyn = '" & vQImgUseYN & "' "
			vQuery = vQuery & ", qimgpc = '" & vQImgPC & "' "
			vQuery = vQuery & ", qimgm = '" & vQImgM & "' "
			vQuery = vQuery & ", htmlcont = '" & vHtmlCont & "' "
			vQuery = vQuery & ", useyn = '" & vUseYN & "' "
			vQuery = vQuery & ", memo = '" & vMemo & "' "
			vQuery = vQuery & ", lastupdateid = '" & vUserID & "' "
			vQuery = vQuery & ", lastupdatedate = getdate() "
			vQuery = vQuery & "where idx = '" & vIdx & "' "
			dbget.Execute vQuery

		End If
		
		If vKeyword <> vKeywordInDB Then	'### ������ �ٸ����� ������Ʈ.
			vQuery = "DELETE [db_sitemaster].[dbo].[tbl_search_keyword] WHERE topidx = '" & vIdx & "' and topgubun = 'q'; "
			For i = LBound(Split(vKeyword,",")) To UBound(Split(vKeyword,","))
				vQuery = vQuery & "INSERT INTO [db_sitemaster].[dbo].[tbl_search_keyword](topidx, topgubun, keyword) "
				vQuery = vQuery & "VALUES('" & vIdx & "', 'q', '" & Trim(Split(vKeyword,",")(i)) & "');"
			Next
			dbget.Execute vQuery
		End If

		Response.Write "<script>alert('ó���Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
		
	ElseIf vAction = "delete" Then
		
		vQuery = "DELETE [db_sitemaster].[dbo].[tbl_search_quicklink] WHERE idx = '" & vIdx & "'; "
		vQuery = vQuery & "DELETE [db_sitemaster].[dbo].[tbl_search_keyword] WHERE topidx = '" & vIdx & "' and topgubun = 'q'; "
		dbget.Execute vQuery
		
    	vQuery = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    	vQuery = vQuery & "VALUES('" & session("ssBctId") & "', 'quicklink', '" & vIdx & "', '0', "
    	vQuery = vQuery & "'����ũ idx="&vIdx&" ����', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vQuery)
		
		Response.Write "<script>alert('�����Ǿ����ϴ�.');parent.location.reload();</script>"
		
	End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->