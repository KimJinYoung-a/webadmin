<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include file="./other_site_iteminfo_cls.asp" -->
<!-- #include virtual="/lib/classes/search/itemCls.asp" -->
<!-- #include virtual="/lib/classes/search/searchCls.asp" -->
<%
	Dim cOSI, vQuery, vSiteCode, vSiteItemID, vItemID, vUserID, vItemName, vAction
	vAction = requestCheckVar(request("action"),50)
	vSiteCode = requestCheckVar(request("sitecode"),50)
	vSiteItemID = requestCheckVar(request("siteitemid"),15)
	vItemID = requestCheckVar(request("itemid"),15)
	vUserID = session("ssBctId")

	If vAction = "delete" Then
		vQuery = "DELETE [db_analyze_etc].[dbo].[tbl_remote_site_price_Item_Match] where sitecode = '" & vSiteCode & "' and siteitemcode = '" & vSiteItemID & "'"
		dbAnalget.execute(vQuery)
	Else
		If vItemID <> "" AND vSiteCode <> "" AND vSiteItemID <> "" Then 
			vQuery = ""
			vQuery = vQuery & "IF EXISTS(select itemid from [db_analyze_etc].[dbo].[tbl_remote_site_price_Item_Match] where sitecode = '" & vSiteCode & "' and siteitemcode = '" & vSiteItemID & "') "
			vQuery = vQuery & "	BEGIN "
			vQuery = vQuery & "		UPDATE [db_analyze_etc].[dbo].[tbl_remote_site_price_Item_Match] "
			vQuery = vQuery & "		SET itemid = '" & vItemID & "', lastmatchdate = getdate(), lastmatchuser = '" & vUserID & "' "
			vQuery = vQuery & "		WHERE sitecode = '" & vSiteCode & "' and siteitemcode = '" & vSiteItemID & "' "
			vQuery = vQuery & "	END "
			vQuery = vQuery & "	ELSE "
			vQuery = vQuery & "	BEGIN "
			vQuery = vQuery & "		INSERT INTO [db_analyze_etc].[dbo].[tbl_remote_site_price_Item_Match](sitecode, siteitemcode, itemid, matchuser, lastmatchuser) "
			vQuery = vQuery & "		VALUES('" & vSiteCode & "', '" & vSiteItemID & "', '" & vItemID & "', '" & vUserID & "', '" & vUserID & "') "
			vQuery = vQuery & "	END"
			dbAnalget.execute(vQuery)
		End If
	End If

	Response.Write "<script>alert('저장되었습니다.');parent.jsItemlist();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->