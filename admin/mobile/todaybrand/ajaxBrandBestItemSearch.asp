<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mobile/today_brandinfoCls.asp" -->
<%
	dim sqlStr, makerid, oaward, ix, itemcnt, resultBest
    makerid = request("makerid")
    itemcnt = request("itemcnt")
	resultBest=""
	set oaward = new CMainbanner
		oaward.FPageSize 			= itemcnt
        oaward.FRectMakerID         = makerid
		oaward.GetBrandItemList
		If oaward.FResultCount>0 Then
			For ix=0 to oaward.FResultCount-1
                if (ix=0) then
                resultBest = "'" & oaward.FItemList(ix).FItemID & "|" & oaward.FItemList(ix).FImageIcon1 & "'"
                else
				resultBest = resultBest & ",'" & oaward.FItemList(ix).FItemID & "|" & oaward.FItemList(ix).FImageIcon1 & "'"
                end if
			Next
		end if
	set oaward = Nothing
	
	response.Write resultBest
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->