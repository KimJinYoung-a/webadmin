<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim mode,cdl,evt_code,viewidx,isusing,page,orgUsing,param, idarr,allusing,idxArrTmp ,cdm , idx, vCateCode, menupos
	mode = Request("mode")
	cdl = Request("cdl")
	cdm = request("cdm")
	evt_code = Request("evt_code")
	viewidx= Request("viewidx")
	isusing=Request("isusing")
	page=Request("page")
	orgUsing=Request("orgUsing")
	allusing = Request("allusing")
	idxArrTmp = request("idxArrTmp")
	idx = request("idx")
	vCateCode = Request("catecode")
	menupos = request("menupos")

dim sqlStr

if mode="add" then

	'저장 실행
	sqlStr = "insert into [db_sitemaster].[dbo].tbl_category_main_eventBanner"
	sqlStr = sqlStr & " (disp1, evt_code, viewidx, isusing,cdm)"
	sqlStr = sqlStr & " values('" & CStr(vCateCode) & "'," & CStr(evt_code) & ", " & viewidx & ",'" & isusing & "','" & CStr(cdm) & "')" & vbcrlf
	rsget.Open sqlStr,dbget,1

	param = "?menupos="&menupos&"&catecode=" & vCateCode
	
elseif mode="edit" then
	sqlStr= "update [db_sitemaster].[dbo].tbl_category_main_eventBanner" & vbcrlf
	sqlStr = sqlStr & " set viewidx=" & CStr(viewidx) & vbcrlf
	sqlStr = sqlStr & " , isusing='" & CStr(isusing) & "'" & vbcrlf
	sqlStr = sqlStr & " , disp1 = '" & CStr(vCateCode) & "'" & vbcrlf
	sqlStr = sqlStr & " , evt_code = '" & CStr(evt_code) & "'" & vbcrlf
	sqlStr = sqlStr & " where idx=" & idx & vbcrlf
	rsget.Open sqlStr,dbget,1

	param = "?menupos="&menupos&"&catecode=" & vCateCode & "&page=" & page

elseif mode="del" then
	sqlStr= "update [db_sitemaster].[dbo].tbl_category_main_eventBanner" & vbcrlf
	sqlStr = sqlStr & " set isusing='N' " & vbcrlf
	sqlStr = sqlStr & " where idx in (" & idxArrTmp & ")" & vbcrlf
	rsget.Open sqlStr,dbget,1

	param = "?menupos="&menupos&"&catecode=" & vCateCode
		
elseif mode="changeUsing" then
	sqlStr= "update [db_sitemaster].[dbo].tbl_category_main_eventBanner" & vbcrlf
	sqlStr = sqlStr & " set isusing='" & allusing & "' " & vbcrlf
	sqlStr = sqlStr & " where idx in (" & idxArrTmp & ")" & vbcrlf
	rsget.Open sqlStr,dbget,1

	param = "?menupos="&menupos&"&catecode=" & vCateCode

end if

'response.write sqlStr
%>

<script language="javascript">

	alert('저장 되었습니다.');
	location.replace('/admin/categorymaster/category_main_EventBanner2015.asp<%=param%>');

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->