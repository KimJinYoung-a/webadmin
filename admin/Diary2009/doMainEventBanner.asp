<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 �ѿ�� 2008����Ʈ�����̵� 2009������ ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim mode,cdl,evt_code,viewidx,isusing,page,orgUsing,param, idarr,allusing,idxArrTmp ,cdm , idx
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

dim sqlStr

if mode="add" then

	'���� ����
	sqlStr = "insert into [db_diary2010].[dbo].tbl_category_main_eventBanner"
	sqlStr = sqlStr & " (cdl, evt_code, viewidx, isusing,cdm)"
	sqlStr = sqlStr & " values('" & CStr(cdl) & "'," & CStr(evt_code) & ", " & viewidx & ",'" & isusing & "','" & CStr(cdm) & "')" & vbcrlf
	rsget.Open sqlStr,dbget,1

	param = "?cdl=" & cdl
	
elseif mode="edit" then
	sqlStr= "update [db_diary2010].[dbo].tbl_category_main_eventBanner" & vbcrlf
	sqlStr = sqlStr & " set viewidx=" & CStr(viewidx) & vbcrlf
	sqlStr = sqlStr & " , isusing='" & CStr(isusing) & "'" & vbcrlf
	sqlStr = sqlStr & " , cdm='" & CStr(cdm) & "'" & vbcrlf
	sqlStr = sqlStr & " where idx=" & idx & vbcrlf
	rsget.Open sqlStr,dbget,1

	param = "?cdl=" & cdl & "&page=" & page & "&isusing=" & orgUsing

elseif mode="del" then
	sqlStr= "update [db_diary2010].[dbo].tbl_category_main_eventBanner" & vbcrlf
	sqlStr = sqlStr & " set isusing='N' " & vbcrlf
	sqlStr = sqlStr & " where idx in (" & idxArrTmp & ")" & vbcrlf
	rsget.Open sqlStr,dbget,1

	param = "?cdl=" & cdl & "&isusing=" & orgUsing
		
elseif mode="changeUsing" then
	sqlStr= "update [db_diary2010].[dbo].tbl_category_main_eventBanner" & vbcrlf
	sqlStr = sqlStr & " set isusing='" & allusing & "' " & vbcrlf
	sqlStr = sqlStr & " where idx in (" & idxArrTmp & ")" & vbcrlf
	rsget.Open sqlStr,dbget,1

	param = "?cdl=" & cdl & "&isusing=" & orgUsing

end if

'response.write sqlStr
%>

<script language="javascript">

	alert('���� �Ǿ����ϴ�.');
	location.replace('/admin/diary2009/category_main_EventBanner.asp<%=param%>');

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->