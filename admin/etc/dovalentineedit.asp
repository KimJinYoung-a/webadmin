<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode, categorycd,idx
dim idxarr,viewidxarr,isusingarr
dim itemid, itemidarr,masterid

masterid = request("masterid")
mode = request.Form("mode")
idx = request.Form("idx")
isusingarr = request.Form("isusingarr")
itemid = request.Form("itemid")
itemidarr = request.Form("itemidarr")

dim sqlStr, i, cnt

if (mode="deleventdetail") then
	sqlStr = " delete from [db_contents].[dbo].tbl_valentine_item" + VbCrlf
	sqlStr = sqlStr + " where  idx=" + idx + VbCrlf
	sqlStr = sqlStr + " and itemid=" + itemid

	rsget.Open sqlStr,dbget,1
elseif (mode="addeventdetailarr") then
	if Right(itemidarr,1)="," then
		itemidarr = Left(itemidarr,Len(itemidarr)-1)
	end if

	itemidarr = split(itemidarr,",")

	cnt = ubound(itemidarr)

	for i=0 to cnt
		sqlStr = " insert into [db_contents].[dbo].tbl_valentine_item" + VbCrlf
		sqlStr = sqlStr + " (masterid,itemid)" + VbCrlf
		sqlStr = sqlStr + " values (" & masterid & "," & itemidarr(i) & ")"
		rsget.Open sqlStr,dbget,1
	next 
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")

if (mode="addcatevrnt") or (mode="addmainevent") then
	refer = refer + "&react=true"
end if
%>

<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->