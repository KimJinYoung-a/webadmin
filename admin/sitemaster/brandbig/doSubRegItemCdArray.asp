<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/mobile/today_brandinfoCls.asp" -->

<%
Dim sqlStr, lp, strRst, strErr, Scnt, Ecnt
Dim listidx, arrItemid, actItemid, subSortNo, subIsUsing
dim tmpArrIid, ptype, mode

mode = requestCheckVar(request("mode"),10)
listidx = request("listidx")

if mode="auto" then
	dim makerid, oaward, ix, itemcnt
	arrItemid=""
    makerid = request("makerid")

	set oaward = new CMainbanner
		oaward.FPageSize 			= 5
        oaward.FRectMakerID         = makerid
		oaward.GetBrandItemList
		If oaward.FResultCount>0 Then
			For ix=0 to oaward.FResultCount-1
                if (ix=0) then
                arrItemid = oaward.FItemList(ix).FItemID
                else
				arrItemid = arrItemid & "," & oaward.FItemList(ix).FItemID
                end if
			Next
		end if
	set oaward = Nothing
	arrItemid = split(arrItemid,",")
else
	arrItemid = split(replace(request("subItemidArray"),vbCrLf,","),",")
end if

if trim(request("itemidarr"))<>"" then
	tmpArrIid = trim(request("itemidarr"))
	if Right(tmpArrIid,1)="," then tmpArrIid=Left(tmpArrIid,Len(tmpArrIid)-1)
	arrItemid = split(tmpArrIid,",")
end if
subSortNo = request("subSortNo")
subIsUsing = request("subIsUsing")
ptype = request("ptype")

if listidx="" then
	Call Alert_Return("���ø� ������ �����ϴ�.")
	dbget.close(): response.End
end if

if Not(isArray(arrItemid)) then
	Call Alert_Return("��ǰ�ڵ� ������ �߸��Ǿ����ϴ�.")
	dbget.close(): response.End
end if

if subSortNo="" then subSortNo="0"
if subIsUsing="" then subIsUsing="Y"
Scnt=0: Ecnt=0

for lp=0 to ubound(arrItemid)
	if isNumeric(arrItemid(lp)) then
		actItemid = actItemid & chkIIF(actItemid<>"",",","") & getNumeric(arrItemid(lp))
		Scnt=Scnt+1
	else
		if trim(arrItemid(lp))<>"" then
			strErr = strErr & chkIIF(strErr<>"",",","") & arrItemid(lp)
			Ecnt=Ecnt+1
		end if
	end if
next

if Scnt>0 then
    sqlStr = " insert into [db_sitemaster].[dbo].tbl_pc_main_brandbig_item" + VbCrlf
    sqlStr = sqlStr + " (listidx, itemid, isusing, sortnum ) " + VbCrlf
    sqlStr = sqlStr + " select '" & listidx & "'" + VbCrlf
    sqlStr = sqlStr + " ,itemid " + VbCrlf
    sqlStr = sqlStr + " ,'" & subIsUsing & "'" + VbCrlf
    sqlStr = sqlStr + " ,'" & subSortNo & "'" + VbCrlf
    sqlStr = sqlStr + " from db_item.dbo.tbl_item" + VbCrlf
    sqlStr = sqlStr + " where itemid in (" & actItemid & ")" + VbCrlf
    sqlStr = sqlStr + " 	and isusing='Y'" + VbCrlf
    sqlStr = sqlStr + " 	and itemid not in (" + VbCrlf
    sqlStr = sqlStr + " 		select itemid" + VbCrlf
    sqlStr = sqlStr + " 		from [db_sitemaster].[dbo].tbl_pc_main_brandbig_item" + VbCrlf
    sqlStr = sqlStr + " 		where listidx='" & listidx & "'" + VbCrlf
    sqlStr = sqlStr + " 	)" + VbCrlf

	dbget.Execute sqlStr

	'// ���������� ���� ������ ������Ʈ
	sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_brandbig_list " + VbCrlf
	sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
	sqlStr = sqlStr + " where idx=" + cstr(listidx)
end if

strRst = "[" & Scnt & "]�� ����"
if Ecnt>0 then strRst = strRst & "\n[" & Ecnt & "]�� ����\n�ؽ��а�: " & strErr

dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('����Ǿ����ϴ�.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->