<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  ����Ʈ
' History : 2015.01.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim sqlStr, lp, strRst, strErr, Scnt, Ecnt, i
Dim themeidx, arrItemid, actItemid, subSortNo, subIsUsing
dim tmpArrIid, executedate, themearr, lastadminid
	lastadminid = session("ssBctId")
	'executedate = requestcheckvar(request("executedate"),10)
	'themeidx = request("themeidx")
	themearr = requestcheckvar(request("themearr"),23)
	arrItemid = split(replace(request("subItemidArray"),vbCrLf,","),",")

if themearr="" then
	Call Alert_Return("�׸� ������ �߸��Ǿ����ϴ�.[1]")
	dbget.close(): response.End
end if

if isarray( split(themearr,"!@@") ) then
	if ubound( split(themearr,"!@@") ) <> 1 then
		Call Alert_Return("�׸� ������ �߸��Ǿ����ϴ�.[3]")
		dbget.close(): response.End
	end if
	
	executedate = split(themearr,"!@@")(0)
	themeidx = split(themearr,"!@@")(1)
else
	Call Alert_Return("�׸� ������ �߸��Ǿ����ϴ�.[2]")
	dbget.close(): response.End
end if

if trim(request("itemidarr"))<>"" then
	tmpArrIid = trim(request("itemidarr"))
	if Right(tmpArrIid,1)="," then tmpArrIid=Left(tmpArrIid,Len(tmpArrIid)-1)
	arrItemid = split(tmpArrIid,",")
end if
subSortNo = request("subSortNo")
subIsUsing = request("subIsUsing")

if themeidx="" then
	Call Alert_Return("�׸� ��ȣ�� �����ϴ�.")
	dbget.close(): response.End
end if
if executedate="" then
	Call Alert_Return("�������� �����ϴ�.")
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
    sqlStr = " insert into db_board.dbo.tbl_gifthint_item" + VbCrlf
    sqlStr = sqlStr + " (themeidx, itemid, executedate, isusing, orderno, lastadminid, lastupdate) " + VbCrlf
    sqlStr = sqlStr + " 	select " & themeidx & "" + VbCrlf
    sqlStr = sqlStr + " 	,itemid " + VbCrlf
    sqlStr = sqlStr + " 	,'" & executedate & "'" + VbCrlf
    sqlStr = sqlStr + " 	,'Y'" + VbCrlf
    sqlStr = sqlStr + " 	,99" + VbCrlf
    sqlStr = sqlStr + " 	,'" + lastadminid + "'" + VbCrlf
    sqlStr = sqlStr + " 	,getdate()" + VbCrlf
    sqlStr = sqlStr + " 	from db_item.dbo.tbl_item" + VbCrlf
    sqlStr = sqlStr + " 	where itemid in (" & actItemid & ")" + VbCrlf
    sqlStr = sqlStr + " 	and isusing='Y'" + VbCrlf
    sqlStr = sqlStr + " 	and itemid not in (" + VbCrlf
    sqlStr = sqlStr + " 		select itemid" + VbCrlf
    sqlStr = sqlStr + " 		from db_board.dbo.tbl_gifthint_item" + VbCrlf
    sqlStr = sqlStr + " 		where themeidx=" & themeidx & " and executedate='"& executedate &"'" + VbCrlf
    sqlStr = sqlStr + " 	)" + VbCrlf

    'response.write sqlStr & "<br>"
	dbget.Execute sqlStr

	sqlStr = "Update db_board.dbo.tbl_gifthint"
	sqlStr = sqlStr & " Set lastupdate=getdate()"
	sqlStr = sqlStr & " ,lastadminid='" & lastadminid & "' Where"		'����Ʈ ����: ��뿩�� > ������� ����
	sqlStr = sqlStr & " themeidx='" & themeidx & "';" & vbCrLf

    'response.write sqlStr & "<br>"
    dbget.Execute sqlStr	
end if

strRst = "[" & Scnt & "]�� ����"
if Ecnt>0 then strRst = strRst & "\n[" & Ecnt & "]�� ����\n�ؽ��а�: " & strErr

Response.Write "<script language='javascript'>" & vbCrLf
Response.Write "alert('" & strRst & "\n����Ǿ����ϴ�.');"& vbCrLf
	if trim(request("itemidarr"))="" then
		Response.Write "opener.location.reload();" & vbCrLf
		Response.Write "window.close();"& vbCrLf
	end if
Response.Write "</script>"
response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->