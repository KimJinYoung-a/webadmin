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
<%
Dim sqlStr, lp, strRst, strErr, Scnt, Ecnt, mode
Dim packIdx, arrItemid, actItemid
dim tmpArrIid

packIdx = request("packIdx")
mode = request("mode")
arrItemid = split(replace(request("subItemidArray"),vbCrLf,","),",")
if trim(request("itemidarr"))<>"" then
	tmpArrIid = trim(request("itemidarr"))
	if Right(tmpArrIid,1)="," then tmpArrIid=Left(tmpArrIid,Len(tmpArrIid)-1)
	arrItemid = split(tmpArrIid,",")
end if

if packIdx="" then
	Call Alert_Return("���屸�� ������ �����ϴ�.")
	dbget.close(): response.End
end if

if Not(isArray(arrItemid)) then
	Call Alert_Return("��ǰ�ڵ� ������ �߸��Ǿ����ϴ�.")
	dbget.close(): response.End
end if

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
	if mode="i" then
		'// ��ǰ�߰�
	    sqlStr = " insert into db_board.dbo.tbl_giftShop_packItem" + VbCrlf
	    sqlStr = sqlStr + " (packIdx, regUserid, itemid) " + VbCrlf
	    sqlStr = sqlStr + " select '" & packIdx & "', '" & session("ssBctId") & "'" + VbCrlf
	    sqlStr = sqlStr + " ,itemid " + VbCrlf
	    sqlStr = sqlStr + " from db_item.dbo.tbl_item" + VbCrlf
	    sqlStr = sqlStr + " where itemid in (" & actItemid & ")" + VbCrlf
	    sqlStr = sqlStr + " 	and itemid not in (" + VbCrlf
	    sqlStr = sqlStr + " 		select itemid" + VbCrlf
	    sqlStr = sqlStr + " 		from db_board.dbo.tbl_giftShop_packItem" + VbCrlf
	    sqlStr = sqlStr + " 		where packIdx='" & packIdx & "'" + VbCrlf
	    sqlStr = sqlStr + " 	)" + VbCrlf
		dbget.Execute(sqlStr)

	elseif mode="d" then
		'// ��ǰ����
	    sqlStr = " delete from db_board.dbo.tbl_giftShop_packItem" + VbCrlf
	    sqlStr = sqlStr + " Where packIdx='" & packIdx & "'" + VbCrlf
	    sqlStr = sqlStr + " 	and itemid in (" & actItemid & ")" + VbCrlf
		dbget.Execute(sqlStr)

	end if

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