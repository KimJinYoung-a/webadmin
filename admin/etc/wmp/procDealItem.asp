<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim itemid, startDate, startDateTime, endDate, endDateTime, mode, i, idx, newItemName, limitCount
Dim sqlStr, AssignedRow, arrLimitCount
Dim arrItemid : arrItemid = request("cksel")
itemid      = request("itemid")
startDate   = request("startDate")
startDateTime = request("startDateTime")
newItemName = request("newItemName")
limitCount = request("limitCount")
endDate     = request("endDate")
endDateTime = request("endDateTime")
mode        = request("mode")
idx         = request("idx")
startDate = startDate & " " & startDateTime
endDate = endDate & " " & endDateTime

If isNumeric(itemid) = False Then
	Response.Write "<script language=javascript>alert('상품코드는 숫자만 가능합니다.');self.close();</script>"
	response.end
End If

If mode = "I" Then
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM db_etcmall.dbo.tbl_wemake_dealItem " & VbCrlf
	sqlStr = sqlStr & " WHERE itemid = '"& itemid &"'" & VbCrlf
	dbget.Execute(sqlStr)

	sqlStr = ""
	sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_wemake_dealItem (itemid, newItemName, limitCount, startDate, endDate, regDate, regUserId) VALUES " & VBCRLF
	sqlStr = sqlStr & " ('"& itemid &"', '"& newItemName &"','"& limitCount &"', '"& startDate &"', '"& endDate &"', getdate(), '"&  session("ssBctID") &"') "
	dbget.Execute(sqlStr)
	Response.Write "<script language=javascript>alert('저장 하였습니다.');opener.location.reload();self.close();</script>"
	dbget.close()	:	response.End
ElseIf mode = "U" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_etcmall.dbo.tbl_wemake_dealItem SET " & VBCRLF
	sqlStr = sqlStr & " newItemName = '"& newItemName &"' " & VBCRLF
	sqlStr = sqlStr & " ,limitCount = '"& limitCount &"' " & VBCRLF
	sqlStr = sqlStr & " ,startDate = '"& startDate &"' " & VBCRLF
	sqlStr = sqlStr & " ,endDate = '"& endDate &"' " & VBCRLF
	sqlStr = sqlStr & " ,lastUpdate = getdate() " & VBCRLF
	sqlStr = sqlStr & " ,lastUpdateUserId = '"& session("ssBctID") &"' " & VBCRLF
	sqlStr = sqlStr & " WHERE idx = '"& idx &"' " & VBCRLF
	dbget.Execute(sqlStr)
	Response.Write "<script language=javascript>alert('수정 하였습니다.');opener.location.reload();self.close();</script>"
	dbget.close()	:	response.End
ElseIf mode = "D" Then
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	'sqlStr = sqlStr & " UPDATE db_etcmall.dbo.tbl_wemake_dealItem " & VbCrlf
	'sqlStr = sqlStr & " SET lastUpdateUserId = '"& session("ssBctID") &"' " & VBCRLF
	sqlStr = sqlStr & " DELETE FROM db_etcmall.dbo.tbl_wemake_dealItem " & VbCrlf
	sqlStr = sqlStr & " WHERE itemid in (" & arrItemid & ")" & VbCrlf
	dbget.Execute sqlStr,AssignedRow

	sqlStr = sqlStr & " DELETE FROM db_etcmall.dbo.tbl_wemake_dealOption " & VbCrlf
	sqlStr = sqlStr & " WHERE itemid in (" & arrItemid & ")" & VbCrlf
	dbget.Execute sqlStr
	Response.Write "<script language=javascript>alert('"& AssignedRow &"건 삭제 하였습니다.');parent.location.reload();</script>"
	dbget.close()	:	response.End
ElseIf mode = "O" Then
	Dim itemoptionarr, optionCountArr, spitemoptionarr, spoptionCountArr
	itemoptionarr  = request("itemoptionarr")
	optionCountArr = request("optionCountArr")

	If Right(itemoptionarr,1) = "," Then
		itemoptionarr = Left(itemoptionarr, Len(itemoptionarr) - 1)
	End If

	If Right(optionCountArr,1) = "," Then
		optionCountArr = Left(optionCountArr, Len(optionCountArr) - 1)
	End If

	spitemoptionarr 	= split(itemoptionarr,",")
	spoptionCountArr	= split(optionCountArr,",")
	limitCount = 0

	For i=0 to UBound(spitemoptionarr)
		if (Len(Trim(spitemoptionarr(i)))=4) then
			sqlStr = ""
			sqlStr = sqlStr & "	If NOT EXISTS(SELECT * FROM db_etcmall.dbo.tbl_wemake_dealOption WHERE itemid = '"& itemid &"' and itemoption = '"& spitemoptionarr(i) &"') "& VBCRLF
			sqlStr = sqlStr & "		BEGIN "& VBCRLF
			sqlStr = sqlStr & "			INSERT INTO db_etcmall.dbo.tbl_wemake_dealOption (itemid, itemoption, quantity, regdate) "& VBCRLF
			sqlStr = sqlStr & "			VALUES('" & itemid & "', '"& spitemoptionarr(i) &"', '"& spoptionCountArr(i) &"', getdate()) "& VBCRLF
			sqlStr = sqlStr & "		END	"& VBCRLF
			sqlStr = sqlStr & "	ELSE "& VBCRLF
			sqlStr = sqlStr & "		BEGIN "& VBCRLF
			sqlStr = sqlStr & "			UPDATE db_etcmall.dbo.tbl_wemake_dealOption SET "& VBCRLF
			sqlStr = sqlStr & "			quantity = '"& spoptionCountArr(i) &"' "& VBCRLF
			sqlStr = sqlStr & "			,lastupdate = getdate() "& VBCRLF
			sqlStr = sqlStr & "			WHERE itemid = '"& itemid &"' and itemoption = '"& spitemoptionarr(i) &"' "& VBCRLF
			sqlStr = sqlStr & "		END	"
			dbget.execute sqlStr
		End If
		limitCount = limitCount + spoptionCountArr(i)
	Next

	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_etcmall.dbo.tbl_wemake_dealItem SET " & VBCRLF
	sqlStr = sqlStr & " limitCount = '"& limitCount &"' " & VBCRLF
	sqlStr = sqlStr & " WHERE itemid = '"& itemid &"' " & VBCRLF
	dbget.Execute(sqlStr)
	Response.Write "<script language=javascript>alert('저장 하였습니다.');opener.location.reload();self.close();</script>"
	dbget.close()	:	response.End
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->