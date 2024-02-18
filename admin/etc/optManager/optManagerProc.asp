<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sqlStr, i, j
Dim arrMallgubun : arrMallgubun = request("chkmall")
Dim cmdparam : cmdparam = requestCheckVar(request("mode"),1)
Dim spItemid, spOptioncode, mallSplit
Dim newitemname
Dim cksel

If cmdparam = "I" Then
	If Right(arrMallgubun,1) = "," Then arrMallgubun = Left(arrMallgubun, Len(arrMallgubun) - 1)
	mallSplit	= Split(arrMallgubun, ",")

	For j = 0 to Ubound(mallSplit)
		For i = 1 to request.form("cksel").count
			cksel 			= request.form("cksel")(i)
			If Trim(request.form("newitemname|"&cksel)) = "" Then
				newitemname = " i.itemname + ' ' + o.optionname "
			Else
				newitemname = " '"&Trim(request.form("newitemname|"&cksel))&"' "
			End If
			spItemid		= Trim(Split(cksel, "_")(0))
			spOptioncode	= Trim(Split(cksel, "_")(1))

			sqlStr = ""
			sqlStr = sqlStr & " IF Not Exists(SELECT * FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] where itemid='"&spItemid&"' and itemoption = '"&spOptioncode&"' and mallid = '"&Trim(mallSplit(j))&"') "
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " 	INSERT INTO db_etcmall.[dbo].[tbl_Outmall_option_Manager] (itemid, itemoption, optionname, mallid, regdate, reguserid, itemname, newitemname) " & VbCRLF
			sqlStr = sqlStr & "  	SELECT TOP 1 o.itemid, o.itemoption, o.optionname, '"&Trim(mallSplit(j))&"', getdate(), '"&session("ssBctID")&"', i.itemname, "&newitemname&"  " & VbCRLF
			sqlStr = sqlStr & "  	FROM db_item.dbo.tbl_item as i " & VbCRLF
			sqlStr = sqlStr & "  	JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid " & VbCRLF
			sqlStr = sqlStr & "  	WHERE i.itemid = '"&spItemid&"'  " & VbCRLF
			sqlStr = sqlStr & "  	and o.itemoption = '"&spOptioncode&"'  " & VbCRLF
			sqlStr = sqlStr & " END "
		    dbget.Execute sqlStr
		Next
	Next
	response.write 	"<script language='javascript'>alert('저장 되었습니다');parent.location.reload();</script>"
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->