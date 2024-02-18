<%@ language = vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

Dim mode
Dim itemgubunarr, itemidarr, itemoptionarr, itemnoarr, barcodearr
Dim itemgubun, itemid, itemoption, itemno, barcode
dim affectedRows

mode 			= requestCheckvar(request("mode"),32)
barcodearr 		= requestCheckvar(request("barcodearr"),50000)
itemgubunarr 	= requestCheckvar(request("itemgubunarr"),15000)
itemidarr 		= requestCheckvar(request("itemidarr"),30000)
itemoptionarr 	= requestCheckvar(request("itemoptionarr"),20000)
itemnoarr 		= requestCheckvar(request("itemnoarr"),20000)

dim sqlStr, refer, i

refer = request.ServerVariables("HTTP_REFERER")

Select Case mode
	Case "setsell2y"
		itemgubunarr = split(itemgubunarr,"|")
		itemidarr	= split(itemidarr,"|")
		itemoptionarr = split(itemoptionarr,"|")

		for i = 0 to UBound(itemgubunarr) - 1
			itemgubun = Trim(itemgubunarr(i))
			itemid = Trim(itemidarr(i))
			itemoption = Trim(itemoptionarr(i))

			If (itemgubun <> "") Then
				If (itemgubun <> "10") Then
					dbget.Close
					response.End
				End If

				sqlStr = " update [db_item].[dbo].tbl_item "
				sqlStr = sqlStr + " set sellyn = 'Y', lastupdate = getdate() "
				sqlStr = sqlStr + " where itemid = " & itemid
				sqlStr = sqlStr + " and sellyn <> 'Y' "
				''response.Write sqlStr & "<br />"
				dbget.execute sqlStr

				response.Write "<script>alert('저장되었습니다.'); location.replace('" & refer & "');</script>"

			End If
		Next
    Case "setbulkstock"
		barcodearr = split(barcodearr,"|")
        itemgubunarr = split(itemgubunarr,"|")
		itemidarr	= split(itemidarr,"|")
		itemoptionarr = split(itemoptionarr,"|")
        itemnoarr = split(itemnoarr,"|")

		for i = 0 to UBound(itemgubunarr) - 1
			barcode = Trim(barcodearr(i))
            itemgubun = Trim(itemgubunarr(i))
			itemid = Trim(itemidarr(i))
			itemoption = Trim(itemoptionarr(i))
            itemno = Trim(itemnoarr(i))

			If (itemgubun <> "") Then
                affectedRows = 0

				sqlStr = " update [db_summary].[dbo].[tbl_current_agvstock_summary] "
				sqlStr = sqlStr + " set bulkstock = " & itemno & ", lastbulkstockdate = getdate() "
				sqlStr = sqlStr + " where itemgubun = '" & itemgubun & "' and itemid = '" & itemid & "' and itemoption = '" & itemoption & "' "
				''response.Write sqlStr & "<br />"
				dbget.execute sqlStr, affectedRows

                if (affectedRows < 1) then
                    sqlStr = " insert into [db_summary].[dbo].[tbl_current_agvstock_summary](itemgubun, itemid, itemoption, skuCd, agvstock, regdate, lastupdate, warehouseCd, totsysstock, errrealcheckno, bulkstock, lastbulkstockdate) "
                    sqlStr = sqlStr + " values('" & itemgubun & "', '" & itemid & "', '" & itemoption & "', '" & barcode & "', 0, getdate(), getdate(), 'BLK', 0, 0, " & itemno & ", getdate()) "
                    ''response.Write sqlStr & "<br />"
                    dbget.execute sqlStr, affectedRows
                end if
			End If
		Next

        response.Write "<script>alert('저장되었습니다.'); location.replace('" & refer & "');</script>"
    Case "setbulkstockerr"
        itemgubunarr = split(itemgubunarr,"|")
		itemidarr	= split(itemidarr,"|")
		itemoptionarr = split(itemoptionarr,"|")
        itemnoarr = split(itemnoarr,"|")

		for i = 0 to UBound(itemgubunarr) - 1
            itemgubun = Trim(itemgubunarr(i))
			itemid = Trim(itemidarr(i))
			itemoption = Trim(itemoptionarr(i))
            itemno = Trim(itemnoarr(i))

			If (itemgubun <> "") Then

                sqlStr = "exec [db_summary].[dbo].[sp_ten_realchekErr_Input_By_CurrentStock] '" & itemgubun & "'," & itemid & ",'" & itemoption & "'," & itemno & ",'" & session("ssBctID") & "'"
                ''response.write sqlStr & "<br />"
                dbget.Execute sqlStr

                ''한정수량 조절
                sqlStr = " exec [db_summary].[dbo].[sp_ten_limitSetByRealStock] '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
                ''response.write sqlStr & "<br />"
                dbget.Execute sqlStr, affectedRows

                ''한정 일시품절->판매조절
                if (itemgubun="10") and (affectedRows>0) then
                    sqlStr = " exec [db_summary].[dbo].sp_Ten_SellYnSetByLimitNo " & itemid
                    ''response.write sqlStr & "<br />"
                    dbget.Execute sqlStr
                end if
			End If
		Next

        response.Write "<script>alert('저장되었습니다.'); location.replace('" & refer & "');</script>"
	Case Else
		''
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
