<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/RackCodeFunction.asp"-->
<%

dim mode
dim itemgubunarr, itemidarr, itemoptionarr
dim itemrackcode
dim previtemgubun, previtemid

mode = request("mode")
itemrackcode = request("itemrackcode")
itemgubunarr = request("itemgubunarr")
itemidarr	= request("itemidadd")
itemoptionarr = request("itemoptionarr")

itemgubunarr = split(itemgubunarr,"|")
itemidarr	= split(itemidarr,"|")
itemoptionarr = split(itemoptionarr,"|")

dim i, j
dim sqlStr

Select Case mode
	Case "modiitem"
		'// 상품별 랙번호 지정
		for i = 0 to UBound(itemgubunarr) - 1
			if (Len(itemgubunarr(i)) > 0) then
				if (itemgubunarr(i) = "90") then
					previtemgubun = itemgubunarr(i)
					previtemid = itemidarr(i)

					sqlStr = "update db_shop.dbo.tbl_shop_item" + VBCrlf
					sqlStr = sqlStr + " set offitemrackcode='" + itemrackcode + "'" + VBCrlf
					sqlStr = sqlStr + " , updt=getdate()" + VBCrlf
					sqlStr = sqlStr + " where itemgubun='" & CStr(itemgubunarr(i)) & "' and shopitemid=" & CStr(itemidarr(i)) & " and itemoption='" & CStr(itemoptionarr(i)) & "' and IsNull(offitemrackcode, '') <> '" + itemrackcode + "' "
					''response.write sqlStr & "<br>"
					''response.end
					''dbget.Execute sqlStr

                    Call RF_SetItemRackCodeByOption(itemgubunarr(i), itemidarr(i), itemoptionarr(i), itemrackcode)
				elseif (itemgubunarr(i) = "10") then
					if (itemgubunarr(i) <> previtemgubun) or (itemidarr(i) <> previtemid) then
						previtemgubun = itemgubunarr(i)
						previtemid = itemidarr(i)

						sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
						sqlStr = sqlStr + " set itemrackcode='" + itemrackcode + "'" + VBCrlf
						sqlStr = sqlStr + " , lastupdate=getdate()" + VBCrlf
						sqlStr = sqlStr + " where itemid=" & CStr(itemidarr(i)) & " and itemrackcode <> '" + itemrackcode + "' "
						''response.write sqlStr & "<br>"
						''dbget.Execute sqlStr

                        Call RF_SetItemRackCode(itemgubunarr(i), itemidarr(i), itemrackcode)
					end if
				else
					response.write "<script>alert('온라인/오프라인(10/90) 상품만 수정가능합니다.')</script>"
					response.write "<br><br>온라인/오프라인(10/90) 상품만 수정가능합니다."
					response.end
				end if
			end if
		next

		response.write "<script>alert('수정되었습니다.'); opener.focus(); window.close();</script>"
	Case "modiopt"
		'// 옵션별 랙번호 지정
		for i = 0 to UBound(itemgubunarr) - 1
			if (Len(itemgubunarr(i)) > 0) then
				if (itemgubunarr(i) = "90") then
					sqlStr = "update db_shop.dbo.tbl_shop_item" + VBCrlf
					sqlStr = sqlStr + " set offitemrackcode='" + itemrackcode + "'" + VBCrlf
					sqlStr = sqlStr + " , updt=getdate()" + VBCrlf
					sqlStr = sqlStr + " where itemgubun='" & CStr(itemgubunarr(i)) & "' and shopitemid=" & CStr(itemidarr(i)) & " and itemoption='" & CStr(itemoptionarr(i)) & "' and IsNull(offitemrackcode, '') <> '" + itemrackcode + "' "
					''response.write sqlStr & "<br>"
					''response.end
					''dbget.Execute sqlStr

                    Call RF_SetItemRackCodeByOption(itemgubunarr(i), itemidarr(i), itemoptionarr(i), itemrackcode)
				elseif (itemgubunarr(i) = "10") then
					sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
					sqlStr = sqlStr + " set optrackcode=" + CHKIIF(itemrackcode="", "NULL", "'" + itemrackcode + "'") + " " + VBCrlf
					sqlStr = sqlStr + " where itemid=" & CStr(itemidarr(i)) & " and itemoption='" & CStr(itemoptionarr(i)) & "' and IsNull(optrackcode, '') <> '" + itemrackcode + "' "
					''response.write sqlStr & "<br>"
					''dbget.Execute sqlStr

                    Call RF_SetItemRackCodeByOption(itemgubunarr(i), itemidarr(i), itemoptionarr(i), itemrackcode)
				else
					response.write "<script>alert('온라인/오프라인(10/90) 상품만 수정가능합니다.')</script>"
					response.write "<br><br>온라인/오프라인(10/90) 상품만 수정가능합니다."
					response.end
				end if
			end if
		next

		response.write "<script>alert('수정되었습니다.'); opener.focus(); window.close();</script>"
	Case Else
		response.write "ERR"
End Select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
