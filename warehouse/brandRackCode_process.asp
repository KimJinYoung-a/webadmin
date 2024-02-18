<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/RackCodeFunction.asp"-->
<%

dim makerid, prtidx, mode, itembarcode, rackboxno
dim makeridArr, warehouseCdArr, warehouseCd

makerid = requestCheckVar(request.Form("makerid"), 32)
prtidx  = requestCheckVar(request.Form("prtidx"), 8)
mode    = requestCheckVar(request.Form("mode"), 32)
itembarcode = requestCheckVar(request.Form("itembarcode"), 32)
rackboxno = requestCheckVar(request.Form("rackboxno"), 8)
makeridArr = requestCheckVar(request.Form("makeridArr"), 8000)
warehouseCdArr = requestCheckVar(request.Form("warehouseCdArr"), 8000)

dim refer
refer = request.ServerVariables("HTTP_REFERER")


dim sqlStr, i, j, k
dim prePrtIdx

Select Case mode
	Case "editprtidx"

        Call RF_SetBrandRackCode(makerid, prtidx)

		'' sqlStr = "select userid,IsNULL(prtidx,'9999') as prtidx from [db_user].[dbo].tbl_user_c" + VbCrlf
		'' sqlStr = sqlStr + " where userid='" + makerid + "'" + VbCrlf
		'' rsget.Open sqlStr, dbget, 1
		'' if Not rsget.Eof then
		'' 	prePrtIdx = rsget("prtidx")
		'' 	prePrtIdx = Format00(4,prePrtIdx)
		'' end if
		'' rsget.Close


		'' sqlStr = "update [db_user].[dbo].tbl_user_c" & VbCrlf
		'' sqlStr = sqlStr & " set prtidx='" & prtidx & "'" & VbCrlf
		'' sqlStr = sqlStr & " where userid='" & makerid & "'"

		'' if (makerid<>"") and (prtidx<>"") then
		'' 	dbget.Execute sqlStr
		'' end if


		'' '// 상품 랙코드 변경전에 보조랙코드(디폴트=상품랙코드) 저장, skyer9, 2016-11-11
		'' sqlStr = " insert into [db_item].[dbo].[tbl_item_logics_addinfo](itemid, subitemrackcode) "
		'' sqlStr = sqlStr + " select i.itemid, i.itemrackcode "
		'' sqlStr = sqlStr + " from "
		'' sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item i "
		'' sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
		'' sqlStr = sqlStr + " 	on "
		'' sqlStr = sqlStr + " 		i.itemid = a.itemid "
		'' sqlStr = sqlStr + " where "
		'' sqlStr = sqlStr + " 	1 = 1 "
		'' sqlStr = sqlStr + " 	and i.makerid = '" + makerid + "' "
		'' sqlStr = sqlStr + " 	and i.itemrackcode = '" + CStr(prePrtIdx) + "' "
		'' sqlStr = sqlStr + " 	and a.itemid is NULL "
		'' dbget.Execute sqlStr

		'' sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
		'' sqlStr = sqlStr + " set itemrackcode='" + CStr(prtidx) + "'" + VbCrlf
		'' sqlStr = sqlStr + " where makerid='" + makerid + "'" + VbCrlf
		'' sqlStr = sqlStr + " and itemrackcode='" + CStr(prePrtIdx) + "'"

		'' dbget.Execute sqlStr


		response.write "<script>alert('수정 되었습니다.');</script>"
		response.write "<script>location.replace('/warehouse/pop_BrandRackCodeEdit.asp?makerid=" & makerid & "&itembarcode=" & itembarcode & "&fcs=itembarcode');</script>"
		dbget.close()	:	response.End
	Case "editrackboxno"
		sqlStr = "update [db_user].[dbo].tbl_user_c" & VbCrlf
		sqlStr = sqlStr & " set rackboxno='" & rackboxno & "'" & VbCrlf
		sqlStr = sqlStr & " where userid='" & makerid & "'"
		''response.write sqlStr
		dbget.Execute sqlStr

		response.write "<script>alert('수정 되었습니다.');</script>"
		response.write "<script>location.replace('/warehouse/pop_BrandRackCodeEdit.asp?makerid=" & makerid & "&itembarcode=" & itembarcode & "&fcs=itembarcode');</script>"
		dbget.close()	:	response.End
	Case "setwarehousecd"
        makeridArr = Split(makeridArr, ",")
        warehouseCdArr = Split(warehouseCdArr, ",")

        for i = 0 to UBOund(makeridArr)
            makerid = makeridArr(i)
            warehouseCd = warehouseCdArr(i)

            if Trim(makerid) <> "" then
		        sqlStr = "update [db_user].[dbo].tbl_user_c" & VbCrlf
		        sqlStr = sqlStr & " set warehouseCd='" & warehouseCd & "'" & VbCrlf
		        sqlStr = sqlStr & " where userid='" & makerid & "'"
		        ''response.write sqlStr & "<br />"
		        dbget.Execute sqlStr
            end if
        next

		response.write "<script>alert('수정 되었습니다.');</script>"
        response.write "<script>document.location.href = '" & refer & "';</script>"
		dbget.close()	:	response.End
	Case Else
		''
End Select

if (mode="editprtidx") then

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
