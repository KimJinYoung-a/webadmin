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
Dim themeIdx, arrItemid, actItemid
dim tmpArrIid

themeIdx = request("themeIdx")
mode = request("mode")
arrItemid = split(replace(request("subItemidArray"),vbCrLf,","),",")
if trim(request("itemidarr"))<>"" then
	tmpArrIid = trim(request("itemidarr"))
	if Right(tmpArrIid,1)="," then tmpArrIid=Left(tmpArrIid,Len(tmpArrIid)-1)
	arrItemid = split(tmpArrIid,",")
end if

if themeIdx="" then
	Call Alert_Return("테마 정보가 없습니다.")
	dbget.close(): response.End
end if

if Not(isArray(arrItemid)) then
	Call Alert_Return("상품코드 정보가 잘못되었습니다.")
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
		'// 상품추가
	    sqlStr = " insert into db_board.dbo.tbl_giftShop_theme_item" + VbCrlf
	    sqlStr = sqlStr + " (themeIdx, itemid) " + VbCrlf
	    sqlStr = sqlStr + " select '" & themeIdx & "'" + VbCrlf
	    sqlStr = sqlStr + " ,itemid " + VbCrlf
	    sqlStr = sqlStr + " from db_item.dbo.tbl_item" + VbCrlf
	    sqlStr = sqlStr + " where itemid in (" & actItemid & ")" + VbCrlf
	    sqlStr = sqlStr + " 	and itemid not in (" + VbCrlf
	    sqlStr = sqlStr + " 		select itemid" + VbCrlf
	    sqlStr = sqlStr + " 		from db_board.dbo.tbl_giftShop_theme_item" + VbCrlf
	    sqlStr = sqlStr + " 		where themeIdx='" & themeIdx & "'" + VbCrlf
	    sqlStr = sqlStr + " 	)" + VbCrlf
		dbget.Execute(sqlStr)

		'// 테마 상품정보 업데이트
		sqlStr = "Update m "
		sqlStr = sqlStr & "Set m.itemCount=d.cnt "
		sqlStr = sqlStr & "From db_board.dbo.tbl_giftShop_theme as m "
		sqlStr = sqlStr & "	join ( "
		sqlStr = sqlStr & "		Select themeIdx, count(themeIdx) as cnt "
		sqlStr = sqlStr & "		From db_board.dbo.tbl_giftShop_theme_item "
		sqlStr = sqlStr & "		Where themeIdx=" & themeIdx
		sqlStr = sqlStr & "		group by themeIdx "
		sqlStr = sqlStr & "	) as d "
		sqlStr = sqlStr & "		on m.themeIdx=d.themeIdx "
		dbget.Execute(sqlStr)

		'// 상품정보 업데이트
		sqlStr = "Update f "
		sqlStr = sqlStr & "set f.themeCount=c.cnt "
		sqlStr = sqlStr & "From db_board.dbo.tbl_gift_itemInfo as f "
		sqlStr = sqlStr & "	join ( "
		sqlStr = sqlStr & "		Select d.itemid, sum(Case When m.isOpen='Y' and m.isUsing='Y' Then 1 Else 0 end) as cnt "
		sqlStr = sqlStr & "		from db_board.dbo.tbl_giftShop_theme as m "
		sqlStr = sqlStr & "			join db_board.dbo.tbl_giftShop_theme_item as d "
		sqlStr = sqlStr & "				on m.themeIdx=d.themeIdx "
		sqlStr = sqlStr & "		where d.itemid in ( "
		sqlStr = sqlStr & "				Select itemid "
		sqlStr = sqlStr & "				from db_board.dbo.tbl_giftShop_theme_item "
		sqlStr = sqlStr & "				where themeIdx=" & themeIdx
		sqlStr = sqlStr & "			) "
		sqlStr = sqlStr & "		group by d.itemid "
		sqlStr = sqlStr & "	) as c "
		sqlStr = sqlStr & "		on f.itemid=c.itemid "
		dbget.Execute(sqlStr)

		sqlStr = "insert into db_board.dbo.tbl_gift_itemInfo (itemid,themeCount) "
		sqlStr = sqlStr & "Select i.itemid, c.cnt "
		sqlStr = sqlStr & "from db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & "	join ( "
		sqlStr = sqlStr & "		Select d.itemid, count(d.themeIdx) as cnt "
		sqlStr = sqlStr & "		from db_board.dbo.tbl_giftShop_theme as m "
		sqlStr = sqlStr & "			join db_board.dbo.tbl_giftShop_theme_item as d "
		sqlStr = sqlStr & "				on m.themeIdx=d.themeIdx "
		sqlStr = sqlStr & "		where m.isOpen='Y' and m.isUsing='Y' "
		sqlStr = sqlStr & "			and m.themeIdx=" & themeIdx
		sqlStr = sqlStr & "		group by d.itemid "
		sqlStr = sqlStr & "	) as c "
		sqlStr = sqlStr & "		on i.itemid=c.itemid "
		sqlStr = sqlStr & "where i.itemid not in ( "
		sqlStr = sqlStr & "		Select itemid "
		sqlStr = sqlStr & "		from db_board.dbo.tbl_gift_itemInfo "
		sqlStr = sqlStr & "	) "
		dbget.Execute(sqlStr)

	elseif mode="d" then
		'// 상품삭제
	    sqlStr = " delete from db_board.dbo.tbl_giftShop_theme_item" + VbCrlf
	    sqlStr = sqlStr + " Where themeIdx='" & themeIdx & "'" + VbCrlf
	    sqlStr = sqlStr + " 	and itemid in (" & actItemid & ")" + VbCrlf
		dbget.Execute(sqlStr)

		'// 테마 상품정보 업데이트
		sqlStr = "Update m "
		sqlStr = sqlStr & "Set m.itemCount=d.cnt "
		sqlStr = sqlStr & "From db_board.dbo.tbl_giftShop_theme as m "
		sqlStr = sqlStr & "	join ( "
		sqlStr = sqlStr & "		Select themeIdx, count(themeIdx) as cnt "
		sqlStr = sqlStr & "		From db_board.dbo.tbl_giftShop_theme_item "
		sqlStr = sqlStr & "		Where themeIdx=" & themeIdx
		sqlStr = sqlStr & "		group by themeIdx "
		sqlStr = sqlStr & "	) as d "
		sqlStr = sqlStr & "		on m.themeIdx=d.themeIdx "
		dbget.Execute(sqlStr)

		'// 상품정보 업데이트
		sqlStr = "Update f "
		sqlStr = sqlStr & "set f.themeCount=c.cnt "
		sqlStr = sqlStr & "From db_board.dbo.tbl_gift_itemInfo as f "
		sqlStr = sqlStr & "	join ( "
		sqlStr = sqlStr & "		Select d.itemid, sum(Case When m.isOpen='Y' and m.isUsing='Y' Then 1 Else 0 end) as cnt "
		sqlStr = sqlStr & "		from db_board.dbo.tbl_giftShop_theme as m "
		sqlStr = sqlStr & "			join db_board.dbo.tbl_giftShop_theme_item as d "
		sqlStr = sqlStr & "				on m.themeIdx=d.themeIdx "
		sqlStr = sqlStr & "		where d.itemid in (" & actItemid & ") "
		sqlStr = sqlStr & "		group by d.itemid "
		sqlStr = sqlStr & "	) as c "
		sqlStr = sqlStr & "		on f.itemid=c.itemid "
		dbget.Execute(sqlStr)

	end if

end if

strRst = "[" & Scnt & "]건 성공"
if Ecnt>0 then strRst = strRst & "\n[" & Ecnt & "]건 실패\n※실패건: " & strErr

Response.Write "<script language='javascript'>" & vbCrLf
Response.Write "alert('" & strRst & "\n저장되었습니다.');"& vbCrLf
	if trim(request("itemidarr"))="" then
		Response.Write "opener.location.reload();" & vbCrLf
		Response.Write "window.close();"& vbCrLf
	end if
Response.Write "</script>"
response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->