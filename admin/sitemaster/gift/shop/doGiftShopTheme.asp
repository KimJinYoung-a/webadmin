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
Dim menupos, mode
Dim themeIdx, subject, subDesc, isOpen, isPick, sortNo, isUsing, pickImage
dim arrKeyIdx, frontItemid, userid, adminid
dim sqlStr

menupos		= requestCheckVar(request("menupos"),8)
mode		= requestCheckVar(request("mode"),1)
themeIdx	= getNumeric(requestCheckVar(request("themeIdx"),10))
subject		= requestCheckVar(request("subject"),36)
subDesc		= requestCheckVar(request("subDesc"),80)
isOpen		= requestCheckVar(request("isOpen"),1)
isPick		= requestCheckVar(request("isPick"),1)
sortNo		= getNumeric(requestCheckVar(request("sortNo"),10))
isUsing		= requestCheckVar(request("isUsing"),1)

pickImage	= requestCheckVar(request("pickImage"),128)
arrKeyIdx	= requestCheckVar(request("arrKeyIdx"),24)
frontItemid	= getNumeric(requestCheckVar(request("frontItemid"),10))

if themeIdx="" and mode="u" then
	Call Alert_Return("테마 정보가 없습니다.")
	dbget.close(): response.End
end if

userid	= session("ssBctId")

if isPick="Y" then
	adminid = userid
elseif isPick="N" then
	pickImage=""
	adminid = ""
end if

Select Case mode
	Case "i"
		'// 테마 저장
		sqlStr = "Insert into db_board.dbo.tbl_giftShop_theme (subject, subDesc, isOpen, userid, adminid, itemCount, device, frontItemid, isPick, sortNo, isUsing, pickImage) "
		sqlStr = sqlStr & " values "
		sqlStr = sqlStr & "('" & subject & "'"
		sqlStr = sqlStr & ",'" & subDesc & "'"
		sqlStr = sqlStr & ",'" & isOpen & "'"
		sqlStr = sqlStr & ",'" & userID & "'"
		sqlStr = sqlStr & ",'" & adminid & "'"
		sqlStr = sqlStr & ",'0','W'"
		sqlStr = sqlStr & ",'" & frontItemid & "'"
		sqlStr = sqlStr & ",'" & isPick & "'"
		sqlStr = sqlStr & ",'" & sortNo & "'"
		sqlStr = sqlStr & ",'" & isUsing & "'"
		sqlStr = sqlStr & ",'" & pickImage & "')"
		dbget.Execute(sqlStr)

		'// 테마번호 접수
		sqlStr = "Select IDENT_CURRENT('db_board.dbo.tbl_giftShop_theme') as themeIdx "
		rsget.Open sqlStr,dbget,1
			themeIdx = rsget("themeIdx")
		rsget.close

		'// 키워드 저장
		sqlStr = "Insert into db_board.dbo.tbl_giftShop_theme_keyword (themeIdx,keywordIdx) "
		sqlStr = sqlStr & "Select " & themeIdx & ", keywordIdx "
		sqlStr = sqlStr & "From db_board.dbo.tbl_gift_keyword "
		sqlStr = sqlStr & "Where keywordIdx in (" & arrKeyIdx & ")"
		dbget.Execute(sqlStr)

	Case "u"
		'// 테마 수정
		sqlStr = "Update db_board.dbo.tbl_giftShop_theme "
		sqlStr = sqlStr & " Set "
		sqlStr = sqlStr & " subject='" & subject & "'"
		sqlStr = sqlStr & ",subDesc='" & subDesc & "'"
		sqlStr = sqlStr & ",isPick='" & isPick & "'"
		sqlStr = sqlStr & ",isOpen='" & isOpen & "'"
		sqlStr = sqlStr & ",sortNo='" & sortNo & "'"
		sqlStr = sqlStr & ",isUsing='" & isUsing & "'"
		sqlStr = sqlStr & ",frontItemid='" & frontItemid & "'"
		sqlStr = sqlStr & ",pickImage='" & pickImage & "'"
		sqlStr = sqlStr & " Where themeIdx=" & themeIdx
		dbget.Execute(sqlStr)

		'// 키워드 저장
		sqlStr = "Delete from db_board.dbo.tbl_giftShop_theme_keyword Where themeIdx=" & themeIdx & "; " & vbCrLf
		sqlStr = sqlStr & "Insert into db_board.dbo.tbl_giftShop_theme_keyword (themeIdx,keywordIdx) "
		sqlStr = sqlStr & "Select " & themeIdx & ", keywordIdx "
		sqlStr = sqlStr & "From db_board.dbo.tbl_gift_keyword "
		sqlStr = sqlStr & "Where keywordIdx in (" & arrKeyIdx & ")"
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


end Select

Response.Write "<script type='text/javascript'>" & vbCrLf
Response.Write "alert('저장되었습니다.');"& vbCrLf
Response.Write "location.href=""/admin/sitemaster/gift/shop/giftshop_themeList.asp?menupos=" & menupos & "&isPick=" & isPick & """;"& vbCrLf
Response.Write "</script>"
response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->