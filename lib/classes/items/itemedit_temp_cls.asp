<%
'####################################################
' Description :  상품 임시 저장 클래스
' History : 2019.04.18 한용민 생성
'####################################################
%>
<%
class Citemedit_temp
    public fidx
    public fitemid
    public forgprice
    public fordertempstatus
    public fregdate
    public fisusing
    public fregadminid
    public ffailtype
	public fitemname
	public fitemname_10x10
	public forgprice_10x10
	public fsellcash_10x10
	public ftempitemid
	public ftempitemoption
	public fmakerid
	public fdispcatecode
	public fbuycash
	public fmwdiv
	public fdeliverytype
	public fitemoptionname
	public fbarcode
	public fupchemanagecode
	public frealitemid
	public frealitemoption
	public fcate_large
	public fcate_mid
	public fcate_small
	public fbuyitemname
	public fbuyitemoptionname
	public fbuycurrencyUnit
	public fbuyitemprice
	public fisbn13
	public fisbn13_10x10
	public ffrontmakerid
	public ffrontmakerid_10x10

	function GetOrderTempStatusName
		GetOrderTempStatusName = Fordertempstatus

		Select Case Fordertempstatus
			Case "0"
				GetOrderTempStatusName = "업로드실패"
			Case "1"
				GetOrderTempStatusName = "업로드완료"
			Case "9"
				GetOrderTempStatusName = "등록완료"
		end Select
	end function

	function GetFailTypeName
		GetFailTypeName = Ffailtype

		Select Case Ffailtype
			Case NULL
				GetFailTypeName = ""
			Case "U"
				GetFailTypeName = "업로드중복"
			Case "S"
				GetFailTypeName = "할인중"
			'Case "F"
			'	GetFailTypeName = "형식요류"
			'Case "B"
			'	GetFailTypeName = "바코드오류"
			'Case "J"
			'	GetFailTypeName = "계약없음"
		end Select
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Citemedit_templist
    public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectRegAdminID
    public FRectExcludeRegFinish
    public FRectCurrentInsertOnly

    ' /admin/itemmaster/pop_itemlist_excel_upload_edit.asp
	public Sub GetsuccessitemList()
		dim sqlStr, addSql, i

		if (FRectExcludeRegFinish = "Y") then
			addSql = addSql & " and t.ordertempstatus <> 9" & vbcrlf
		end if
		if (FRectRegAdminID <> "") then
			addSql = addSql & " and t.regadminid = '"& FRectRegAdminID &"'" & vbcrlf
		end if

		sqlStr = "select count(t.idx) as cnt, CEILING(CAST(Count(t.idx) AS FLOAT)/'"&FPageSize&"' ) as totPg" & vbcrlf
		sqlStr = sqlStr & " from db_temp.dbo.tbl_item_edit_temp t with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item i with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.itemid = i.itemid" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_contents c with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.itemid = c.itemid" & vbcrlf
		sqlStr = sqlStr & " where t.isusing = 'Y'" & vbcrlf
		sqlStr = sqlStr & " and t.ordertempstatus <> 0 " & addSql

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		if FTotalCount < 1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " t.idx, t.itemid, t.itemname, t.orgprice, t.ordertempstatus, t.regdate, t.isusing, t.regadminid, t.failtype" & vbcrlf
		sqlStr = sqlStr & " , c.isbn13 as isbn13_10x10, t.isbn13" & vbcrlf
		sqlStr = sqlStr & " , i.itemname as itemname_10x10, i.orgprice as orgprice_10x10, i.sellcash as sellcash_10x10" & vbcrlf
		sqlStr = sqlStr & " , i.frontmakerid as frontmakerid_10x10, t.frontmakerid as frontmakerid" & vbcrlf
		sqlStr = sqlStr & " from db_temp.dbo.tbl_item_edit_temp t with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item i with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.itemid = i.itemid" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_contents c with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.itemid = c.itemid" & vbcrlf
		sqlStr = sqlStr & " where t.isusing = 'Y'" & vbcrlf
		sqlStr = sqlStr & " and t.ordertempstatus <> 0 " & addSql
		sqlStr = sqlStr & " order by t.idx desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Citemedit_temp

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).fitemid				= rsget("itemid")
				FItemList(i).forgprice				= rsget("orgprice")
				FItemList(i).fitemname_10x10				= db2html(rsget("itemname_10x10"))
				FItemList(i).fsellcash_10x10				= rsget("sellcash_10x10")
				FItemList(i).forgprice_10x10				= rsget("orgprice_10x10")
				FItemList(i).fordertempstatus				= rsget("ordertempstatus")
				FItemList(i).fregdate				= rsget("regdate")
				FItemList(i).fisusing				= rsget("isusing")
				FItemList(i).fregadminid				= rsget("regadminid")
				FItemList(i).ffailtype				= rsget("failtype")
				FItemList(i).fitemname				= db2html(rsget("itemname"))
				FItemList(i).fisbn13_10x10				= rsget("isbn13_10x10")
				FItemList(i).fisbn13				= rsget("isbn13")
				FItemList(i).ffrontmakerid_10x10				= rsget("frontmakerid_10x10")
				FItemList(i).ffrontmakerid				= rsget("frontmakerid")

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

    ' /admin/itemmaster/pop_itemlist_excel_upload_edit.asp
	public Sub GetFailList()
		dim sqlStr, addSql, i

		if (FRectExcludeRegFinish = "Y") then
			addSql = addSql & " and t.ordertempstatus <> 9" & vbcrlf
		end if
		if (FRectRegAdminID <> "") then
			addSql = addSql & " and t.regadminid = '"& FRectRegAdminID &"'" & vbcrlf
		end if

		sqlStr = "select count(t.idx) as cnt, CEILING(CAST(Count(t.idx) AS FLOAT)/'"&FPageSize&"' ) as totPg" & vbcrlf
		sqlStr = sqlStr & " from db_temp.dbo.tbl_item_edit_temp t with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item i with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.itemid = i.itemid" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_contents c with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.itemid = c.itemid" & vbcrlf
		sqlStr = sqlStr & " where t.isusing = 'Y'" & vbcrlf
		sqlStr = sqlStr & " and t.ordertempstatus = 0 " & addSql

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close
		
		if FTotalCount < 1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " t.idx, t.itemid, t.itemname, t.orgprice, t.ordertempstatus, t.regdate, t.isusing, t.regadminid, t.failtype" & vbcrlf
		sqlStr = sqlStr & " , c.isbn13 as isbn13_10x10, t.isbn13" & vbcrlf
		sqlStr = sqlStr & " , i.itemname as itemname_10x10" & vbcrlf
		sqlStr = sqlStr & " , i.frontmakerid as frontmakerid_10x10, t.frontmakerid as frontmakerid" & vbcrlf
		sqlStr = sqlStr & " from db_temp.dbo.tbl_item_edit_temp t with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item i with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.itemid = i.itemid" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_contents c with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.itemid = c.itemid" & vbcrlf
		sqlStr = sqlStr & " where t.isusing = 'Y'" & vbcrlf
		sqlStr = sqlStr & " and t.ordertempstatus = 0 " & addSql
		sqlStr = sqlStr & " order by t.idx desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Citemedit_temp

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).fitemid				= rsget("itemid")
				FItemList(i).forgprice				= rsget("orgprice")
				FItemList(i).fordertempstatus				= rsget("ordertempstatus")
				FItemList(i).fregdate				= rsget("regdate")
				FItemList(i).fisusing				= rsget("isusing")
				FItemList(i).fregadminid				= rsget("regadminid")
				FItemList(i).ffailtype				= rsget("failtype")
				FItemList(i).fitemname				= db2html(rsget("itemname"))
				FItemList(i).fitemname_10x10				= db2html(rsget("itemname_10x10"))
				FItemList(i).fisbn13_10x10				= rsget("isbn13_10x10")
				FItemList(i).fisbn13				= rsget("isbn13")
				FItemList(i).ffrontmakerid_10x10				= rsget("frontmakerid_10x10")
				FItemList(i).ffrontmakerid				= rsget("frontmakerid")

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

	' 엑셀 신규 상품 등록
	' /admin/itemmaster/pop_itemlist_excel_upload_reg.asp
	public Sub GetsuccessitemregList()
		dim sqlStr, addSql, i

		if (FRectExcludeRegFinish = "Y") then
			addSql = addSql & " and t.ordertempstatus <> 9" & vbcrlf
		end if
		if (FRectRegAdminID <> "") then
			addSql = addSql & " and t.regadminid = '"& FRectRegAdminID &"'" & vbcrlf
		end if

		sqlStr = "select count(t.idx) as cnt, CEILING(CAST(Count(t.idx) AS FLOAT)/'"&FPageSize&"' ) as totPg" & vbcrlf
		sqlStr = sqlStr & " from db_temp.dbo.tbl_item_reg_temp t with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where t.isusing = 'Y'" & vbcrlf
		sqlStr = sqlStr & " and t.ordertempstatus <> 0 " & addSql

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		if FTotalCount < 1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " t.idx, t.tempitemid, isnull(t.tempitemoption,'') as tempitemoption, t.makerid, t.dispcatecode" & vbcrlf
		sqlStr = sqlStr & " , replace(replace(replace( t.itemname ,char(9),''),char(10),''),char(13),'') as itemname, t.orgprice, t.buycash, t.mwdiv" & vbcrlf
		sqlStr = sqlStr & " , t.deliverytype, replace(replace(replace( isnull(t.itemoptionname,'') ,char(9),''),char(10),''),char(13),'') as itemoptionname" & vbcrlf
		sqlStr = sqlStr & " , t.barcode, t.upchemanagecode, t.ordertempstatus, t.regdate, t.isusing, t.regadminid, t.failtype, t.realitemid" & vbcrlf
		sqlStr = sqlStr & " , t.realitemoption, t.cate_large, t.cate_mid, t.cate_small, t.buyitemname, t.buyitemoptionname, t.buycurrencyUnit, t.buyitemprice" & vbcrlf
		sqlStr = sqlStr & " , t.frontmakerid" & vbcrlf
		sqlStr = sqlStr & " from db_temp.dbo.tbl_item_reg_temp t with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where t.isusing = 'Y'" & vbcrlf
		sqlStr = sqlStr & " and t.ordertempstatus <> 0 " & addSql
		sqlStr = sqlStr & " order by t.idx desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Citemedit_temp

				FItemList(i).Fidx = rsget("idx")
				FItemList(i).Ftempitemid = rsget("tempitemid")
				FItemList(i).Ftempitemoption = rsget("tempitemoption")
				FItemList(i).Fmakerid = rsget("makerid")
				FItemList(i).Fdispcatecode = rsget("dispcatecode")
				FItemList(i).Fitemname = db2html(rsget("itemname"))
				FItemList(i).Forgprice = rsget("orgprice")
				FItemList(i).Fbuycash = rsget("buycash")
				FItemList(i).Fmwdiv = rsget("mwdiv")
				FItemList(i).Fdeliverytype = rsget("deliverytype")
				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).Fbarcode = db2html(rsget("barcode"))
				FItemList(i).Fupchemanagecode = db2html(rsget("upchemanagecode"))
				FItemList(i).Fordertempstatus = rsget("ordertempstatus")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Fisusing = rsget("isusing")
				FItemList(i).Fregadminid = rsget("regadminid")
				FItemList(i).Ffailtype = rsget("failtype")
				FItemList(i).Frealitemid = rsget("realitemid")
				FItemList(i).Frealitemoption = rsget("realitemoption")
				FItemList(i).fcate_large = rsget("cate_large")
				FItemList(i).fcate_mid = rsget("cate_mid")
				FItemList(i).fcate_small = rsget("cate_small")
				FItemList(i).fbuyitemname = db2html(rsget("buyitemname"))
				FItemList(i).fbuyitemoptionname = db2html(rsget("buyitemoptionname"))
				FItemList(i).fbuycurrencyUnit = rsget("buycurrencyUnit")
				FItemList(i).fbuyitemprice = rsget("buyitemprice")
				FItemList(i).ffrontmakerid = rsget("frontmakerid")

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

	' 엑셀 신규 상품 등록
	' /admin/itemmaster/pop_itemlist_excel_upload_reg.asp
	public Sub GetregFailList()
		dim sqlStr, addSql, i

		if (FRectExcludeRegFinish = "Y") then
			addSql = addSql & " and t.ordertempstatus <> 9" & vbcrlf
		end if
		if (FRectRegAdminID <> "") then
			addSql = addSql & " and t.regadminid = '"& FRectRegAdminID &"'" & vbcrlf
		end if

		sqlStr = "select count(t.idx) as cnt, CEILING(CAST(Count(t.idx) AS FLOAT)/'"&FPageSize&"' ) as totPg" & vbcrlf
		sqlStr = sqlStr & " from db_temp.dbo.tbl_item_reg_temp t with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where t.isusing = 'Y'" & vbcrlf
		sqlStr = sqlStr & " and t.ordertempstatus = 0 " & addSql

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		if FTotalCount < 1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " t.idx, t.tempitemid, isnull(t.tempitemoption,'') as tempitemoption, t.makerid, t.dispcatecode" & vbcrlf
		sqlStr = sqlStr & " , replace(replace(replace( t.itemname ,char(9),''),char(10),''),char(13),'') as itemname, t.orgprice, t.buycash, t.mwdiv" & vbcrlf
		sqlStr = sqlStr & " , t.deliverytype, replace(replace(replace( isnull(t.itemoptionname,'') ,char(9),''),char(10),''),char(13),'') as itemoptionname" & vbcrlf
		sqlStr = sqlStr & " , t.barcode, t.upchemanagecode, t.ordertempstatus, t.regdate, t.isusing, t.regadminid, t.failtype, t.realitemid" & vbcrlf
		sqlStr = sqlStr & " , t.realitemoption, t.cate_large, t.cate_mid, t.cate_small, t.buyitemname, t.buyitemoptionname, t.buycurrencyUnit, t.buyitemprice" & vbcrlf
		sqlStr = sqlStr & " , t.frontmakerid" & vbcrlf
		sqlStr = sqlStr & " from db_temp.dbo.tbl_item_reg_temp t with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where t.isusing = 'Y'" & vbcrlf
		sqlStr = sqlStr & " and t.ordertempstatus = 0 " & addSql
		sqlStr = sqlStr & " order by t.idx desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Citemedit_temp

				FItemList(i).Fidx = rsget("idx")
				FItemList(i).Ftempitemid = rsget("tempitemid")
				FItemList(i).Ftempitemoption = rsget("tempitemoption")
				FItemList(i).Fmakerid = rsget("makerid")
				FItemList(i).Fdispcatecode = rsget("dispcatecode")
				FItemList(i).Fitemname = db2html(rsget("itemname"))
				FItemList(i).Forgprice = rsget("orgprice")
				FItemList(i).Fbuycash = rsget("buycash")
				FItemList(i).Fmwdiv = rsget("mwdiv")
				FItemList(i).Fdeliverytype = rsget("deliverytype")
				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).Fbarcode = db2html(rsget("barcode"))
				FItemList(i).Fupchemanagecode = db2html(rsget("upchemanagecode"))
				FItemList(i).Fordertempstatus = rsget("ordertempstatus")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Fisusing = rsget("isusing")
				FItemList(i).Fregadminid = rsget("regadminid")
				FItemList(i).Ffailtype = rsget("failtype")
				FItemList(i).Frealitemid = rsget("realitemid")
				FItemList(i).Frealitemoption = rsget("realitemoption")
				FItemList(i).fcate_large = rsget("cate_large")
				FItemList(i).fcate_mid = rsget("cate_mid")
				FItemList(i).fcate_small = rsget("cate_small")
				FItemList(i).fbuyitemname = db2html(rsget("buyitemname"))
				FItemList(i).fbuyitemoptionname = db2html(rsget("buyitemoptionname"))
				FItemList(i).fbuycurrencyUnit = rsget("buycurrencyUnit")
				FItemList(i).fbuyitemprice = rsget("buyitemprice")
				FItemList(i).ffrontmakerid = rsget("frontmakerid")

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

	Private Sub Class_Initialize()
		redim preserve FInsureList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>