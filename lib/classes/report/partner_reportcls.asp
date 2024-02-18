<%
class CDesignerJumunList

 	public FMasterItemList()
    public Fselltotal
    public Fseldate
    public Fsellcnt
	public maxt
	public maxc
	public FResultCount
	public FCancelyn
    public FItemCount
	public FItemID
	public FItemName
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class



class CJumunMasterItem
	public FMasterItemList()
    public Fselltotal
    public Fseldate
    public Fsellcnt
	public maxt
	public maxc
	public FResultCount
    public FItemCount
	public FItemID
	public FItemName

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub


end Class

class CJumunMaster
	public FMasterItemList()
	public maxt
	public maxc
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FRectRegStart
	public FRectRegEnd
    public FRectItemid
	public FCurrPage
	public FRectSettle
	public FRectSettle2
	public FRectFromDate
	public FRectToDate
	public FRectIpkumDiv4
    public FRectDesignerID
    public FRectRdSite

    public FItemCount
	public FItemID
	public FItemName
	public FItemimgsmall
	public FTotalFavoriteCount
	public FSubtotal

	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		redim  FMasterItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	public sub SearchSellrePort()

    Dim sql, sqltmp, wheredetail, i

    maxt = -1
    maxc = -1

		if FRectsettle2="m" then
			sqltmp = " convert(varchar(7),regdate,120)"
		else
			sqltmp = " convert(varchar(10),regdate,120) "
		end if

		if (FRectsettle2="m") then
		wheredetail = " and (" & sqltmp & " >= '" & left(FRectFromDate,7) & "') and (" & sqltmp & " <='" & left(FRectToDate,7) & "')"
		else
		wheredetail = " and (" & sqltmp & " >= '" & FRectFromDate & "') and (" & sqltmp & " <='" & FRectToDate & "')"
		end if

		if (FRectIpkumDiv4<>"") then
			wheredetail = wheredetail + " and ipkumdiv>=4"
		else
		    wheredetail = wheredetail + " and ipkumdiv>1"
		end if
		wheredetail = wheredetail + " and cancelyn='N'"

		''#################################################
		''µ¥ÀÌÅ¸.
		''#################################################


			sql = "select " & sqltmp & " as ipdate,"
			sql = sql + " sum(subtotalprice) as sumtotal,"
			sql = sql + " count(orderserial) as sellcnt"
			sql = sql + " from [db_order].[dbo].tbl_order_master"
			if FRectRdSite<>"" then
				sql = sql + " where rdsite='" + FRectRdSite + "'"
			else
				sql = sql + " where sitename='" + FRectDesignerID + "'"
			end if

			sql = sql + wheredetail
            sql = sql + " Group by " + sqltmp

			rsget.Open sql,dbget,1

					FResultCount = rsget.RecordCount


		        redim preserve FMasterItemList(FResultCount)


					do until rsget.eof

				set FMasterItemList(i) = new CDesignerJumunList
						FMasterItemList(i).Fseldate = rsget("ipdate")
						FMasterItemList(i).Fselltotal = rsget("sumtotal")
						FMasterItemList(i).Fsellcnt = rsget("sellcnt")


						if Not IsNull(FMasterItemList(i).Fselltotal) then
							maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
							maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
						end if

						rsget.MoveNext
						i = i + 1
					loop

			rsget.close

	end sub


	public Sub SearchJumunListByfavitemlist()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""


		if FRectsettle = "self" then
		wheredetail = wheredetail + " and i.makerid='" + FRectDesignerID + "'"
		end if

		if (FRectRegStart<>"") then
			wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectItemid<>"") then
		wheredetail = wheredetail + " and i.itemid=" + FRectItemid
		end if


		''#################################################
		''ÃÑ °¹¼ö. ÃÑ±Ý¾×
		''#################################################
		sqlStr = "select DISTINCT count(m.itemid) as Titemid"
		sqlStr = sqlStr + " from  [db_user].[dbo].tbl_myfavorite m, [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where m.itemid = i.itemid"
		sqlStr = sqlStr + wheredetail


		rsget.Open sqlStr,dbget,1
		FSubtotal = rsget("Titemid")
		rsget.Close


		''#################################################
		''µ¥ÀÌÅ¸.
		''#################################################


		sqlStr = "select "
		sqlStr = sqlStr + "DISTINCT m.itemid, count(m.itemid) as itemcount, i.itemname, i.smallimage as imgsmall"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_myfavorite m, [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where m.itemid = i.itemid"
'		sqlStr = sqlStr + " and m.itemid<>0"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by m.itemid, i.itemname, i.smallimage"
		sqlStr = sqlStr + " order by itemcount desc"

'response.write sqlStr

		rsget.PageSize = FPageSize

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
'response.write FPageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster
	'			FMasterItemList(i).Forderserial = rsget("orderserial")
	'			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
	'			FMasterItemList(i).Fuserid		= rsget("userid")
	'			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
	'			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
	'			FMasterItemList(i).Fregdate		= rsget("regdate")
	'			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
	'			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
	'			FMasterItemList(i).Freqphone	= rsget("reqphone")
	'			FMasterItemList(i).Freqhp		= rsget("reqhp")
	'			FMasterItemList(i).Fdeliverno	= rsget("deliverno")
	'			FMasterItemList(i).Fsitename	= rsget("sitename")
	'			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
	'			FMasterItemList(i).FCancelyn	= rsget("cancelyn")
				FMasterItemList(i).FItemID       = rsget("itemid")
			    FMasterItemList(i).FItemCount       = rsget("itemcount")
    			FMasterItemList(i).FItemimgsmall      = rsget("imgsmall")
				FMasterItemList(i).FItemName     = rsget("itemname")
	'			FMasterItemList(i).FItemOption   = rsget("itemoption")


						if Not IsNull(FMasterItemList(i).FItemCount) then
							maxc = MaxVal(maxc,FMasterItemList(i).FItemCount)
						end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub SearchJumunListBybestfavitemlist()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""


		if FRectsettle = "self" then
		wheredetail = wheredetail + " and i.makerid='" + FRectDesignerID + "'"
		end if

		if (FRectItemid<>"") then
		wheredetail = wheredetail + " and i.itemid=" + FRectItemid
		end if


		''#################################################
		''ÃÑ °¹¼ö. ÃÑ±Ý¾×
		''#################################################

		sqlStr = "select sum(c.favcount) as favtotalcount"
		sqlStr = sqlStr + " from [db_const].[dbo].tbl_const_category c, [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where c.itemid = i.itemid"
		sqlStr = sqlStr + wheredetail
		rsget.Open sqlStr,dbget,1

		FSubtotal = rsget("favtotalcount")


		rsget.Close



		''#################################################
		''µ¥ÀÌÅ¸.
		''#################################################


		sqlStr = "select "
		sqlStr = sqlStr + "c.itemid, c.favcount, i.itemname, g.imgsmall"
		sqlStr = sqlStr + " from [db_const].[dbo].tbl_const_category c, [db_item].[dbo].tbl_item i, [db_item].[dbo].tbl_item_image g"
		sqlStr = sqlStr + " where c.itemid = i.itemid"
		sqlStr = sqlStr + " and c.itemid = g.itemid"
		sqlStr = sqlStr + wheredetail
'		sqlStr = sqlStr + " group by c.itemid, c.favcount, i.itemname, g.imgsmall"
		sqlStr = sqlStr + " order by c.favcount desc"


		rsget.PageSize = FPageSize

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if


		FPageCount = rsget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster
	'			FMasterItemList(i).Forderserial = rsget("orderserial")
	'			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
	'			FMasterItemList(i).Fuserid		= rsget("userid")
	'			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
	'			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
	'			FMasterItemList(i).Fregdate		= rsget("regdate")
	'			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
	'			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
	'			FMasterItemList(i).Freqphone	= rsget("reqphone")
	'			FMasterItemList(i).Freqhp		= rsget("reqhp")
	'			FMasterItemList(i).Fdeliverno	= rsget("deliverno")
	'			FMasterItemList(i).Fsitename	= rsget("sitename")
	'			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
	'			FMasterItemList(i).FCancelyn	= rsget("cancelyn")
				FMasterItemList(i).FItemID       = rsget("itemid")
			    FMasterItemList(i).FItemCount       = rsget("favcount")
    			FMasterItemList(i).FItemimgsmall      = rsget("imgsmall")
				FMasterItemList(i).FItemName     = rsget("itemname")
	'			FMasterItemList(i).FItemOption   = rsget("itemoption")

						if Not IsNull(FMasterItemList(i).FItemCount) then
							maxc = MaxVal(maxc,FMasterItemList(i).FItemCount)
						end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub


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
