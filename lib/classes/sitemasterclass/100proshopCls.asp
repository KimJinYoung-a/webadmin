<%

class C100ProShopItem

	public Fidx
	public Ftypegubun
	public Fitemid
	public Ftitleimg
	public Flistimg
	public Fstartdate
	public Fenddate
	public Fimg1
	public Fimg2
	public Fimg3
	public Fimg4
	public Fimg5
	public Fisusing
	public Fregdate
	public Fmileage

	public FCouponStartDate
	public FCouponExpireDate
	public FCouponName
	public Fminbuyprice
	public FCouponValue

	public Fcontents
	public Fisopen
	public Fdetailidx

    public Fmdname1
	public Fmdcomment1
	public Fmdname2
	public Fmdcomment2
	public Fmdname3
	public Fmdcomment3

	public FItemImageSmall

	public Fopendate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	Public function GetTypeName()
		if Ftypegubun="E" then
			GetTypeName="이벤트"
		elseif Ftypegubun="I" then
			GetTypeName="상품"
		end if
	end function

end Class

Class C100ProShopCouponRegItem
	public FITemID
	public FImgSmall
	public FRegCount
	public FUseCount
	public FTotalSellCount
	public FMatchCount


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class C100ProShop
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FrectMallType
	public FRectCDL

	public FRectIdx

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub getCouponOpenList()
		dim sql, i
		sql = "select s.itemid,  "
		sql = sql + " m.smallimage, IsNULL(T.dcnt,0) as dcnt, IsNULL(T.fcnt,0) as fcnt, IsNULL(K.regcnt,0) as regcnt, IsNULL(K.usecnt,0) as usecnt"
		sql = sql + " from "
		sql = sql + " [db_sitemaster].[dbo].tbl_100proshop s"

		sql = sql + " left join ( select s.itemid, count(c.idx) as regcnt,"
		sql = sql + " sum(case c.isusing when 'Y' then 1 else 0 end ) as usecnt"
		sql = sql + " from [db_sitemaster].[dbo].tbl_100proshop s, "
		sql = sql + " [db_user].[dbo].tbl_user_coupon c "
		sql = sql + " where s.evt_code=" + CStr(FRectIdx)
		sql = sql + " and c.masteridx=0"
		sql = sql + " and c.exitemid=s.itemid"
		sql = sql + " and c.deleteyn='N'"
		sql = sql + " group by s.itemid"
		sql = sql + " ) K on s.itemid=K.itemid"

		sql = sql + " left join [db_item].[dbo].tbl_item m"
		sql = sql + " on s.itemid=m.itemid"

		sql = sql + " left join (select d.itemid, sum(d.itemno) as dcnt, "
		sql = sql + " sum(case when m.ipkumdiv>6 then d.itemno else 0 end ) as fcnt "
		sql = sql + " from [db_order].[dbo].tbl_order_master m,"
		sql = sql + " [db_order].[dbo].tbl_order_detail d,"
		sql = sql + " [db_sitemaster].[dbo].tbl_100proshop s"
		sql = sql + " where m.orderserial=d.orderserial"
		sql = sql + " and m.regdate >= s.startdate"
		sql = sql + " and m.regdate < s.enddate"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and m.ipkumdiv>3"
		sql = sql + " and d.itemid=s.itemid"
		sql = sql + " and s.evt_code=" + CStr(FRectIdx)
		sql = sql + " and d.cancelyn<>'Y'"
		sql = sql + " and d.itemid<>0"
		sql = sql + " group by d.itemid"
		sql = sql + " ) as T on T.itemid=s.itemid"

		sql = sql + " where s.evt_code=" + CStr(FRectIdx)

''response.write sql
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			do until rsget.eof
				set FItemList(i) = new C100ProShopCouponRegItem

				FItemList(i).FItemID   = rsget("itemid")
				FItemList(i).FImgSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("smallimage")
				FItemList(i).FRegCount = rsget("regcnt")
				FItemList(i).FUseCount = rsget("usecnt")

				FItemList(i).FTotalSellCount = rsget("dcnt")
				FItemList(i).FMatchCount = rsget("fcnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub


	public Sub getMasterList()
		dim sql, i
		sql = " SELECT count(t1.evt_code) as cnt "&_
  			  " FROM [db_event].[dbo].[tbl_event] as t1 INNER JOIN  [db_event].[dbo].[tbl_event_display] as t2 On t1.evt_code = t2.evt_code "&_
			  " 	left OUTER join [db_sitemaster].[dbo].tbl_100proshop as t3 On t1.evt_code = t3.evt_code "&_
			  "		WHERE t1.evt_kind = 3 and t1.evt_using ='Y' "
		
		rsget.Open sql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sql = " SELECT  top " + CStr(FPageSize * FCurrPage) + " t1.evt_code,t2.evt_icon, t3.idx as detailidx, t3.itemid, t3.startdate, t3.enddate, t3.couponstartdate, t3.couponexpiredate"&_
			  " 	, t3.couponname, IsNull(t3.minbuyprice,0) as minbuyprice, IsNull(t3.couponvalue,0) as couponvalue ,t4.smallimage "&_
			  " FROM [db_event].[dbo].[tbl_event] as t1 INNER JOIN  [db_event].[dbo].[tbl_event_display] as t2 On t1.evt_code = t2.evt_code "&_
			  "  	left OUTER join [db_sitemaster].[dbo].tbl_100proshop as t3 On t1.evt_code = t3.evt_code left OUTER join "&_
			  "		[db_item].[dbo].tbl_item as t4  On t3.itemid = t4.itemid "&_
			  "	WHERE t1.evt_kind = 3 and t1.evt_using ='Y' ORDER BY t1.evt_code DESC, t3.idx "
		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new C100ProShopItem

				FItemList(i).Fidx          = rsget("evt_code")
				FItemList(i).Fdetailidx		= rsget("detailidx")		       
		        if Not(rsget("evt_icon")="" or isNull(rsget("evt_icon"))) then
		        	FItemList(i).Flistimg       = rsget("evt_icon")
		        else
		        	FItemList(i).Flistimg	= "http://fiximage.10x10.co.kr/camerashop/img50x50.gif"
		        end if

				FItemList(i).Fitemid   = rsget("itemid")
				FItemList(i).Fstartdate   = rsget("startdate")
				FItemList(i).Fenddate   = rsget("enddate")
				FItemList(i).FCouponStartDate = rsget("couponstartdate")
				FItemList(i).FCouponExpireDate = rsget("couponexpiredate")
				FItemList(i).FCouponName = db2html(rsget("couponname"))
				FItemList(i).Fminbuyprice = rsget("minbuyprice")
				FItemList(i).FCouponValue = rsget("couponvalue")
				if FItemList(i).FItemId <> "" then
				FItemList(i).FItemImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("smallimage")
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub
	
	
	public Sub list()
		dim sql, i

		sql = "select count(idx) as cnt "
		sql = sql + " from [db_sitemaster].[dbo].tbl_100proshop "
'		sql = sql + " where isusing = 'Y'"

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


		sql = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_sitemaster].[dbo].tbl_100proshop "
'		sql = sql + " where isusing = 'Y'"

		sql = sql + " order by idx desc "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new C100ProShopItem

				FItemList(i).Fidx          = rsget("idx")
				FItemList(i).Ftypegubun   = rsget("typegubun")
				FItemList(i).Fitemid   = rsget("itemid")
		        FItemList(i).Ftitleimg       = rsget("titleimg")
		        FItemList(i).Flistimg       = rsget("listimg")
				FItemList(i).Fstartdate   = rsget("startdate")
				FItemList(i).Fenddate   = rsget("enddate")
				FItemList(i).Fimg1   = "http://webimage.10x10.co.kr/image/100proshop/" + rsget("img1")
				FItemList(i).Fimg2   = "http://webimage.10x10.co.kr/image/100proshop/" + rsget("img2")
				FItemList(i).Fimg3   = "http://webimage.10x10.co.kr/image/100proshop/" + rsget("img3")
				FItemList(i).Fimg4   = "http://webimage.10x10.co.kr/image/100proshop/" + rsget("img4")
				FItemList(i).Fimg5   = "http://webimage.10x10.co.kr/image/100proshop/" + rsget("img5")

				FItemList(i).Fisusing   = rsget("isusing")
				FItemList(i).Fregdate      = rsget("regdate")

				FItemList(i).FCouponStartDate = rsget("couponstartdate")
				FItemList(i).FCouponExpireDate = rsget("couponexpiredate")
				FItemList(i).FCouponName = db2html(rsget("couponname"))
				FItemList(i).Fminbuyprice = rsget("minbuyprice")
				FItemList(i).FCouponValue = rsget("couponvalue")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub read(byVal v)
		dim sql, i
		IF idx = "" OR isNull(idx) THEN Exit Sub
		sql = "select top 1 * from [db_sitemaster].[dbo].tbl_100proshop "
'		sql = sql + " where isusing <> 'N' "
		sql = sql + " where (idx = " + CStr(v) + ") "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FTotalCount = rsget.RecordCount
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
			set FItemList(0) = new C100ProShopItem

			FItemList(0).Fidx          = rsget("idx")
			FItemList(0).Fitemid   = rsget("itemid")
			FItemList(0).Fstartdate   = rsget("startdate")
			FItemList(0).Fenddate   = rsget("enddate")

			FItemList(0).Fisusing   = rsget("isusing")
			FItemList(0).Fregdate      = rsget("regdate")

			FItemList(0).FCouponStartDate = rsget("couponstartdate")
			FItemList(0).FCouponExpireDate = rsget("couponexpiredate")
			FItemList(0).FCouponName = db2html(rsget("couponname"))
			FItemList(0).Fminbuyprice = rsget("minbuyprice")
			FItemList(0).FCouponValue = rsget("couponvalue")

			FItemList(0).Fmdname1      = db2html(rsget("mdname1"))
			FItemList(0).Fmdcomment1	= db2html(rsget("mdcomment1"))
			FItemList(0).Fmdname2      = db2html(rsget("mdname2"))
			FItemList(0).Fmdcomment2	= db2html(rsget("mdcomment2"))
			FItemList(0).Fmdname3      = db2html(rsget("mdname3"))
			FItemList(0).Fmdcomment3	= db2html(rsget("mdcomment3"))
		end if
		rsget.close
	end sub

	

	public Sub getItemList(byval eCode) '해당 100프로샵 상품 리스트 가져오기

		dim sql,i


		sql = " SELECT  s.idx, s.itemid " &_
					" ,s.couponname,s.couponvalue,s.coupontype,s.minbuyprice " &_
					" ,s.startdate,s.enddate,s.couponstartdate,s.couponexpiredate " &_
					" ,i.smallimage " &_
					" ,s.mdname1,s.mdcomment1 ,s.mdname2,s.mdcomment2 ,s.mdname3,s.mdcomment3 " &_
					" FROM [db_sitemaster].[dbo].tbl_100proshop s " &_
					" JOIN db_item.[dbo].tbl_item i " &_
					" 	on s.itemid = i.itemid " &_
					" WHERE s.evt_code='" & eCode & "' "


				rsget.open sql ,dbget ,1

				if not rsget.eof then

					FResultCount = rsget.RecordCount

					i = 0
					redim preserve FItemList(FResultCount)

						do until rsget.eof
							set FItemList(i) = new C100ProShopItem

							FItemList(i).Fidx       = rsget("idx")
							FItemList(i).Fitemid   	= rsget("itemid")
							FItemList(i).Fstartdate = rsget("startdate")
							FItemList(i).Fenddate   = rsget("enddate")

							FItemList(i).FCouponStartDate 	= rsget("couponstartdate")
							FItemList(i).FCouponExpireDate 	= rsget("couponexpiredate")
							FItemList(i).FCouponName 	= db2html(rsget("couponname"))
							FItemList(i).Fminbuyprice = rsget("minbuyprice")
							FItemList(i).FCouponValue = rsget("couponvalue")

							FItemList(i).Fmdname1			= db2html(rsget("mdname1"))
							FItemList(i).Fmdcomment1	= db2html(rsget("mdcomment1"))
							FItemList(i).Fmdname2			= db2html(rsget("mdname2"))
							FItemList(i).Fmdcomment2	= db2html(rsget("mdcomment2"))
							FItemList(i).Fmdname3			= db2html(rsget("mdname3"))
							FItemList(i).Fmdcomment3	= db2html(rsget("mdcomment3"))

							FItemList(i).FItemImageSmall ="http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("smallimage")

							rsget.movenext
							i=i+1
						loop
				end if
				rsget.close
	end Sub

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



class C100ProshopCommentItem

	public Fid
	public Fmasterid
	public Fuserid
	public Fcomments
	public Fisusing
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class C100ProshopComment
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FrectMallType
	public FRectCDL


	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub list(byval idx)
		dim sql, i

		sql = "select count(id) as cnt "
		sql = sql + " from [db_sitemaster].[dbo].tbl_100proshop_comment "
		sql = sql + " where masterid = '" + Cstr(idx) + "'"

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


		sql = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_sitemaster].[dbo].tbl_100proshop_comment "
		sql = sql + " where masterid = '" + Cstr(idx) + "'"
		sql = sql + " order by id desc"

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new C100ProshopCommentItem

				FItemList(i).Fid          = rsget("id")
				FItemList(i).Fmasterid   = rsget("masterid")
		        FItemList(i).Fuserid       = rsget("userid")
				FItemList(i).Fcomments   = db2html(rsget("comments"))
				FItemList(i).Fisusing   = rsget("isusing")
				FItemList(i).Fregdate      = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub read(byVal v)
		dim sql, i

		sql = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_sitemaster].[dbo].tbl_100proshop_comment "
'		sql = sql + " where isusing <> 'N' "
		sql = sql + " where id = " + CStr(v)

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FTotalCount = rsget.RecordCount
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new C100ProshopCommentItem

				FItemList(i).Fid          = rsget("id")
				FItemList(i).Fmasterid   = rsget("masterid")
		        FItemList(i).Fuserid       = rsget("userid")
				FItemList(i).Fcomments   = db2html(rsget("comments"))
				FItemList(i).Fisusing   = rsget("isusing")
				FItemList(i).Fregdate      = rsget("regdate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
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
