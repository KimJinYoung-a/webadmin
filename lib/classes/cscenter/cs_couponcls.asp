<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'###########################################################

Class CCSCenterCouponItem
	public Fidx
	public Fmasteridx
	public Fuserid
	public Fcoupontype
	public Fcouponvalue
	public Fcouponname
	public Fminbuyprice
	public Ftargetitemlist
	public Fcouponimage
	public Fregdate
	public Fstartdate
	public Fexpiredate
	public Fisusing
	public Fdeleteyn
	public Forderserial
	public Fexitemid
	public Fvalidsitename
	public Fnotvalid10x10
	public Fcouponmeaipprice
	public Fssnkey
	public Fscratchcouponidx
	public FuseLevel

	public Freguserid
	public Fcsorderserial
	public FmxCpnDiscount

	public FprevCopiedCouponCount				'// 복사발행한 쿠폰 수

	public function GetCouponTypeName()
		if Fcoupontype="1" then
			GetCouponTypeName = "정률할인쿠폰"
		elseif Fcoupontype="2" then
			GetCouponTypeName = "정액할인쿠폰"
		else
			GetCouponTypeName = "-"
		end if
	end function

	public function GetCouponTypeUnit()
		if Fcoupontype="1" then
			GetCouponTypeUnit = "%"
		elseif Fcoupontype="2" then
			GetCouponTypeUnit = "원"
		else
			GetCouponTypeUnit = "-"
		end if
	end function

	public function GetCouponStatus()
		if (Fexpiredate < Now) then
			GetCouponStatus = "<font color='red'>유효기간 경과</font>"
		elseif (FprevCopiedCouponCount > 0) then
			GetCouponStatus = "<font color='red'>재발급불가(재발급은 1장만 가능)</font>"
		else
			GetCouponStatus = "재발급 가능(1장만 발급가능)"
		end if
	end function

	public function IsCouponCopyValid()
		if (Fexpiredate < Now) then
			IsCouponCopyValid = False
		elseif (FprevCopiedCouponCount > 0) then
			IsCouponCopyValid = False
		else
			IsCouponCopyValid = True
		end if
	end function

	Private Sub Class_Initialize()
		FprevCopiedCouponCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCSCenterCoupon
        public FItemList()
        public FOneItem

        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

        public FRectUserID
        public FRectRegUserID
        public FRectExcludeUnavailable
        public FRectExcludeDelete
        public FRectBonusCouponIdx

        public Sub GetCSCenterCouponList()
			dim i,sqlStr

			sqlStr = " select top 500 idx,masteridx,userid,couponname,couponvalue,coupontype, "
			sqlStr = sqlStr + "     minbuyprice,convert(varchar(19),startdate,21) as startdate,convert(varchar(19),expiredate,21) as expiredate,regdate,isusing,orderserial,exitemid,deleteyn,reguserid, csorderserial, uselevel "
			sqlStr = sqlStr + "		,isNull(mxCpnDiscount,0) as mxCpnDiscount "
			sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon with (nolock)"
			sqlStr = sqlStr + " where 1 = 1 "

			if (FRectUserID <> "") then
					sqlStr = sqlStr + " and userid='" + CStr(FRectUserID) + "' "
			end if

			if (FRectRegUserID <> "") then
					sqlStr = sqlStr + " and reguserid='" + CStr(FRectRegUserID) + "' "
			end if

			if (FRectExcludeUnavailable = "Y") then
					sqlStr = sqlStr + " and expiredate >= getdate() "
			end if

			if (FRectExcludeDelete = "Y") then
					sqlStr = sqlStr + " and deleteyn <> 'Y' "
			end if

			sqlStr = sqlStr + " order by idx desc "
			''response.write sqlStr
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

			FResultCount = rsget.RecordCount

			redim preserve FItemList(FResultCount)
			if not rsget.EOF then
				i = 0
				do until rsget.eof
				set FItemList(i) = new CCSCenterCouponItem

				FItemList(i).Fidx               = rsget("idx")
				FItemList(i).Fmasteridx         = rsget("masteridx")
				FItemList(i).Fuserid            = rsget("userid")
				FItemList(i).Fcouponname        = rsget("couponname")
				FItemList(i).Fcouponvalue       = rsget("couponvalue")
				FItemList(i).Fcoupontype        = rsget("coupontype")
				FItemList(i).Fminbuyprice       = rsget("minbuyprice")
				FItemList(i).Fstartdate         = rsget("startdate")
				FItemList(i).Fexpiredate        = rsget("expiredate")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Forderserial       = rsget("orderserial")
				FItemList(i).Fexitemid          = rsget("exitemid")
				FItemList(i).Fdeleteyn          = rsget("deleteyn")

				FItemList(i).Freguserid         = rsget("reguserid")
				FItemList(i).Fcsorderserial     = rsget("csorderserial")
				FItemList(i).FuseLevel			= rsget("uselevel")
				FItemList(i).FmxCpnDiscount		= rsget("mxCpnDiscount")

				rsget.MoveNext
				i = i + 1
				loop
			end if
			rsget.close
        end sub

		public Sub GetOneCSCenterCoupon
			dim sqlStr,i

	        sqlStr = " select top 1 idx,masteridx,userid,couponname,couponvalue,coupontype, "
	        sqlStr = sqlStr + "     minbuyprice,convert(varchar(19),startdate,21) as startdate,convert(varchar(19),expiredate,21) as expiredate,regdate,isusing,orderserial,exitemid,deleteyn,reguserid "
	        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon "
	        sqlStr = sqlStr + " where idx=" + CStr(FRectBonusCouponIdx)

			rsget.Open sqlStr, dbget, 1
			FResultCount = rsget.RecordCount

			set FOneItem = new CCSCenterCouponItem

			If not Rsget.Eof then

	            FOneItem.Fidx               = rsget("idx")
	            FOneItem.Fmasteridx         = rsget("masteridx")
	            FOneItem.Fuserid            = rsget("userid")
	            FOneItem.Fcouponname        = rsget("couponname")
	            FOneItem.Fcouponvalue       = rsget("couponvalue")
	            FOneItem.Fcoupontype        = rsget("coupontype")
	            FOneItem.Fminbuyprice       = rsget("minbuyprice")
	            FOneItem.Fstartdate         = rsget("startdate")
	            FOneItem.Fexpiredate        = rsget("expiredate")
	            FOneItem.Fregdate           = rsget("regdate")
	            FOneItem.Fisusing           = rsget("isusing")
	            FOneItem.Forderserial       = rsget("orderserial")
	            FOneItem.Fexitemid          = rsget("exitemid")
	            FOneItem.Fdeleteyn          = rsget("deleteyn")

	            FOneItem.Freguserid         = rsget("reguserid")

			end if
			rsget.close

			FOneItem.FprevCopiedCouponCount = 0
			if Not IsNull(FOneItem.Forderserial) and (FOneItem.Fmasteridx <> "") then
				'// 복사 발행된 쿠폰 있는지
				sqlStr = " select count(*) as cnt "
				sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and csorderserial = '" + CStr(FOneItem.Forderserial) + "' "
				sqlStr = sqlStr + " 	and masteridx = " + CStr(FOneItem.Fmasteridx) + " "
				sqlStr = sqlStr + " 	and masteridx <> 287 "										'// 기타 CS쿠폰(배송비 쿠폰 등)
				sqlStr = sqlStr + " 	and deleteyn <> 'Y' "
				sqlStr = sqlStr + " 	and userid = '" + CStr(FOneItem.Fuserid) + "' "

				rsget.Open sqlStr, dbget, 1
					FOneItem.FprevCopiedCouponCount = rsget("cnt")
				rsget.close
			end if
		end sub

        Private Sub Class_Initialize()
                FCurrPage       = 1
                FPageSize       = 20
                FResultCount    = 0
                FScrollCount    = 10
                FTotalCount     = 0
        End Sub

        Private Sub Class_Terminate()

        End Sub

        public Function HasPreScroll()
                HasPreScroll = StartScrollPage > 1
        end Function

        public Function HasNextScroll()
                HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
        end Function

        public Function StartScrollPage()
                StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
        end Function

end Class

Class CCSCenterItemCouponItem
	public fcouponidx
	public fuserid
	public fitemcouponidx
	public fissuedno
	public fitemcoupontype
	public fitemcouponvalue
	public fitemcouponstartdate
	public fitemcouponexpiredate
	public fitemcouponname
	public fitemcouponimage
	public fregdate
	public fusedyn
	public forderserial
	public fcouponGubun
	public fcsorderserial
	public FprevCopiedItemCouponCount				'// 복사발행한 상품 쿠폰수
	public Fopenstate

	public function GetDiscountStr()
		GetDiscountStr = CStr(Fitemcouponvalue) + GetItemCouponTypeName + " 할인"
	end function

	public function GetItemCouponTypeName
		Select Case Fitemcoupontype
			Case "1"
				GetItemCouponTypeName = "%"
			Case "2"
				GetItemCouponTypeName = "원"
			Case "3"
				GetItemCouponTypeName = "배송료"
			Case Else
				GetItemCouponTypeName = Fitemcoupontype
		end Select
	end function

	public function GetOpenStateName()
		Select Case Fopenstate
			case "0"
				GetOpenStateName = "발급대기"
			case "6"
				GetOpenStateName = "발급예약"
			case "7"
				GetOpenStateName = "오픈"
			case "9"
				GetOpenStateName = "발급강제종료"
			case else
				GetOpenStateName = Fopenstate
		end Select

    end function

	public function IsItemCouponCopyValid()
		if (fitemcouponexpiredate < dateconvert(Now())) then
			IsItemCouponCopyValid = False
		elseif (FprevCopiedItemCouponCount > 0) then
			IsItemCouponCopyValid = False
		else
			IsItemCouponCopyValid = True
		end if
	end function

	Private Sub Class_Initialize()
		FprevCopiedItemCouponCount = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CCSCenterItemCoupon
        public FItemList()
        public FOneItem
        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

		public frectcouponGubun
		public FRectorderserial
		public FRectuserid

        Private Sub Class_Initialize()
			redim  FItemList(0)

			FCurrPage       = 1
			FPageSize       = 20
			FResultCount    = 0
			FScrollCount    = 10
			FTotalCount     = 0
        End Sub
        Private Sub Class_Terminate()
        End Sub

        public Function HasPreScroll()
			HasPreScroll = StartScrollPage > 1
        end Function
        public Function HasNextScroll()
			HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
        end Function
        public Function StartScrollPage()
			StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
        end Function

		' 상품쿠폰정보	' 2023.10.16 한용민 생성
		public Sub GetCSCenterorderItemCoupon
			dim sqlStr,i, sqlsearch

			sqlsearch=""

			if FRectorderserial="" or isnull(FRectorderserial) then exit Sub

			if frectcouponGubun<>"" then
				sqlsearch = sqlsearch & " and c.couponGubun='"& frectcouponGubun &"'"
			end if
			if FRectuserid<>"" then
				sqlsearch = sqlsearch & " and c.userid='"& FRectuserid &"'"
			end if

	        sqlStr = " select top ("& FPageSize*FCurrPage &")"
			sqlStr = sqlStr & " c.couponidx, c.userid, c.itemcouponidx, c.issuedno, c.itemcoupontype, c.itemcouponvalue"
			sqlStr = sqlStr & " , convert(varchar(19),c.itemcouponstartdate,21) as itemcouponstartdate, convert(varchar(19),c.itemcouponexpiredate,21) as itemcouponexpiredate"
			sqlStr = sqlStr & " , c.itemcouponname, c.itemcouponimage, c.regdate, c.usedyn, c.orderserial, c.couponGubun, c.csorderserial"
			sqlStr = sqlStr & " , isnull((select count(cc.couponidx)"
			sqlStr = sqlStr & " 	from db_item.dbo.tbl_user_item_coupon cc with (nolock)"
			sqlStr = sqlStr & " 	where cc.itemcouponidx = c.itemcouponidx"
			sqlStr = sqlStr & " 	and cc.userid = c.userid"
			sqlStr = sqlStr & " 	and cc.couponidx <> c.couponidx),0) as prevCopiedItemCouponCount"
			sqlStr = sqlStr & " , cm.openstate"
	        sqlStr = sqlStr & " from db_item.dbo.tbl_user_item_coupon c with (nolock)"
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_coupon_master as cm with (noLock)"
			sqlStr = sqlStr & " 	on c.itemcouponidx=cm.itemcouponidx"
	        sqlStr = sqlStr & " where (c.orderserial='"& FRectorderserial &"' or c.csorderserial='"& FRectorderserial &"') " & sqlsearch
			sqlStr = sqlStr & " order by c.couponidx desc"

			'response.write sqlStr & "<br>"
			rsget.pagesize = FPageSize
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

			FTotalCount = rsget.RecordCount
			FResultCount = rsget.RecordCount
			FTotalPage = 1

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = 1
				do until rsget.eof
					set FItemList(i) = new CCSCenterItemCouponItem

					FItemList(i).fcouponidx = rsget("couponidx")
					FItemList(i).fuserid = rsget("userid")
					FItemList(i).fitemcouponidx = rsget("itemcouponidx")
					FItemList(i).fissuedno = rsget("issuedno")
					FItemList(i).fitemcoupontype = rsget("itemcoupontype")
					FItemList(i).fitemcouponvalue = rsget("itemcouponvalue")
					FItemList(i).fitemcouponstartdate = rsget("itemcouponstartdate")
					FItemList(i).fitemcouponexpiredate = rsget("itemcouponexpiredate")
					FItemList(i).fitemcouponname = rsget("itemcouponname")
					FItemList(i).fitemcouponimage = rsget("itemcouponimage")
					FItemList(i).fregdate = rsget("regdate")
					FItemList(i).fusedyn = rsget("usedyn")
					FItemList(i).forderserial = rsget("orderserial")
					FItemList(i).fcouponGubun = rsget("couponGubun")
					FItemList(i).fcsorderserial = rsget("csorderserial")
					FItemList(i).fprevCopiedItemCouponCount = rsget("prevCopiedItemCouponCount")
					FItemList(i).Fopenstate            = rsget("openstate")

					i=i+1
					rsget.moveNext
				loop
			end if
			rsget.Close
		end Sub
end Class
%>
