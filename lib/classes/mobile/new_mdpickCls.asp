<%
'###############################################
' PageName :mdpickCls
' Discription : 사이트 메인 mdpick
' History : 2014.01.28 이종화 생성
'###############################################

Class CmdpickItem
	public fidx
	Public fgubun
	Public fmdpicktitle
	Public Fstartdate
	Public Fenddate 
	Public Fadminid
	Public Flastadminid
	public Fisusing 
	Public Fregdate
	Public Fusername
	Public Flastupdate

	Public FsubIdx
	Public Flistidx
	Public Fsortnum
	Public FitemName
	Public FsmallImage
	public FFrontImage
	public FLowestPrice
	Public FItemid
	public FTentenImg '// tenten이미지

	public Forgprice
	public Fsailprice
	public Fsailyn
	public Fitemcouponyn
	public Fitemcoupontype
	public Fitemcouponvalue
	public Fsailsuplycash
	public Forgsuplycash
	public Fcouponbuyprice
	public FmwDiv
	public Fdeliverytype
	public Fsellcash
	public Fbuycash

	Public Fxmlregdate

	Public Ftopview

	public Fuserlevelgubun
	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// 세일포함 실제가격
	public Function getRealPrice() '!
		getRealPrice = FSellCash
	end Function

	'// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// 쿠폰 할인가
	public Function GetCouponDiscountPrice() '?
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function

	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function
	
	public function saleCouponPriceCheck(saleYN , couponYN , orgPrice , salePrice , couponType) 
		'할인가
		if saleYN="Y" then
			Response.Write " / <font color=#F08050>("&CLng((orgPrice-salePrice)/orgPrice*100) & "%할)" & FormatNumber(salePrice,0) & "</font>"
		end if
		'쿠폰가
		if couponYN="Y" then
			Select Case couponType
				Case "1"
					Response.Write " / <font color=#5080F0>(쿠)" & FormatNumber(GetCouponDiscountPrice(),0) & "</font>"
				Case "2"
					Response.Write " / <font color=#5080F0>(쿠)" & FormatNumber(GetCouponDiscountPrice(),0) & "</font>"
			end Select
		end if
	end function

	public function priceMarginCheck(saleYN , couponYN  , couponType , saleSuplyCash , salePrice , couponBuyPrice , buycash)
		if saleYN="Y" then
			Response.Write " / <font color=#F08050>" & fnPercent(saleSuplyCash,salePrice,1) & "</font>"
		end if
		'쿠폰가
		if couponYN="Y" then
			Select Case couponType
				Case "1"
					if couponBuyPrice=0 or isNull(couponBuyPrice) then
						Response.Write " / <font color=#5080F0>" & fnPercent(buycash,GetCouponAssignPrice(),1) & "</font>"
					else
						Response.Write " / <font color=#5080F0>" & fnPercent(couponBuyPrice,GetCouponAssignPrice(),1) & "</font>"
					end if
				Case "2"
					if couponBuyPrice=0 or isNull(couponBuyPrice) then
						Response.Write " / <font color=#5080F0>" & fnPercent(buycash,GetCouponAssignPrice(),1) & "</font>"
					else
						Response.Write " / <font color=#5080F0>" & fnPercent(couponBuyPrice,GetCouponAssignPrice(),1) & "</font>"
					end if
			end Select
		end if
	end function

	public function deliveryTypeName(deliveryType)
		Select Case deliveryType
			Case "1"
				response.write "텐배"
			Case "2"
				Response.Write "무료"
			Case "4"
				Response.Write "텐무"
			Case "9"
				Response.Write "조건"
			Case "7"
				Response.Write "착불"
		end Select
	end function 
end Class

Class Cmdpick
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
       
    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	Public FRectSubIdx
	Public FRectlistidx
	
	'//admin/mobile/tpobanner/tpo_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_mdpick_list "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CmdpickItem
        
        if Not rsget.Eof then
    		FOneItem.fidx			= rsget("idx")
    		FOneItem.fmdpicktitle	= rsget("mdpicktitle")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fadminid		= rsget("adminid")
			FOneItem.Flastadminid	= rsget("lastadminid")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Ftopview		= rsget("topview")
			FOneItem.Fuserlevelgubun= rsget("userlevelgubun")
        end If
        
        rsget.Close
    end Sub

	'//subitem
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = "Select top 1 s.*, i.itemname, i.smallImage "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.Itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        SqlStr = SqlStr + " where subIdx=" + CStr(FRectSubIdx)

		'rw SqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CmdpickItem
        if Not rsget.Eof then
            FOneItem.FsubIdx			= rsget("subIdx")
            FOneItem.Flistidx			= rsget("listIdx")
            FOneItem.FItemid			= rsget("Itemid")
            FOneItem.Fsortnum			= rsget("sortnum")
            FOneItem.Fisusing			= rsget("isusing")
            FOneItem.FitemName			= rsget("itemname")
            FOneItem.FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FOneItem.FItemid) & "/" & rsget("smallImage"),"")
			FOneItem.FFrontImage		= rsget("frontimg")
			FOneItem.FLowestPrice		= rsget("islowestprice")
			If ISNULL(FOneItem.FLowestPrice) Then
				FOneItem.FLowestPrice = ""
			End If
        end if
        rsget.close
	End Sub
	
	'//admin/mobile/cateimg/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_mobile_main_mdpick_list "
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >='" & Fsdt & " 00:00:00' and  EndDate <='" & Fsdt & " 23:59:59' "
		'if Fedt<>"" then sqlStr = sqlStr & " and  EndDate <='" & Fedt & " 23:59:59' "

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr + " * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_mdpick_list "
        sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >='" & Fsdt & " 00:00:00' and  EndDate <='" & Fsdt & " 23:59:59' "
		'if Fedt<>"" then sqlStr = sqlStr & " and  EndDate <='" & Fedt & " 23:59:59' "
        
		sqlStr = sqlStr + " order by  idx desc" 

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

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
				set FItemList(i) = new CmdpickItem
				
				FItemList(i).fidx				= rsget("idx")
				FItemList(i).fmdpicktitle		= rsget("mdpicktitle")
				FItemList(i).Fstartdate			= rsget("startdate")
				FItemList(i).Fenddate			= rsget("enddate")
				FItemList(i).Fadminid			= rsget("adminid")
				FItemList(i).Flastadminid		= rsget("lastadminid")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fxmlregdate		= rsget("xmlregdate")
				FItemList(i).Fuserlevelgubun	= rsget("userlevelgubun")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	'//admin/mobile/cateimg/index.asp
    public Sub GetContentsItemList()
       dim sqlStr, addSql, i

		sqlStr = " select count(listidx) as cnt from db_sitemaster.dbo.tbl_mobile_main_mdpick_item "
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr & " and  listidx='" & FRectlistidx & "'"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " s.subidx , s.listidx , s.itemid , s.isusing as itemusing , s.sortnum, isnull(s.itemname,i.itemname) as itemname , i.smallImage , s.gubun , i.itemdiv, s.frontimg, s.islowestprice "
		sqlStr = sqlStr + " , isnull(i.orgprice,0) as orgprice , isnull(i.sailprice,0) as sailprice , i.sailyn , i.itemcouponyn , i.itemcoupontype , isnull(i.sailsuplycash,0) as sailsuplycash , isnull(i.orgsuplycash ,0) as orgsuplycash " 
		sqlStr = sqlStr + " , Case i.itemCouponyn When 'Y' then ( Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid ) end as couponbuyprice , i.mwdiv , i.deliverytype , isnull(i.sellcash,0) as sellcash  , isnull(i.buycash,0) as buycash , i.itemcouponvalue , i.tentenimage50"
        sqlStr = sqlStr & " From [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        sqlStr = sqlStr & " Where listidx='" & FRectlistidx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		sqlStr = sqlStr + " order by sortnum asc" 

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

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
				set FItemList(i) = new CmdpickItem
				
				FItemList(i).FsubIdx				= rsget("subidx")
	            FItemList(i).Flistidx				= rsget("listidx")
	            FItemList(i).Fitemid				= rsget("itemid")
	            FItemList(i).Fsortnum				= rsget("sortnum")
	            FItemList(i).FIsUsing				= rsget("itemusing")
	            FItemList(i).FitemName				= rsget("itemname")
				If rsget("itemdiv") = "21" Then '// Deal 상품
	            FItemList(i).FsmallImage			= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & rsget("smallImage"),"")
				Else
				FItemList(i).FsmallImage			= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallImage"),"")
				End If 
				FItemList(i).Fgubun					= rsget("gubun")
				FItemList(i).FFrontImage				= rsget("frontimg")
				FItemList(i).FLowestPrice				= rsget("islowestprice")
				If ISNULL(FItemList(i).FLowestPrice) Then
					FItemList(i).FLowestPrice = ""
				End If

				FItemLIst(i).Forgprice 		= rsget("orgprice")
				FItemLIst(i).Fsailprice 	= rsget("sailprice")
				FItemLIst(i).Fsailyn 		= rsget("sailyn")
				FItemLIst(i).Fitemcouponyn  = rsget("itemcouponyn")
				FItemLIst(i).Fitemcoupontype= rsget("itemcoupontype")
				FItemLIst(i).Fsailsuplycash = rsget("sailsuplycash")
				FItemLIst(i).Forgsuplycash 	= rsget("orgsuplycash")
				FItemLIst(i).Fcouponbuyprice= rsget("couponbuyprice")
				FItemLIst(i).FmwDiv 		= rsget("mwDiv")
				FItemLIst(i).Fdeliverytype 	= rsget("deliverytype")
				FItemLIst(i).Fsellcash 		= rsget("sellcash")
				FItemLIst(i).Fbuycash 		= rsget("buycash")
				FItemLIst(i).Fitemcouponvalue = rsget("itemcouponvalue")

				if rsget("tentenimage50") <> "" then
					IF application("Svr_Info") = "Dev" THEN						
						FItemList(i).FTentenImg		= "http://testwebimage.10x10.co.kr/image/tenten50/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("tentenimage50")
					else
						FItemList(i).FTentenImg		= "http://webimage.10x10.co.kr/image/tenten50/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("tentenimage50")
					end if											
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

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

'// STAFF 이름 접수
public Function getStaffUserName(uid)
	if uid="" or isNull(uid) then
		exit Function
	end if

	Dim strSql
	strSql = "Select top 1 username From db_partner.dbo.tbl_user_tenbyten Where userid='" & uid & "'"
	rsget.Open strSql, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		getStaffUserName = rsget("username")
	end if
	rsget.Close
End Function
%>