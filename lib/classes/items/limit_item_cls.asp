<%

'옵션이 있는 상품에 대해, 옵션 한정수량이 변경된것을 상품테이블에 반영하는 함수
function UpdateItemLimitSum(byval itemid)
    dim i, sqlStr

    i = 0
    sqlStr = " select top 1 optioncnt from [db_item].[dbo].tbl_item where itemid = " + CStr(itemid) + " "
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        i = rsget("optioncnt")
    end if
    rsget.close

    if (i > 0) then
        sqlStr = " update [db_item].[dbo].tbl_item "
        sqlStr = sqlStr + " set optioncnt = T.optcnt, limitno = T.optlimitno, limitsold = T.optlimitsold "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + "     select count(itemid) as optcnt, sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold "
        sqlStr = sqlStr + "     from [db_item].[dbo].tbl_item_option where itemid = " + CStr(itemid) + " and isusing = 'Y' "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "
        'response.write sqlStr
        rsget.Open sqlStr,dbget,1
    end if

    UpdateItemLimitSum = i
end function

'특정상품의 특정옵션의 수량을 수정하고, 상품테이블을 업데이트한다.
function UpdateItemLimitCount(byval itemid, byval itemoption, byval limitno, byval limitsold)
    dim sqlStr

    if (itemoption = "0000") then
        sqlStr = " update [db_item].[dbo].tbl_item "
        sqlStr = sqlStr + " set limitno = " + CStr(limitno) + ", limitsold = " + CStr(limitsold) + " "
        sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "
        'response.write sqlStr
        rsget.Open sqlStr,dbget,1

        UpdateItemLimitCount = 0
    else
        sqlStr = " update [db_item].[dbo].tbl_item_option "
        sqlStr = sqlStr + " set optlimitno = " + CStr(limitno) + ", optlimitsold = " + CStr(limitsold) + " "
        sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
        'response.write sqlStr
        rsget.Open sqlStr,dbget,1

        UpdateItemLimitCount = UpdateItemLimitSum(itemid)
    end if
end function

'특정상품의 전달된 옵션들의 수량을 수정하고, 상품테이블을 업데이트한다.
function UpdateItemLimitArrayCount(byval itemid, byval arritemoption, byval arrlimitno, byval arrlimitsold)
    dim sqlStr, i

    arritemoption   = SPlit(arritemoption, "|")
    arrlimitno      = SPlit(arrlimitno, "|")
    arrlimitsold    = SPlit(arrlimitsold, "|")

    for i = 0 to UBound(arritemoption) - 1
        if (Trim(arritemoption(i)) <> "") then
            if (Trim(arritemoption(i)) = "0000") then
                sqlStr = " update [db_item].[dbo].tbl_item "
                sqlStr = sqlStr + " set limitno = " + CStr(Trim(arrlimitno(i))) + ", limitsold = " + CStr(Trim(arrlimitsold(i))) + " "
                sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "
                'response.write sqlStr
                rsget.Open sqlStr,dbget,1

                exit for
            else
                sqlStr = " update [db_item].[dbo].tbl_item_option "
                sqlStr = sqlStr + " set optlimitno = " + CStr(Trim(arrlimitno(i))) + ", optlimitsold = " + CStr(Trim(arrlimitsold(i))) + " "
                sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " and itemoption = '" + CStr(Trim(arritemoption(i))) + "' "
                'response.write sqlStr
                rsget.Open sqlStr,dbget,1
            end if
        end if
    next

    UpdateItemLimitArrayCount = UpdateItemLimitSum(itemid)
end function

'특정상품의 특정옵션의 한정합계수량을 추가하고, 상품테이블을 업데이트한다.
function AddItemLimitNo(byval itemid, byval itemoption, byval limitno)
    dim sqlStr

    if (itemoption = "0000") then
        sqlStr = " update [db_item].[dbo].tbl_item "
        sqlStr = sqlStr + " set limitno = limitno + " + CStr(limitno) + " "
        sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "
        'response.write sqlStr
        rsget.Open sqlStr,dbget,1

        AddItemLimitNo = 0
    else
        sqlStr = " update [db_item].[dbo].tbl_item_option "
        sqlStr = sqlStr + " set optlimitno = optlimitno + " + CStr(limitno) + " "
        sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
        'response.write sqlStr
        rsget.Open sqlStr,dbget,1

        AddItemLimitNo = UpdateItemLimitSum(itemid)
    end if
end function

'특정상품의 특정옵션의 한정판매수량을 추가하고, 상품테이블을 업데이트한다.
function AddItemLimitSold(byval itemid, byval itemoption, byval limitsold)
    dim sqlStr

    if (itemoption = "0000") then
        sqlStr = " update [db_item].[dbo].tbl_item "
        sqlStr = sqlStr + " set limitsold = limitsold + " + CStr(limitsold) + " "
        sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "
        'response.write sqlStr
        rsget.Open sqlStr,dbget,1

        AddItemLimitSold = 0
    else
        sqlStr = " update [db_item].[dbo].tbl_item_option "
        sqlStr = sqlStr + " set optlimitsold = optlimitsold + " + CStr(limitsold) + " "
        sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
        'response.write sqlStr
        rsget.Open sqlStr,dbget,1

        AddItemLimitSold = UpdateItemLimitSum(itemid)
    end if
end function

class CLimitItemItem
	public FItemID
	public FItemName
	public FMakerID
	public FDisplayYn
	public FSellYn
	public FBaesongGB
	public FLimitYn
	public FLimitDiv
	public FLimitNo
	public FLimitSold

	public FImageSmall
	public FMwDiv

	public Fsailyn

	public Fisusing
	public FitemOption
	public FitemOptionName
	public FitemOptionCount
	public FCurrNo
	public FSellCnt


	public function GetMwDivName
		if Fmwdiv="M" then
			GetMwDivName = "매입"
		elseif Fmwdiv="W" then
			GetMwDivName = "위탁"
		elseif Fmwdiv="U" then
			GetMwDivName = "업체배송"
		end if
	end function

	public function GetMwDivColor
		if Fmwdiv="M" then
			GetMwDivColor = "#FF0000"
		elseif Fmwdiv="W" then
			GetMwDivColor = "#0000FF"
		elseif Fmwdiv="U" then
			GetMwDivColor = "#000000"
		end if
	end function

	public function IsSoldOut()
		IsSoldOut = (FSellYn<>"Y") or (FDisplayYn<>"Y") or ((FLimitYn<>"N") and (FLimitNo-FLimitSold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CLimitItem
	public FItemList()

	public FSearchItemid
	public FSearchItemName
	public FSearchDesigner
	public FSearchDispYn
	public FSearchSellYn
	public FSearchLimitYn
	public FSearchUsingYn
	public FSearchSailYn
	public FSearchBaedalDiv

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectOnlyTenBeasong
	public FRectMwDiv

	Private Sub Class_Initialize()
	redim  FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function GetImageFolerName(byval i)
		GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public sub GetLimitItemList()
		dim sqlStr
		dim sqlrect
		dim i

        sqlStr = " select count(i.itemid) as cnt "
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
        'sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on i.itemid = o.itemid "
        sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on i.itemid = o.itemid and isnull(o.itemoption,'0000') <> '0000' and o.isusing <> 'N' "
        sqlStr = sqlStr + " where i.itemid<>0"


		if (FSearchItemid<>"") then
			sqlrect = sqlrect + " and i.itemid=" + CStr(FSearchItemid)
		end if

		if (FSearchItemName<>"") then
			sqlrect = sqlrect + " and i.itemname like '%" + CStr(FSearchItemname) + "%'"
		end if

		if (FSearchDesigner<>"") then
			sqlrect = sqlrect + " and i.makerid = '" + CStr(FSearchDesigner) + "'"
		end if

		if (FSearchDispYn<>"") then
			sqlrect = sqlrect + " and i.dispyn = '" + CStr(FSearchDispYn) + "'"
		end if

		if (FSearchSellYn<>"") then
			sqlrect = sqlrect + " and i.sellyn = '" + CStr(FSearchSellYn) + "'"
		end if

		if (FSearchUsingyn<>"") then
			sqlrect = sqlrect + " and i.isusing = '" + CStr(FSearchusingyn) + "'"
		end if

		if (FSearchSailyn <> "") then
			sqlrect = sqlrect + " and i.sailyn = '" + CStr(FSearchsailyn) + "'"
		end if

		if (FSearchLimitYn<>"") then
			sqlrect = sqlrect + " and i.limityn = '" + CStr(FSearchLimitYn) + "'"
		end if

		if (FSearchBaedalDiv<>"") then
			sqlrect = sqlrect + " and i.deliverytype = '" + CStr(FSearchBaedalDiv) + "'"
		end if

		rsget.Open sqlStr + sqlrect,dbget,1

		'response.write sqlStr + sqlrect
		FTotalCount = rsget("cnt")
		rsget.close

		sqlrect = sqlrect + " order by i.itemid desc"

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemdiv, i.itemname, i.makerid, i.buycash,"
		sqlStr = sqlStr + " i.buyvat, i.sellcash, i.sellvat, i.margin, "
		sqlStr = sqlStr + " i.sellyn, i.dispyn, i.deliverytype , i.vatinclude, "
		sqlStr = sqlStr + " i.pojangok, i.limityn, i.limitdiv, i.limitno, i.limitsold, i.smallimage as imgsmall,"
		sqlStr = sqlStr + " i.mwdiv ,i.sailyn, i.isusing,"
		sqlStr = sqlStr + " i.sailprice, i.sailsuplycash, i.optioncnt, isnull(o.itemoption,'0000') as itemoption, isnull(o.optionname,'') as itemoptionname, o.optlimitno, o.optlimitsold "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on i.itemid = o.itemid and isnull(o.itemoption,'0000') <> '0000' and o.isusing <> 'N' "
		sqlStr = sqlStr + " where i.itemid<>0 "

		rsget.pagesize = FPageSize
		'response.write sqlStr + sqlrect
		rsget.Open sqlStr + sqlrect,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLimitItemItem
				FItemList(i).FItemID    = rsget("itemid")
				FItemList(i).FItemName  = db2html(rsget("itemname"))
				FItemList(i).FMakerID   = rsget("makerid")
				FItemList(i).FDisplayYn = rsget("dispyn")
				FItemList(i).FSellYn    = rsget("sellyn")
				FItemList(i).FBaesongGB = rsget("deliverytype")

				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("imgsmall")
				FItemList(i).FMWDiv		= rsget("mwdiv")
				FItemList(i).Fsailyn	= rsget("sailyn")
				FItemList(i).Fisusing	= rsget("isusing")

				FItemList(i).FitemOption= rsget("itemoption")
				FItemList(i).FitemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FitemOptionCount=rsget("optioncnt")

				FItemList(i).FLimitYn = rsget("limityn")
				if (FItemList(i).FitemOption = "0000") then
    				FItemList(i).FLimitNo = rsget("limitno")
    				FItemList(i).FLimitSold = rsget("limitsold")
    			else
    				FItemList(i).FLimitNo = rsget("optlimitno")
    				FItemList(i).FLimitSold = rsget("optlimitsold")
    			end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.close
	end Sub

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
%>