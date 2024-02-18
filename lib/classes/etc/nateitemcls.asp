<%
class CNateItem
	public FItemId
	public FTitle
	public FlistImage
	public Fbasicimage
	public FPFrom
	public FSTerm
	public FDesc

	public FNmLarge
	public FNmMid
	public FNmSmall

	public FSellCount
	public FFavCount
	public FSellcash
	public Fmakername
	public FItemName

	public FitemMaker
	public FsourceArea
	public Fmileage
    
    public Fdeliverytype
    
    ''쿠폰 100 이하의 경우 : 할인율로 표시됨[ex : 25 -> 25%],  100 이상의 경우 : 할인금액으로 표시됨[ex : 2500 -> 2,500원]
    public function GetMMCouponStr()
        GetMMCouponStr =""
    end function

    public function GetDeliverPay()
        if (Fdeliverytype="2") or (Fdeliverytype="5") then
            GetDeliverPay = 0
        else
            if FSellcash>=30000 then
                GetDeliverPay = 0
            else
                GetDeliverPay = 2000
            end if
        end if
    end function

	public function Getmakername()
		Getmakername = Fmakername
	end function

	public function GetModelname()
		GetModelname = FItemName
	end function

	public function GetItemUrl()
		GetItemUrl = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & CStr(FItemId) & "&rdsite=yahoo"
	end function
	
	public function GetNateItemUrl()
		GetNateItemUrl = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & CStr(FItemId) & "&rdsite=nate"
	end function

	public function GetListImageUrl()
		if IsNULL(FListImage) then
			GetListImageUrl = ""
		else
			GetListImageUrl = FListImage
		end if
	end function

	public function GetBasicImageUrl()
		if IsNULL(FbasicImage) then
			GetBasicImageUrl = ""
		else
			GetBasicImageUrl = FbasicImage
		end if
	end function

	public function GetPrice()
		GetPrice = FSellcash
	end function

	public function GetTitle()
		dim re
		re = replace(replace(FTitle,"<","["),">","]")
		GetTitle = "<title:" & re & ">"
	end function

	public function getNateBBPath()
		getNateBBPath = FNmLarge & "@" & FNmMid & "@" & FNmSmall
	end function

	public function getWeightCalcu()
		dim p
		p = CLng((FSellCount + FFavCount)/2)


		if (p=>100) then
			getWeightCalcu = 100
		else
			getWeightCalcu = p
		end if
	end function

	public function DelInSide(byval v)
		dim re
		dim pos1,pos2
		re = v
		pos1 = InStr(re,"<")
		pos2 = InStr(re,">")

		if (pos1<1) or (pos2<1) then
			DelInSide = re
		else
			on error resume next
			re = Left(v,pos1-1) + Mid(v,pos2+1,512)
			DelInSide = DelInSide(re)
			if err then DelInSide = re
			on error goto 0

		end if
	end function

	public function getDesc()
		dim re
		re = FDesc
		re = replace(re,vbcrlf,"",1,-1,1)
		re = replace(re,vbcr,"",1,-1,1)
		re = replace(re,"   "," ",1,-1,1)
		getDesc = "<desc:" & LeftB(stripHTML(re),180) & ">"
		exit function

		re = replace(re,"<br>"," ",1,-1,1)
		re = replace(re,"<p>"," ",1,-1,1)
		re = replace(re,"</p>"," ",1,-1,1)
		re = replace(re,"<b>"," ",1,-1,1)
		re = replace(re,"</b>"," ",1,-1,1)
		re = replace(re,"<font>"," ",1,-1,1)
		re = replace(re,"</font>"," ",1,-1,1)
		re = replace(re,"<center>"," ",1,-1,1)
		re = replace(re,"</center>"," ",1,-1,1)
		re = replace(re,vbcrlf,"",1,-1,1)
		re = replace(re,vbcr,"",1,-1,1)

		re = DelInSide(re)
		re = replace(re,"<"," ",1,-1,1)
		re = replace(re,">"," ",1,-1,1)
		re = replace(re,"   "," ",1,-1,1)
		re = replace(re," "," ",1,-1,1)
		getDesc = "<desc:" & LeftB(re,180) & ">"
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CNateItemList
	public FItemList()

	public FResultCount
	public FTotalCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

    public Sub GetNateItemCountDB3
		dim sqlStr,i
		sqlStr = " select count(i.itemid) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + "From [db_datamart].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.sellyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.isextusing='Y'"
        sqlStr = sqlStr + " and (i.recentsellcount>1 or i.sellcount>15)"
        sqlStr = sqlStr + " and i.deliverytype<6"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<'999'"
		sqlStr = sqlStr + " and i.itemid<>203302"

		'sqlStr = sqlStr + " and datediff(d,i.lastupdate,getdate())<4"

		db3_rsget.Open sqlStr,db3_dbget,1
			FTotalCount = db3_rsget("cnt")
			FTotalPage = db3_rsget("totPg")
		db3_rsget.Close
	end Sub

    public Sub GetNateItemDB3
		dim sqlStr,i
		sqlStr = " select count(i.itemid) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + "From [db_datamart].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.sellyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.isextusing='Y'"
        sqlStr = sqlStr + " and (i.recentsellcount>1 or i.sellcount>15)"
        sqlStr = sqlStr + " and i.deliverytype<6"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<'999'"
		sqlStr = sqlStr + " and i.itemid<>203302"
		'sqlStr = sqlStr + " and datediff(d,i.lastupdate,getdate())<4"

		db3_rsget.Open sqlStr,db3_dbget,1
			FTotalCount = db3_rsget("cnt")
			FTotalPage = db3_rsget("totPg")
		db3_rsget.Close

'''/// db_datamart.[dbo].[sp_Ten_DB_ITEM_copy_hour] 에 tbl_item_contents 추가 해야됨

		sqlStr = " select  distinct top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.listimage, i.mainimage, i.basicimage, i.sellcash, i.makername, i.deliverytype, i.mileage,"
		sqlStr = sqlStr + " v.nmlarge, v.nmmid, v.nmsmall,"
		sqlStr = sqlStr + " i.itemMaker, i.sourceArea "
		sqlStr = sqlStr + " from [db_datamart].[dbo].tbl_item i "
		sqlStr = sqlStr + "     left join [db_datamart].[dbo].tbl_item_cate_all v "
		sqlStr = sqlStr + "     on  v.cdlarge=i.cate_large"
		sqlStr = sqlStr + "     and v.cdmid=i.cate_mid"
		sqlStr = sqlStr + "     and v.cdsmall=i.cate_small"
		sqlStr = sqlStr + " where i.sellyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.isextusing='Y'"
        sqlStr = sqlStr + " and (i.recentsellcount>1 or i.sellcount>15)"
        sqlStr = sqlStr + " and i.deliverytype<6"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<'999'"
		sqlStr = sqlStr + " and i.itemid<>203302"
		'sqlStr = sqlStr + " and datediff(d,i.lastupdate,getdate())<4"
		sqlStr = sqlStr + " order by i.itemid desc"

		db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,1

		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CNateItem
				FItemList(i).FItemId = db3_rsget("itemid")
				FItemList(i).FTitle = db2html(db3_rsget("itemname"))
				FItemList(i).Fbasicimage = db3_rsget("basicimage")
				FItemList(i).FListImage = db3_rsget("listimage")

				if IsNULL(FItemList(i).Fbasicimage) then FItemList(i).Fbasicimage=""
				if IsNULL(FItemList(i).FlistImage) then FItemList(i).FlistImage=""

				if FItemList(i).Fbasicimage<>"" then
					FItemList(i).Fbasicimage = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).Fbasicimage
				end if

				if FItemList(i).FListImage<>"" then
					FItemList(i).FListImage = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).Flistimage
				end if

				''FItemList(i).FDesc  = db2html(db3_rsget("des"))
				FItemList(i).FNmLarge  = db2html(db3_rsget("nmlarge"))
				FItemList(i).FNmMid    = db2html(db3_rsget("nmmid"))
				FItemList(i).FNmSmall  = db2html(db3_rsget("nmsmall"))

				if IsNULL(FItemList(i).FNmLarge) then FItemList(i).FNmLarge=""
				if IsNULL(FItemList(i).FNmMid) then FItemList(i).FNmMid=""
				if IsNULL(FItemList(i).FNmSmall) then FItemList(i).FNmSmall=""


				FItemList(i).FSellcash = db3_rsget("sellcash")
				FItemList(i).Fmakername = db2html(db3_rsget("makername"))
				FItemList(i).FItemName = db2html(db3_rsget("itemname"))

				FItemList(i).FitemMaker = db2html(db3_rsget("itemMaker"))
				FItemList(i).FsourceArea = db2html(db3_rsget("sourcearea"))
				FItemList(i).Fmileage = db2html(db3_rsget("mileage"))

                FItemList(i).Fdeliverytype = db3_rsget("deliverytype")
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
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
end class
%>