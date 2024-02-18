<%

class CLgeshopItem
	public FItemId
	public FItemOption
	public FItemName
	public FItemOptionName
	public FSellcash
	public FMakerID

	public function GetItemIdNOption()
		GetItemIdNOption = Cstr(FItemId) + "_" + CStr(FItemOption)
	end function

	public function GetItemIdNOptionName()
		if FItemOptionName="" then
			GetItemIdNOptionName = FItemName
		else
			GetItemIdNOptionName = FItemName + "(" + FItemOptionName + ")"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CYahooItem
	public FItemId
	public FTitle
	public FImage
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

	public function GetImageUrl()
		if IsNULL(FImage) then
			GetImageUrl = ""
		else
			GetImageUrl = FImage
		end if
	end function

	public function GetPrice()
		GetPrice = FSellcash
	end function

	public function GetLargeNameKor()
		if FNmLarge="LIVING" then
			GetLargeNameKor = "거실,가구"
		elseif FNmLarge="FASHION" then
			GetLargeNameKor = "패션,잡화"
		elseif FNmLarge="JEWERLY" then
			GetLargeNameKor = "보석,악세사리"
		elseif FNmLarge="OFFICE/DESK" then
			GetLargeNameKor = "문구,사무용품"
		elseif FNmLarge="KITCHEN/BATH" then
			GetLargeNameKor = "주방,욕실용품"
		elseif FNmLarge="PERSONAL" then
			GetLargeNameKor = "휴대,개인용품"
		elseif FNmLarge="MANIA/HOBBY" then
			GetLargeNameKor = "장난감,취미"
		elseif FNmLarge="ANNIVERSARY" then
			GetLargeNameKor = "기념일"
		elseif Left(FNmLarge,4)="[애견]" then
			GetLargeNameKor = "애견"
		elseif FNmLarge="BOARDGAME" then
			GetLargeNameKor = "보드게임"
		else
			GetLargeNameKor = FNmLarge
		end if
	end function

	public function GetTitle()
		dim re
		re = replace(replace(FTitle,"<","["),">","]")
		GetTitle = "<title:" & re & ">"
	end function

	public function getSUrl()
		getSUrl = "<surl:" & "http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & CStr(FItemId) & "&rdsite=yahoo>"
	end function

	public function getImage()
		if IsNULL(FImage) then
			getImage = "<image:" & "" & ">"
		else
			getImage = "<image:" & FImage & ">"
		end if
	end function

	public function getPFrom()
		getPFrom = "<pfrom:" & CStr(FPFrom) & ">"
	end function

	public function getYPath()
		getYPath = "<ypath:" & GetLargeNameKor & ":" & CStr(FNmMid) & ":" & CStr(FNmSmall) & ">"
	end function

	public function getNatePath()
		getNatePath = FNmLarge & ">" & FNmMid & ">" & FNmSmall
	end function

	public function getSterm()
		getSterm = "<sterm:" & CStr(FNmMid) & " " & CStr(FNmSmall) & ">"
	end function

	public function getWeight()
		getWeight = "<weight:" & CStr(getWeightCalcu) & ">"
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

Class CEmpasItem
	public FItemId
	public FItemName
	public FImage
	public FSellcash
	public FOrgSellcash
	public FCateName
	public FDesc
	public Fsourcearea
	public Fmakername
 	public FRegDate

	public FNmLarge
	public FNmMid
	public FNmSmall

	public FCdL
	public FCdM
	public FCdS

	public function GetEmpasAllCategory()
		if (Fcdl=10) and (Fcdm=40) then
			GetEmpasAllCategory = "AC05"
		elseif (Fcdl=10) and (Fcdm=45) then
			GetEmpasAllCategory = "AB07"
		elseif (Fcdl=10) and (Fcdm=50) then
			GetEmpasAllCategory = "AC05"
		elseif (Fcdl=10) and (Fcdm=60) then
			GetEmpasAllCategory = "AA01"
		elseif (Fcdl=10) and (Fcdm=65) then
			GetEmpasAllCategory = "AA08"
		elseif (Fcdl=10) and (Fcdm=70) then
			GetEmpasAllCategory = "AA01"
		elseif (Fcdl=20) and (Fcdm=45) then
			GetEmpasAllCategory = "AC04"
		elseif (Fcdl=25) and (Fcdm=90) then
			GetEmpasAllCategory = "AB09"
		elseif (Fcdl=30) and (Fcdm=11) then
			GetEmpasAllCategory = "AB05"
		elseif (Fcdl=30) and (Fcdm=50) then
			GetEmpasAllCategory = "AB02"
		elseif (Fcdl=30) and (Fcdm=55) then
			GetEmpasAllCategory = "AB06"
		elseif (Fcdl=30) and (Fcdm=60) then
			GetEmpasAllCategory = "AB07"
		elseif (Fcdl=30) and (Fcdm=65) then
			GetEmpasAllCategory = "AA00"
		elseif (Fcdl=30) and (Fcdm=70) then
			GetEmpasAllCategory = "AB02"
		elseif (Fcdl=30) and (Fcdm=90) then
			GetEmpasAllCategory = "AB08"
		elseif (Fcdl=30) and (Fcdm=91) then
			GetEmpasAllCategory = "AB04"
		elseif (Fcdl=40) and (Fcdm=40) then
			GetEmpasAllCategory = "AA04"
		elseif (Fcdl=40) and (Fcdm=60) then
			GetEmpasAllCategory = "AA01"
		elseif (Fcdl=40) and (Fcdm=65) then
			GetEmpasAllCategory = "AA06"
		elseif (Fcdl=40) and (Fcdm=70) then
			GetEmpasAllCategory = "AA01"
		elseif (Fcdl=40) and (Fcdm=75) then
			GetEmpasAllCategory = "AA06"
		elseif (Fcdl=40) and (Fcdm=90) then
			GetEmpasAllCategory = "AA06"
		elseif (Fcdl=45) and (Fcdm=10) then
			GetEmpasAllCategory = "AE02"
		elseif (Fcdl=10) then
			GetEmpasAllCategory = "AA00"
		elseif (Fcdl=20) then
			GetEmpasAllCategory = "AA08"
		elseif (Fcdl=25) then
			GetEmpasAllCategory = "AB08"
		elseif (Fcdl=30) then
			GetEmpasAllCategory = "AB09"
		elseif (Fcdl=40) then
			GetEmpasAllCategory = "AA02"
		elseif (Fcdl=42) then
			GetEmpasAllCategory = "AB09"
		elseif (Fcdl=45) then
			GetEmpasAllCategory = "AE01"
		elseif (Fcdl=48) then
			GetEmpasAllCategory = "AE01"
		elseif (Fcdl=50) then
			GetEmpasAllCategory = "AB00"
		elseif (Fcdl=52) then
			GetEmpasAllCategory = "AE00"
		elseif (Fcdl=53) then
			GetEmpasAllCategory = "AE00"
		elseif (Fcdl=60) then
			GetEmpasAllCategory = "AE03"
		elseif (Fcdl=70) then
			GetEmpasAllCategory = "AE01"
		elseif (Fcdl=80) then
			GetEmpasAllCategory = "AB09"
		else
			GetEmpasAllCategory = "AA00"
		end if
	end function

	public function GetEmpasLargeCode()
		GetEmpasLargeCode = Left(GetEmpasAllCategory,2)
	end function

	public function GetEmpasMidCode()
		GetEmpasMidCode = Right(GetEmpasAllCategory,2)
	end function

	public function GetEmpasSmallCode()
		GetEmpasSmallCode =""
	end function

	public function GetEmpasSeCode()
		GetEmpasSeCode = ""
	end function

	public function GetTenbytenCategoryName()
		GetTenbytenCategoryName = GetLargeNameKor() + ">" + FNMmid + ">" + FNMSmall
	end function

	public function getImage()
		if IsNULL(FImage) then
			getImage = ""
		else
			getImage = FImage
		end if
	end function

	public function GetEmpasItemName()
		dim re
		re = FItemName
		re = replace(re,vbcrlf,"",1,-1,1)
		re = replace(re,vbcr,"",1,-1,1)
		re = replace(re,"   "," ",1,-1,1)
		GetEmpasItemName = LeftB(stripHTML(re),64)
	end function

	public function GetModelName()

	end function

	public function GetSourceArea()
		GetSourceArea = Fsourcearea
	end function

	public function GetEmpasUrl()
		GetEmpasUrl = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & CStr(FItemId) & "&rdsite=empas"
	end function

	public function GetEmpasDesc()
		dim re
		re = Fdesc
		re = replace(re,vbcrlf,"")
		re = replace(re,vbcr,"")
		're = replace(re,"    "," ")
		're = replace(re,"   "," ")
		re = replace(re,"  "," ")
		re = stripHTML(re)
		re = replace(re,"<","")
		re = replace(re,">","")
		GetEmpasDesc = LeftB(re,128)
	end function

	public function GetJejosa()
		GetJejosa = Fmakername
	end function

	public function GetBrandName()

	end function

	public function GetOrgSellcash()
		if FOrgSellcash>FSellcash then
			GetOrgSellcash = FOrgSellcash
		else
			GetOrgSellcash = FSellcash
		end if
	end function

	public function GetRealSellcash()
		GetRealSellcash = FSellcash
	end function

	public function GetLastEditDate()
		GetLastEditDate = Left(Replace(FRegDate,"-",""),8)
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

	public function GetLargeNameKor()
		if FNmLarge="LIVING" then
			GetLargeNameKor = "거실,가구"
		elseif FNmLarge="FASHION" then
			GetLargeNameKor = "패션,잡화"
		elseif FNmLarge="JEWERLY" then
			GetLargeNameKor = "보석,악세사리"
		elseif FNmLarge="OFFICE/DESK" then
			GetLargeNameKor = "문구,사무용품"
		elseif FNmLarge="KITCHEN/BATH" then
			GetLargeNameKor = "주방,욕실용품"
		elseif FNmLarge="PERSONAL" then
			GetLargeNameKor = "휴대,개인용품"
		elseif FNmLarge="MANIA/HOBBY" then
			GetLargeNameKor = "장난감,취미"
		elseif FNmLarge="ANNIVERSARY" then
			GetLargeNameKor = "기념일"
		elseif Left(FNmLarge,4)="[애견]" then
			GetLargeNameKor = "애견"
		elseif FNmLarge="BOARDGAME" then
			GetLargeNameKor = "보드게임"
		else
			GetLargeNameKor = FNmLarge
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CYahooItemList
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

	public Sub GetEmpasItem
		dim sqlStr,i
		sqlStr = " select count(i.itemid) as cnt from"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.dispyn='Y'"
		sqlStr = sqlStr + " and i.sellyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.sellcount>=5"
		sqlStr = sqlStr + " and i.cate_large<'90'"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		sqlStr = " select  distinct top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.listimage120 as imglist, i.sellcash,"
		sqlStr = sqlStr + " '' as des,"
		sqlStr = sqlStr + " IsNull(v.nmlarge,'') as nmlarge, IsNull(v.nmmid,'') as nmmid, IsNull(v.nmsmall,'') as nmsmall,"
		sqlStr = sqlStr + " cate_large, itemserial_mid, itemserial_small,"
		sqlStr = sqlStr + " i.orgprice, i.sourcearea, i.makername, i.regdate"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " left join [db_const].[dbo].tbl_const_category c on c.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].vw_category v on v.cdlarge=i.cate_large"
		sqlStr = sqlStr + " and v.cdmid=i.itemserial_mid"
		sqlStr = sqlStr + " and v.cdsmall=i.itemserial_small"

		sqlStr = sqlStr + " where i.dispyn='Y'"
		sqlStr = sqlStr + " and i.sellyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.sellcount>=5"
		sqlStr = sqlStr + " and i.cate_large<'90'"
		sqlStr = sqlStr + " order by i.itemid desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CEmpasItem
				FItemList(i).FItemId = rsget("itemid")
				FItemList(i).FItemName = db2html(rsget("itemname"))
				FItemList(i).FImage = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imglist")
				FItemList(i).FSellcash = rsget("sellcash")
				FItemList(i).FOrgSellcash = rsget("orgprice")
				FItemList(i).FNmLarge = db2html(rsget("nmlarge"))
				FItemList(i).FNmMid = db2html(rsget("nmmid"))
				FItemList(i).FNmSmall = db2html(rsget("nmsmall"))
				FItemList(i).FDesc = db2html(rsget("des"))
				FItemList(i).Fsourcearea = db2html(rsget("sourcearea"))
				FItemList(i).Fmakername = db2html(rsget("makername"))
				FItemList(i).FRegDate = rsget("regdate")

				FItemList(i).FCDL = rsget("cate_large")
				FItemList(i).FCDM = rsget("itemserial_mid")
				FItemList(i).FCDS = rsget("itemserial_small")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public Sub GetLgEshopItem
		dim sqlStr,i
		sqlStr = " select count(i.itemid) as cnt from"
		sqlStr = sqlStr + " tbl_item i"
		sqlStr = sqlStr + " left join  vw_itemoption v on (i.itemid=v.itemid) and (v.isusing='Y')"
		sqlStr = sqlStr + " where i.dispyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.cate_large<'90'"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select  distinct top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.sellcash, i.makerid, IsNull(v.itemoption,'0000') as itemoption, IsNull(v.opt2name,'') as opt2name"
		sqlStr = sqlStr + " from tbl_item i"
		sqlStr = sqlStr + " left join  vw_itemoption v on (i.itemid=v.itemid) and (v.isusing='Y')"
		sqlStr = sqlStr + " where i.dispyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.cate_large<'90'"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLgeshopItem
				FItemList(i).FItemId = rsget("itemid")
				FItemList(i).FItemOption = rsget("itemoption")
				FItemList(i).FItemName = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("opt2name"))
				FItemList(i).FMakerID = rsget("makerid")
				FItemList(i).Fsellcash = rsget("sellcash")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub

	public Sub GetYahooItem
		dim sqlStr,i
		sqlStr = " select count(i.itemid) as cnt from"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.sellyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.isextusing='Y'"
        sqlStr = sqlStr + " and i.itemscore>10"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.cate_large<'999'"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		sqlStr = " select  distinct top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.listimage, i.mainimage, i.basicimage, i.sellcash, i.makername,"
		''sqlStr = sqlStr + " convert(varchar(300),i.itemcontent) as des,"
		sqlStr = sqlStr + " v.nmlarge, v.nmmid, v.nmsmall"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " left join [db_item].[dbo].vw_category v on v.cdlarge=i.cate_large"
		sqlStr = sqlStr + " and v.cdmid=i.itemserial_mid"
		sqlStr = sqlStr + " and v.cdsmall=i.itemserial_small"
		sqlStr = sqlStr + " where i.dispyn='Y'"
		sqlStr = sqlStr + " and i.sellyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.isextusing='Y'"
        sqlStr = sqlStr + " and (i.recentsellcount>1 or i.sellcount>10)"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.cate_large<'90'"
		sqlStr = sqlStr + " order by i.itemid desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CYahooItem
				FItemList(i).FItemId = rsget("itemid")
				FItemList(i).FTitle = db2html(rsget("itemname"))
				FItemList(i).Fbasicimage = rsget("basicimage")
				FItemList(i).FImage = rsget("mainimage")

				if IsNULL(FItemList(i).Fbasicimage) then FItemList(i).Fbasicimage=""
				if IsNULL(FItemList(i).FImage) then FItemList(i).FImage=""

				if FItemList(i).Fbasicimage<>"" then
					FItemList(i).FImage = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).Fbasicimage
				else
					FItemList(i).FImage = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FImage
				end if

				''FItemList(i).FDesc  = db2html(rsget("des"))
				FItemList(i).FNmLarge  = rsget("nmlarge")
				FItemList(i).FNmMid    = rsget("nmmid")
				FItemList(i).FNmSmall  = rsget("nmsmall")

				if IsNULL(FItemList(i).FNmLarge) then FItemList(i).FNmLarge=""
				if IsNULL(FItemList(i).FNmMid) then FItemList(i).FNmMid=""
				if IsNULL(FItemList(i).FNmSmall) then FItemList(i).FNmSmall=""


				FItemList(i).FSellcash = rsget("sellcash")
				FItemList(i).Fmakername = db2html(rsget("makername"))
				FItemList(i).FItemName = db2html(rsget("itemname"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

    public Sub GetYahooItemDB2
		dim sqlStr,i
		sqlStr = " select count(i.itemid) as cnt from"
		sqlStr = sqlStr + " [db_search].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.sellyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.isextusing='Y'"
        sqlStr = sqlStr + " and i.itemscore>1"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.cate_large<'999'"

		db2_rsget.Open sqlStr,db2_dbget,1
			FTotalCount = db2_rsget("cnt")
		db2_rsget.Close


		sqlStr = " select  distinct top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.listimage, i.mainimage, i.basicimage, i.sellcash, i.makername,"
		''sqlStr = sqlStr + " convert(varchar(300),i.itemcontent) as des,"
		sqlStr = sqlStr + " v.nmlarge, v.nmmid, v.nmsmall"
		sqlStr = sqlStr + " from [db_search].[dbo].tbl_item i "
		sqlStr = sqlStr + " left join [db_search].[dbo].tbl_item_category v on v.cdlarge=i.cate_large"
		sqlStr = sqlStr + " and v.cdmid=i.itemserial_mid"
		sqlStr = sqlStr + " and v.cdsmall=i.itemserial_small"
		sqlStr = sqlStr + " where i.dispyn='Y'"
		sqlStr = sqlStr + " and i.sellyn='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.isextusing='Y'"
        sqlStr = sqlStr + " and (i.recentsellcount>1 or i.sellcount>10)"
		sqlStr = sqlStr + " and i.itemdiv<'60'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.cate_large<'90'"
		sqlStr = sqlStr + " order by i.itemid desc"

		db2_rsget.pagesize = FPageSize
		db2_rsget.Open sqlStr,db2_dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db2_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db2_rsget.EOF  then
			db2_rsget.absolutepage = FCurrPage
			do until db2_rsget.eof
				set FItemList(i) = new CYahooItem
				FItemList(i).FItemId = db2_rsget("itemid")
				FItemList(i).FTitle = db2html(db2_rsget("itemname"))
				FItemList(i).Fbasicimage = db2_rsget("basicimage")
				FItemList(i).FImage = db2_rsget("mainimage")

				if IsNULL(FItemList(i).Fbasicimage) then FItemList(i).Fbasicimage=""
				if IsNULL(FItemList(i).FImage) then FItemList(i).FImage=""

				if FItemList(i).Fbasicimage<>"" then
					FItemList(i).FImage = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).Fbasicimage
				else
					FItemList(i).FImage = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FImage
				end if

				''FItemList(i).FDesc  = db2html(db2_rsget("des"))
				FItemList(i).FNmLarge  = db2html(db2_rsget("nmlarge"))
				FItemList(i).FNmMid    = db2html(db2_rsget("nmmid"))
				FItemList(i).FNmSmall  = db2html(db2_rsget("nmsmall"))

				if IsNULL(FItemList(i).FNmLarge) then FItemList(i).FNmLarge=""
				if IsNULL(FItemList(i).FNmMid) then FItemList(i).FNmMid=""
				if IsNULL(FItemList(i).FNmSmall) then FItemList(i).FNmSmall=""


				FItemList(i).FSellcash = db2_rsget("sellcash")
				FItemList(i).Fmakername = db2html(db2_rsget("makername"))
				FItemList(i).FItemName = db2html(db2_rsget("itemname"))

				i=i+1
				db2_rsget.moveNext
			loop
		end if
		db2_rsget.Close
	end Sub
    
    public Sub GetYahooItemDB3
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


		sqlStr = " select  distinct top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.listimage, i.mainimage, i.basicimage, i.sellcash, i.makername, i.deliverytype,"
		sqlStr = sqlStr + " v.nmlarge, v.nmmid, v.nmsmall"
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
				set FItemList(i) = new CYahooItem
				FItemList(i).FItemId = db3_rsget("itemid")
				FItemList(i).FTitle = db2html(db3_rsget("itemname"))
				FItemList(i).Fbasicimage = db3_rsget("basicimage")
				FItemList(i).FImage = db3_rsget("mainimage")

				if IsNULL(FItemList(i).Fbasicimage) then FItemList(i).Fbasicimage=""
				if IsNULL(FItemList(i).FImage) then FItemList(i).FImage=""

				if FItemList(i).Fbasicimage<>"" then
					FItemList(i).FImage = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).Fbasicimage
				else
					FItemList(i).FImage = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FImage
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
                
                FItemList(i).Fdeliverytype = db3_rsget("deliverytype")
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
	end Sub


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

		sqlStr = " select  distinct top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.listimage, i.mainimage, i.basicimage, i.sellcash, i.makername, i.deliverytype,"
		sqlStr = sqlStr + " v.nmlarge, v.nmmid, v.nmsmall"
		sqlStr = sqlStr + " from [db_datamart].[dbo].tbl_item i "
		sqlStr = sqlStr + "     left join [db_datamart].[dbo].tbl_item_cate_ALL v "
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
				set FItemList(i) = new CYahooItem
				FItemList(i).FItemId = db3_rsget("itemid")
				FItemList(i).FTitle = db2html(db3_rsget("itemname"))
				FItemList(i).Fbasicimage = db3_rsget("basicimage")
				FItemList(i).FImage = db3_rsget("mainimage")

				if IsNULL(FItemList(i).Fbasicimage) then FItemList(i).Fbasicimage=""
				if IsNULL(FItemList(i).FImage) then FItemList(i).FImage=""

				if FItemList(i).Fbasicimage<>"" then
					FItemList(i).FImage = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).Fbasicimage
				else
					FItemList(i).FImage = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FImage
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