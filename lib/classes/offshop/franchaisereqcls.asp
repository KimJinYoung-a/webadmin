<%
class CFranChaiseReqItem
	public FIdx
	public Fuserid
	public Fusername
	public Fjuminno
	public Faddress
	public Fuserphone
	public Fuseremail
	public Fconsulttype


	public Fage
	public Fsex
	public Fhphone
	public Fkyungro
	public Fetc


	public Finvest_money
	public Fgain_percent
	public Fgain_money
	public Finvest_year
	public Finvest_etc

	public Ffran_open
	public Fshop_exists
	public Fshop_upjong
	public Fshop_currarea
	public Fshop_pyng
	public Fshop_opertype
	public Fshop_mayarea
	public Fshop_maypyng
	public Fshop_mayfund
	public Fshop_maymonthgain
	public Fshop_etc
	public Fregdate
	public Ffinishflag
	public Fadmincomment
	public Fetcfile

	public function Getinvest_yearName()
		if Finvest_year="1" then
			Getinvest_yearName ="3년이내"
		elseif Finvest_year="3" then
			Getinvest_yearName ="3년~5년"
		elseif Finvest_year="5" then
			Getinvest_yearName ="5년이상"
		end if
	end function

	public function Getshop_opertypeName()
		if Fshop_opertype="1" then
			Getshop_opertypeName ="직접"
		elseif Fshop_opertype="3" then
			Getshop_opertypeName ="친인척"
		elseif Fshop_opertype="5" then
			Getshop_opertypeName ="판매관리자 채용"
		end if
	end function
	
	public function GetKyungro()
		if Fkyungro="1" then
			GetKyungro ="텐바이텐매장"
		elseif Fkyungro="2" then
			GetKyungro ="온라인 쇼핑몰"
		elseif Fkyungro="3" then
			GetKyungro ="신문, 인터넷"
		elseif Fkyungro="4" then
			GetKyungro ="기타 - " &  Fetc
		end if
	end function

	public function GetconsulttypeName()
		if Fconsulttype="1" then
			GetconsulttypeName = "투자상담"
		elseif Fconsulttype="2" then
			GetconsulttypeName = "가맹점상담"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CFranChaiseReqList
	public FItemList()

	public FResultCount
	public FTotalCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectOnlymifinish
	public FRectIdx
	public FrectGubun
	public FRectNotDel
	public upfolder

	public Sub GetReqList()
		dim sqlStr,i
		sqlStr = " select count(idx) as cnt from"
		sqlStr = sqlStr + " [db_cs].[dbo].tbl_franchaise"

		sqlStr = sqlStr + " where idx<>0"

		if FRectIdx<>"" then
			sqlStr = sqlStr + " and idx=" + CStr(FRectIdx)
		end if

		if FrectGubun<>"" then
			sqlStr = sqlStr + " and consulttype='" + CStr(FrectGubun) + "'"
		end if

		if FRectNotDel<>"" then
			sqlStr = sqlStr + " and deleteyn='N'"
		end if

		if FRectOnlymifinish="on" then
			sqlStr = sqlStr + " and finishflag<>'7'"
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " *, convert(varchar(20),regdate,20) as regdate2 from"
		sqlStr = sqlStr + " [db_cs].[dbo].tbl_franchaise"

		sqlStr = sqlStr + " where idx<>0"

		if FRectIdx<>"" then
			sqlStr = sqlStr + " and idx=" + CStr(FRectIdx)
		end if

		if FrectGubun<>"" then
			sqlStr = sqlStr + " and consulttype='" + CStr(FrectGubun) + "'"
		end if

		if FRectNotDel<>"" then
			sqlStr = sqlStr + " and deleteyn='N'"
		end if

		if FRectOnlymifinish="on" then
			sqlStr = sqlStr + " and finishflag<>'7'"
		end if

		sqlStr = sqlStr + " order by idx desc"

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
				set FItemList(i) = new CFranChaiseReqItem

				FItemList(i).FIdx = rsget("idx")
				FItemList(i).Fuserid = rsget("userid")
				FItemList(i).Fusername = db2html(rsget("username"))
				FItemList(i).Fjuminno = rsget("juminno")
				FItemList(i).Faddress = rsget("address")
				FItemList(i).Fuserphone = rsget("userphone")
				FItemList(i).Fuseremail = rsget("useremail")
				FItemList(i).Fconsulttype = rsget("consulttype")

				FItemList(i).Finvest_money = rsget("invest_money")
				FItemList(i).Fgain_percent = rsget("gain_percent")
				FItemList(i).Fgain_money	 = rsget("gain_money")
				FItemList(i).Finvest_year = rsget("invest_year")
				FItemList(i).Finvest_etc = rsget("invest_etc")

				FItemList(i).Ffran_open = rsget("fran_open")
				FItemList(i).Fshop_exists = rsget("shop_exists")
				FItemList(i).Fshop_upjong = rsget("shop_upjong")
				FItemList(i).Fshop_currarea = rsget("shop_currarea")
				FItemList(i).Fshop_pyng = rsget("shop_pyng")
				FItemList(i).Fshop_opertype = rsget("shop_opertype")
				FItemList(i).Fshop_mayarea = rsget("shop_mayarea")
				FItemList(i).Fshop_maypyng = rsget("shop_maypyng")
				FItemList(i).Fshop_mayfund = rsget("shop_mayfund")
				FItemList(i).Fshop_maymonthgain = rsget("shop_maymonthgain")
				FItemList(i).Fshop_etc = db2html(rsget("shop_etc"))
				FItemList(i).Fregdate = rsget("regdate2")
				FItemList(i).Ffinishflag = rsget("finishflag")
				FItemList(i).Fadmincomment = db2html(rsget("admincomment"))
				
				
				FItemList(i).Fhphone = rsget("hphone")
				FItemList(i).Fkyungro = rsget("kyungro")
				FItemList(i).Fetc	= rsget("etc")
				FItemList(i).Fetcfile = db2html(rsget("etcfile"))
				
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		upfolder = "/uploadimg/"
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
end class
%>