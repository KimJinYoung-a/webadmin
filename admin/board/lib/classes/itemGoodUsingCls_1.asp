<%
Class CGoodUsingItem
	public FUsingId
	public Fuserid
	public Fgubun
	public Fcontents
	public Fmakerid
	public Fitemid
	public Fitemname
	public FisUsing
	public FtotalPoint
	public FPointFunction
	public FPointDesign
	public FPointPrice
	public FPointSatisfy
	public Fregdate
	public FCDL
	public FCDM
	public FCDS
	public FFile1
	public FFile2
	public FImageIcon1
	public FImageIcon2


	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class


Class CGoodUsing
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectUsingId
	public FRectSearchKey1
	public FRectSearchKey2
	public FRectselStatus
	public FRectStartDt
	public FRectEndDt
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectPoint
	public FRectSort
	public FRectPhotoMode
    public FRectMakerid

	public FAvgTotalPoint
	public FAvgFunctionPoint
	public FAvgDesignPoint
	public FAvgPricePoint
	public FAvgSatisfyPoint

	Private Sub Class_Initialize()
		redim  FitemList(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public function GetImageFolerName(byval i)
		'GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
		GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
	end function

	public Sub GetGoodUsingList()
		dim SQL, AddSQL, i, strTemp

		if FRectSearchKey1<>"" then
				AddSQL = AddSQL & " and t1.userid='" & FRectSearchKey1 & "' "
		end if
		if FRectSearchKey2<>"" then
				AddSQL = AddSQL & " and t1.itemid=" & FRectSearchKey2 & " "
		end if

		if FRectselStatus<>"" then
			AddSQL = AddSQL & " and t1.isUsing='" & FRectselStatus & "' "
		end if

		if Not(FRectStartDt="" or FRectEndDt="") then
			AddSQL = AddSQL & " and t1.regdate between '" & FRectStartDt & "' and DateAdd(day,1,'" & FRectEndDt & "') "
		end if

		if FRectCDL<>"" then
			AddSQL = AddSQL & " and t2.cate_large='" & FRectCDL & "' "
		end if
		if FRectCDM<>"" then
			AddSQL = AddSQL & " and t2.cate_mid='" & FRectCDM & "' "
		end if
		if FRectCDS<>"" then
			AddSQL = AddSQL & " and t2.cate_small='" & FRectCDS & "' "
		end if

		if FRectPoint<>"" then
			AddSQL = AddSQL & " and t1.totalPoint='" & FRectPoint & "' "
		end if

		IF FRectPhotoMode="on" then
			AddSQL = AddSQL & " and t1.File1 is not null "
		End IF

		IF (FRectMakerid<>"") then
		    AddSQL = AddSQL & " and t2.makerid='"&FRectMakerid&"'"
		end if

		'// 개수 파악 //
		SQL =	"Select count(t1.itemid), CEILING(CAST(Count(t1.itemid) AS FLOAT)/" & FPageSize & ") " &_
				"	,avg(cast(t1.TotalPoint as float)) as tp " &_
				"	,avg(cast(t1.Point_Function as float)) as pf " &_
				"	,avg(cast(t1.Point_Design as float)) as pd " &_
				"	,avg(cast(t1.Point_Price as float)) as pp " &_
				"	,avg(cast(t1.Point_Satisfy as float)) as ps " &_
				"from db_board.dbo.tbl_Item_Evaluate as t1 " &_
				"	join db_item.[dbo].tbl_item as t2 " &_
				"		on t1.itemid=t2.itemid " &_
				"where 1=1 " & AddSQL
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)

			FAvgTotalPoint = rsget(2)
			FAvgFunctionPoint = rsget(3)
			FAvgDesignPoint = rsget(4)
			FAvgPricePoint = rsget(5)
			FAvgSatisfyPoint = rsget(6)
		rsget.Close

		'// 목록 접수 //
		SQL =	"Select top " & CStr(FPageSize*FCurrPage) &_
				"	t1.idx, t1.userid " &_
				"	, Case t1.gubun when '0' then '일반' " &_
				"		when '14' then '매니아' " &_
				"		else '기타' " &_
				"	  end as gubun " &_
				"	, t1.contents, t1.isUsing, t2.makerid, t2.itemname, t1.itemid " &_
				"	, t1.totalPoint, t1.Point_Function, t1.Point_Design, t1.Point_Price, t1.Point_Satisfy, t1.regdate ,t1.File1 , t1.File2 " &_
				"	, cL.code_nm as cdlName, cM.code_nm as cdmName, cS.code_nm as cdsName  " &_
				"From db_board.dbo.tbl_Item_Evaluate as t1 " &_
				"	join db_item.[dbo].tbl_item as t2 " &_
				"		on t1.itemid=t2.itemid " &_
				"	left Join db_item.dbo.tbl_Cate_large as cL " &_
				"		on t2.cate_large=cL.code_large " &_
				"	left Join db_item.dbo.tbl_Cate_mid as cM " &_
				"		on t2.cate_large=cM.code_large " &_
				"			and t2.cate_mid=cM.code_mid " &_
				"	left Join db_item.dbo.tbl_Cate_small as cS " &_
				"		on t2.cate_large=cS.code_large " &_
				"			and t2.cate_mid=cS.code_mid " &_
				"			and t2.cate_small=cS.code_small " &_
				"where 1=1 " & AddSQL
				'response.write SQL

		'#정렬방법 선택(2008.07.21;허진원 추가)
		Select Case FRectSort
			Case "idxAcd"
				SQL = SQL & "order by idx asc"
			Case "idxDcd"
				SQL = SQL & "order by idx desc"
			Case "pntAcd"
				SQL = SQL & "order by t1.totalPoint asc"
			Case "pntDcd"
				SQL = SQL & "order by t1.totalPoint desc"
			Case "iidAcd"
				SQL = SQL & "order by t1.itemid asc"
			Case "iidDcd"
				SQL = SQL & "order by t1.itemid desc"
			Case Else
				SQL = SQL & "order by idx desc"
		end Select

		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CGoodUsingItem

				FItemList(i).FUsingId		= rsget("idx")
				FItemList(i).Fuserid		= rsget("userid")
				FItemList(i).Fgubun			= rsget("gubun")
				FItemList(i).Fcontents		= rsget("contents")
				FItemList(i).Fmakerid		= rsget("makerid")
				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fitemname		= rsget("itemname")
				FItemList(i).FisUsing		= rsget("isUsing")
				FItemList(i).FtotalPoint	= rsget("totalPoint")
				FItemList(i).FPointFunction	= rsget("Point_Function")
				FItemList(i).FPointDesign	= rsget("Point_Design")
				FItemList(i).FPointPrice	= rsget("Point_Price")
				FItemList(i).FPointSatisfy	= rsget("Point_Satisfy")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FCDL			= rsget("cdlName")
				FItemList(i).FCDM			= rsget("cdmName")
				FItemList(i).FCDS			= rsget("cdsName")
				FItemList(i).FFile1			= rsget("file1")
				FItemList(i).FFile2			= rsget("file2")

				IF Not(rsget("File1")="" or isNull(rsget("File1"))) Then
					FItemList(i).FImageIcon1		= "http://imgstatic.10x10.co.kr/goodsimage/" + GetImageFolerName(i) + "/" + rsget("File1")
				End IF
				IF Not(rsget("File2")="" or isNull(rsget("File2"))) Then
					FItemList(i).FImageIcon2		= "http://imgstatic.10x10.co.kr/goodsimage/" + GetImageFolerName(i) + "/" + rsget("File2")
				End IF
				'FItemList(i).FImageList120    = "http://webimage.10x10.co.kr/image/list120/" + GetImageFolerName(i) + "/" + rsget("listimage120")
				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

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