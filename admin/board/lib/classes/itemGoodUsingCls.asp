<%
Class CGoodUsingItem
	public FUsingId
	public Fuserid
	public Fgubun
	public Fcontents
	public Fmakerid
	public Fitemid
	public Fitemname
	public Fitemoption
	public Fitemoptionname
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
	public FCateName
	public FFile1
	public FFile2
	public FImageIcon1
	public FImageIcon2
	public FSellPrice


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
	public FRectDispcate
	public FRectPoint
	public FRectSort
	public FRectPhotoMode
    public FRectMakerid
	public FRectOrderserial
	public FRectFirst
	public FRectKeyword

	public FAvgTotalPoint
	public FAvgFunctionPoint
	public FAvgDesignPoint
	public FAvgPricePoint
	public FAvgSatisfyPoint
	public farrlist
    
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

	' 밑에 함수를 수정할경우 GetGoodUsingList 도 동일하게 수정해 주셔야 합니다.
	public Sub GetGoodUsingList_excel()
		dim SQL, AddSQL, i, strTemp

		if FRectSearchKey1<>"" then
				AddSQL = AddSQL & " and t1.userid='" & FRectSearchKey1 & "' "
		end if
		if FRectSearchKey2<>"" then
				AddSQL = AddSQL & " and t1.itemid=" & FRectSearchKey2 & " "
		end if

		if FRectselStatus<>"A" then
			AddSQL = AddSQL & " and t1.isUsing='" & FRectselStatus & "' "
		end if

		if Not(FRectStartDt="" or FRectEndDt="") then
			AddSQL = AddSQL & " and t1.regdate between '" & FRectStartDt & "' and DateAdd(day,1,'" & FRectEndDt & "') "
		end if

		if FRectDispcate<>"" then
			AddSQL = AddSQL & " and t3.catecode like '" & FRectDispcate & "%' "
		end if

		if FRectPoint<>"" then
			AddSQL = AddSQL & " and t1.totalPoint='" & FRectPoint & "' "
		end if

		IF FRectPhotoMode="on" then
			AddSQL = AddSQL & " and (isnull(t1.File1,'')<>'' or isnull(t1.File2,'')<>'' or isnull(t1.File3,'')<>'') "
		End IF

		IF (FRectMakerid<>"") then
		    AddSQL = AddSQL & " and t2.makerid='"&FRectMakerid&"'"
		end if

		IF (FRectOrderserial<>"") then
		    AddSQL = AddSQL & " and t1.orderserial='"&FRectOrderserial&"'"
		end if

		IF (FRectFirst<>"") then
		    AddSQL = AddSQL & " and t1.isFirst='Y' "
		end if

		IF(FRectKeyword<>"") Then
			AddSQL = AddSQL & " and t1.contents like '%" & FRectKeyword & "%' "
		End IF

		'// 목록 접수 //
		SQL =	"Select " &_
				"	t1.idx, t1.userid " &_
				"	, Case t1.gubun when '0' then '일반' " &_
				"		when '14' then '매니아' " &_
				"		else '기타' " &_
				"	  end as gubun " &_
				"	, replace(replace(replace(replace(replace(convert(nvarchar(max),t1.contents),char(9),''),char(10),''),char(13),''),'""',''),'''','') as contents" &_
				"	, t1.isUsing, t2.makerid, t2.itemname, t1.itemid " &_
				"	, t1.totalPoint, t1.Point_Function, t1.Point_Design, t1.Point_Price, t1.Point_Satisfy, t1.regdate ,t1.File1 , t1.File2 " &_
				"	, t2.sellcash " &_
				"	, db_item.dbo.getCateCodeFullDepthName(t3.catecode) as cateName " &_
				"	, t1.itemoption, t1.itemOptionName " &_
				" From db_board.dbo.tbl_Item_Evaluate as t1 with (nolock)" &_
				"	join db_item.[dbo].tbl_item as t2 with (nolock)" &_
				"		on t1.itemid=t2.itemid " &_
				"	left join db_item.dbo.tbl_display_cate_item as t3 with (nolock)" &_
				"		on t1.itemid=t3.itemid and t3.isDefault='y' " &_
				"where 1=1 " & AddSQL

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
			Case "sprcAcd"
				SQL = SQL & "order by t2.sellcash asc"
			Case "sprcDcd"
				SQL = SQL & "order by t2.sellcash desc"
			Case Else
				SQL = SQL & "order by idx desc"
		end Select

		'페이징
		SQL = SQL & " OFFSET " & cStr((FCurrPage-1)*FPagesize) & " ROWS FETCH NEXT " & cStr(FPagesize) & " ROWS ONLY "

		'response.write SQL & "<Br>"
		'rsget.pagesize = FPageSize
    	rsget.CursorLocation = adUseClient
    	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

		'FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount

		if FResultCount<1 then FResultCount=0
		i=0
		if  not rsget.EOF  then
			'rsget.absolutepage = FCurrPage
			farrlist		= rsget.getRows()
		end if
		rsget.Close
	end Sub

	' 밑에 함수를 수정할경우 GetGoodUsingList_excel 도 동일하게 수정해 주셔야 합니다.
	public Sub GetGoodUsingList()
		dim SQL, AddSQL, i, strTemp

		if FRectSearchKey1<>"" then
				AddSQL = AddSQL & " and t1.userid='" & FRectSearchKey1 & "' "
		end if
		if FRectSearchKey2<>"" then
				AddSQL = AddSQL & " and t1.itemid=" & FRectSearchKey2 & " "
		end if

		if FRectselStatus<>"A" then
			AddSQL = AddSQL & " and t1.isUsing='" & FRectselStatus & "' "
		end if

		if Not(FRectStartDt="" or FRectEndDt="") then
			AddSQL = AddSQL & " and t1.regdate between '" & FRectStartDt & "' and DateAdd(day,1,'" & FRectEndDt & "') "
		end if

		if FRectDispcate<>"" then
			AddSQL = AddSQL & " and t3.catecode like '" & FRectDispcate & "%' "
		end if

		if FRectPoint<>"" then
			AddSQL = AddSQL & " and t1.totalPoint='" & FRectPoint & "' "
		end if

		IF FRectPhotoMode="on" then
			AddSQL = AddSQL & " and (isnull(t1.File1,'')<>'' or isnull(t1.File2,'')<>'' or isnull(t1.File3,'')<>'') "
		End IF
		
		IF (FRectMakerid<>"") then
		    AddSQL = AddSQL & " and t2.makerid='"&FRectMakerid&"'"
		end if

		IF (FRectOrderserial<>"") then
		    AddSQL = AddSQL & " and t1.orderserial='"&FRectOrderserial&"'"
		end if

		IF (FRectFirst<>"") then
		    AddSQL = AddSQL & " and t1.isFirst='Y' "
		end if

		IF(FRectKeyword<>"") Then
			AddSQL = AddSQL & " and t1.contents like '%" & FRectKeyword & "%' "
		End IF
	
		
		'// 개수 파악 //
		SQL =	"Select count(t1.itemid), CEILING(CAST(Count(t1.itemid) AS FLOAT)/" & FPageSize & ") " &_
				"	,avg(cast(t1.TotalPoint as float)) as tp " &_
				"	,avg(cast(t1.Point_Function as float)) as pf " &_
				"	,avg(cast(t1.Point_Design as float)) as pd " &_
				"	,avg(cast(t1.Point_Price as float)) as pp " &_
				"	,avg(cast(t1.Point_Satisfy as float)) as ps " &_
				" from db_board.dbo.tbl_Item_Evaluate as t1 with (nolock)" &_
				"	join db_item.[dbo].tbl_item as t2 with (nolock)" &_
				"		on t1.itemid=t2.itemid " &_
				"	left join db_item.dbo.tbl_display_cate_item as t3 with (nolock)" &_
				"		on t1.itemid=t3.itemid and t3.isDefault='y' " &_
				"where 1=1 " & AddSQL
    	rsget.CursorLocation = adUseClient
    	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)

			FAvgTotalPoint = rsget(2)
			FAvgFunctionPoint = rsget(3)
			FAvgDesignPoint = rsget(4)
			FAvgPricePoint = rsget(5)
			FAvgSatisfyPoint = rsget(6)
		rsget.Close

		'// 목록 접수 //
		'SQL =	"Select top " & CStr(FPageSize*FCurrPage) &_
		SQL =	"Select " &_
				"	t1.idx, t1.userid " &_
				"	, Case t1.gubun when '0' then '일반' " &_
				"		when '14' then '매니아' " &_
				"		else '기타' " &_
				"	  end as gubun " &_
				"	, t1.contents, t1.isUsing, t2.makerid, t2.itemname, t1.itemid " &_
				"	, t1.totalPoint, t1.Point_Function, t1.Point_Design, t1.Point_Price, t1.Point_Satisfy, t1.regdate ,t1.File1 , t1.File2 " &_
				"	, t2.sellcash " &_
				"	, db_item.dbo.getCateCodeFullDepthName(t3.catecode) as cateName " &_
				"	, t1.itemoption, t1.itemOptionName " &_
				" From db_board.dbo.tbl_Item_Evaluate as t1 with (nolock)" &_
				"	join db_item.[dbo].tbl_item as t2 with (nolock)" &_
				"		on t1.itemid=t2.itemid " &_
				"	left join db_item.dbo.tbl_display_cate_item as t3 with (nolock)" &_
				"		on t1.itemid=t3.itemid and t3.isDefault='y' " &_
				"where 1=1 " & AddSQL

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
			Case "sprcAcd"
				SQL = SQL & "order by t2.sellcash asc"
			Case "sprcDcd"
				SQL = SQL & "order by t2.sellcash desc"
			Case Else
				SQL = SQL & "order by idx desc"
		end Select

		'페이징
		SQL = SQL & " OFFSET " & cStr((FCurrPage-1)*FPagesize) & " ROWS FETCH NEXT " & cStr(FPagesize) & " ROWS ONLY "

		'rsget.pagesize = FPageSize
    	rsget.CursorLocation = adUseClient
    	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

		'FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		FResultCount = rsget.RecordCount

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			'rsget.absolutepage = FCurrPage
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
				FItemList(i).FCateName		= replace(rsget("cateName"),"^^","》")
				FItemList(i).FFile1			= rsget("file1")
				FItemList(i).FFile2			= rsget("file2")
				FItemList(i).FSellPrice		= rsget("sellcash")
				FItemList(i).Fitemoption	= rsget("itemoption")
				FItemList(i).Fitemoptionname= rsget("itemoptionname")

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