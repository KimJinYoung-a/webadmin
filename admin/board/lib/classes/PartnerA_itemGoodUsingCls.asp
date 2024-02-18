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
	public FRectItemid
	public FRectItemName
	public FRectRegID
	public FRectselStatus
	public FRectStartDt
	public FRectEndDt
	public FRectDispCate 
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
		dim SQL, AddSQL, i, strTemp,strSubSort

		if FRectRegID<>"" then
				AddSQL = AddSQL & " and t1.userid='" & FRectRegID & "' "
		end if
		if FRectItemid<>"" then
				AddSQL = AddSQL & " and t1.itemid=" & FRectItemid & " "
		end if
		if FRectItemName<>"" then
				AddSQL = AddSQL & " and t2.itemname='" & FRectItemName & "' "   '' no! like
		end if
		
		if FRectselStatus<>"" then
			AddSQL = AddSQL & " and t1.isUsing='" & FRectselStatus & "' "
		end if

		if Not(FRectStartDt="" or FRectEndDt="") then
			AddSQL = AddSQL & " and t1.regdate between '" & FRectStartDt & "' and DateAdd(day,1,'" & FRectEndDt & "') "
		end if

		if FRectDispCate<>"" then
			AddSQL = AddSQL & " and dci.catecode  like '" & FRectDispCate & "%' "
		end if
 

		if FRectPoint<>"" then
			AddSQL = AddSQL & " and t1.totalPoint='" & FRectPoint & "' "
		end if

		IF FRectPhotoMode="on" then
			AddSQL = AddSQL & " and t1.File1 is not null "
		End IF
		
	 
		
		'// 개수 파악 //
		SQL =	"Select count(t1.itemid) " &vbcrlf
		SQL =		SQL &"	,avg(cast(t1.TotalPoint as float)) as tp "&vbcrlf
		SQL =		SQL &		"	,avg(cast(t1.Point_Function as float)) as pf " &vbcrlf
		SQL =		SQL &		"	,avg(cast(t1.Point_Design as float)) as pd " &vbcrlf
		SQL =		SQL &		"	,avg(cast(t1.Point_Price as float)) as pp " &vbcrlf
		SQL =		SQL &		"	,avg(cast(t1.Point_Satisfy as float)) as ps " &vbcrlf
		SQL =		SQL &		"from db_board.dbo.tbl_Item_Evaluate as t1 " &vbcrlf
		SQL =		SQL &		"	inner join db_item.[dbo].tbl_item as t2 " &vbcrlf
		SQL =		SQL &		"		on t1.itemid=t2.itemid " &vbcrlf
			if FRectDispCate <>"" then
		SQL =		SQL &		"	left outer join  db_item.dbo.tbl_display_cate_item as dci on t1.itemid = dci.itemid and isdefault ='y' "				
			end if
		SQL =		SQL &		"where t2.makerid = '"&FRectMakerid &"'  " & AddSQL

		rsget.CursorLocation = adUseClient
        rsget.Open SQL,dbget,adOpenForwardOnly, adLockReadOnly

		if not rsget.eof then
			FTotalCount = rsget(0) 
			FAvgTotalPoint = rsget(1)
			FAvgFunctionPoint = rsget(2)
			FAvgDesignPoint = rsget(3)
			FAvgPricePoint = rsget(4)
			FAvgSatisfyPoint = rsget(5)
			FTotalPage =  int((FTotalCount-1)/FPageSize) +1
		end if	
		rsget.Close
 
	    if (FTotalCount<1) then Exit Sub '' 2017/01/11 추가
	        
'#정렬방법 선택(2008.07.21;허진원 추가)
		Select Case FRectSort
			Case "idxAcd"
				strSubSort= " idx asc"
			Case "idxDcd"   
				strSubSort= " idx desc"
			Case "pntAcd"   
				strSubSort= " t1.totalPoint asc"
			Case "pntDcd"   
				strSubSort= " t1.totalPoint desc"
			Case "iidAcd"   
				strSubSort= " t1.itemid asc"
			Case "iidDcd"   
				strSubSort= " t1.itemid desc"
			Case Else       
				strSubSort= " idx desc"
		end Select
		dim iSPageNo, iEPageNo
		iSPageNo = (FPageSize*(FCurrPage-1)) + 1
		iEPageNo = FPageSize*FCurrPage	
		
		'// 목록 접수 //
		SQL =	"Select TB.* FROM ( "&vbCrlf
		SQL = 	SQL & " SELECT ROW_NUMBER() OVER (ORDER BY  "&strSubSort&" ) as RowNum ,"
		SQL = 	SQL &	"	t1.idx, t1.userid " & vbCrlf
		SQL = 	SQL &		"	, Case t1.gubun when '0' then '일반' " & vbCrlf
		SQL = 	SQL & "		when '14' then '매니아' " & vbCrlf
		SQL = 	SQL &		"		else '기타' " & vbCrlf
		SQL = 	SQL &		"	  end as gubun " & vbCrlf
		SQL = 	SQL &		"	, t1.contents, t1.isUsing, t2.makerid, t2.itemname, t1.itemid " & vbCrlf
		SQL = 	SQL &		"	, t1.totalPoint, t1.Point_Function, t1.Point_Design, t1.Point_Price, t1.Point_Satisfy, t1.regdate ,t1.File1 , t1.File2 " & vbCrlf	
		SQL = 	SQL &		"	From db_board.dbo.tbl_Item_Evaluate as t1 " & vbCrlf
		SQL = 	SQL &		"	inner join db_item.[dbo].tbl_item as t2 on t1.itemid=t2.itemid " & vbCrlf
		SQL = 	SQL &		"	left outer join db_item.dbo.tbl_display_cate_item as dci on t2.itemid = dci.itemid and dci.isdefault ='Y'	 " & vbCrlf
		SQL = 	SQL &	 	" where t2.makerid = '"&FRectMakerid &"' "
		SQL = 	SQL &		AddSQL&vbCrlf
		SQL = 	SQL &	 ") AS TB "&vbCrlf
		SQL = 	SQL &			" WHERE TB.RowNum Between "&iSPageNo&" AND "  &iEPageNo & " "& vbCrlf
		Select Case FRectSort
			Case "idxAcd"
				SQL = SQL & "order by idx asc"
			Case "idxDcd"
				SQL = SQL & "order by idx desc"
			Case "pntAcd"
				SQL = SQL & "order by totalPoint asc"
			Case "pntDcd"
				SQL = SQL & "order by totalPoint desc"
			Case "iidAcd"
				SQL = SQL & "order by itemid asc"
			Case "iidDcd"
				SQL = SQL & "order by itemid desc"
			Case Else
				SQL = SQL & "order by idx desc"
		end Select
 
		rsget.CursorLocation = adUseClient
		rsget.Open SQL,dbget,adOpenForwardOnly, adLockReadOnly
		 FResultCount =  rsget.RecordCount

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then 
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

	 
end Class
%>