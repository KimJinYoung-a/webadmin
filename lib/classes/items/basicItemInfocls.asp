<%
'##### 제품정보 저장 처리 #####
Class CItemListItems
	public Fitemid
	public Fitemname
	public Fsellcash
	public Fmakername
	public Fregdate
	public Fmakerid

	public FLinkitemid
	public FImgSmall
	public FSellyn
    public FisUsing
    
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

'##### 제품 목록 #####
class CItemlist
	'// 변수 선언 //
	public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectCurrState
	public FRectItemId
	public FRectMakerId

	public FRectRegState

	'// 세팅 초기화 //
	Private Sub Class_Initialize()
	redim FItemList(0)
		FCurrPage =1
		FPageSize = 5
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub


	Private Sub Class_Terminate()

	End Sub


	'// 제품 목록 접수 //
	public sub ProductList()
		dim SQL, addSQL, i,wheredetail

		if Not(FRectItemId="" or isNull(FRectItemId)) then
			addSQL = " and i.itemid='" & FRectItemId & "'"
		end if

		if Not(FRectMakerId="" or isNull(FRectMakerId)) then
			addSQL = addSQL & " and i.makerid='" & FRectMakerId & "'"
		end if

		'##### 등록대기 상품 총 갯수 구하기 #####
		if FRectRegState="W" then
			SQL =	"select count(i.itemid) as cnt " & VbCrlf
			SQL =	SQL + " from [db_temp].[dbo].tbl_wait_item i" & VbCrlf
			SQL =	SQL + " Where i.isusing='Y' " & addSQL
		else
			SQL =	"select count(i.itemid) as cnt " & VbCrlf
			SQL =	SQL + " from [db_item].[dbo].tbl_item i" & VbCrlf
			SQL =	SQL + " Where i.isusing='Y' " & VbCrlf
			SQL =	SQL + " and i.itemdiv<>21 " & addSQL
		end if

		rsget.Open SQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		'##### 등록대기 상품 데이터 #####
		if FRectRegState="W" then
			SQL =	"select top " + Cstr(FPageSize * FCurrPage) & VbCrlf
			SQL =	SQL + "	i.itemid, i.makerid, m.imgsmall as smallimage, i.itemname, i.sellcash " & VbCrlf
			SQL =	SQL + "	, i.makername, i.regdate, i.sellyn, i.isusing " & VbCrlf
			SQL =	SQL + " from [db_temp].[dbo].tbl_wait_item i left join [db_temp].[dbo].tbl_wait_item_image m on i.itemid=m.itemid" & VbCrlf
			SQL =	SQL + " Where 1=1 " & addSQL & VbCrlf
			SQL =	SQL + " order by i.itemid Desc "
		else
			SQL =	"select top " + Cstr(FPageSize * FCurrPage) & VbCrlf
			SQL =	SQL + "	i.itemid, i.makerid, i.smallimage, i.itemname, i.sellcash " & VbCrlf
			SQL =	SQL + "	, c.makername, i.regdate, i.sellyn, i.isusing " & VbCrlf
			SQL =	SQL + " from [db_item].[dbo].tbl_item i" & VbCrlf
			SQL =	SQL + "     left join [db_item].[dbo].tbl_item_Contents C" & VbCrlf
			SQL =	SQL + "     on i.itemid=C.itemid" & VbCrlf
			SQL =	SQL + " Where i.isusing='Y' " & VbCrlf
			SQL =	SQL + " and i.itemdiv<>21 " & addSQL
			SQL =	SQL + " order by i.itemid Desc "
		end if

		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))

		FTotalPage = clng(FTotalCount\FPageSize) + 1


		redim preserve FItemList(FResultCount)	'배열을 크기를 결과 수많큼 늘린다.

		i=0
		if Not(rsget.EOF or rsget.BOF) then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CItemListItems
				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).Fmakerid = rsget("makerid")
			    FItemList(i).Fitemname = db2html(rsget("itemname"))
				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).Fmakername = rsget("makername")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).FSellyn = rsget("sellyn")
				FItemList(i).FisUsing = rsget("isusing")
				
				if FRectRegState="W" then
					FItemList(i).FImgSmall = "http://partner.10x10.co.kr/waitimage/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallimage")
				else
					FItemList(i).FImgSmall = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallimage")
				end if
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub


	'// 이전 페이지 검사 //
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function


	'// 다음 페이지 검사 //
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function


	'// 초기 페이지 반환 //
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

'##### 상품 조건 검색  #####
Class CSearchItemList

	public FBrand
	public FLCategory
	public FMCategory
	public FSCategory
	public FitemId
	public FitemName
	public FdispYN
	public FsellYN
	public FImgUrl
	public FsortDiv
	
	public FCPage	'Set 현재 페이지
	public FPSize	'Set 페이지 사이즈
	public FTotCnt
	
	public Function fnGetItemList
		Dim strSqlCnt, strSql,strSqlAdd,iDelCnt, ordSQL
		
		'// 조건 쿼리
		IF FBrand <> "" THEN
			strSqlAdd = " and makerid = '"&FBrand&"' "
		END IF
		IF 	FLCategory <> "" THEN
			strSqlAdd = strSqlAdd & " and cate_large = '"&FLCategory&"'"
		END IF	
		IF 	FMCategory <> "" THEN
			strSqlAdd = strSqlAdd & " and cate_mid = '"&FMCategory&"'"
		END IF	
		IF 	FSCategory <> "" THEN
			strSqlAdd = strSqlAdd & " and cate_small = '"&FSCategory&"'"
		END IF	
		IF 	FitemId <> "" THEN
			strSqlAdd = strSqlAdd & " and itemid in ("&FitemId&")"
		END IF	
		IF 	FitemName <> "" THEN
			strSqlAdd = strSqlAdd & " and itemname like '%"&FitemName&"%'"
		END IF	
		IF 	FdispYN <> "" THEN
			strSqlAdd = strSqlAdd & " and dispyn = '"&FdispYN&"'"
		END IF	
		IF 	FsellYN <> "" THEN
			strSqlAdd = strSqlAdd & " and sellyn = '"&FsellYN&"' "
		END IF	

		'// 정렬방법 선택
		Select Case FsortDiv
			Case "new"			'신상품
				ordSQL = " ORDER by itemid DESC "
			Case "highprice"	'고가격순
				ordSQL = " ORDER by sellcash DESC "
			Case "lowprice"		'저가격순
				ordSQL = " ORDER by sellcash ASC "
			Case "best"			'베스트셀러
				ordSQL = " ORDER by recentsellcount desc, sellcount desc, itemid desc "
			Case "brand"		'브랜드순
				ordSQL = " ORDER by makerid ASC "
			Case "sale"			'할인상품순
				ordSQL = " ORDER by sellcash/orgprice "
			Case Else
				ordSQL = " ORDER by itemid DESC "
		End Select

		'// 결과 카운트
		strSqlCnt = "SELECT count(itemid) FROM [db_item].[dbo].tbl_item WHERE itemid <> 0 "&strSqlAdd
		rsget.Open strSqlCnt,dbget,1
		IF Not rsget.EOF THEN
			FTotCnt = rsget(0)
		END IF
	
		rsget.Close	

		'// 목록 접수
		IF FTotCnt > 0 THEN
			iDelCnt =  (FCPage - 1) * FPSize
				
			strSql = " SELECT TOP  "&FPSize&" itemid, makerid, itemname, sellcash, buycash, dispyn, sellyn, isusing, mwdiv, limityn, limitno, limitsold, "&_
					" 		IsNull(makername,'') as makername, regdate, IsNull(smallimage,'') as imgsmall ,deliverytype "&_
					"  FROM [db_item].[dbo].tbl_item  "&_
					"	WHERE itemid not in ( SELECT TOP "&iDelCnt&" itemid FROM [db_item].[dbo].tbl_item WHERE itemid <> 0 "&_
					"	" & strSqlAdd & ordSQL & ")" & strSqlAdd & ordSQL
			rsget.Open strSql,dbget,1
			IF Not rsget.EOF THEN		
				fnGetItemList = rsget.getRows()				
			END IF	
			rsget.Close
		END IF
		
	End Function
	
		'// 판매종료  여부 
	public Function IsSoldOut(ByVal dispYN, ByVal sellYN, ByVal limitYN, ByVal limitNo, ByVal limitsold)
	 	 IsSoldOut = (dispYN = "N" or sellYN= "N") or (limitYN = "Y" and (clng(limitNo)-clng(limitsold)<= 0))
	end Function
	
	public Function fnGetImgUrl(ByVal itemid, ByVal imgname)
		fnGetImgUrl = "http://webimage.10x10.co.kr/image/small/"&GetImageSubFolderByItemid(itemid)&"/"&imgname		
	End Function	
End Class
%>