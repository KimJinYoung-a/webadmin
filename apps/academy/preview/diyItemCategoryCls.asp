<%
'#######################################################
'	History	: 2010.10.25 허진원 생성
'			  2010.11.10 한용민 수정	
'	Description : DIY샵 카테고리 페이지 클래스
'#######################################################

Class CCatemanageItem
	public Fidx
	public Fmakerid
	public Ftitleimgurl
	Public Fmodelitem
	Public FImgSmall
	Public FImgList
	Public Fitemid
	Public Fitemname
	Public FSellCash
	Public FSailYN
	Public FSailPrice
	Public FOrgPrice
	Public FGiftyn
	Public Fitemcouponyn
	Public Fitemcoupontype
	Public Fitemcouponvalue
	Public NowEventDoing

	Private Sub Class_Initialize()
		NowEventDoing = false   
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CCatemanager
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	Public FRectCDL
	Public FRectCDM

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'//chtml/make_category_left_bestBrand_JS.asp
	public Function GetCategoryLeftbestBrand()
		dim sqlStr,i

		sqlStr = "select top 3" + vbcrlf
		sqlStr = sqlStr + " makerid, imgfile" + vbcrlf
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_category_left_bestbrand"
		sqlStr = sqlStr + " where 1=1 and isUsing='Y' " + vbcrlf
		
		If FRectCDL <> "" then
			sqlStr = sqlStr + " and cdl = '" + CStr(FRectCDL) + "'"
		End If
		
		sqlStr = sqlStr + " order by sortNo"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fmakerid		= db2html(rsACADEMYget("makerid"))
				FItemList(i).Ftitleimgurl	= fingersImgUrl + "/left/bestbrand/" + rsACADEMYget("imgfile")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end Function

	'// 페이지 이동
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

'// 카테고리 Histoty 출력(중분류까지만 표현)
Sub printCategoryHistory(cd1,cd2)
	dim strHistory, strLink, SQL
    dim StrLogTrack : StrLogTrack ="" ''logger Tracking
    
	'히스토리 기본
	strHistory = "<a href='/'>HOME</a>" &_
				" &gt; <a href='/diyshop/shop_main.asp'>DIY SHOP</a>"

	'// 카테고리 이름 접수
	SQL =	"exec [db_academy].[dbo].sp_academy_category_history_name '" & cd1 & "', '" & cd2 & "'"
	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.CursorType = adOpenStatic
	rsACADEMYget.LockType = adLockOptimistic
	rsACADEMYget.Open SQL, dbACADEMYget

	if NOT(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		'#대분류 출력
		strLink = "/diyshop/shop_list.asp"
		if cdl<>"" then
			if cdm<>"" then
				strHistory = strHistory & " &gt; <a href='" & strLink & "?cdl=" & cdl & "' >" & db2html(rsACADEMYget(0)) & "</a>"
			else
				strHistory = strHistory & " &gt; <a href='" & strLink & "?cdl=" & cdl & "' ><strong>" & db2html(rsACADEMYget(0)) & "</strong></a>"
			end if
			
			StrLogTrack = db2html(rsACADEMYget(0))
		end if

		'#중분류 출력
		strLink = "/diyshop/shop_list.asp"
		if cdm<>"" then
			if cds<>"" then
				strHistory = strHistory & " &gt; <a href='" & strLink & "?cdl=" & cdl & "&cdm=" & cdm & "' >" & db2html(rsACADEMYget(1)) & "</a>"
			else
				strHistory = strHistory & " &gt; <a href='" & strLink & "?cdl=" & cdl & "&cdm=" & cdm & "' ><strong>" & db2html(rsACADEMYget(1)) & "</strong></a>"
			end if
			
			StrLogTrack = StrLogTrack & ">" & db2html(rsACADEMYget(1))
		end if

	end if

	rsACADEMYget.Close

	Response.Write strHistory
end Sub

'// 카테고리 Histoty 출력(중분류까지만 표현)
Sub printMidCategoryName(cd1,cd2)
	dim strHistory, SQL
    dim StrLogTrack : StrLogTrack ="" ''logger Tracking
    
	'// 카테고리 이름 접수
	SQL =	"exec [db_academy].[dbo].sp_academy_category_history_name '" & cd1 & "', '" & cd2 & "'"
	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.CursorType = adOpenStatic
	rsACADEMYget.LockType = adLockOptimistic
	rsACADEMYget.Open SQL, dbACADEMYget

	if NOT(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		'#중분류 출력
		if cdm<>"" then
				strHistory = db2html(rsACADEMYget(1)) 
		end if
	end if

	rsACADEMYget.Close

	Response.Write strHistory
end Sub

'// 카테고리 Histoty 출력(중분류까지만 표현)
Sub printLargeCategoryTitle(cd1,cd2)
	dim strHistory, SQL
    dim StrLogTrack : StrLogTrack ="" ''logger Tracking
    
	'// 카테고리 이름 접수
	SQL =	"exec [db_academy].[dbo].sp_academy_category_history_name '" & cd1 & "', '" & cd2 & "'"
	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.CursorType = adOpenStatic
	rsACADEMYget.LockType = adLockOptimistic
	rsACADEMYget.Open SQL, dbACADEMYget

	if NOT(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		'#대분류 출력
		if cdl<>"" then
				strHistory = db2html(rsACADEMYget(0)) 
		end if
	end if

	rsACADEMYget.Close

	Response.Write strHistory
end Sub


Function getDisplayCateNameDB(disp)
	Dim SQL

	'유효성 검사
	if disp="" then
		getDisplayCateNameDB = "전체보기"
		Exit Function
	end if

	SQL = "select [db_academy].[dbo].getDisplayCateName_Academy('" & disp & "')"
	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open SQL, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

		if NOT(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			getDisplayCateNameDB = db2html(rsACADEMYget(0))
		else
			getDisplayCateNameDB = "전체보기"
		end if
	rsACADEMYget.Close
End Function

'//위시 상품
Function getIsMyFavItem(uid,iid)
	dim sqlStr
	sqlStr = "select count(f.itemid) as cnt from db_academy.dbo.tbl_diy_myfavorite as f inner join db_academy.dbo.tbl_diy_item as i on f.itemid = i.itemid where f.userid='" & uid & "' and f.itemid = '" & iid & "'"

	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

	if not rsACADEMYget.EOF then
		if rsACADEMYget("cnt")>0 then
			getIsMyFavItem = true
		else
			getIsMyFavItem = false
		end If
	End If 
	rsACADEMYget.Close
end Function

'//my 좋은작가 여부
Function getIsMyFavauthor(uid,iid)
	dim sqlStr
	sqlStr = "select count(makerid) as cnt from db_academy.dbo.tbl_diy_user_teacherlist where userid='" & uid & "' and makerid = '" & iid & "'"

	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  

	if not rsACADEMYget.EOF then
		if rsACADEMYget("cnt")>0 then
			getIsMyFavauthor = true
		else
			getIsMyFavauthor = false
		end If
	End If 
	rsACADEMYget.Close
end Function

'구매후기 카운트
Function getIsEvaluateCnt(itemid)
	dim sqlStr
	sqlStr = "select count(*) as cnt FROM db_academy.dbo.tbl_diy_item_Evaluate as e inner join db_academy.dbo.tbl_diy_item as i on e.itemid = i.itemid where e.isusing = 'Y' and e.itemid = " & itemid

	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

	if not rsACADEMYget.EOF Then
		getIsEvaluateCnt = rsACADEMYget("cnt")
	Else
		getIsEvaluateCnt = 0
	End If 
	rsACADEMYget.Close
end Function

'QnA카운트
Function getItemIsQnACnt(itemid)
	dim sqlStr
	sqlStr = "select count(*) as cnt from db_academy.dbo.tbl_academy_qna_new as q inner join db_academy.dbo.tbl_diy_item as i on q.itemid = i.itemid where q.pagegubun = 'D' and q.reply_depth = 0 and q.isusing = 'Y' and q.itemid = " & itemid

	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

	if not rsACADEMYget.EOF Then
		getItemIsQnACnt = rsACADEMYget("cnt")
	Else
		getItemIsQnACnt = 0
	End If 
	rsACADEMYget.Close
end Function

''new 아이콘 설정
Function getIsNewIcon(makerid)
	dim sqlStr
	sqlStr = "select top 1 datediff(DD,regdate ,getdate()) as limitday from db_academy.dbo.tbl_diy_item where makerid = '"& makerid &"' order by regdate asc"

	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly, adLockReadOnly

	if not rsACADEMYget.EOF then
		if rsACADEMYget("limitday") < 30 then
			getIsNewIcon = true
		else
			getIsNewIcon = false
		end If
	End If 
	rsACADEMYget.Close
end Function

'구매후기 카운트
Function getIsEvalavgCnt(itemid)
	dim sqlStr
	sqlStr = "SELECT isnull(AVG(totalpoint),0) AS avgpoint FROM db_academy.[dbo].[tbl_diy_Item_Evaluate] where isusing = 'Y' and itemid = " & itemid

	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

	if not rsACADEMYget.EOF Then
		getIsEvalavgCnt = rsACADEMYget("avgpoint")
	Else
		getIsEvalavgCnt = 0
	End If 
	rsACADEMYget.Close
end Function
%>