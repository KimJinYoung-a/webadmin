<%
'###########################################################
' Description : HITCHHIKER ADMIN 프리뷰 클래스
' Hieditor : 2014.08.05 유태욱 생성
'###########################################################
%>
<%
class CHitchhikerItem
	public Fidx
	public FReqcash
	public FReqTitle
	public FReqSdate
	public FReqEdate
	public FReqIsusing
	public FReqSortnum
	public FreqRegdate
	public FReqmileage
	public FReqpreview_detail
	public FReqpreview_thumbimg
	
	public FimgCnt

	public FMimgCnt

	public FIsusing
	public Fsortnum
	public FRegdate
	public Fdetailidx
	public Fmasteridx
	public Fpreviewimg
end class

class CHitchhikerPreview
	public FItemList()
	Public Fgubun
	Public Ftitle
	Public FIsusing
	public FrectIdx
	Public FOneItem
	public FCurrPage
	public FPageSize
	public FPageCount
	public FTotalPage
	public FTotalCount
	public FScrollCount
	public FResultCount
	public FValiddate	

	public FrectDevice
	'//admin/hitchhiker/preview/index.asp
	public sub fnGetHitchhikerList
		dim sqlStr,i, sqlsearch
				
		if Ftitle <> "" Then
			sqlsearch = sqlsearch & " AND m.title like'%"& Ftitle &"%'"
		end if

		if FIsusing <> "" Then
			sqlsearch = sqlsearch & " AND m.isusing ='"& FIsusing &"'"
		end if
		
        if FValiddate<>"" then
            sqlsearch = sqlsearch + " AND m.edate > getdate()"
        end if
		
		'글의 총 갯수 구하기
		sqlStr = "SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_hitchhiker_preview_list m"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
		
		'DB 데이터 리스트
		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.idx, m.title, m.preview_detail, m.preview_thumbimg, m.sdate, m.edate, m.sortnum, m.isusing, m.cash, m.mileage, m.regdate"
		sqlStr = sqlStr & " ,(select count(*) from db_sitemaster.dbo.tbl_hitchhiker_preview_detail as d where m.idx = d.masteridx and d.device='W' and d.isusing='Y') as imgCnt "
		sqlStr = sqlStr & " ,(select count(*) from db_sitemaster.dbo.tbl_hitchhiker_preview_detail as d where m.idx = d.masteridx and d.device='M' and d.isusing='Y') as MimgCnt "
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_hitchhiker_preview_list as m"	
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		sqlStr = sqlStr & " ORDER BY m.sortnum ASC, m.idx DESC"
		rsget.pagesize = FPageSize

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CHitchhikerItem
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).FReqcash = rsget("cash")
					FItemList(i).FimgCnt = rsget("imgCnt")
					FItemList(i).FMimgCnt = rsget("MimgCnt")
					FItemList(i).FReqTitle = rsget("title")
					FItemList(i).FReqSdate = rsget("sdate")
					FItemList(i).FReqEdate = rsget("edate")
					FItemList(i).FReqIsusing = rsget("isusing")
					FItemList(i).FReqSortnum = rsget("sortnum")
					FItemList(i).FReqmileage = rsget("mileage")
					FItemList(i).FreqRegdate = rsget("regdate")
					FItemList(i).FReqpreview_detail = rsget("preview_detail")
					FItemList(i).FReqpreview_thumbimg = rsget("preview_thumbimg")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
	
	'//admin/hitchhiker/preview/hitchhiker_preview_write.asp		'//admin/hitchhiker/preview/iframe_hitchhiker_preview.asp
	Public Sub sbpreviewwrite
		Dim sqlStr, i, sqlsearch
		
		if FrectIdx="" then exit Sub
		
		if FrectIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx = "&FrectIdx&""
		end if
		
		sqlStr = "SELECT TOP 1"
		sqlStr = sqlStr & " m.idx, m.title, m.preview_detail, m.preview_thumbimg, m.sdate, m.edate, m.sortnum, m.isusing, m.regdate, m.cash, m.mileage"
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_hitchhiker_preview_list m"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
		
		ftotalcount = rsget.recordcount
        SET FOneItem = new CHitchhikerItem
	        If Not rsget.Eof then
	        	FOneItem.Fidx = rsget("idx")
	        	FOneItem.FReqcash = rsget("cash")
	        	FOneItem.FReqTitle = rsget("title")
	        	FOneItem.FReqSdate = rsget("sdate")
				FOneItem.FReqEdate = rsget("edate")
	        	FOneItem.FReqSortnum = rsget("sortnum")
	        	FOneItem.FReqIsusing = rsget("isusing")
	        	FOneItem.FReqRegdate = rsget("regdate")
	        	FOneItem.FReqmileage = rsget("mileage")
	        	FOneItem.FReqpreview_detail = rsget("preview_detail")
	        	FOneItem.FReqpreview_thumbimg = rsget("preview_thumbimg")
        	End If
        rsget.Close
	End Sub

	'/admin/hitchhiker/preview/iframe_hitchhiker_preview.asp
	Public Sub sbpreviewDetaillist
		Dim sqlStr, i, sqlsearch
		
		if FrectIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx='"& FrectIdx &"'"
		end if

		if FrectDevice<>"" then
			sqlsearch = sqlsearch & " and d.device='"& FrectDevice &"'"
		end if

		sqlStr = "SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_hitchhiker_preview_list as m"
		sqlStr = sqlStr & " JOIN db_sitemaster.dbo.tbl_hitchhiker_preview_detail as d"
		sqlStr = sqlStr & " 	on m.idx=d.masteridx and d.isusing='Y'"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
		
		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " d.detailidx, d.masteridx, d.previewimg, d.isusing, d.regdate, d.sortnum"
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_hitchhiker_preview_list as m"
		sqlStr = sqlStr & " JOIN db_sitemaster.dbo.tbl_hitchhiker_preview_detail as d"
		sqlStr = sqlStr & " 	on m.idx=d.masteridx and d.isusing='Y'"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		sqlStr = sqlStr & " ORDER BY d.sortnum ASC"
		rsget.pagesize = FPageSize
	
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CHitchhikerItem
					FItemList(i).FIsusing = rsget("isusing")
					FItemList(i).Fsortnum = rsget("sortnum")
					FItemList(i).FRegdate = rsget("regdate")
					FItemList(i).Fdetailidx	= rsget("detailidx")
					FItemList(i).Fmasteridx	= rsget("masteridx")
					FItemList(i).fpreviewimg= rsget("previewimg")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
end class
%>