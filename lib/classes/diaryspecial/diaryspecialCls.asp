<%
'###########################################################
' Description : 다이어리 스페셜 관리
' Hieditor : 2015.10.05 유태욱 생성
'##########################################################

class CDiaryspecialItem
	public Fidx
	public Fpcmainimage
	public Fpcoverimage
	public Fpctext
	public Fmobileimage
	public Fmobiletext
	public Fitemid1
	public Fitemid2
	public Fitemid3
	public Fitemid4
	public Fitemid5
	public Flinkgubun
	public Flinkcode
	public FSortnum
	public FIsusing
	public FRegdate
	
	public Fdetailitemimage1
	public Fdetailitemimage2
	public Fdetailitemimage3
	public Fdetailitemimage4
	public Fdetailitemimage5
end class

class CDiaryspecial
	public FItemList()
	public Foneitem
	public Frecidx
	public FPageSize
	public FCurrPage
	public FTotalPage
	public FPageCount
	public FTotalCount
	public FScrollCount
	public FrectIsusing
	public FResultCount
	
	public FRectitemid
	public FRectevtcode
	
	
	public Sub fnGetDiaryspecial_oneitem()
	    dim sqlStr, sqlsearch
	
	if Frecidx <> "" Then
		sqlsearch = sqlsearch & " AND d.idx ='"& Frecidx &"'"
	end if

	    sqlStr = "Select top 1" & vbcrlf
		sqlStr = sqlStr & " d.idx, d.pcmainimage, d.pcoverimage, d.pctext, d.mobileimage, d.mobiletext"
		sqlStr = sqlStr & " ,d.linkgubun, d.linkcode, d.sortnum, d.isusing, d.regdate"
		sqlStr = sqlStr & " ,(select top 1 itemid from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=1) as itemid1"
		sqlStr = sqlStr & " ,(select top 1 itemid from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=2) as itemid2"
		sqlStr = sqlStr & " ,(select top 1 itemid from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=3) as itemid3"
		sqlStr = sqlStr & " ,(select top 1 itemid from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=4) as itemid4"
		sqlStr = sqlStr & " ,(select top 1 itemid from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=5) as itemid5"
		sqlStr = sqlStr & " ,(select top 1 detailitemimage from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=1) as detailitemimage1"
		sqlStr = sqlStr & " ,(select top 1 detailitemimage from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=2) as detailitemimage2"
		sqlStr = sqlStr & " ,(select top 1 detailitemimage from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=3) as detailitemimage3"
		sqlStr = sqlStr & " ,(select top 1 detailitemimage from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=4) as detailitemimage4"
		sqlStr = sqlStr & " ,(select top 1 detailitemimage from [db_diary2010].[dbo].[tbl_diaryspecial_detail] as dd where dd.midx = d.idx And itemordernum=5) as detailitemimage5"
		sqlStr = sqlStr & "  From [db_diary2010].[dbo].[tbl_diaryspecial] d"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by d.idx Desc"

'	    response.write sqlStr&"<br>"
	    rsget.Open SqlStr, dbget, 1
	    FResultCount = rsget.RecordCount
	    
	    set FOneItem = new CDiaryspecialItem
	    
	    if Not rsget.Eof then
			Foneitem.Fidx = rsget("idx")

			Foneitem.Fpcmainimage = rsget("pcmainimage")
			Foneitem.Fpcoverimage = rsget("pcoverimage")
			Foneitem.Fmobileimage = rsget("mobileimage")
			
			Foneitem.Fpctext = db2html(rsget("pctext"))
			Foneitem.Fmobiletext = db2html(rsget("mobiletext"))
			Foneitem.Fitemid1 = rsget("itemid1")
			Foneitem.Fitemid2 = rsget("itemid2")
			Foneitem.Fitemid3 = rsget("itemid3")
			Foneitem.Fitemid4 = rsget("itemid4")
			Foneitem.Fitemid5 = rsget("itemid5")
			Foneitem.Flinkgubun = rsget("linkgubun")
			Foneitem.Flinkcode = rsget("linkcode")
			Foneitem.Fsortnum = rsget("sortnum")
			Foneitem.Fisusing = rsget("isusing")
			Foneitem.Fregdate = rsget("regdate")

			Foneitem.Fdetailitemimage1 = rsget("detailitemimage1")
			Foneitem.Fdetailitemimage2 = rsget("detailitemimage2")
			Foneitem.Fdetailitemimage3 = rsget("detailitemimage3")
			Foneitem.Fdetailitemimage4 = rsget("detailitemimage4")
			Foneitem.Fdetailitemimage5 = rsget("detailitemimage5")
	    end if
	    rsget.Close
	end Sub
    
	public sub fnGetDiaryspecial
		dim sqlStr,i, sqlsearch

'		if Frectitemid <> "" Then
'			sqlsearch = sqlsearch & " AND itemid1 ='"& Frectitemid &"' Or itemid2 ='"& Frectitemid &"' Or itemid3 ='"& Frectitemid &"' Or itemid4 ='"& Frectitemid &"' Or itemid5 ='"& Frectitemid &"'"
'		end if

		if Frectevtcode <> "" Then
			sqlsearch = sqlsearch & " AND linkcode ='"& Frectevtcode &"'"
		end if

		if FrectIsusing <> "" Then
			sqlsearch = sqlsearch & " AND isusing ='"& FrectIsusing &"'"
		end if

		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from [db_diary2010].[dbo].[tbl_diaryspecial] "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, pcmainimage, pcoverimage, pctext, mobileimage, mobiletext"
		sqlStr = sqlStr & " ,linkgubun, linkcode, sortnum, isusing, regdate"
		sqlStr = sqlStr & " from [db_diary2010].[dbo].[tbl_diaryspecial] "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"

'		response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize		
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CDiaryspecialItem

					FItemList(i).Fidx = rsget("idx")

					FItemList(i).Fpcmainimage = rsget("pcmainimage")
					FItemList(i).Fpcoverimage = rsget("pcoverimage")
					FItemList(i).Fmobileimage = rsget("mobileimage")
					
					FItemList(i).Fpctext = db2html(rsget("pctext"))
					FItemList(i).Fmobiletext = db2html(rsget("mobiletext"))
'					FItemList(i).Fitemid1 = rsget("itemid1")
'					FItemList(i).Fitemid2 = rsget("itemid2")
'					FItemList(i).Fitemid3 = rsget("itemid3")
'					FItemList(i).Fitemid4 = rsget("itemid4")
'					FItemList(i).Fitemid5 = rsget("itemid5")
					FItemList(i).Flinkgubun = rsget("linkgubun")
					FItemList(i).Flinkcode = rsget("linkcode")
					FItemList(i).Fsortnum = rsget("sortnum")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).Fregdate = rsget("regdate")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	

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






	

		