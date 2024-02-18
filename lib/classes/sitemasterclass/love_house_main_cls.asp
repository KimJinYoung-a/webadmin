<%

Class LoveHouse
	public FIdx()
	public FImage()
	public FLink()
	public FIsusing()
	public FViewYn()
	public FRectIdx
	
	public FTotalCount
	public FTotalPage
	public FResultcount
	
	
	public FScrollCount
	public FPageSize
	public FCurrPage
	
	
	Private Sub Class_Initialize()
	
	End Sub
	
	Private Sub Class_Terminate()
	
	End Sub
	
	Public Sub GetLoveMainList()
	
		dim sqlStr,i
		
		sqlStr = "select count(idx) as ccnt" + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_love_house_winner" + vbcrlf
		'sqlStr = sqlStr + " where isusing='Y'" + vbcrlf
		if FRectIdx<>"" then
		sql = sql + " where idx='" + CStr(FRectIdx) + "'" + vbcrlf
		end if
		rsget.open sqlStr,dbget,1
		
		FTotalCount=rsget("ccnt")
		
		rsget.close
		
		sqlStr = "select top " + CStr(FCurrPage*FPageSize ) + " idx, mainimage, viewYN, linkidx,isusing" + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_love_house_winner" + vbcrlf
		
		'sqlStr = sqlStr + " where isusing='Y'" + vbcrlf
		if FRectIdx<>"" then
		sql = sql + " where idx='" + CStr(FRectIdx) + "'" + vbcrlf
		end if
		sqlStr = sqlStr + " order by windate desc" + vbcrlf
		
		'response.write sqlStr
		'dbget.close()	:	response.End
		rsget.pagesize=FPageSize
		rsget.open sqlStr,dbget,1
		
		FResultCount=rsget.recordcount-((FCurrPage-1)*FPageSize)
		
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		
		if not rsget.eof then
			i=0 
			Redim Preserve FIdx(FResultCount)
			Redim Preserve FImage(FResultCount)
			Redim Preserve FLink(FResultCount)
			Redim Preserve FIsusing(FResultCount)
			Redim Preserve FViewYn(FResultCount)
			rsget.absolutepage=FCurrPage
			
			Do until rsget.eof 
				FIdx(i) =rsget("idx")
				FImage(i)="http://imgstatic.10x10.co.kr/contents/lovehousewin/" + rsget("mainimage")
				Flink(i)= db2html(rsget("linkidx"))
				FViewYn(i) = rsget("viewYN")
				FIsusing(i)=rsget("isusing")
				rsget.movenext
				i=i+1
			loop
		end if
		
		
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


End Class

Class LoveHouseOne
	public FIdx
	public FImage
	public FLink
	public FViewYn
	public FWinImage
	public FLoveMap
	public FIsusing
	Public Fuserid	
	Public Fwindate
	public FRectIdx
	public FResultCount
	public Fitemid1
	public Fitemid2
	public Fitemid3
	public Fitemid4
	Private Sub Class_Initialize()
	
	End Sub
	
	Private Sub Class_Terminate()
	
	End Sub
	
	Public Sub GetLoveMainOne()
		dim sqlStr
		sqlStr = "select top 1 idx, userid, windate, mainimage, viewYN, linkidx, winimage, lovemap, isusing, itemid1,itemid2,itemid3,itemid4" + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_love_house_winner" + vbcrlf
		sqlStr = sqlStr + " where idx='" + CStr(FRectIdx) + "'" + vbcrlf
		
		
		rsget.open sqlStr,dbget,1
		
		FResultCount=rsget.recordcount
		
		if not rsget.eof then
			FIdx =rsget("idx")
			FImage="http://imgstatic.10x10.co.kr/contents/lovehousewin/" + rsget("mainimage")
			if rsget("linkidx")<>"" then
				FLink= db2html(rsget("linkidx"))
			else
				FLink="winner_love_house_" + CStr(idx) + ".asp"
			end if
			
			Fuserid = rsget("userid")
			Fwindate = rsget("windate")
			FViewYn = rsget("viewYN")
			FWinImage= "http://imgstatic.10x10.co.kr/contents/lovehousewin/" + db2html(rsget("winimage"))
			FLoveMap=db2html(rsget("lovemap"))
			FIsusing=rsget("isusing")
			Fitemid1=rsget("itemid1")
			Fitemid2=rsget("itemid2")
			Fitemid3=rsget("itemid3")
			Fitemid4=rsget("itemid4")
			i=i+1
		end if
		
		
	End Sub
end Class

class CLoveHouseMainItem

	public Fidx
	public Fsmall_img
	public FIsusing

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CLoveHouseMain
	public FMasterItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FIsusing
	
    Private Sub Class_Initialize()
		redim FMasterItemList(0)
		FCurrPage = 1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub GetLoveHouseMain()
		dim sqlStr
		dim i

		'###########################################################################
		'상품 총 갯수 구하기
		'###########################################################################
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].tbl_lovehouse_main" + vbcrlf
'		sqlStr = sqlStr & " where title <> ''"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		''#################################################
		''데이타
		''#################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage) & "" + vbcrlf
		sqlStr = sqlStr & " idx,title,isusing" + vbcrlf
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].tbl_lovehouse_main" + vbcrlf
		sqlStr = sqlStr & " order by idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
		do until rsget.EOF
				
					set FMasterItemList(i) = new CLoveHouseMainItem

					FMasterItemList(i).Fidx = rsget("idx")
					FMasterItemList(i).Fsmall_img = "http://imgstatic.10x10.co.kr/contents/interior_theme/" & rsget("smallimg")
					FMasterItemList(i).FIsusing = rsget("isusing")
				rsget.movenext
				i=i+1

			loop
		end if
		rsget.Close
	end sub


	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function
	
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function
	
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1 
	end Function
end Class
%>