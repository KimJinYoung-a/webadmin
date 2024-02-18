<%
'###########################################################
' Description : 베스트 위시 리스트 통계
' History : 2008.06.23 한용민 생성
'###########################################################
%>
<%
class cwishlist_oneitem
    public fitemid_count
	public fregdate_count
	public fitemid
	public fsmallimage
	public fitemname
	public fmakerid
	public fmwdiv
	public FImageSmall
	public fsellcash
	public forgsuplycash
	public FLowprice
	public FNvRegdate

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

end Class

Class cwishlist
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount

	public frectstartdate
	public frectenddate
	public frectcdl
	public FRectSellY
	public frectordertype
	public frectmincash
	public frectmaxcash
	public frectipgocheck
	public frectnewitem
	public frectdisp1
	public FRectNvshop
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub
	
	public Sub fwishlist()
		dim sqlStr, i
		
		sqlStr = ""
		sqlStr = "select"	+vbcrlf 
		sqlStr = sqlStr & " count(b.itemid) as cnt"	+vbcrlf 
		sqlStr = sqlStr & " from ("	+vbcrlf 
		sqlStr = sqlStr & " 	select itemid , count(itemid) as itemid_count"	+vbcrlf 
		sqlStr = sqlStr & " 	,sum(datediff(day, regdate , getdate())) as regdate_count"	+vbcrlf 
		sqlStr = sqlStr & " 	from db_my10x10.dbo.tbl_myfavorite with (noLock)"	+vbcrlf 
		sqlStr = sqlStr & " 	where regdate between '"& frectstartdate &"' and '"& dateadd("d",1,frectenddate) &"'"	+vbcrlf
		sqlStr = sqlStr & " 	group by itemid" 	+vbcrlf 
		sqlStr = sqlStr & " 	) as a"	+vbcrlf 
		sqlStr = sqlStr & " join db_item.dbo.tbl_item b with (noLock)"	+vbcrlf 
		sqlStr = sqlStr & " on a.itemid = b.itemid"	+vbcrlf  
		sqlStr = sqlStr & " left join db_temp.[dbo].[tbl_tmp_naver_lowprice] as l with (noLock) on a.itemid = l.itemid "	+vbcrlf 
		sqlStr = sqlStr & " where 1=1" 	+vbcrlf 
			if frectcdl <> "" then
				sqlStr = sqlStr & " and b.cate_large = '"& frectcdl &"'"	+vbcrlf 		
			end if
			if FRectSellY <> "" then
				sqlStr = sqlStr & " and b.sellyn = 'Y'"	+vbcrlf 		
			end if
			if frectipgocheck <> ""then
				sqlStr = sqlStr & " and b.sellcash between "&frectmincash&" and "&frectmaxcash&""	+vbcrlf 				
			end if
			if frectnewitem <> "" Then
				sqlStr = sqlStr & " and datediff(d,b.regdate,getdate()) <= 14 "	+vbcrlf
			End if
			if frectdisp1 <> "" Then
				sqlStr = sqlStr & " and b.dispcate1 = '" & frectdisp1 & "' "	+vbcrlf
			End If
			Select Case FRectNvshop
				Case "nvshopY"		sqlStr = sqlStr & " and isnull(l.itemid, '') <> '' "	+vbcrlf
				Case "nvshopN"		sqlStr = sqlStr & " and isnull(l.itemid, '') = '' "	+vbcrlf
			End Select				
			
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.close
		
		sqlStr = ""
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " a.itemid_count ,a.regdate_count"	+vbcrlf 
		sqlStr = sqlStr & " , b.itemid , b.smallimage , b.itemname , b.makerid"	+vbcrlf 
		sqlStr = sqlStr & " , b.makerid , b.sellcash , b.orgsuplycash"	+vbcrlf 
		sqlStr = sqlStr & " ,(case when b.mwdiv='M' or b.mwdiv='W' then '텐배송' else '업체배송' end ) as mwdiv" 	+vbcrlf 
		sqlStr = sqlStr & " ,isnull(l.lowprice, '') as lowprice, l.regdate as nvRegdate " 	+vbcrlf 
		sqlStr = sqlStr & " from ("	+vbcrlf 
		sqlStr = sqlStr & " 	select itemid , count(itemid) as itemid_count"	+vbcrlf 
		sqlStr = sqlStr & " 	,sum(datediff(day, regdate , getdate())) as regdate_count"	+vbcrlf 
		sqlStr = sqlStr & " 	from db_my10x10.dbo.tbl_myfavorite with (noLock)"	+vbcrlf 
		sqlStr = sqlStr & " 	where regdate between '"& frectstartdate &"' and '"& dateadd("d",1,frectenddate) &"'"	+vbcrlf
		sqlStr = sqlStr & " 	group by itemid" 	+vbcrlf 
		sqlStr = sqlStr & " 	) as a"	+vbcrlf 
		sqlStr = sqlStr & " join db_item.dbo.tbl_item b with (noLock)"	+vbcrlf 
		sqlStr = sqlStr & " on a.itemid = b.itemid" 	+vbcrlf 
		sqlStr = sqlStr & " left join db_temp.[dbo].[tbl_tmp_naver_lowprice] as l with (noLock) on a.itemid = l.itemid "	+vbcrlf 
		sqlStr = sqlStr & " where 1=1" 	+vbcrlf 
			if frectcdl <> "" then
				sqlStr = sqlStr & " and b.cate_large = '"& frectcdl &"'"	+vbcrlf 		
			end if
			if FRectSellY <> "" then
				sqlStr = sqlStr & " and b.sellyn = 'Y'"	+vbcrlf 		
			end if	
			if frectipgocheck <> ""then
				sqlStr = sqlStr & " and b.sellcash between "&frectmincash&" and "&frectmaxcash&""	+vbcrlf 				
			end if	
			if frectnewitem <> "" Then
				sqlStr = sqlStr & " and datediff(d,b.regdate,getdate()) <= 14 "	+vbcrlf
			End if
			if frectdisp1 <> "" Then
				sqlStr = sqlStr & " and b.dispcate1 = '" & frectdisp1 & "' "	+vbcrlf
			End If
			Select Case FRectNvshop
				Case "nvshopY"		sqlStr = sqlStr & " and isnull(l.itemid, '') <> '' "	+vbcrlf
				Case "nvshopN"		sqlStr = sqlStr & " and isnull(l.itemid, '') = '' "	+vbcrlf
			End Select	

			if frectordertype = "select" then
				''sqlStr = sqlStr & " order by a.regdate_count desc"	+vbcrlf
				sqlStr = sqlStr & " order by a.itemid_count desc"	+vbcrlf
			else
				sqlStr = sqlStr & " order by b.sellcash desc"	+vbcrlf 
			end if	

		'response.write sqlStr &"<br>"
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
				
				set FItemList(i) = new cwishlist_oneitem
				
					FItemList(i).fitemid_count = rsget("itemid_count")
					FItemList(i).fregdate_count = rsget("regdate_count")
					FItemList(i).fitemid = rsget("itemid")				
					FItemList(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")				
					FItemList(i).fitemname = rsget("itemname")				
					FItemList(i).fmakerid = rsget("makerid")
					FItemList(i).fmwdiv = rsget("mwdiv")											
					FItemList(i).fsellcash = rsget("sellcash")
					FItemList(i).fmwdiv = rsget("mwdiv")	
					FItemList(i).forgsuplycash = rsget("orgsuplycash")
					FItemList(i).FLowprice = rsget("lowprice")
					FItemList(i).FNvRegdate = rsget("nvRegdate")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

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