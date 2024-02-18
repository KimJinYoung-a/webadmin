<%
'###########################################################
' Description : 감성모모 클래스
' Hieditor : 2009.10.28 한용민 생성
'###########################################################

Class chonor_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Fidx
	public Fyyyymm
	public Fgubun
	public Forderno
	public Fuserid
	public Fcontents
end class

class chonor_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public frectyyyymm
	public frectgubun

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'/admin/momo/honor/honor_list.asp '/frectgubun  1 명예의전당
	public sub fhonor_winner()
		dim sqlStr,i , sqlsearch
		
		if frectyyyymm <> "" then
			sqlsearch = sqlsearch & " and yyyymm = '"&frectyyyymm&"'" + vbcrlf
		end if
		if frectgubun <> "" then
			sqlsearch = sqlsearch & " and gubun = "&frectgubun&"" + vbcrlf
		end if
		
		'데이터 리스트 
		sqlStr = "select top 6"
		sqlStr = sqlStr & " idx ,yyyymm ,gubun ,orderno ,userid ,contents" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_winner" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by orderno asc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount

		redim preserve FItemList(ftotalcount)
	
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new chonor_oneitem
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fyyyymm = rsget("yyyymm")
				FItemList(i).fgubun = rsget("gubun")
				FItemList(i).forderno = rsget("orderno")	
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fcontents = db2html(rsget("contents"))
																
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
end class	

Class coneline_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fstartdate
	public fenddate
	public fwinnerdate
	public fcomment
	public fregdate
	public fisusing
	public fidx
	public fgubun
	public fuserid	
	public fcoinyn
	public fstats
	public fcommentcount
	public fonelineid	
	public fwinnercomment
	public fwinner
end class

class coneline_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem	
	public frectonelineid
	public frectisusing

	''//admin/momo/oneline/oneline_comment_list.asp
	public sub fonelinecomment_list()
		dim sqlStr,i , sqlsearch

		if frectonelineid <> "" then
			sqlsearch = sqlsearch & " and onelineid="&frectonelineid&"" + vbcrlf
		end if

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_oneline_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y'" & sqlsearch
		
		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,onelineid,gubun,userid,comment,regdate,isusing,coinyn" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_oneline_comment" + vbcrlf		
		sqlStr = sqlStr & " where isusing='Y'" & sqlsearch
		sqlStr = sqlStr & " order by idx desc" 
		
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
				set FItemList(i) = new coneline_oneitem
	
				FItemList(i).fidx = rsget("idx")				
				FItemList(i).fonelineid = rsget("onelineid")
				FItemList(i).fgubun = rsget("gubun")
				FItemList(i).fuserid = rsget("userid")						
				FItemList(i).fcomment = db2html(rsget("comment"))		
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")			
				FItemList(i).fcoinyn = rsget("coinyn")								
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	''/admin/momo/oneline/oneline_reg.asp
    public Sub foneline_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " onelineid,startdate,enddate,winnerdate,comment,regdate,isusing" + vbcrlf
		sqlStr = sqlStr & " ,winner,winnercomment" + vbcrlf	
		sqlStr = sqlStr & " from db_momo.dbo.tbl_oneline" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectonelineid <> "" then
			sqlStr = sqlStr & " and onelineid="&frectonelineid&"" + vbcrlf
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new coneline_oneitem
        
        if Not rsget.Eof then
				
				FOneItem.fwinner = db2html(rsget("winner"))
				FOneItem.fwinnercomment = db2html(rsget("winnercomment"))	
				FOneItem.fonelineid = rsget("onelineid")
				FOneItem.fstartdate = db2html(rsget("startdate"))				
				FOneItem.fenddate = db2html(rsget("enddate"))
				FOneItem.fwinnerdate = db2html(rsget("winnerdate"))
				FOneItem.fregdate = db2html(rsget("regdate"))
				FOneItem.fcomment = db2html(rsget("comment"))				
				FOneItem.fisusing = rsget("isusing")

        end if
        rsget.Close
    end Sub
        
	
	'/admin/momo/oneline/oneline.asp
	public sub foneline_list()
		dim sqlStr,i , sqlsearch

		if frectonelineid <> "" then
			sqlsearch = sqlsearch & " and onelineid="&frectonelineid&"" + vbcrlf
		end if
	
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing='"&frectisusing&"'" + vbcrlf
		end if

		'총 갯수 구하기
		sqlStr = "select count(onelineid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_oneline a" + vbcrlf
		sqlStr = sqlStr & " where onelineid <> 0 " & sqlsearch
		
		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " (case when startdate <= getdate() and enddate <= getdate() then 1 " + vbcrlf
		sqlStr = sqlStr & " when getdate() between startdate and enddate then 2 else 3 end) as stats " + vbcrlf	
		sqlStr = sqlStr & " ,onelineid,startdate,enddate,winnerdate,comment,regdate,isusing" + vbcrlf
		sqlStr = sqlStr & " ,(select count(onelineid) from db_momo.dbo.tbl_oneline_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y' and a.onelineid = onelineid) as commentcount " + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_oneline a" + vbcrlf
		sqlStr = sqlStr & " where onelineid <> 0 " & sqlsearch	
		sqlStr = sqlStr & " order by stats desc , onelineid desc"
		
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
				set FItemList(i) = new coneline_oneitem
				
				FItemList(i).fstats = rsget("stats")
				FItemList(i).fonelineid = rsget("onelineid")
				FItemList(i).fstartdate = db2html(rsget("startdate"))				
				FItemList(i).fenddate = db2html(rsget("enddate"))
				FItemList(i).fwinnerdate = db2html(rsget("winnerdate"))
				FItemList(i).fregdate = db2html(rsget("regdate"))
				FItemList(i).fcomment = db2html(rsget("comment"))				
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fcommentcount = rsget("commentcount")				
																
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

Class cwith_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fwithid
	public fstartdate
	public fenddate
	public fregdate
	public fisusing
	public fidx
	public fwithgubun
	public fuserid
	public fcomment
	public fwithimage_small
	public fwithimage_large
	public forderno
	public fstats
	public fcommentcount
end class
	

class cwith_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem	
	public frectwithid	
	public frectisusing
	public frectwithgubun
	public frectidx
	
	''/admin/momo/with/with_snsreg.asp
    public Sub fwith_snsoneitem()
        dim sqlStr
        
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " idx,withid,withgubun,userid,comment,regdate,isusing" + vbcrlf
		sqlStr = sqlStr & " ,withimage_small,withimage_large,orderno" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_with_comment" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectwithid <> "" then
			sqlStr = sqlStr & " and withid="&frectwithid&"" + vbcrlf
		end if
		if frectidx <> "" then
			sqlStr = sqlStr & " and idx="&frectidx&"" + vbcrlf
		end if
		
        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cwith_oneitem
        
        if Not rsget.Eof then
    			
			FOneItem.fidx = rsget("idx")
			FOneItem.fwithid = rsget("withid")
			FOneItem.fwithgubun = rsget("withgubun")
			FOneItem.fuserid = rsget("userid")
			FOneItem.fcomment = db2html(rsget("comment"))
			FOneItem.fregdate = db2html(rsget("regdate"))								           
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fwithimage_small = db2html(rsget("withimage_small"))
			FOneItem.fwithimage_large = db2html(rsget("withimage_large"))
			FOneItem.forderno = rsget("orderno")
																	
        end if
        rsget.Close
    end Sub

	''/admin/momo/with/with_sns.asp
	public sub fwith_snslist()
		dim sqlStr,i , sqlsearch

		if frectwithgubun <> "" then
			sqlsearch = sqlsearch & " and withgubun="&frectwithgubun&"" + vbcrlf
		end if

		if frectwithid <> "" then
			sqlsearch = sqlsearch & " and withid="&frectwithid&"" + vbcrlf
		end if
	
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing='"&frectisusing&"'" + vbcrlf
		end if
		
		'총 갯수 구하기
		sqlStr = "select count(withid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_with_comment a" + vbcrlf
		sqlStr = sqlStr & " where withid <> 0 " & sqlsearch

		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,withid,withgubun,userid,comment,regdate,isusing" + vbcrlf
		sqlStr = sqlStr & " ,withimage_small,withimage_large,orderno" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_with_comment a" + vbcrlf
		sqlStr = sqlStr & " where withid <> 0 " & sqlsearch	
		sqlStr = sqlStr & " order by orderno asc"
		
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
				set FItemList(i) = new cwith_oneitem
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fwithid = rsget("withid")
				FItemList(i).fwithgubun = rsget("withgubun")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fcomment = db2html(rsget("comment"))
				FItemList(i).fregdate = db2html(rsget("regdate"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).forderno = rsget("orderno")
				FItemList(i).fwithimage_small = db2html(rsget("withimage_small"))
				FItemList(i).fwithimage_large = db2html(rsget("withimage_large"))
																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	

	''/admin/momo/with/with_reg.asp
    public Sub fwith_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " withid,startdate,enddate,regdate,isusing" + vbcrlf	
		sqlStr = sqlStr & " from db_momo.dbo.tbl_with" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectwithid <> "" then
			sqlStr = sqlStr & " and withid="&frectwithid&"" + vbcrlf
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cwith_oneitem
        
        if Not rsget.Eof then
    			
			FOneItem.fwithid = rsget("withid")
			FOneItem.fstartdate = db2html(rsget("startdate"))
			FOneItem.fenddate = db2html(rsget("enddate"))			
			FOneItem.fregdate = db2html(rsget("regdate"))
			FOneItem.fisusing = rsget("isusing")
							           
        end if
        rsget.Close
    end Sub

	''/admin/momo/with/with.asp
	public sub fwith_list()
		dim sqlStr,i , sqlsearch

		if frectwithid <> "" then
			sqlsearch = sqlsearch & " and withid="&frectwithid&"" + vbcrlf
		end if
	
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing='"&frectisusing&"'" + vbcrlf
		end if

		'총 갯수 구하기
		sqlStr = "select count(withid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_with a" + vbcrlf
		sqlStr = sqlStr & " where withid <> 0 " & sqlsearch
		
		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " (case when startdate <= getdate() and enddate <= getdate() then 1 " + vbcrlf
		sqlStr = sqlStr & " when getdate() between startdate and enddate then 2 else 3 end) as stats " + vbcrlf	
		sqlStr = sqlStr & " ,withid,startdate,enddate,regdate,isusing" + vbcrlf
		sqlStr = sqlStr & " ,(select count(withid) from db_momo.dbo.tbl_with_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y' and a.withid = withid) as commentcount " + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_with a" + vbcrlf
		sqlStr = sqlStr & " where withid <> 0 " & sqlsearch	
		sqlStr = sqlStr & " order by stats desc , withid desc"
		
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
				set FItemList(i) = new cwith_oneitem
				
				FItemList(i).fstats = rsget("stats")
				FItemList(i).fwithid = rsget("withid")
				FItemList(i).fstartdate = db2html(rsget("startdate"))				
				FItemList(i).fenddate = db2html(rsget("enddate"))
				FItemList(i).fregdate = db2html(rsget("regdate"))				
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fcommentcount = rsget("commentcount")				
																
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

Class cforecast_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Fcardidx
	public Fstartdate
	public Fenddate
	public Fisusing
	public Fregdate
	public Fidx
	public Fforecastgubun
	public Fimage_url
	public Fcontents
	public Fyyyymmdd
	public Fuserid
	public Ftemperature
	public Fcoinyn
	public fcardcount
	public fusercount
	public fyyyymm
	public fgubun
	public forderno
	public fstats
	public flink_url
	public fcouponidx
end class

class cforecast_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public frectisusing
	public frectcardidx
	public frectidx
	public frectyyyymm
	public frectgubun

	'/admin/momo/forecast/card_list.asp '/frectgubun  0 감성예보
	public sub fuser_winner()
		dim sqlStr,i , sqlsearch
		
		if frectyyyymm <> "" then
			sqlsearch = sqlsearch & " and yyyymm = '"&frectyyyymm&"'" + vbcrlf
		end if
		if frectgubun <> "" then
			sqlsearch = sqlsearch & " and gubun = "&frectgubun&"" + vbcrlf
		end if
		
		'데이터 리스트 
		sqlStr = "select top 3"
		sqlStr = sqlStr & " idx ,yyyymm ,gubun ,orderno ,userid ,contents" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_winner" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by orderno asc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount

		redim preserve FItemList(ftotalcount)
	
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cforecast_oneitem
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fyyyymm = rsget("yyyymm")
				FItemList(i).fgubun = rsget("gubun")
				FItemList(i).forderno = rsget("orderno")	
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fcontents = db2html(rsget("contents"))
																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'/admin/momo/forecast/card_list.asp
	public sub fuser_list()
		dim sqlStr,i , sqlsearch
		
		if frectyyyymm <> "" then
			sqlsearch = sqlsearch & " and convert(varchar(7),yyyymmdd,121) = '"&frectyyyymm&"'" + vbcrlf
		end if
		
		'데이터 리스트 
		sqlStr = "select top 100"
		sqlStr = sqlStr & " userid , count(userid) as usercount" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_forecast_user" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " group by userid" + vbcrlf
		sqlStr = sqlStr & " order by count(userid) desc" + vbcrlf
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount

		redim preserve FItemList(ftotalcount)
	
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cforecast_oneitem
				
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fusercount = rsget("usercount")													
																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'/admin/momo/forecast/card_reg.asp
    public Sub fcard_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " cardidx ,startdate ,enddate ,isusing ,regdate" + vbcrlf	
		sqlStr = sqlStr & " from db_momo.dbo.tbl_forecast_card" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectcardidx <> "" then
			sqlStr = sqlStr & " and cardidx="&frectcardidx&"" + vbcrlf
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cforecast_oneitem
        
        if Not rsget.Eof then
    			
			FOneItem.fcardidx = rsget("cardidx")
			FOneItem.fstartdate = db2html(rsget("startdate"))
			FOneItem.fenddate = db2html(rsget("enddate"))
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fregdate = db2html(rsget("regdate"))				
							           
        end if
        rsget.Close
    end Sub

	'/admin/momo/forecast/card_reg.asp
    public Sub fcarddetail_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " idx ,cardidx ,forecastgubun ,image_url ,contents ,isusing, link_url, couponidx " + vbcrlf	
		sqlStr = sqlStr & " from db_momo.dbo.tbl_forecast_card_detail" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectidx <> "" then
			sqlStr = sqlStr & " and idx="&frectidx&"" + vbcrlf
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cforecast_oneitem
        
        if Not rsget.Eof then
    		
    		FOneItem.fidx = rsget("idx")	
			FOneItem.fcardidx = rsget("cardidx")
			FOneItem.fforecastgubun = rsget("forecastgubun")			
			FOneItem.fimage_url = db2html(rsget("image_url"))
			FOneItem.fcontents = db2html(rsget("contents"))
			FOneItem.fisusing = rsget("isusing")
			FOneItem.flink_url = rsget("link_url")
			FOneItem.fcouponidx = rsget("couponidx")
							           
        end if
        rsget.Close
    end Sub

	'/admin/momo/forecast/card_list.asp
	public sub fcard_detaillist()
		dim sqlStr,i , sqlsearch
		
		if frectcardidx <> "" then
			sqlsearch = sqlsearch & " and cardidx="&frectcardidx&""
		end if
		
		'총 갯수 구하기
		sqlStr = "select count(cardidx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_forecast_card_detail a" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx ,cardidx ,forecastgubun ,image_url ,contents ,isusing" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_forecast_card_detail a" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch	
		
		sqlStr = sqlStr & " order by cardidx desc" + vbcrlf
		
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
				set FItemList(i) = new cforecast_oneitem
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fcardidx = rsget("cardidx")
				FItemList(i).fforecastgubun = rsget("forecastgubun")
				FItemList(i).fimage_url = db2html(rsget("image_url"))
				FItemList(i).fcontents = db2html(rsget("contents"))
				FItemList(i).fisusing = rsget("isusing")																
																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'/admin/momo/forecast/card_list.asp
	public sub fcard_list()
		dim sqlStr,i , sqlsearch
		
		if frectcardidx <> "" then
			sqlsearch = sqlsearch & " and cardidx = "&frectcardidx&""
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing = '"&frectisusing&"'"
		end if		
		
		'총 갯수 구하기
		sqlStr = "select count(cardidx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_forecast_card a" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " (case when startdate <= getdate() and enddate <= getdate() then 1 " + vbcrlf
		sqlStr = sqlStr & " when getdate() between startdate and enddate then 2 else 3 end) as stats " + vbcrlf			
		sqlStr = sqlStr & " ,cardidx ,startdate ,enddate ,isusing ,regdate" + vbcrlf
		sqlStr = sqlStr & " ,(select count(sb.cardidx) from db_momo.dbo.tbl_forecast_card_detail sb" + vbcrlf
		sqlStr = sqlStr & " where a.cardidx = sb.cardidx and sb.isusing ='Y') as cardcount" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_forecast_card a" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		sqlStr = sqlStr & " order by cardidx desc" + vbcrlf
		
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
				set FItemList(i) = new cforecast_oneitem
				
				FItemList(i).fstats = rsget("stats")
				FItemList(i).fcardidx = rsget("cardidx")
				FItemList(i).fstartdate = db2html(rsget("startdate"))
				FItemList(i).fenddate = db2html(rsget("enddate"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = db2html(rsget("regdate"))
				FItemList(i).fcardcount = rsget("cardcount")																
																
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

class Cdiary_oneitem
	public Fidx
	public Fdiary_date
	public Ftitle
	public Fcontents
	public Fmainimage1
	public Fmainimage2
	public Fmainimage3
	public Fisusing
	public Fregdate
	public Fdiary_order
	public fdiarytype
	public fcommentcount
	public fdiaryidx
	public fuserid
	public fcomment
	public fcoinyn
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class Cdiary_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public frectisusing
	public frectidx
	public frectdiarytype
	public frectdiary_date
						
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

	''//admin/momo/diary/diary_comment_list.asp
	public sub fdiarycomment_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_diary_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectidx <> "" then
			sqlStr = sqlStr & " and diaryidx="&frectidx&"" + vbcrlf
		end if
						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,diaryidx,userid,comment,regdate,isusing,coinyn" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_diary_comment" + vbcrlf		
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectidx <> "" then
			sqlStr = sqlStr & " and diaryidx="&frectidx&"" + vbcrlf
		end if
		
		sqlStr = sqlStr & " order by idx desc" 
		

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
				set FItemList(i) = new cdiary_oneitem
	
				FItemList(i).fidx = rsget("idx")				
				FItemList(i).fdiaryidx = rsget("diaryidx")
				FItemList(i).fuserid = db2html(rsget("userid"))							
				FItemList(i).fcomment = db2html(rsget("comment"))		
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")			
				FItemList(i).fcoinyn = rsget("coinyn")								
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	''/admin/momo/diary/diary_list.asp
	public sub fdiary_contents_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_diary" + vbcrlf
		sqlStr = sqlStr & " where idx<>0 " + vbcrlf		
		
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing = 'Y'" + vbcrlf
		end if
		if frectdiarytype <> "" then
			sqlStr = sqlStr & " and diarytype = '"&frectdiarytype&"'" + vbcrlf	
		end if
		if frectdiary_date <> "" then
			sqlStr = sqlStr & " and convert(varchar(10),diary_date,121) = '"&frectdiary_date&"'" + vbcrlf			
		end if
		
		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,diary_date,title,contents,mainimage1,isusing" + vbcrlf
		sqlStr = sqlStr & " ,diary_order,mainimage2,mainimage3,diarytype" + vbcrlf
		sqlStr = sqlStr & " ,(select count(idx) from db_momo.dbo.tbl_diary_comment" + vbcrlf 
		sqlStr = sqlStr & " where isusing = 'Y' and a.idx = diaryidx) as commentcount" + vbcrlf	
		sqlStr = sqlStr & " from db_momo.dbo.tbl_diary a" + vbcrlf
		sqlStr = sqlStr & " where idx<>0 " + vbcrlf
		
		if frectisusing <> "" then
		sqlStr = sqlStr & " and isusing = 'Y'" + vbcrlf
		end if
		if frectdiarytype <> "" then
			sqlStr = sqlStr & " and diarytype = '"&frectdiarytype&"'" + vbcrlf	
		end if	
		if frectdiary_date <> "" then
			sqlStr = sqlStr & " and convert(varchar(10),diary_date,121) = '"&frectdiary_date&"'" + vbcrlf			
		end if
			
	
		sqlStr = sqlStr & " order by diary_order asc ,idx Desc"

		
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
				set FItemList(i) = new Cdiary_oneitem
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fdiary_date = rsget("diary_date")
				FItemList(i).ftitle = db2html(rsget("title"))
				FItemList(i).fcontents = db2html(rsget("contents"))
				FItemList(i).fmainimage1 = db2html(rsget("mainimage1"))
				FItemList(i).fmainimage2 = db2html(rsget("mainimage2"))
				FItemList(i).fmainimage3 = db2html(rsget("mainimage3"))
				FItemList(i).fisusing = rsget("isusing")				
				FItemList(i).fdiary_order = rsget("diary_order")
				FItemList(i).fdiarytype = rsget("diarytype")
				FItemList(i).fcommentcount = rsget("commentcount")
																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	''/admin/momo/diary/diary_reg.asp
    public Sub fdiarycontents_oneitem()
        dim sqlStr

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " idx,diary_date,title,contents,mainimage1,isusing" & vbcrlf
		sqlStr = sqlStr & " ,diary_order,mainimage2,mainimage3,diarytype" & vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_diary" & vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectidx <> "" then
			sqlStr = sqlStr & " and idx="&frectidx&"" + vbcrlf
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new Cdiary_oneitem
        
        if Not rsget.Eof then
    					
			FOneItem.fidx = rsget("idx")
			FOneItem.fdiary_date = rsget("diary_date")
			FOneItem.ftitle = db2html(rsget("title"))
			FOneItem.fcontents = db2html(rsget("contents"))
			FOneItem.fmainimage1 = db2html(rsget("mainimage1"))
			FOneItem.fmainimage2 = db2html(rsget("mainimage2"))
			FOneItem.fmainimage3 = db2html(rsget("mainimage3"))
			FOneItem.fisusing = rsget("isusing")			
			FOneItem.fdiary_order = rsget("diary_order")						
			FOneItem.fdiarytype = rsget("diarytype")
							           
        end if
        rsget.Close
    end Sub

end class

class CqnaItem

	public Fqnaid
	public FqstTitle
	public FqstContents
	public FansTitle
	public FansContents
	public FcommCd
	public FqstUserid
	public Fusername
	public FqstUserMail
	public FmailOk
	public Fisanswer
	public FlecIdx
	public FlecTitle
	public Fregdate
	public fbestviewcount
	public fisusing
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

end Class

Class Cqna

	public FqnaList()
	public FlecList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectqnaid
	public FRectuserid
	public FRectTopCnt
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString
	public FRectisAnswer
	public FRectlecIdx
	public FRectSearchLecturer

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FqnaList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'// QnA 내용 보기
	public Sub GetQnARead()
		dim SQL

		SQL =	" Select qnaid, qstTitle, qstContents, qstUserid, qstUsername, qstUserMail " &_
				"		,Case isanswer When 'Y' Then '<font color=darkred>완료</font>' When 'N' Then '<font color=darkblue>대기</font>' End isanswer " &_			
				"		, ansTitle, ansContents " &_
				"		,Case isanswer When 'Y' Then '완료' When 'N' Then '대기' End isanswer " &_
				"		, commCd, regdate " &_
				" From db_momo.dbo.tbl_QnA" &_
				" Where qnaid = " & FRectqnaid

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim FqnaList(0)

		if Not(rsget.EOF or rsget.BOF) then

			set FqnaList(0) = new CqnaItem

			FqnaList(0).Fqnaid			= rsget("qnaid")
			FqnaList(0).FqstTitle		= rsget("qstTitle")
			FqnaList(0).FqstContents	= rsget("qstContents")
			FqnaList(0).FansTitle		= rsget("ansTitle")
			FqnaList(0).FansContents	= rsget("ansContents")
			FqnaList(0).FcommCd			= rsget("commCd")			
			FqnaList(0).FqstUserid		= rsget("qstUserid")
			FqnaList(0).Fusername		= rsget("qstUsername")
			FqnaList(0).FqstUserMail	= rsget("qstUserMail")
			FqnaList(0).Fisanswer		= rsget("isanswer")
			FqnaList(0).Fregdate		= rsget("regdate")
			FqnaList(0).Fisanswer		= rsget("isanswer")

		end if
		rsget.close
	end sub
	
	'// QnA 분류별 목록 출력 '///admin/momo/qna/qna_list.asp
	public Sub GetQnAList()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and commCd='" & FRectsearchDiv & "' "
		end if

		if FRectisAnswer<>"" then
			AddSQL = AddSQL & " and isanswer='" & FRectisAnswer & "' "
		end if
		
		if FRectuserid <> "" then
			AddSQL = AddSQL & " and qstUserId='" & FRectuserid & "' "
		end if

		'@ 총데이터수
		SQL =	" Select count(qnaid) as cnt " &_
				" From db_momo.dbo.tbl_QnA as t1 " &_				
				" Where 1=1 " & AddSQL

		rsget.Open sql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		qnaid, qstUserId ,commCd , isusing" &_
				"		, isNull(qstTitle, Cast(qstContents as varchar(50))) as qstTitle " &_
				"		,Case isanswer When 'Y' Then '<font color=darkred>완료</font>' When 'N' Then '<font color=darkblue>대기</font>' End isanswer " &_
				"		, regdate ,qstContents" &_
				" From db_momo.dbo.tbl_QnA" &_
				" Where 1=1" & AddSQL &_
				" Order by qnaid desc "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FqnaList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    lp = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FqnaList(lp) = new CqnaItem

				FqnaList(lp).Fqnaid			= rsget("qnaid")
				FqnaList(lp).FqstTitle		= db2html(rsget("qstTitle"))				
				FqnaList(lp).fqstContents		= db2html(rsget("qstContents"))
				FqnaList(lp).FqstUserId		= rsget("qstUserId")
				FqnaList(lp).Fisanswer		= rsget("isanswer")
				FqnaList(lp).Fregdate		= rsget("regdate")
				FqnaList(lp).fcommCd		= rsget("commCd")				
				FqnaList(lp).fisusing		= rsget("isusing")
				
				lp=lp+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub	
	
	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// 첫페이지 산출
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class	

class CNoticeItem
	public FntcId
	public Ftitle
	public Fcontents
	public Fuserid
	public Fusername
	public FcommCd
	public Fregdate
	public fisusing
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CNotice
	public FNoticeList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount	
	public FRectNtcId
	public FRectTopCnt
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString
	
	Private Sub Class_Initialize()
		redim preserve FNoticeList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'//목록 출력 '//admin/momo/notice/notice_list.asp
	public Sub GetNoitceList()
		dim SQL, AddSQL, lp

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if
		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and commCd = "& FRectsearchDiv &""
		end if

		'@ 총데이터수
		SQL =	" Select count(ntcId) as cnt " &_
				" From db_momo.dbo.tbl_notice" &_
				" Where 1=1 " & AddSQL
		
		'response.write sql &"<br>"
		rsget.Open sql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				" ntcId, title, contents, regdate,commCd,isusing,userid" &_				
				" From db_momo.dbo.tbl_notice" &_
				" Where 1=1 " & AddSQL &_
				" Order by ntcId desc "
		
		'response.write sql &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FNoticeList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    lp = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FNoticeList(lp) = new CNoticeItem

				FNoticeList(lp).FntcId		= rsget("ntcId")
				FNoticeList(lp).fcommCd		= rsget("commCd")
				FNoticeList(lp).Ftitle		= db2html(rsget("title"))
				FNoticeList(lp).Fregdate	= rsget("regdate")
				FNoticeList(lp).fisusing	= rsget("isusing")
				FNoticeList(lp).fuserid	= rsget("userid")
				
				lp=lp+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

	'//내용 보기 '//admin/momo/notice/notice_modi.asp
	public Sub GetNoitceRead()
		dim SQL

		SQL =	" Select ntcId, title, contents, commCd,userid, regdate " &_
				" ,isusing" &_
				" From db_momo.dbo.tbl_notice" &_
				" Where ntcId = " & FRectNtcId

		rsget.Open sql, dbget, 1
		FTotalCount = rsget.recordcount
		
		redim FNoticeList(0)

		if Not(rsget.EOF or rsget.BOF) then

			set FNoticeList(0) = new CNoticeItem
			
			FNoticeList(0).FntcId		= rsget("ntcId")
			FNoticeList(0).Ftitle		= db2html(rsget("title"))
			FNoticeList(0).Fcontents	= db2html(rsget("contents"))
			FNoticeList(0).Fuserid		= rsget("userid")
			FNoticeList(0).FcommCd		= rsget("commCd")
			FNoticeList(0).Fregdate		= rsget("regdate")
			FNoticeList(0).fisusing		= rsget("isusing")
			
		end if
		rsget.close

	end sub

	public FPrevID
	public FNextID

	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// 첫페이지 산출
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

Class cvote_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
		
	public fvote_num
	public ftitle
	public fquestion
	public fstartdate
	public fenddate
	public fregdate
	public fisusing
	public fidx
	public fcontents_num
	public fcontents
	public fuserid
	public fcoinyn
	public fstats
	public fcontentscount
	public fvotecount
	public fmainimage
	
end class

class cvote_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

	public frectvote_num
	public frecttitle
	public frectisusing
									
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

	''/admin/momo/vote/vote_contents.asp
    public Sub fvote_contents()
		dim sqlStr,i
	
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,vote_num,contents_num,contents,isusing,regdate" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_vote_contents" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf		
		
		if frectvote_num <> "" then
			sqlStr = sqlStr & " and vote_num="&frectvote_num&"" + vbcrlf
		end if		
	
		sqlStr = sqlStr & " order by contents_num asc" + vbcrlf
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		ftotalcount = rsget.RecordCount
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
				set FItemList(i) = new cvote_oneitem
			
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fvote_num = rsget("vote_num")
				FItemList(i).fcontents_num = rsget("contents_num")
				FItemList(i).fcontents = db2html(rsget("contents"))		
				FItemList(i).fisusing = rsget("isusing")			
				FItemList(i).fregdate = rsget("regdate")
																				
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	

	''/admin/momo/vote/vote_reg.asp
    public Sub fvote_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " vote_num, title, question, startdate, enddate, isusing , mainimage" + vbcrlf	
		sqlStr = sqlStr & " from db_momo.dbo.tbl_vote" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectvote_num <> "" then
			sqlStr = sqlStr & " and vote_num="&frectvote_num&"" + vbcrlf
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cvote_oneitem
        
        if Not rsget.Eof then
    			
			FOneItem.fvote_num = rsget("vote_num")
			FOneItem.ftitle = db2html(rsget("title"))
			FOneItem.fquestion = db2html(rsget("question"))		
			FOneItem.fstartdate = rsget("startdate")
			FOneItem.fenddate = rsget("enddate")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fmainimage = rsget("mainimage")						
							           
        end if
        rsget.Close
    end Sub

	''/admin/momo/vote/vote_list.asp
	public sub fvote_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(vote_num) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_vote a" + vbcrlf
		sqlStr = sqlStr & " where vote_num <> 0" + vbcrlf		
		
		if frectvote_num <> "" then
			sqlStr = sqlStr & " and novelid="&frectvote_num&"" + vbcrlf
		end if
		if frecttitle <> "" then
			sqlStr = sqlStr & " and title like '%"&frecttitle&"%'" + vbcrlf
		end if			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if

		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " (case when startdate <= getdate() and enddate <= getdate() then 1 " + vbcrlf
		sqlStr = sqlStr & " when getdate() between startdate and enddate then 2 else 3 end) as stats " + vbcrlf
		sqlStr = sqlStr & " ,vote_num,title,question,startdate,enddate,regdate,isusing,mainimage" + vbcrlf
		sqlStr = sqlStr & " ,(select count(vote_num) from db_momo.dbo.tbl_vote_contents" + vbcrlf
		sqlStr = sqlStr & " where a.vote_num = vote_num and isusing ='Y') as contentscount" + vbcrlf
		sqlStr = sqlStr & " ,(select count(vote_num) from db_momo.dbo.tbl_vote_count" + vbcrlf
		sqlStr = sqlStr & " where a.vote_num = vote_num and isusing ='Y') as votecount" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_vote a" + vbcrlf
		sqlStr = sqlStr & " where vote_num <> 0" + vbcrlf			
		
		if frectvote_num <> "" then
			sqlStr = sqlStr & " and novelid="&frectvote_num&"" + vbcrlf
		end if
		if frecttitle <> "" then
			sqlStr = sqlStr & " and title like '%"&frecttitle&"%'" + vbcrlf
		end if			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if	
		
		sqlStr = sqlStr & " order by stats desc, vote_num desc" + vbcrlf
		
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
				set FItemList(i) = new cvote_oneitem
				
				FItemList(i).fstats = rsget("stats")
				FItemList(i).fmainimage = db2html(rsget("mainimage"))
				FItemList(i).fvote_num = rsget("vote_num")
				FItemList(i).ftitle = db2html(rsget("title"))	
				FItemList(i).fquestion = db2html(rsget("question"))			
				FItemList(i).fstartdate = rsget("startdate")				
				FItemList(i).fenddate = rsget("enddate")
				FItemList(i).fregdate = rsget("regdate")				
				FItemList(i).fisusing = rsget("isusing")		
				FItemList(i).fcontentscount = rsget("contentscount")
				FItemList(i).fvotecount = rsget("votecount")
																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	
end class

Class cbookmark_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
		
	public fideafileid
	public fbookmarkid
	public fgubun
	public fsitename
	public fsiteaddress
	public fsiteinfo
	public fuserid
	public fregdate
	public fisusing
	public fcoinyn
end class

class cbookmark_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public frectbookmarkid
	public frectisusing
	public frectcoinyn
						
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
	
	'///momo/bookmark/bookmark.asp
	public sub fbookmark_list()
		dim sqlStr,i
		
		'총 갯수 구하기
		sqlStr = "select count(bookmarkid) as cnt" + vbcrlf		
		sqlStr = sqlStr & " from db_momo.dbo.tbl_bookmark" + vbcrlf		
		sqlStr = sqlStr & " where 1=1"
		
		if frectbookmarkid <> "" then
			sqlStr = sqlStr & " and bookmarkid="&frectbookmarkid&"" + vbcrlf
		end if
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if	
			
		if frectcoinyn <> "" then
			sqlStr = sqlStr & " and coinyn='"&frectcoinyn&"'" + vbcrlf
		end if	
			
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
				
		'데이터 리스트 	
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " bookmarkid,gubun,sitename,siteaddress,siteinfo" + vbcrlf
		sqlStr = sqlStr & " ,userid,regdate,isusing,coinyn" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_bookmark" + vbcrlf
		sqlStr = sqlStr & " where 1=1"

		if frectbookmarkid <> "" then
			sqlStr = sqlStr & " and bookmarkid="&frectbookmarkid&"" + vbcrlf
		end if
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if	
		if frectcoinyn <> "" then
			sqlStr = sqlStr & " and coinyn='"&frectcoinyn&"'" + vbcrlf
		end if	
			
	
		sqlStr = sqlStr & " order by bookmarkid desc" + vbcrlf		
		
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
				set FItemList(i) = new cbookmark_oneitem

				FItemList(i).fbookmarkid = rsget("bookmarkid")
				FItemList(i).fgubun = rsget("gubun")
				FItemList(i).fsitename = db2html(rsget("sitename"))
				FItemList(i).fsiteaddress = db2html(rsget("siteaddress"))
				FItemList(i).fsiteinfo = db2html(rsget("siteinfo"))
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fcoinyn = rsget("coinyn")
																				
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

end class	

Class cideafile_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
		
	public fideafileid
	public fitemid
	public fcomment
	public fuserid
	public fbest
	public fregdate
	public fisusing
	public fcoinyn	
	public fImageList
	public fImageList120
	public fitemname
	public fcate_large
	public fbestyn
end class

class cideafile_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
		
	public frectideafileid
	public frectcate_large
	public frectisusing
	public frectbestyn
		
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

	public sub fideafile_list()
		dim sqlStr,i
		
		'총 갯수 구하기
		sqlStr = " select count(a.ideafileid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_ideafile a" + vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr & " on a.itemid = i.itemid" + vbcrlf		
		sqlStr = sqlStr & " where a.isusing = 'Y'" + vbcrlf		

		if frectcate_large <> "" and frectcate_large <> "max" then
			sqlStr = sqlStr & " and i.cate_large = '"&frectcate_large&"' " + vbcrlf		
		end if
			
		if frectideafileid <> "" then
			sqlStr = sqlStr & " and a.ideafileid="&frectideafileid&"" + vbcrlf
		end if
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and a.isusing='"&frectisusing&"'" + vbcrlf
		end if
		if frectbestyn <> "" then
			sqlStr = sqlStr & " and a.bestyn='"&frectbestyn&"'" + vbcrlf
		end if	
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
				
		'데이터 리스트 	
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " a.ideafileid,a.itemid,a.comment,a.userid,a.best,a.bestyn" + vbcrlf
		sqlStr = sqlStr & " ,a.regdate,a.isusing,a.coinyn , i.listimage,i.listimage120" + vbcrlf
		sqlStr = sqlStr & " ,i.itemname ,i.cate_large" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_ideafile a" + vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr & " on a.itemid = i.itemid" + vbcrlf
		sqlStr = sqlStr & " where a.isusing = 'Y'" + vbcrlf	
		
		if frectcate_large <> "" and frectcate_large <> "max" then
			sqlStr = sqlStr & " and i.cate_large = '"&frectcate_large&"' " + vbcrlf		
		end if
			
		if frectideafileid <> "" then
			sqlStr = sqlStr & " and a.ideafileid="&frectideafileid&"" + vbcrlf
		end if
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and a.isusing='"&frectisusing&"'" + vbcrlf
		end if	
		if frectbestyn <> "" then
			sqlStr = sqlStr & " and a.bestyn='"&frectbestyn&"'" + vbcrlf
		end if	
		
		if frectcate_large = "max" then					
			sqlStr = sqlStr & " order by a.best desc , a.ideafileid desc " + vbcrlf		
		else	
			sqlStr = sqlStr & " order by a.ideafileid desc " + vbcrlf		
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
				set FItemList(i) = new cideafile_oneitem
				
				FItemList(i).fbestyn = rsget("bestyn")
				FItemList(i).fcate_large = rsget("cate_large")
				FItemList(i).fideafileid = rsget("ideafileid")
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fcomment = db2html(rsget("comment"))
				FItemList(i).fitemname = db2html(rsget("itemname"))
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fbest = rsget("best")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fcoinyn = rsget("coinyn")	
				FItemList(i).fImageList	= "http://webimage.10x10.co.kr/image/list/" & GetImageSubFolderByItemid(FItemList(i).fitemid) & "/" &db2html(rsget("ListImage"))
				FItemList(i).fImageList120	= "http://webimage.10x10.co.kr/image/list120/" & GetImageSubFolderByItemid(FItemList(i).fitemid) & "/" &db2html(rsget("listimage120"))
																				
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

end class

Class cnovel_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fnovelid
	public fstartdate
	public fenddate
	public fregdate
	public fprolog
	public ftitle
	public fgenre
	public fisusing
	public fidx
	public fuserid
	public fcomment
	public fcoinyn
	public fcommentcount
	public fwordimage
	public fstats
	public fwinner
	public fcontents
end class

class cnovel_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem	
	public frectnovelid
	public frecttitle
	public frectisusing
	
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

	''/admin/momo/novel/novel_reg.asp
    public Sub fnovel_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " novelid,startdate,enddate,regdate,prolog,title,genre,isusing,wordimage , winner" + vbcrlf	
		sqlStr = sqlStr & " from db_momo.dbo.tbl_novel" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectnovelid <> "" then
			sqlStr = sqlStr & " and novelid="&frectnovelid&"" + vbcrlf
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cnovel_oneitem
        
        if Not rsget.Eof then
    			
			FOneItem.fnovelid = rsget("novelid")
			FOneItem.fstartdate = rsget("startdate")
			FOneItem.fenddate = rsget("enddate")			
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fprolog = db2html(rsget("prolog"))			
			FOneItem.ftitle = db2html(rsget("title"))							
			FOneItem.fgenre = rsget("genre")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fwinner = rsget("winner")	
			FOneItem.fwordimage = db2html(rsget("wordimage"))	
							           
        end if
        rsget.Close
    end Sub

	'/admin/momo/novel/proposal_list.asp
	public sub fproposal_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_proposal" + vbcrlf

		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx ,title ,contents ,userid ,regdate ,isusing" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_proposal" + vbcrlf		
		sqlStr = sqlStr & " order by idx desc"
		
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
				set FItemList(i) = new cnovel_oneitem
				
				FItemList(i).fidx = rsget("idx")			
				FItemList(i).ftitle = db2html(rsget("title"))		
				FItemList(i).fcontents = db2html(rsget("contents"))
				FItemList(i).fuserid = rsget("userid")		
				FItemList(i).fisusing = rsget("isusing")	
				FItemList(i).fregdate = rsget("regdate")					
																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	

	''/admin/momo/novel/novel_list.asp
	public sub fnovel_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(novelid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_novel a" + vbcrlf
		sqlStr = sqlStr & " where novelid <> 0" + vbcrlf	
		
		if frectnovelid <> "" then
			sqlStr = sqlStr & " and novelid="&frectnovelid&"" + vbcrlf
		end if
		if frecttitle <> "" then
			sqlStr = sqlStr & " and title like '%"&frecttitle&"%'" + vbcrlf
		end if			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if

		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " (case when startdate <= getdate() and enddate <= getdate() then 1 " + vbcrlf
		sqlStr = sqlStr & " when getdate() between startdate and enddate then 2 else 3 end) as stats " + vbcrlf	
		sqlStr = sqlStr & " ,novelid,startdate,enddate,regdate,prolog,title,genre,isusing,wordimage" + vbcrlf
		sqlStr = sqlStr & " ,(select count(novelid) from db_momo.dbo.tbl_novel_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y' and a.novelid = novelid) as commentcount " + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_novel a" + vbcrlf
		sqlStr = sqlStr & " where novelid <> 0" + vbcrlf				
		
		if frectnovelid <> "" then
			sqlStr = sqlStr & " and novelid="&frectnovelid&"" + vbcrlf
		end if
		if frecttitle <> "" then
			sqlStr = sqlStr & " and title like '%"&frecttitle&"%'" + vbcrlf
		end if			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if		
		
		sqlStr = sqlStr & " order by stats desc , novelid desc"
		
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
				set FItemList(i) = new cnovel_oneitem
				
				FItemList(i).fstats = rsget("stats")
				FItemList(i).fnovelid = rsget("novelid")
				FItemList(i).fstartdate = rsget("startdate")				
				FItemList(i).fenddate = rsget("enddate")
				FItemList(i).fregdate = rsget("regdate")				
				FItemList(i).fprolog = db2html(rsget("prolog"))		
				FItemList(i).ftitle = db2html(rsget("title"))											
				FItemList(i).fgenre = db2html(rsget("genre"))
				FItemList(i).fisusing = rsget("isusing")		
				FItemList(i).fcommentcount = rsget("commentcount")	
				FItemList(i).fwordimage = rsget("wordimage")
																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	

	''//admin/momo/novel/novel_comment_list.asp
	public sub fnovelcomment_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_novel_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectnovelid <> "" then
			sqlStr = sqlStr & " and novelid="&frectnovelid&"" + vbcrlf
		end if
						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx, novelid, userid, comment, regdate, isusing ,coinyn" + vbcrlf	
		sqlStr = sqlStr & " from db_momo.dbo.tbl_novel_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectnovelid <> "" then
			sqlStr = sqlStr & " and novelid="&frectnovelid&"" + vbcrlf
		end if
		
		sqlStr = sqlStr & " order by idx desc" + vbcrlf
		

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
				set FItemList(i) = new cnovel_oneitem
	
				FItemList(i).fidx = rsget("idx")				
				FItemList(i).fnovelid = rsget("novelid")
				FItemList(i).fuserid = db2html(rsget("userid"))							
				FItemList(i).fcomment = db2html(rsget("comment"))		
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

end class	

Class ctabloid_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public ftabloid
	public fyyyymmdd
	public ftitle
	public fuserid
	public ftempid
	public fbest
	public fisusing
	public fidx
	public fitemid
	public fitemorder
	public fcomment
	public FImageList
	public FImageicon1
	public FImageBasic
	public fitemname
	public fitemcount
	public fcoinyn	
end class

class ctabloid_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public frecttabloid
	public frectitemid
	public frecttitle
	public frectisusing
	
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
	
	''/admin/momo/tabloid/tabloid_list.asp
	public sub ftabloid_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(tabloid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_tabloid a" + vbcrlf
		sqlStr = sqlStr & " where tabloid<>0" + vbcrlf
		
		if frecttabloid <> "" then
			sqlStr = sqlStr & " and tabloid="&frecttabloid&"" + vbcrlf
		end if
		if frecttitle <> "" then
			sqlStr = sqlStr & " and title like '%"&frecttitle&"%'" + vbcrlf
		end if			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if		

		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " tabloid, yyyymmdd, title, userid, tempid, best, isusing, coinyn" + vbcrlf
		sqlStr = sqlStr & " ,(select count(*) from db_momo.dbo.tbl_tabloid_item" + vbcrlf 
		sqlStr = sqlStr & " where isusing = 'Y' and a.tabloid = tabloid ) as itemcount" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_tabloid a" + vbcrlf	
		sqlStr = sqlStr & " where tabloid<>0" + vbcrlf

		
		if frecttabloid <> "" then
			sqlStr = sqlStr & " and tabloid="&frecttabloid&"" + vbcrlf
		end if
		if frecttitle <> "" then
			sqlStr = sqlStr & " and title like '%"&frecttitle&"%'" + vbcrlf
		end if			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if		
		
		sqlStr = sqlStr & " order by tabloid desc" + vbcrlf
		
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
				set FItemList(i) = new ctabloid_oneitem
	
				FItemList(i).ftabloid = rsget("tabloid")				
				FItemList(i).fyyyymmdd = rsget("yyyymmdd")
				FItemList(i).ftitle = db2html(rsget("title"))								
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).ftempid = rsget("tempid")		
				FItemList(i).fbest = rsget("best")
				FItemList(i).fisusing = rsget("isusing")		
				FItemList(i).fcoinyn = rsget("coinyn")																							
				FItemList(i).fitemcount = rsget("itemcount")
												
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	
	
end Class	

Class cphoto_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fphotoid
	public fphotoword
	public fmainimage
	public fregdate
	public fisusing
	public fidx
	public fuserid
	public fcomment
	public fcommentcount
	public fdetailimage
	public ftag
	public fingimage
	public fwordimage
	public fwordovimage
end class
	
 
class cphoto_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public frectphotoid
	public frectphotoword
	public frectisusing

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

	''//감성포토상세 /admin/momo/photo/photo_reg.asp
    public Sub fphoto_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " photoid, photoword, mainimage, regdate, isusing, detailimage" + vbcrlf
		sqlStr = sqlStr & " ,ingimage, wordimage, wordovimage" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_photo a" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectphotoid <> "" then
		sqlStr = sqlStr & " and photoid= "&frectphotoid&"" + vbcrlf	
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cphoto_oneitem
        
        if Not rsget.Eof then
    			
			FOneItem.fphotoid = rsget("photoid")
			FOneItem.fphotoword = db2html(rsget("photoword"))							
			FOneItem.fmainimage = db2html(rsget("mainimage"))
			FOneItem.fdetailimage = db2html(rsget("detailimage"))
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fisusing = rsget("isusing")						
			FOneItem.fingimage = db2html(rsget("ingimage"))							
			FOneItem.fwordimage = db2html(rsget("wordimage"))
			FOneItem.fwordovimage = db2html(rsget("wordovimage"))		           
        end if
        rsget.Close
    end Sub

	''//감성포토리스트 /admin/momo/photo/photo_list.asp
	public sub fphoto_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(photoid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_photo" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		

		if frectphotoid <> "" then
			sqlStr = sqlStr & " and photoid = "&frectphotoid&"" + vbcrlf
		end if		
		if frectphotoword <> "" then
			sqlStr = sqlStr & " and photoword = '"&frectphotoword&"'" + vbcrlf
		end if	
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing = '"&frectisusing&"'" + vbcrlf
		end if			
					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " photoid,photoword,mainimage,regdate,isusing, detailimage" + vbcrlf
		sqlStr = sqlStr & " ,ingimage, wordimage, wordovimage" + vbcrlf		
		sqlStr = sqlStr & " ,(select count(idx) from db_momo.dbo.tbl_photo_contents" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y' and a.photoid = photoid) as commentcount" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_photo a" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
		
		if frectphotoid <> "" then
			sqlStr = sqlStr & " and photoid = "&frectphotoid&"" + vbcrlf
		end if		
		if frectphotoword <> "" then
			sqlStr = sqlStr & " and photoword = '"&frectphotoword&"'" + vbcrlf
		end if	
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing = '"&frectisusing&"'" + vbcrlf
		end if	
			
		sqlStr = sqlStr & " order by regdate desc" + vbcrlf
		

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
				set FItemList(i) = new cphoto_oneitem
			
				FItemList(i).fphotoid = rsget("photoid")
				FItemList(i).fphotoword = db2html(rsget("photoword"))							
				FItemList(i).fmainimage = db2html(rsget("mainimage"))
				FItemList(i).fdetailimage = db2html(rsget("detailimage"))				
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")			
				FItemList(i).fcommentcount = rsget("commentcount")													
				FItemList(i).fingimage = db2html(rsget("ingimage"))							
				FItemList(i).fwordimage = db2html(rsget("wordimage"))								
				FItemList(i).fwordovimage = db2html(rsget("wordovimage"))
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	''//감성포토코맨트리스트 /admin/momo/photo/photo_comment_list.asp
	public sub fphotocomment_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_photo_contents" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectphotoid <> "" then
			sqlStr = sqlStr & " and photoid="&frectphotoid&"" + vbcrlf
		end if
						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,photoid,userid,comment,regdate,isusing,mainimage,tag" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_photo_contents" + vbcrlf		
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectphotoid <> "" then
			sqlStr = sqlStr & " and photoid="&frectphotoid&"" + vbcrlf
		end if
		
		sqlStr = sqlStr & " order by idx desc" + vbcrlf
		

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
				set FItemList(i) = new cphoto_oneitem
	
				FItemList(i).fidx = rsget("idx")				
				FItemList(i).fphotoid = rsget("photoid")
				FItemList(i).fuserid = db2html(rsget("userid"))							
				FItemList(i).fcomment = db2html(rsget("comment"))		
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")			
				FItemList(i).fmainimage = db2html(rsget("mainimage"))	
				FItemList(i).ftag = db2html(rsget("tag"))													
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

end Class	

Class cyesno_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fyesnoid
	public fyesnoword
	public fmainimage
	public fdetailimage
	public fregdate
	public fisusing
	public fidx
	public fuserid
	public fyes
	public fno
	public fcommentcount
	public fingimage
	public fwordimage
	public fwordovimage
end class

class cyesno_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public frectyesnoid
	public frectyesnoword
	public frectisusing

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

	''//감성yesno코맨트리스트 /admin/momo/yesno/yesno_comment_list.asp
	public sub fyesnocomment_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_yesno_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectyesnoid <> "" then
			sqlStr = sqlStr & " and yesnoid="&frectyesnoid&"" + vbcrlf
		end if
						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,yesnoid,userid,yes,no,regdate,isusing" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_yesno_comment" + vbcrlf		
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectyesnoid <> "" then
			sqlStr = sqlStr & " and yesnoid="&frectyesnoid&"" + vbcrlf
		end if
		
		sqlStr = sqlStr & " order by idx desc" + vbcrlf
		

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
				set FItemList(i) = new cyesno_oneitem
	
				FItemList(i).fidx = rsget("idx")				
				FItemList(i).fyesnoid = rsget("yesnoid")
				FItemList(i).fuserid = db2html(rsget("userid"))							
				FItemList(i).fyes = db2html(rsget("yes"))		
				FItemList(i).fno = db2html(rsget("no"))
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")			
												
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	''//감성yesno리스트 /admin/momo/yesno/yesno_list.asp
    public Sub fyesno_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " yesnoid,yesnoword,mainimage,regdate,isusing, detailimage" + vbcrlf
		sqlStr = sqlStr & " ,ingimage, wordimage, wordovimage " + vbcrlf		
		sqlStr = sqlStr & " from db_momo.dbo.tbl_yesno a" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectyesnoid <> "" then
		sqlStr = sqlStr & " and yesnoid= "&frectyesnoid&"" + vbcrlf	
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cyesno_oneitem
        
        if Not rsget.Eof then
    			
			FOneItem.fyesnoid = rsget("yesnoid")
			FOneItem.fyesnoword = db2html(rsget("yesnoword"))							
			FOneItem.fmainimage = db2html(rsget("mainimage"))
			FOneItem.fdetailimage = db2html(rsget("detailimage"))
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fisusing = rsget("isusing")						
			FOneItem.fingimage = db2html(rsget("ingimage"))							
			FOneItem.fwordimage = db2html(rsget("wordimage"))				           
        	FOneItem.fwordovimage = db2html(rsget("wordovimage"))
        end if
        rsget.Close
    end Sub

	''//감성yesno리스트 /admin/momo/yesno/yesno_list.asp
	public sub fyesno_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(yesnoid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_yesno" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		

		if frectyesnoid <> "" then
			sqlStr = sqlStr & " and yesnoid = "&frectyesnoid&"" + vbcrlf
		end if		
		if frectyesnoword <> "" then
			sqlStr = sqlStr & " and yesnoword = '"&frectyesnoword&"'" + vbcrlf
		end if	
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing = '"&frectisusing&"'" + vbcrlf
		end if			
					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " yesnoid,yesnoword,mainimage,regdate,isusing, detailimage" + vbcrlf
		sqlStr = sqlStr & " ,ingimage, wordimage , wordovimage" + vbcrlf		
		sqlStr = sqlStr & " ,(select count(idx) from db_momo.dbo.tbl_yesno_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y' and a.yesnoid = yesnoid) as commentcount" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_yesno a" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
		
		if frectyesnoid <> "" then
			sqlStr = sqlStr & " and yesnoid = "&frectyesnoid&"" + vbcrlf
		end if		
		if frectyesnoword <> "" then
			sqlStr = sqlStr & " and yesnoword = '"&frectyesnoword&"'" + vbcrlf
		end if	
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing = '"&frectisusing&"'" + vbcrlf
		end if	
			
		sqlStr = sqlStr & " order by regdate desc" + vbcrlf
		

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
				set FItemList(i) = new cyesno_oneitem
			
				FItemList(i).fyesnoid = rsget("yesnoid")
				FItemList(i).fyesnoword = db2html(rsget("yesnoword"))							
				FItemList(i).fmainimage = db2html(rsget("mainimage"))
				FItemList(i).fdetailimage = db2html(rsget("detailimage"))				
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")			
				FItemList(i).fcommentcount = rsget("commentcount")													
				FItemList(i).fingimage = db2html(rsget("ingimage"))							
				FItemList(i).fwordimage = db2html(rsget("wordimage"))		
				FItemList(i).fwordovimage = db2html(rsget("wordovimage"))											
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
end Class	

Class cword_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fkeyid
	public fkeyword
	public fmainimage
	public fregdate
	public fisusing
	public fidx
	public fuserid
	public fcomment
	public fcommentcount
	public fdetailimage
	public fingimage
	public fwordimage	
	public fwordovimage		
	public fmainimage_small
	public ftag
	public fwinner
	public fisbest
	public fprizedate
end class
	

class cword_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public frectkeyid
	public frectgubun
	public frectkeyword
	public frectisusing

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

	''//감성사전상세 /admin/momo/word/word_reg.asp
    public Sub fword_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " keyid,keyword,mainimage,regdate,isusing, detailimage" + vbcrlf
		sqlStr = sqlStr & " ,ingimage, wordimage , wordovimage , winner, prizedate" + vbcrlf		
		sqlStr = sqlStr & " from db_momo.dbo.tbl_word a" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		
		
		if frectkeyid <> "" then
		sqlStr = sqlStr & " and keyid= "&frectkeyid&"" + vbcrlf	
		end if

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cword_oneitem
        
        if Not rsget.Eof then
    		
    		FOneItem.fwinner = db2html(rsget("winner"))
			FOneItem.fkeyid = rsget("keyid")
			FOneItem.fkeyword = db2html(rsget("keyword"))
			FOneItem.fmainimage = db2html(rsget("mainimage"))
			FOneItem.fdetailimage = db2html(rsget("detailimage"))
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fingimage = db2html(rsget("ingimage"))
			FOneItem.fwordimage = db2html(rsget("wordimage"))
			FOneItem.fwordovimage = db2html(rsget("wordovimage"))
			FOneItem.fprizedate = rsget("prizedate")
        end if
        rsget.Close
    end Sub

	''//감성사전리스트 /admin/momo/word/word_list.asp
	public sub fword_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(keyid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_word" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf		

		if frectkeyid <> "" then
			sqlStr = sqlStr & " and keyid = "&frectkeyid&"" + vbcrlf
		end if		
		if frectkeyword <> "" then
			sqlStr = sqlStr & " and keyword = '"&frectkeyword&"'" + vbcrlf
		end if	
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing = '"&frectisusing&"'" + vbcrlf
		end if			
					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " keyid,keyword,mainimage,regdate,isusing, detailimage" + vbcrlf
		sqlStr = sqlStr & " ,ingimage, wordimage , wordovimage" + vbcrlf		
		sqlStr = sqlStr & " ,(select count(idx) from db_momo.dbo.tbl_word_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y' and a.keyid = keyid) as commentcount" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_word a" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
		
		if frectkeyid <> "" then
			sqlStr = sqlStr & " and keyid = "&frectkeyid&"" + vbcrlf
		end if		
		if frectkeyword <> "" then
			sqlStr = sqlStr & " and keyword = '"&frectkeyword&"'" + vbcrlf
		end if	
			
		if frectisusing <> "" then
			sqlStr = sqlStr & " and isusing = '"&frectisusing&"'" + vbcrlf
		end if	
			
		sqlStr = sqlStr & " order by regdate desc" + vbcrlf
		

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
				set FItemList(i) = new cword_oneitem
			
				FItemList(i).fkeyid = rsget("keyid")
				FItemList(i).fkeyword = db2html(rsget("keyword"))							
				FItemList(i).fmainimage = db2html(rsget("mainimage"))
				FItemList(i).fdetailimage = db2html(rsget("detailimage"))				
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")			
				FItemList(i).fcommentcount = rsget("commentcount")													
				FItemList(i).fingimage = db2html(rsget("ingimage"))							
				FItemList(i).fwordimage = db2html(rsget("wordimage"))	
				FItemList(i).fwordovimage = db2html(rsget("wordovimage"))											
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	''//감성사전코맨트리스트 /admin/momo/word/word_comment_list.asp
	public sub fwordcomment_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_word_comment" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectkeyid <> "" then
			sqlStr = sqlStr & " and keyid = "&frectkeyid&"" + vbcrlf
		end if
		
		If frectgubun <> "" Then
			If frectgubun = "p" Then
				sqlStr = sqlStr & " and mainimage is Not Null and mainimage <> '' " + vbcrlf
			ElseIf frectgubun = "n" Then
				sqlStr = sqlStr & " and (mainimage is Null or mainimage = '') " + vbcrlf
			End If
		End IF
						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,keyid,userid,comment,regdate,isusing,mainimage,mainimage_small,tag, isNull(isbest,'') AS isbest" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_word_comment" + vbcrlf		
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf
		
		if frectkeyid <> "" then
			sqlStr = sqlStr & " and keyid = "&frectkeyid&"" + vbcrlf
		end if
		
		If frectgubun <> "" Then
			If frectgubun = "p" Then
				sqlStr = sqlStr & " and mainimage is Not Null and mainimage <> '' " + vbcrlf
			ElseIf frectgubun = "n" Then
				sqlStr = sqlStr & " and (mainimage is Null or mainimage = '') " + vbcrlf
			End If
		End IF
		
		sqlStr = sqlStr & " order by idx desc" + vbcrlf
		

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
				set FItemList(i) = new cword_oneitem
	
				FItemList(i).fidx = rsget("idx")				
				FItemList(i).fkeyid = rsget("keyid")
				FItemList(i).fuserid = db2html(rsget("userid"))							
				FItemList(i).fcomment = db2html(rsget("comment"))		
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")			
				FItemList(i).fmainimage = db2html(rsget("mainimage"))	
				FItemList(i).fmainimage_small = db2html(rsget("mainimage_small"))	
				FItemList(i).ftag = db2html(rsget("tag"))
				FItemList(i).fisbest = rsget("isbest")
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

end Class	

Class cmomo_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fposcode
	public fposname
	public fimagetype
	public fimagewidth
	public fimageheight
	public fisusing
	public fimagepath
	public flinkpath
	public fevt_code
	public fregdate
	public fimagecount
	public fimage_order
	
end class

class cmomo_list

	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

	public FRectPoscode
	public FRectIsusing
	public FRectvaliddate
	public FRectIdx
	public frecttoplimit
	
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
	
	'/admin/momo/imagemake_contents.asp
    public Sub fcontents_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " a.posname,a.imagetype,a.imagewidth,a.imageheight,a.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,b.idx,b.imagepath,b.linkpath,b.regdate,b.poscode,b.isusing,b.image_order" & vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_poscode a" & vbcrlf
		sqlStr = sqlStr & " left join db_momo.dbo.tbl_momo_poscode_image b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf
        sqlStr = sqlStr & " and b.idx = "& FRectIdx&""

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new cmomo_item
        
        if Not rsget.Eof then
    
			FOneItem.fposcode = rsget("poscode")
			FOneItem.fposname = db2html(rsget("posname"))
			FOneItem.fimagetype = rsget("imagetype")
			FOneItem.fimagewidth = rsget("imagewidth")
			FOneItem.fimageheight = rsget("imageheight")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fidx = rsget("idx")
			FOneItem.fimagepath = db2html(rsget("imagepath"))
			FOneItem.flinkpath = db2html(rsget("linkpath"))
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fimagecount = rsget("imagecount") 
			FOneItem.fimage_order = rsget("image_order") 
            
        end if
        rsget.Close
    end Sub

	'//admin/momo/imagemake_poscode.asp
	public sub fposcode_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " count(poscode) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_poscode" + vbcrlf
					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " poscode,isusing,posname,imagetype,imagewidth,imageheight,imagecount" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_poscode" + vbcrlf			
		sqlStr = sqlStr & " where 1=1" + vbcrlf

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
				set FItemList(i) = new cmomo_item
				
				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fimagecount = rsget("imagecount")
														
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/momo/imagemake_poscode.asp
    public Sub fposcode_oneitem()		
        dim SqlStr
        SqlStr = "select"
		sqlStr = sqlStr & " poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount" + vbcrlf        
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_poscode" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
        SqlStr = SqlStr + " and poscode=" + CStr(FRectPoscode)
         
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new cmomo_item
        if Not rsget.Eof then
            
            FOneItem.fposcode = rsget("poscode")
            FOneItem.fposname = db2html(rsget("posname"))
            FOneItem.fimagetype	= rsget("imagetype")
            FOneItem.fimagewidth = rsget("imagewidth")
            FOneItem.fimageheight = rsget("imageheight")
            FOneItem.fisusing = rsget("isusing")
            FOneItem.fimagecount = rsget("imagecount")
                       
        end if
        rsget.close
    end Sub
	
	
	'//admin/momo/imagemake_list.asp
	public sub fcontents_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_momo.dbo.tbl_momo_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"& FRectIsusing &"'" & vbcrlf		
			end if	

			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if	
				
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " b.posname,b.imagetype,b.imagewidth,b.imageheight,b.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,a.idx,a.imagepath,a.linkpath,a.regdate,a.poscode,a.isusing,a.image_order" & vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_momo.dbo.tbl_momo_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"&FRectIsusing&"'" & vbcrlf		
			end if	
			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if	

		sqlStr = sqlStr & " order by a.image_order asc" + vbcrlf

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
				set FItemList(i) = new cmomo_item
				
				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fimagepath = rsget("imagepath")
				FItemList(i).flinkpath = rsget("linkpath")
				FItemList(i).fregdate = rsget("regdate")		
				FItemList(i).fimagecount = rsget("imagecount")
				FItemList(i).fimage_order = rsget("image_order")													
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	 

end class

Class cPlay_Item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fstats
	public fplaySn
	public fstartdate
	public fenddate
	public fplayLinkType
	public flinkURL
	public fevtCode
	public fisusing
	public fregdate
	public fitemCount

	public fplyItemSn
	public fitemid
	public fmakerid
	public fitemname
	public fsellcash
	public fsellyn
	public FImageSmall

end class

class cPlayList
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FOneItem
	public frectisusing
	public frectPlaySn
	public frectyyyymm
	
	'/admin/momo/play/play_list.asp
	public sub fplay_list()
		dim sqlStr,i , sqlsearch
		
		if frectPlaySn <> "" then
			sqlsearch = sqlsearch & " and playSn = "&frectPlaySn&""
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing = '"&frectisusing&"'"
		end if		
		
		'총 갯수 구하기
		sqlStr = "select count(playSn) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_playInfo a" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " (case when playStartDate <= getdate() and playEndDate <= getdate() then 1 " + vbcrlf
		sqlStr = sqlStr & " when getdate() between playStartDate and playEndDate then 2 else 3 end) as stats " + vbcrlf			
		sqlStr = sqlStr & " ,playSn ,playStartDate ,playEndDate, playLinkType, linkURL, evt_code, isusing ,regdate" + vbcrlf
		sqlStr = sqlStr & " ,(select count(sb.plyItemSn) from db_momo.dbo.tbl_momo_playItem sb" + vbcrlf
		sqlStr = sqlStr & " where a.playSn = sb.playSn) as itemCount" + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_playInfo a" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		sqlStr = sqlStr & " order by playSn desc" + vbcrlf
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cPlay_Item
				
				FItemList(i).fstats = rsget("stats")
				FItemList(i).fplaySn = rsget("playSn")
				FItemList(i).fstartdate = db2html(rsget("playStartDate"))
				FItemList(i).fenddate = db2html(rsget("playEndDate"))
				FItemList(i).fplayLinkType = rsget("playLinkType")
				FItemList(i).flinkURL = rsget("linkURL")
				FItemList(i).fevtCode = rsget("evt_code")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = db2html(rsget("regdate"))
				FItemList(i).fitemCount = rsget("itemCount")																
																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub fitem_list()
		dim sqlStr,i , sqlsearch
		
		if frectPlaySn <> "" then
			sqlsearch = sqlsearch & " and playSn = "&frectPlaySn&""
		end if

		'총 갯수 구하기
		sqlStr = "select count(plyItemSn) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_playItem a" + vbcrlf
		sqlStr = sqlStr & "		Join db_item.dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr & "			on a.itemid=i.itemid" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " plyItemSn, playSn , a.itemid, i.makerid, i.itemname, i.sellcash, i.isusing, i.sellyn, i.smallimage " + vbcrlf
		sqlStr = sqlStr & " from db_momo.dbo.tbl_momo_playItem a" + vbcrlf
		sqlStr = sqlStr & "		Join db_item.dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr & "			on a.itemid=i.itemid" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		sqlStr = sqlStr & " order by playSn desc" + vbcrlf
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cPlay_Item

				FItemList(i).fplaySn = rsget("playSn")
				FItemList(i).fplyItemSn = rsget("plyItemSn")
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fmakerid = rsget("makerid")
				FItemList(i).fitemname = db2html(rsget("itemname"))
				FItemList(i).fsellcash = rsget("sellcash")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fsellyn = rsget("sellyn")
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")

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

function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select poscode,posname from db_momo.dbo.tbl_momo_poscode"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("poscode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("poscode")&"' "&tmp_str&">" + db2html(rsget("posname")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

'카테고리 & 추천 명 가져오기
function Drawcate(selectBoxName,selectedId)
dim tmp_str,sql
	%>
	<select name="<%=selectBoxName%>" class='input_02' style='width:130px;height:20px;' onchange='change_id(this.value);'>
	<option value=''>전체보기</option>
	<option value='max' <% if selectedId="max" then response.write " selected"%>>추천많은 순서보기</option>
	<%
	sql = " select code_large, code_nm"
	sql = sql & " from db_item.dbo.tbl_Cate_large"
	sql = sql & " where display_yn = 'Y' and code_large<>999"
		
	'response.write sql &"<br>"
	rsget.Open sql,dbget,1
	
	if  not rsget.EOF  then
	   do until rsget.EOF
	       if Lcase(selectedId) = Lcase(rsget("code_large")) then
	           tmp_str = " selected"
	       end if
	       response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">" + db2html(rsget("code_nm")) + "</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   loop
	end if
	rsget.close
	response.write("</select>")
	
end function

'//admin/momo/bookmark/bookmark_list.asp
function imggubun(tmp)
	if tmp = 0 then
		imggubun = "<img src='http://fiximage.10x10.co.kr/web2009/momo/images/bookmark_ico01.gif' width=50 height=50>"
	elseif tmp = 1 then
		imggubun = "<img src='http://fiximage.10x10.co.kr/web2009/momo/images/bookmark_ico02.gif' width=50 height=50>"
	elseif tmp = 2 then
		imggubun = "<img src='http://fiximage.10x10.co.kr/web2009/momo/images/bookmark_ico03.gif' width=50 height=50>"
	elseif tmp = 3 then
		imggubun = "<img src='http://fiximage.10x10.co.kr/web2009/momo/images/bookmark_ico04.gif' width=50 height=50>"
	end if														
end function

'//상태값 반환
function statsgubun(stats)
	if stats = 1 then
		statsgubun = "마감"
	elseif stats = 2 then
		statsgubun = "진행중"
	elseif stats = 3 then
		statsgubun = "시작전"		
	end if	
end function

'// 이메일을 보낸다. //
Sub Send_mail(FromMail,ToMail,strTitle,MainCont)
	Dim iMsg
	Dim iConf
	Dim Flds
	Dim strHTML

	set iMsg	= CreateObject("CDO.Message")
	set iConf	= CreateObject("CDO.Configuration")

	Set Flds	= iConf.Fields
    
    if (ToMail<>"") and (FromMail<>"") then
		With iMsg
			Set .Configuration = iConf
			.To			= ToMail
			.From		= FromMail
			.Subject	= strTitle
			.HTMLBody	= MainCont
			.Send
		End With
	end if

	Set iMsg	= Nothing
	Set iConf	= Nothing
	Set Flds	= Nothing
End Sub

'// 로컬 디스크의 파일을 읽어 변수에 저장 //
Function ReadLocalFile(file_name, path_name)
	dim vPath, Filecont
	dim fso, file

	vPath = Server.MapPath (path_name) & "\"	'로컬 디렉토리를 얻는다.

	Set fso = Server.CreateObject("Scripting.FileSystemObject")

		Set file = fso.OpenTextFile(vPath & file_name)

			Filecont = file.ReadAll

		file.close

		Set file = Nothing

	Set fso = Nothing

	ReadLocalFile = Filecont
End Function

'/감성예보 카드 구분
function drawforecastgubun(boxname , stats , flg)
%>
	<select name="<%=boxname%>" <%=flg%>>
		<option value="" <% if stats = "" then response.write " selected"%>>전체</option>
		<option value="0" <% if stats = "0" then response.write " selected"%>>sunny</option>
		<option value="1" <% if stats = "1" then response.write " selected"%>>cloudy</option>
		<option value="2" <% if stats = "2" then response.write " selected"%>>thunder</option>
		<option value="3" <% if stats = "3" then response.write " selected"%>>rainy</option>
	</select>
<%
end function

'/감성예보 카드 구분
function getforecastgubun(tmp)
	if tmp = "0" then
		getforecastgubun = "sunny"
	elseif tmp = "1" then
		getforecastgubun = "cloudy"
	elseif tmp = "2" then
		getforecastgubun = "thunder"
	elseif tmp = "3" then
		getforecastgubun = "rainy"
	else
		getforecastgubun = tmp
	end if		
end function

'/공지사항 구분
function drawnotics_gubun(boxname , stats , flg)
%>
	<select name="<%=boxname%>" <%=flg%>>
		<option value="" <% if stats = "" then response.write " selected"%>>전체</option>
		<option value="1" <% if stats = "1" then response.write " selected"%>>공지사항</option>
		<option value="2" <% if stats = "2" then response.write " selected"%>>FAQ</option>
		<option value="3" <% if stats = "3" then response.write " selected"%>>트위터</option>
		<option value="4" <% if stats = "4" then response.write " selected"%>>미투데이</option>
		<option value="5" <% if stats = "5" then response.write " selected"%>>오프캐스트</option>	
	</select>
<%
end function

'/공지사항 구분
function getnotics_gubun(tmp)
	if tmp = "1" then
		getnotics_gubun = "공지사항"
	elseif tmp = "2" then
		getnotics_gubun = "FAQ"
	elseif tmp = "3" then
		getnotics_gubun = "트위터"
	elseif tmp = "4" then
		getnotics_gubun = "미투데이"
	elseif tmp = "5" then
		getnotics_gubun = "오프캐스트"		
	else
		getnotics_gubun = tmp
	end if		
end function

Function getWeekSerial(dt)
	dim startWeek, totalWeek
	totalWeek = DatePart("ww", dt)	'전체 주차
	startWeek = DatePart("ww", DateSerial(year(dt),month(dt),"01"))		'첫째일 주차

	'계산 및 값 반환
	getWeekSerial = totalWeek - startWeek + 1	
end Function

'/공지사항 구분
function getwith_gubun(tmp)
	if tmp = "0" then
		getwith_gubun = "트위터"
	elseif tmp = "1" then
		getwith_gubun = "미투데이"
	elseif tmp = "2" then
		getwith_gubun = "복합"	
	else
		getwith_gubun = tmp
	end if		
end function
%>