<%
Class cvideo_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public ftitle
	public fcatecd2
	public flecturer
	public fmakerid
	public fkeyword
	public fimage_url
	public fimage2_url
	public fyoutube_url
	public fyoutube_source
	public fisusing
	public fregdate
	
end class

class cvideo
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public foneitem
	
	public frectidx
	public frectisusing
	public frectcate2
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub


	public sub fvideo_list()
		dim sqlStr, addStr, i

		if frectisusing = "Y" then 
			addStr = addStr & " and isusing = 'Y'" + vbcrlf 
		elseif 	frectisusing = "N" then 
			addStr = addStr & " and isusing = 'N'" + vbcrlf 		 
		end if
		
		If frectcate2 <> "" Then
			addStr = addStr & " and cateCD2 = '" & frectcate2 & "'" + vbcrlf
		End If

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_video " + vbcrlf 
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 
		sqlStr = sqlStr & " 	" & addStr & " " + vbcrlf 
			
		rsacademyget.Open sqlStr,dbacademyget,1
			FTotalCount = rsacademyget("cnt")
		rsacademyget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx, title, cateCD2, lecturer, makerid, keyword, image_url, image2_url, " + vbcrlf
		sqlStr = sqlStr & " youtube_url, youtube_source, isusing, regdate " + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_video " + vbcrlf
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 
		sqlStr = sqlStr & " 	" & addStr & " " + vbcrlf 
		sqlStr = sqlStr & " order by idx desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsacademyget.pagesize = FPageSize
		rsacademyget.Open sqlStr,dbacademyget,1

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
		if  not rsacademyget.EOF  then
			rsacademyget.absolutepage = FCurrPage
			do until rsacademyget.EOF
				set FItemList(i) = new cvideo_oneitem

				FItemList(i).fidx				= rsacademyget("idx")
				FItemList(i).ftitle				= db2html(rsacademyget("title"))
				FItemList(i).fcatecd2			= rsacademyget("cateCD2")
				FItemList(i).flecturer			= rsacademyget("lecturer")
				FItemList(i).fmakerid			= rsacademyget("makerid")
				FItemList(i).fkeyword			= db2html(rsacademyget("keyword"))
				FItemList(i).fimage_url			= rsacademyget("image_url")
				FItemList(i).fimage2_url		= rsacademyget("image2_url")
				FItemList(i).fyoutube_url		= rsacademyget("youtube_url")
				FItemList(i).fyoutube_source	= db2html(rsacademyget("youtube_source"))
				FItemList(i).fisusing			= rsacademyget("isusing")
				FItemList(i).fregdate			= rsacademyget("regdate")
		
				rsacademyget.movenext
				i=i+1
			loop
		end if
		rsacademyget.Close
	end sub
	
	
    public Sub video_view()
        dim sqlStr , i
        
		sqlStr = "select " + vbcrlf
		sqlStr = sqlStr & " idx, title, cateCD2, lecturer, makerid, keyword, image_url, image2_url, " + vbcrlf
		sqlStr = sqlStr & " youtube_url, youtube_source, isusing, regdate " + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_video " + vbcrlf
		sqlStr = sqlStr & " where idx = '" & frectidx & "'" + vbcrlf 		 

		
        'response.write sqlStr&"<br>"
        rsacademyget.Open sqlStr,dbacademyget,1
        ftotalcount = rsacademyget.RecordCount
        
        set FOneItem = new cvideo_oneitem
        
        if Not rsacademyget.Eof then

			foneitem.fidx				= rsacademyget("idx")
			foneitem.ftitle				= db2html(rsacademyget("title"))
			foneitem.fcatecd2			= rsacademyget("cateCD2")
			foneitem.flecturer			= rsacademyget("lecturer")
			foneitem.fmakerid			= rsacademyget("makerid")
			foneitem.fkeyword			= db2html(rsacademyget("keyword"))
			foneitem.fimage_url			= rsacademyget("image_url")
			foneitem.fimage2_url		= rsacademyget("image2_url")
			foneitem.fyoutube_url		= rsacademyget("youtube_url")
			foneitem.fyoutube_source	= db2html(rsacademyget("youtube_source"))
			foneitem.fisusing			= rsacademyget("isusing")
			foneitem.fregdate			= rsacademyget("regdate")
			   						   
        end if
        rsacademyget.Close
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