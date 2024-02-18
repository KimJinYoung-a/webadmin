 <%
'###########################################################
' Description : 인스타그램 이벤트용 수동 등록 클래스
' Hieditor : 2016.06.23 유태욱 생성
'###########################################################
class CinstagrameventItem
	public Fidx
	public Fgubun
	public FIsusing
	public Fcontentsidx
	public Fevt_code
	public Fimgurl
	public Fuserid
	public Flinkurl
	public FRegdate		
end class

class CInstagramevent
	public FItemList()
	public Foneitem
	public FPageSize
	public FCurrPage
	public FTotalPage
	public FPageCount
	public FTotalCount
	public FScrollCount
	public FrectIsusing
	public FResultCount
	public Frectcontentsidx
	Public Feventid
	
	public Sub fnGetinstagramevent_oneitem()
	    dim sqlStr, sqlsearch
	
	if Frectcontentsidx <> "" Then
		sqlsearch = sqlsearch & " AND idx ='"& Frectcontentsidx &"'"
	end if

	    sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " idx,evt_code,imgurl,userid,linkurl,isusing,regdate"
		sqlStr = sqlStr & " from [db_temp].[dbo].[tbl_event_instagram]"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"
	
	    'response.write sqlStr&"<br>"
	    rsget.Open SqlStr, dbget, 1
	    FResultCount = rsget.RecordCount
	    
	    set FOneItem = new CinstagrameventItem
	    
	    if Not rsget.Eof then
	    	
		Foneitem.Fcontentsidx = rsget("idx")
		Foneitem.Fevt_code = rsget("evt_code")
		Foneitem.Fimgurl = rsget("imgurl")
		Foneitem.Fuserid = rsget("userid")
		Foneitem.Flinkurl = rsget("linkurl")
		Foneitem.FIsusing = rsget("isusing")
		Foneitem.FRegdate = rsget("regdate")
	    end if
	    rsget.Close
	end Sub

	public sub fnGetInstagrameventList
		dim sqlStr,i, sqlsearch

		if FrectIsusing <> "" Then
			sqlsearch = sqlsearch & " AND isusing ='"& FrectIsusing &"'"
		end If
		
		if Feventid <> "" Then
			sqlsearch = sqlsearch & " AND evt_code ='"& Feventid &"'"
		end If

		'글의 총 갯수 구하기
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from [db_temp].[dbo].[tbl_event_instagram]"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'DB 데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx,evt_code,imgurl,userid,linkurl,isusing,regdate"
		sqlStr = sqlStr & " from [db_temp].[dbo].[tbl_event_instagram]"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"
		
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
				set FItemList(i) = new CinstagrameventItem

				FItemList(i).Fcontentsidx = rsget("idx")
				FItemList(i).Fevt_code = rsget("evt_code")
				FItemList(i).Fimgurl = rsget("imgurl")
				FItemList(i).Fuserid = rsget("userid")
				FItemList(i).Flinkurl = rsget("linkurl")
				FItemList(i).FIsusing = rsget("isusing")
				FItemList(i).FRegdate = rsget("regdate")

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






	

		