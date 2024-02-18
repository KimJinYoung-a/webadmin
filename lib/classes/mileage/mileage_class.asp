<%
'###########################################################
' Description :  마일리지 구분 
' History : 2007.10.23 한용민 생성
'###########################################################

Class Cmileageoneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fjukyocd
	public fjukyoname
	public fisusing
	
end class

class Cmileagelist
	Private Sub Class_Initialize()
		redim flist(0)
		FCurrPage = 1
		FPageSize = 0
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0	
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public frectjukyocd
	public frectisusing
	public frectseachjukyocd
	
'##############################################################################################		
	public sub fmileagelist									
	dim sqlcount, cnt, i
	sqlcount = "select count(jukyocd) as cnt"			'검색 선택값에 해당하는 인덱스수를 가져온다
	sqlcount = sqlcount & " from db_user.dbo.tbl_mileage_gubun"
	sqlcount = sqlcount & " where 1=1 "	

	if frectisusing <> "" then 
		sqlcount = sqlcount & "and isusing = '" & frectisusing & "'"
	end if			
	if frectseachjukyocd <> "" then 
		sqlcount = sqlcount & "and jukyocd = '" & frectseachjukyocd & "'"
	end if	
		
	rsget.open sqlcount,dbget,1
	'response.write sqlcount&"<br>"
	FTotalCount = rsget("cnt")				'총레코드 수에 인덱스카운트를 넣고
	rsget.close	
	
	dim sql 
	sql = "select top "& FPageSize*FCurrpage &" jukyocd,jukyoname,isusing"
	sql = sql & " from db_user.dbo.tbl_mileage_gubun"
	sql = sql & " where 1=1 "
	
	if frectisusing <> "" then 
		sql = sql & "and isusing = '" & frectisusing & "'"
	end if			
	if frectseachjukyocd <> "" then 
		sql = sql & "and jukyocd = '" & frectseachjukyocd & "'"
	end if	

	sql = sql & " order by jukyocd desc"
	
	rsget.pagesize = FPageSize
	rsget.open sql,dbget,1
	'response.write sql&"<br>"
	
	FResultCount = rsget.RecordCount - (FPageSize*(FCurrPage-1))
	FTotalPage = CInt(FTotalCount\FPageSize) + 1	
	
	redim flist(FResultCount)
	i = 0 
	
	if not rsget.eof then
		rsget.absolutepage = FCurrPage
		do until rsget.eof
			set flist(i) = new Cmileageoneitem
			
			flist(i).fjukyocd = rsget("jukyocd")
			flist(i).fjukyoname = rsget("jukyoname")
			flist(i).fisusing = rsget("isusing")
			
		rsget.movenext
		i = i + 1
		loop		
	end if
	rsget.close
	end sub
'##############################################################################################	
	public sub fmileage_add
	
	dim sql 
	sql = "select jukyocd,jukyoname,isusing"
	sql = sql & " from db_user.dbo.tbl_mileage_gubun"
	sql = sql & " where 1=1 and jukyocd = '" & frectjukyocd & "'"
	
	rsget.open sql,dbget,1
	
	FTotalCount = rsget.RecordCount
	redim flist(FTotalCount)
	i = 0 
	
	if not rsget.eof then
		do until rsget.eof
			set flist(i) = new Cmileageoneitem
			
			flist(i).fjukyocd = rsget("jukyocd")
			flist(i).fjukyoname = rsget("jukyoname")
			flist(i).fisusing = rsget("isusing")
			
		rsget.movenext
		i = i + 1
		loop		
	end if
	rsget.close
	end sub
'##############################################################################################		
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1								'//시작 페이지가 1보다 크면 생성
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1	'//전체 페이지가 시작페이지+전체페이지링크수-1의 수보다 크면 생성
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1	'//시작 페이지는 현재페이지에서 1을 빼고 전체페이지링크수로 나눈후 전체페이지링크수를 곱한후 +1을 하면 생김. 
	end Function	
end class
%>