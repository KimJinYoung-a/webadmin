<%
'###########################################################
' Description :  텐바이텐 traffic analysis  클래스
' History : 2007.09.04 한용민 생성
'###########################################################

Class CtrafficOne
	
	public fyyyymmdd			'날짜
	public ftotalcount			'방문자수
	public fpageview			'페이지뷰
	public fnewcount			'신규방문자수
	public frecount				'재방문자수
	public frealcount			'실제방문자수
	
   Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class
'##################################################################
class Ctrafficlist							'텐바이텐 트래픽내역  디비에서 가져오기

	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public frectbuy_date
	public frectbuy_date1		
	
	public sub Ftrafficlist
	dim sql , i
	
	sql = "select yyyymmdd,totalcount,pageview,newcount,recount,realcount"
	sql = sql & " from db_datamart.dbo.tbl_traffic_analysis"
	sql = sql & " where 1=1 and yyyymmdd between '"&frectbuy_date&"' and '"&frectbuy_date1&"'"
	sql = sql & " order by yyyymmdd desc"
	
	'response.write sql			'오류시 뿌려본다.
	db3_rsget.open sql,db3_dbget,1
	
	FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0
	
	if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
		do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
			set flist(i) = new CtrafficOne 			'클래스를 넣고
		
			flist(i).fyyyymmdd = db3_rsget("yyyymmdd")				'날짜
			flist(i).ftotalcount = db3_rsget("totalcount")			'방문자수
			flist(i).fpageview = db3_rsget("pageview")				'페이지뷰
			flist(i).fnewcount = db3_rsget("newcount")				'신규방문자수
			flist(i).frecount = db3_rsget("recount")				'재방문자수
			flist(i).frealcount = db3_rsget("realcount")			'실제방문자수
		
			db3_rsget.movenext
			i = i+1
			loop		
		end if
	db3_rsget.close
	end sub
		
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class
'##################################################################
class Ctrafficgraph							'텐바이텐 트래픽내역  그래프용

	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public frectyyyy					'년을 가져오기 위한 변수
	public frectmm						'달을 가져오기 위한 변수
	public function frecttot()			'년수과 달을 폼에서 받아와서 합침...
	frecttot = frectyyyy & frectmm
	end function
	public function frecttotnew()
	
		if frectyyyy <>"" and frectmm <> "" then												'날짜값이 있다면
			frecttotnew = " and left(yyyymmdd,6) = "& frecttot &""		'위에 쿼리에 검색 옵션을 붙인다
		else 
			frecttotnew = " and left(yyyymmdd,6) = "& 0 &""				'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		end if	
	end function
	
	public sub Ftrafficlist
	dim sql , i
	
	sql = "select yyyymmdd,totalcount,pageview,newcount,recount,realcount"
	sql = sql & " from db_datamart.dbo.tbl_traffic_analysis"
	sql = sql & " where 1=1 "& frecttotnew &""
	sql = sql & " order by yyyymmdd asc"
	
	'response.write sql			'오류시 뿌려본다.
	db3_rsget.open sql,db3_dbget,1
	
	FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0
	
	if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
		do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
			set flist(i) = new CtrafficOne 			'클래스를 넣고
		
			flist(i).fyyyymmdd = db3_rsget("yyyymmdd")				'날짜
			flist(i).ftotalcount = db3_rsget("totalcount")			'방문자수
			flist(i).fpageview = db3_rsget("pageview")				'페이지뷰
			flist(i).fnewcount = db3_rsget("newcount")				'신규방문자수
			flist(i).frecount = db3_rsget("recount")				'재방문자수
			flist(i).frealcount = db3_rsget("realcount")			'실제방문자수
		
			db3_rsget.movenext
			i = i+1
			loop		
		end if
	db3_rsget.close
	end sub
		
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

%>
