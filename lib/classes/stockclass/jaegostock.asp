<%

'###########################################################
' Description :  재고파악
' History : 2007.07.13 한용민 개발
' History : 2007.11.28 한용민 수정
'###########################################################

class Cfitem					'클래스 선언
	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx				'인덱스번호
	public fitemgubun		'상품구분
	public fitemid			'상품번호
	public fitemoption		'옵션코드	
	public fitemname		'상품명
	public fitemoptionname	'옵션명
	public fmakerid			'브랜드id
	public fregdate			'등록일
	public freguserid		'지시자id	
	public factionusername	'재고파악한사람
	public factionstartdate	'재고파악일시
	public fbasicstock		'재고파악재고
	public frealstock		'재고파악 실사갯수
	public ferrstock		'오차
	public ffinishuserid	'완료자id
	public fstatecd			'상태코드
	public deleteyn			'삭제여부
	public makerid			'검색에필요한브랜드id
	public fstats			'상태
	public fsmallimage		'상품이미지
	public fbigo			'비고
	public foptioncnt		'옵션비교시 필요한 변수
	public frealstocks		'realstock + offconfirmno + ipkumdiv5
	
	public function getbigoName()
		if fbigo = 1 then
		 	getbigoName = "작업지시"
		elseif fbigo = 5 then
			getbigoName = "재고파악완료"
		elseif fbigo = 7 then
			getbigoName = "완료(반영됨)"
		elseif fbigo = 8 then
			getbigoName = "완료(미반영)"
		else
			getbigoName =""
		end if
	end function
end class

class Cfitemlist					'클래스 선언
	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public Frectidx					'인덱스 값을 받기 위한 변수
	public Frectitemid				'상품id 값을 받기 위한 변수
	public frectmakerid				'브랜드 값을 받기 위한 변수
	public frectstats				'상태 값을 받기 위한 변수 
	public frectorderingdate		'작업지시일을 받기 위한 변수
	public frectguestlist			'일반사용자용 리스트를 뿌리기 위한 변수
	public frectitemoption
	
	public Sub fjaegoinsert()			'재고입력용
		dim sqlStr ,i 
		
		sqlStr = "select" 
		sqlstr = sqlstr & " isnull(b.realstock,'0') as realstock,"
		sqlstr = sqlstr & " a.itemid, isnull(c.itemoption,'0000') as itemoption," 
		sqlstr = sqlstr & " a.itemgubun,isnull(a.itemgubun,'10') as itemgubun,"
		sqlstr = sqlstr & " a.smallimage,a.makerid ,a.itemname"
		sqlstr = sqlstr & " from db_item.[dbo].tbl_item a"
		sqlstr = sqlstr & " left join [db_summary].dbo.tbl_current_logisstock_summary b" 
		sqlstr = sqlstr & " on a.itemid = b.itemid" 
		sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item_option c" 
		sqlstr = sqlstr & " on b.itemid = c.itemid and b.itemoption = c.itemoption" 
		sqlstr = sqlstr & " where 1=1 and a.itemid = '" & frectitemid &"'"
			if frectitemoption <> "0000" then 		
				sqlstr = sqlstr & " and c.itemoption = '" & frectitemoption &"'"
			end if
		rsget.Open sqlStr,dbget,1					'재고파악지시테이블(f)와 재고테이블(r)를 조인해서 상품 리스트를 가져온다.
		'response.write sqlstr&"<br>"
				   	
	   	FTotalCount = rsget.recordcount
	   	redim flist(FTotalCount)
		i=0
			
		if  not rsget.EOF  then
			do until rsget.eof
				set flist(i) = new Cfitem
				
						
				flist(i).fitemgubun = rsget("itemgubun")				'상품구분
				flist(i).fitemid = rsget("itemid")						'상품번호
				flist(i).fitemoption = rsget("itemoption")				'옵션코드	
				flist(i).fitemname = rsget("itemname")					'상품명
				flist(i).fmakerid = rsget("makerid")					'브랜드id
				flist(i).frealstock = rsget("realstock")				'재고파악 실사갯수
				flist(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")				'상품이미지
				
				rsget.moveNext		   
				i=i+1
				
			loop
		end if
		rsget.close
	end Sub	
	
	public Sub fprintlist()			' 프린트 출력 부분
		dim sqlStr ,i

		sqlStr = "select" 
		sqlstr = sqlstr & " a.idx, b.smallimage, a.itemid, b.makerid ,a.basicstock," 
		sqlstr = sqlstr & " b.itemname, a.itemoption, c.optionname, a.statecd, a.actiondate," 
		sqlstr = sqlstr & " isnull(d.realstock,'0') as realstock" 
		sqlstr = sqlstr & " from [db_summary].[dbo].tbl_req_realstock a" 
		sqlstr = sqlstr & " join db_item.[dbo].tbl_item b" 
		sqlstr = sqlstr & " on a.itemid = b.itemid" 
		sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item_option c" 
		sqlstr = sqlstr & " on a.itemid = c.itemid and a.itemoption = c.itemoption" 
		sqlstr = sqlstr & " left join db_summary.dbo.tbl_current_logisstock_summary d" 
		sqlstr = sqlstr & " on a.itemid = d.itemid and a.itemoption = d.itemoption" 			 		
		sqlstr = sqlstr & " Where a.idx in (" + Frectidx + ")"
		sqlstr = sqlstr & " order by idx desc"					
		rsget.Open sqlStr,dbget,1				'재고파악지시테이블(f)와 재고테이블(r)에서 상품id와 상품옵션이 값은 것을 가져온다.
		'response.write sqlstr&"<br>"
					   	
	   	FTotalCount = rsget.recordcount
	   	redim flist(FTotalCount)
		i=0
			
		if  not rsget.EOF  then
			do until rsget.eof
				set flist(i) = new Cfitem								'클래스를 넣고
				
				flist(i).fidx = rsget("idx")		
				flist(i).fitemid = rsget("itemid")						'상품번호
				flist(i).fitemoption = rsget("itemoption")				'옵션코드	
				flist(i).fitemname = rsget("itemname")					'상품명
				flist(i).fitemoptionname = rsget("optionname")		'옵션명
				flist(i).fmakerid = rsget("makerid")					'브랜드id
				flist(i).factionstartdate = rsget("actiondate")			'재고파악일시
				flist(i).fbasicstock = rsget("basicstock")				'재고파악재고
				flist(i).frealstock = rsget("realstock")				'재고파악 실사갯수
				flist(i).fstatecd = rsget("statecd")					'상태코드
				flist(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")				'상품이미지
				
				rsget.moveNext		   
				i=i+1
				
			loop
		end if
		rsget.close
	end Sub	
	
	public Sub fwritelist()				'상품 옵션별로 검색
		dim sql554 ,i 
			
			sql554 = "select o.optionname,isnull(o.itemoption,'0000') as itemoption,r.optioncnt from [db_item].[dbo].tbl_item r" 
			sql554 = sql554 & " left join [db_item].[dbo].tbl_item_option o on r.itemid = o.itemid" 
			sql554 = sql554 & " where r.itemid = '"& Frectitemid &"'"
			rsget.open sql554,dbget,1			'상품테이블(r)과 상품옵션테이블(o)를 비교해서 상품이 같은것으르 가져온다.
			'response.write sql554&"<br>"
			
			FTotalCount = rsget.recordcount
		   	redim flist(FTotalCount)
		   	i = 0
		   	
		   	if not rsget.EOF  then
				do until rsget.eof
					set flist(i) = new Cfitem							'클래스넣고
					flist(i).fitemoptionname = rsget("optionname")		'상품옵션이름넣고	
					flist(i).fitemoption = rsget("itemoption")			'상품옵션코드넣고
					flist(i).foptioncnt = rsget("optioncnt")			'한 상품에 총 옵션이 몇개인지 수 넣고
					rsget.moveNext		   
					i=i+1
				loop
			end if
			rsget.close
	end sub
	
	public Sub fjonglist()						'상품 리스트 뿌려주는 부분
		dim sql , i , sqlcount , cnt
		
		sqlcount = "select count(idx) as cnt from [db_summary].[dbo].tbl_req_realstock"		'검색 선택값에 해당하는 인덱스수를 가져온다
		sqlcount = sqlcount & " where 1=1"
		'response.write sqlcount
		
		if Frectguestlist <> "" then
			sqlcount = sqlcount & " and statecd in (" & Frectguestlist & ")"
		end if
		
		if frectstats <> "" then												'위에서 request 값에 상태 값이 있다면
			sqlcount = sqlcount & " and statecd = " & frectstats & ""			'위에 쿼리에 검색 옵션을 붙인다
		end if
		
		rsget.open sqlcount,dbget,1
		FTotalCount = rsget("cnt")				'총레코드 수에 인덱스카운트를 넣고
		rsget.close
		
		sql = "select top "& FPageSize*FCurrpage &""
		sql = sql & " (isnull(d.realstock,'0')+isnull(d.offconfirmno,'0')+isnull(d.ipkumdiv5,'0')) as realstocks,"		
		sql = sql & " b.smallimage,b.itemname,b.makerid,b.smallimage,"
		sql = sql & " c.optionname , a.*"
		sql = sql & " from [db_summary].[dbo].tbl_req_realstock a"
		sql = sql & " join db_item.[dbo].tbl_item b"
		sql = sql & " on a.itemid = b.itemid"
		sql = sql & " left join [db_item].[dbo].tbl_item_option c" 
		sql = sql & " on a.itemid = c.itemid and a.itemoption = c.itemoption"
		sql = sql & " left join db_summary.dbo.tbl_current_logisstock_summary d"
		sql = sql & " on a.itemid = d.itemid and a.itemoption = d.itemoption"				
		sql = sql & " where 1=1"
	
		if Frectguestlist <> "" then
			sql = sql & " and statecd in (" & Frectguestlist & ")"
		end if	
		
		if frectmakerid <> "" then									'위에서 request 값에 메이커 값이 있다면
			sql = sql & " and makerid = '" & frectmakerid & "'"		'위에 쿼리에 검색 옵션을 붙인다
		end if 
		
		if frectstats <> "" then								'위에서 request 값에 상태 값이 있다면
			sql = sql & " and statecd = " & frectstats & ""		'위에 쿼리에 검색 옵션을 붙인다
		end if

		sql = sql & " order by idx desc"
		'response.write sql			'삑살시 뿌려본다
		rsget.pagesize = FPageSize
		rsget.open sql,dbget,1
		
		FResultCount = rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1	
		
		redim flist(FResultCount)
		i=0			'루프돌 i 값에 o넣고
			
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set flist(i) = new Cfitem		'클래스넣고
				
				flist(i).fidx = rsget("idx")		
				flist(i).fitemgubun = rsget("itemgubun")				'상품구분
				flist(i).fitemid = rsget("itemid")						'상품번호
				flist(i).fitemoption = rsget("itemoption")				'옵션코드	
				flist(i).fitemname = rsget("itemname")					'상품명
				flist(i).fitemoptionname = rsget("optionname")		'옵션명
				flist(i).fmakerid = rsget("makerid")					'브랜드id
				flist(i).fregdate = rsget("regdate")					'등록일
				flist(i).freguserid = rsget("reguserid")				'지시자id	
				flist(i).factionstartdate = rsget("actiondate")			'재고파악일시
				flist(i).fbasicstock = rsget("basicstock")				'재고파악재고
				flist(i).frealstock = rsget("realstock")				'재고파악 실사갯수
				flist(i).ferrstock = rsget("errstock")					'오차
				flist(i).ffinishuserid = rsget("finishuserid")			'완료자id
				flist(i).fstatecd = rsget("statecd")					'상태코드
				flist(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")				'상품이미지
				flist(i).fbigo = rsget("statecd")						'삭제나 입력등등을 할 상태값
				flist(i).frealstocks = rsget("realstocks")						'삭제나 입력등등을 할 상태값				
				rsget.moveNext		   
				i=i+1
				
			loop
		end if
		rsget.close
	end Sub
	
	
	public Sub fbrandinsert()			'브랜드 재고 검색 저장을 위한 클래스
		dim sqlstr ,i 

		sqlstr = "select" 
		sqlstr = sqlstr & " isnull(b.realstock,'0') as realstock,"
		sqlstr = sqlstr & " a.itemid , a.makerid , a.smallimage,a.itemname"
		sqlstr = sqlstr & " ,isnull(c.itemoption,'0000') as itemoption,"
		sqlstr = sqlstr & " c.optionname"
		sqlstr = sqlstr & " from db_item.[dbo].tbl_item a"
		sqlstr = sqlstr & " join [db_summary].dbo.tbl_current_logisstock_summary b" 
		sqlstr = sqlstr & " on a.itemid = b.itemid" 
		sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item_option c" 
		sqlstr = sqlstr & " on b.itemid = c.itemid and b.itemoption = c.itemoption" 
		sqlstr = sqlstr & " where 1=1"
		
		if frectmakerid <> "" then
			sqlstr = sqlstr & " and a.makerid= '" & frectmakerid &"'"
		end if 
		if Frectitemid <> "" then
			sqlstr = sqlstr & " and a.itemid= '" & Frectitemid &"'"
		end if 
	
		sqlstr = sqlstr & " order by a.itemid desc"
		
		rsget.Open sqlstr,dbget,1					'
		'response.write sqlstr&"<br>"	   	
	   	FTotalCount = rsget.recordcount
		redim flist(FTotalCount)
		i=0			'루프돌 i 값에 o넣고
			
		if not rsget.EOF then
			do until rsget.eof
				set flist(i) = new Cfitem		'클래스넣고
				
				flist(i).fitemid = rsget("itemid")		
				flist(i).fmakerid = rsget("makerid")				
				flist(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")			
				flist(i).fitemname = rsget("itemname")					
				flist(i).fitemoption = rsget("itemoption")		
				flist(i).fitemoptionname = rsget("optionname")					
				flist(i).fbasicstock = rsget("realstock")					
						
				rsget.moveNext		   
				i=i+1
				
			loop
		end if
		rsget.close
	end Sub
			
	Private Sub Class_Initialize()
		redim flist(0)
		FCurrPage = 1
		FPageSize = 11
		FResultCount = 0
		FScrollCount = 11
		FTotalCount =0
	end sub

	Private Sub Class_Terminate()

	End Sub

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

'// 상품 이미지 경로를 계산하여 반환 //
function GetImageSubFolderByItemid(byval iitemid)
    if (iitemid <> "") then
	    GetImageSubFolderByItemid = Num2Str(CStr(Clng(iitemid) \ 10000),2,"0","R")
	else
	    GetImageSubFolderByItemid = ""
	end if
end function
%>