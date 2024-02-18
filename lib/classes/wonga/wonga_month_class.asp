<%
'###########################################################
' Description :  월간원가보고서용 클래스
' History : 2007.09.10 한용민 생성
'###########################################################
Class Cwongaitem
	public groupname				'그룹명
	public yyyymm						'날짜
	public fcategory
	public fcategory_isusing
	public fcategory_name
	public ffield
	public ffield_name
	public ffield_value	
	public fgijun_value
		
	public gubun0_isusing		'첫번째 카테코리 사용여부
	public gubun1_isusing		'두번째 카테고리 사용여부
	public gubun2_isusing
	public gubun3_isusing
	public gubun4_isusing
	public gubun5_isusing
	public gubun0_name			'첫번째 카테고리명
	public gubun1_name			'두번째 카테고리명
	public gubun2_name
	public gubun3_name
	public gubun4_name
	public gubun5_name
	public gubunsum					'카테고리필드값의 합산값
	public chulgocount			'계산을위한 총출고수나 계산수
	
	public category0_yyyy_sum
	public category1_yyyy_sum
	public category2_yyyy_sum
	public category3_yyyy_sum
	public category4_yyyy_sum
	public category5_yyyy_sum
	public category_yyyy_sum
	
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

'##################################################################
class Cwongalist							

	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public frectgubun
	public frectyyyy					'년을 가져오기 위한 변수
	public frectmm						'달을 가져오기 위한 변수
	
	public function frectyyyymm()		'년도와 달 합치기
		frectyyyymm = frectyyyy&frectmm
	end function
	
	public function frectyyyy_re()		'이전 년도를 검색하기 위한 변수
		frectyyyy_re = frectyyyy - 1
	end function
	public function frectmm_re()			'이전달 을 검색하기 위한 변수
		if frectmm = "12" then
			frectmm_re = "11"
		elseif frectmm = "11" then
			frectmm_re = "10"
		elseif frectmm = "10" then
			frectmm_re = "09"
		elseif frectmm = "09" then
			frectmm_re = "08"
		elseif frectmm = "08" then
			frectmm_re = "07"
		elseif frectmm = "07" then
			frectmm_re = "06"
		elseif frectmm = "06" then
			frectmm_re = "05"
		elseif frectmm = "05" then
			frectmm_re = "04"
		elseif frectmm = "04" then
			frectmm_re = "03"
		elseif frectmm = "03" then
			frectmm_re = "02"
		elseif frectmm = "02" then
			frectmm_re = "01"
		elseif frectmm = "01" then
			frectmm_re = "12"
		end if
	end function
	public function frectyyyymm_re()
		if frectmm = "01" then		'달이 1월일경우 그 전달이 그전년도 12월 임으로 그전년도를 붙인다.
			frectyyyymm_re = frectyyyy_re&frectmm_re
		else
			frectyyyymm_re = frectyyyy&frectmm_re
		end if
	end function

'##################################################################
public sub fwongamonth						'월별 비용 내역 클래스
dim sql , i

sql = "select"
sql = sql & " a.groupname,a.category,a.category_name,a.category_isusing"
sql = sql & " ,a.field,a.field_name,a.gijun_value"
sql = sql & " ,isnull(b.yyyymm,'') as yyyymm,isnull(b.field_value,'') as field_value,isnull(b.count,'') as count"
sql = sql & " from [db_datamart].[dbo].tbl_month_wonga_category a"
sql = sql & " left join [db_datamart].[dbo].tbl_month_wonga b"
sql = sql & " on a.groupname = b.groupname and a.category = b.category and a.field = b.field"
sql = sql & " where 1=1 and a.groupname= '"& frectgubun &"'" 

	if frectyyyy <> "" then
		sql = sql & " and b.yyyymm = '"& frectyyyymm &"'"
	else
		sql = sql & " and b.yyyymm = '0'"	  
	end if 

sql = sql & " order by b.yyyymm asc"
db3_rsget.open sql,db3_dbget,1
'response.write sql&"<br>"			'삑살시 뿌려본다.

FTotalCount = db3_rsget.recordcount
redim flist(FTotalCount)
i = 0
		
	if not db3_rsget.eof then				'레코드의 첫번째가 아니라면
		do until db3_rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
			set flist(i) = new Cwongaitem 			'클래스를 넣고
			
				flist(i).groupname = db3_rsget("groupname")
				flist(i).yyyymm = db3_rsget("yyyymm")
				flist(i).fcategory = db3_rsget("category")
				flist(i).fcategory_isusing = db3_rsget("category_isusing")
				flist(i).fcategory_name = db3_rsget("category_name")
				flist(i).chulgocount = clng(db3_rsget("count"))				
				flist(i).ffield = db3_rsget("field")
				flist(i).ffield_name = db3_rsget("field_name")
				flist(i).ffield_value = clng(db3_rsget("field_value"))
				flist(i).fgijun_value = db3_rsget("gijun_value")
			db3_rsget.movenext
			i = i+1
		loop		
	end if
db3_rsget.close			
end sub

'##################################################################
public sub fwongamonth_re						'검색 년도에서 이전달 값을 가져온다
dim sql , i

sql = "select"
sql = sql & " a.groupname,a.category,a.category_name,a.category_isusing"
sql = sql & " ,a.field,a.field_name,a.gijun_value"
sql = sql & " ,isnull(b.yyyymm,'') as yyyymm,isnull(b.field_value,'') as field_value,isnull(b.count,'') as count"
sql = sql & " from db_datamart.dbo.tbl_month_wonga_category a"
sql = sql & " left join db_datamart.dbo.tbl_month_wonga b"
sql = sql & " on a.groupname = b.groupname and a.category = b.category and a.field = b.field"
sql = sql & " where 1=1 and a.groupname= '"& frectgubun &"'" 

	if frectmm <> "" then
		sql = sql & " and b.yyyymm = '"& frectyyyymm_re &"'"
	else
		sql = sql & " and b.yyyymm = '0'"		  
	end if 

sql = sql & " order by b.yyyymm asc"
db3_rsget.open sql,db3_dbget,1
'response.write sql&"<br>"			'삑살시 뿌려본다.

FTotalCount = db3_rsget.recordcount
redim flist(FTotalCount)
i = 0
		
	if not db3_rsget.eof then				'레코드의 첫번째가 아니라면
		do until db3_rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
			set flist(i) = new Cwongaitem 			'클래스를 넣고
			
				flist(i).groupname = db3_rsget("groupname")
				flist(i).yyyymm = db3_rsget("yyyymm")
				flist(i).fcategory = db3_rsget("category")
				flist(i).fcategory_isusing = db3_rsget("category_isusing")
				flist(i).fcategory_name = db3_rsget("category_name")
				flist(i).chulgocount = db3_rsget("count")				
				flist(i).ffield = db3_rsget("field")
				flist(i).ffield_name = db3_rsget("field_name")
				flist(i).ffield_value = db3_rsget("field_value")

			db3_rsget.movenext
			i = i+1
		loop		
	end if
db3_rsget.close			
end sub
'##################################################################
public sub fwongalist						'년 총비용 클래스

dim sql , i

sql = "select"
sql = sql & " sum(case when b.category = '0' then field_value end) as category0,"
sql = sql & " sum(case when b.category = '1' then field_value end) as category1,"
sql = sql & " sum(case when b.category = '2' then field_value end) as category2,"
sql = sql & " sum(case when b.category = '3' then field_value end) as category3,"
sql = sql & " sum(case when b.category = '4' then field_value end) as category4,"
sql = sql & " sum(case when b.category = '5' then field_value end) as category5,"
sql = sql & " sum(field_value) as categorysum,"
sql = sql & " b.yyyymm,a.groupname,"
sql = sql & " (select category_isusing from db_datamart.dbo.tbl_month_wonga_category where category='0' and groupname= '"& frectgubun &"' group by category_isusing)"
sql = sql & " as category0_isusing,"
sql = sql & " (select category_isusing from db_datamart.dbo.tbl_month_wonga_category where category='1' and groupname= '"& frectgubun &"' group by category_isusing)"
sql = sql & " as category1_isusing,"
sql = sql & " (select category_isusing from db_datamart.dbo.tbl_month_wonga_category where category='2' and groupname= '"& frectgubun &"' group by category_isusing)"
sql = sql & " as category2_isusing,"
sql = sql & " (select category_isusing from db_datamart.dbo.tbl_month_wonga_category where category='3' and groupname='"& frectgubun &"' group by category_isusing)"
sql = sql & " as category3_isusing,"
sql = sql & " (select category_isusing from db_datamart.dbo.tbl_month_wonga_category where category='4' and groupname= '"& frectgubun &"' group by category_isusing)"
sql = sql & " as category4_isusing,"
sql = sql & " (select category_isusing from db_datamart.dbo.tbl_month_wonga_category where category='5' and groupname= '"& frectgubun &"' group by category_isusing)"
sql = sql & " as category5_isusing,"
sql = sql & " (select category_name from db_datamart.dbo.tbl_month_wonga_category where category='0' and groupname= '"& frectgubun &"' group by category_name)"
sql = sql & " as category0_name,"
sql = sql & " (select category_name from db_datamart.dbo.tbl_month_wonga_category where category='1' and groupname= '"& frectgubun &"' group by category_name)"
sql = sql & " as category1_name,"
sql = sql & " (select category_name from db_datamart.dbo.tbl_month_wonga_category where category='2' and groupname= '"& frectgubun &"' group by category_name)"
sql = sql & " as category2_name,"
sql = sql & " (select category_name from db_datamart.dbo.tbl_month_wonga_category where category='3' and groupname= '"& frectgubun &"' group by category_name)"
sql = sql & " as category3_name,"
sql = sql & " (select category_name from db_datamart.dbo.tbl_month_wonga_category where category='4' and groupname= '"& frectgubun &"' group by category_name)"
sql = sql & " as category4_name,"
sql = sql & " (select category_name from db_datamart.dbo.tbl_month_wonga_category where category='5' and groupname= '"& frectgubun &"' group by category_name)"
sql = sql & " as category5_name"
sql = sql & " from db_datamart.dbo.tbl_month_wonga_category a"
sql = sql & " left join db_datamart.dbo.tbl_month_wonga b"
sql = sql & " on a.groupname = b.groupname and a.category = b.category and a.field = b.field"
sql = sql & " where 1=1 and a.groupname= '"& frectgubun &"'" 
		if frectyyyy <> "" then
		sql = sql & " and left(b.yyyymm,4) = '"& frectyyyy &"'"
	else
		sql = sql & " and b.yyyymm = '0'"		  
	end if 
sql = sql & " group by b.yyyymm ,a.groupname"

db3_rsget.open sql,db3_dbget,1
'response.write sql&"<br>"		'삑살시 뿌려본다.

FTotalCount = db3_rsget.recordcount
redim flist(FTotalCount)
i = 0
		
	if not db3_rsget.eof then				'레코드의 첫번째가 아니라면
		do until db3_rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
			set flist(i) = new Cwongaitem 			'클래스를 넣고
			
				flist(i).groupname = db3_rsget("groupname")
				flist(i).yyyymm = db3_rsget("yyyymm")
				flist(i).category0_yyyy_sum = db3_rsget("category0")
				flist(i).category1_yyyy_sum = db3_rsget("category1")
				flist(i).category2_yyyy_sum = db3_rsget("category2")
				flist(i).category3_yyyy_sum = db3_rsget("category3")
				flist(i).category4_yyyy_sum = db3_rsget("category4")
				flist(i).category5_yyyy_sum = db3_rsget("category5")
				flist(i).category5_yyyy_sum = db3_rsget("category5")
				flist(i).category_yyyy_sum = db3_rsget("categorysum")
				if db3_rsget("category0_name") <> "" then
				flist(i).gubun0_name = db3_rsget("category0_name")
				end if
				flist(i).gubun1_name = db3_rsget("category1_name")
				flist(i).gubun2_name = db3_rsget("category2_name")
				flist(i).gubun3_name = db3_rsget("category3_name")
				flist(i).gubun4_name = db3_rsget("category4_name")
				if db3_rsget("category5_name") <> "" then
				flist(i).gubun5_name = db3_rsget("category5_name")
				end if				
				flist(i).gubun0_isusing	= db3_rsget("category0_isusing")
				flist(i).gubun1_isusing	= db3_rsget("category1_isusing")
				flist(i).gubun2_isusing	= db3_rsget("category2_isusing")
				flist(i).gubun3_isusing	= db3_rsget("category3_isusing")
				flist(i).gubun4_isusing	= db3_rsget("category4_isusing")
				flist(i).gubun5_isusing	= db3_rsget("category5_isusing")
				
			db3_rsget.movenext
			i = i+1
		loop		
	end if
db3_rsget.close			
end sub
'##################################################################
public sub fwongamonth_add						'신규등록용 클래스
dim sql , i
sql = "select"
sql = sql & " a.groupname,a.category,a.category_name,a.category_isusing"
sql = sql & " ,a.field,a.field_name,a.gijun_value"
sql = sql & " ,isnull(b.yyyymm,'') as yyyymm,isnull(b.field_value,'') as field_value,isnull(b.count,'') as count"
sql = sql & " from db_datamart.dbo.tbl_month_wonga_category a"
sql = sql & " left join db_datamart.dbo.tbl_month_wonga b"
sql = sql & " on a.groupname = b.groupname and a.category = b.category and a.field = b.field"
sql = sql & " where 1=1 and a.groupname= '"& frectgubun &"'" 

db3_rsget.open sql,db3_dbget,1
'response.write sql&"<br>"		'삑살시 뿌려본다.

FTotalCount = db3_rsget.recordcount
redim flist(FTotalCount)
i = 0
		
	if not db3_rsget.eof then				'레코드의 첫번째가 아니라면
		do until db3_rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
			set flist(i) = new Cwongaitem 			'클래스를 넣고
			
			flist(i).groupname = db3_rsget("groupname")
				flist(i).yyyymm = db3_rsget("yyyymm")
				flist(i).fcategory = db3_rsget("category")
				flist(i).fcategory_isusing = db3_rsget("category_isusing")
				flist(i).fcategory_name = db3_rsget("category_name")
				flist(i).chulgocount = db3_rsget("count")				
				flist(i).ffield = db3_rsget("field")
				flist(i).ffield_name = db3_rsget("field_name")
				flist(i).ffield_value = clng(db3_rsget("field_value"))
				flist(i).fgijun_value = db3_rsget("gijun_value")

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
public function frectcategoryname(categoty,v)		'카테고리명
	dim intU	
		for intU = 0 to owongamonth_re.ftotalcount -1
			if owongamonth_re.flist(intU).fcategory = categoty then
				frectcategoryname = owongamonth_re.flist(intU).fcategory_name
			end if
		next
end function 
public function frectfieldname(categoty,v)		'필드명	
	dim intU	
		for intU = 0 to owongamonth_re.ftotalcount -1
			if owongamonth_re.flist(intU).fcategory = categoty then
				if owongamonth_re.flist(intU).ffield = v then
					frectfieldname = owongamonth_re.flist(intU).ffield_name
				end if 
			end if
		next
end function
public function frectfieldvalue(categoty,v)		'필드값	
	dim intU	
		for intU = 0 to owongamonth_re.ftotalcount -1
			if owongamonth_re.flist(intU).fcategory = categoty then
				if owongamonth_re.flist(intU).ffield = v then
					if owongamonth_re.flist(intU).fcategory_isusing = "y" then
						if owongamonth_re.flist(intU).ffield_value = "" then
						frectfieldvalue = 0
						else
						frectfieldvalue = owongamonth_re.flist(intU).ffield_value
						end if
					end if
				end if 
			end if
		next
end function 		
public function frectchulgovalue(categoty,v)		'출고건당비용	
	dim intU	
		for intU = 0 to owongamonth.ftotalcount -1
			if owongamonth.flist(intU).fcategory = categoty then
				if owongamonth.flist(intU).ffield = v then
					if owongamonth.flist(intU).ffield_value = 0 then
						frectchulgovalue = 0
					else	
						frectchulgovalue =owongamonth.flist(intU).ffield_value / owongamonth.flist(intU).chulgocount
					end if
				end if 
			end if
		next
end function 	
public function frectgijunvalue(categoty,v)		'기준값
	dim intU	
		for intU = 0 to owongamonth_re.ftotalcount -1
			if owongamonth_re.flist(intU).fcategory = categoty then
				if owongamonth_re.flist(intU).ffield = v then
					frectgijunvalue = owongamonth_re.flist(intU).fgijun_value
				end if 
			end if
		next
end function 

public function frectfieldvaluesum(groupbox,yyyymmbox,categoty)		'총운영비
dim sql
	sql = "select b.yyyymm , sum(b.field_value) as field_value_sum"
	sql = sql & " from db_datamart.dbo.tbl_month_wonga_category a"
	sql = sql & " left join db_datamart.dbo.tbl_month_wonga b" 
	sql = sql & " on a.groupname = b.groupname and a.category = b.category and a.field = b.field" 
	sql = sql & " where 1=1 and a.groupname= '"&groupbox&"' and b.yyyymm = '"& yyyymmbox &"' and b.category in ("& categoty &")"
	sql = sql & " group by b.yyyymm"	
	
	db3_rsget.open sql,db3_dbget,1
	'response.write sql&"<br>"		'삑살시 뿌려본다.
		if not db3_rsget.eof then				
			do until db3_rsget.eof				
				frectfieldvaluesum = clng(db3_rsget("field_value_sum"))				
				db3_rsget.movenext
			
			loop			
		end if
	db3_rsget.close			
end function 

%>
