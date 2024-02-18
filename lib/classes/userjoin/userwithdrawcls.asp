<%
'###########################################################
' Description :  텐바이텐 회원탈퇴 현황
' History : 2008.02.15 한용민 개발
'###########################################################

Class cuserwithdrawoneitem		'회원가입현황
	public fwdrawDate		'날짜
	public fwdrawSex		'성별
	public fwdrawAreaSido	'주소(지역)
	public fwdrawAreaGugun	'상세주소
	public fwdrawAge		'나이
	public fwdrawReason		'탈퇴사유
	public fwdrawCount		'탈퇴수
	public fwdrawReason_01	'상품품질불만
	public fwdrawReason_02	'이용빈도낮음
	public fwdrawReason_03	'배송지연
	public fwdrawReason_04	'개인정보유출우려
	public fwdrawReason_05	'교환/환불/품질불만
	public fwdrawReason_06	'기타
	public fwdrawReason_07	'a/s불만
	public fmancount		'남자수
	public fgirlcount		'여자수
	public fwithdrowtotalcount		
	
    Private Sub Class_Initialize()
	end sub

	Private Sub Class_Terminate()
	End Sub
end Class

class cuserwithdrawlist		
	public FItemList

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	
	public FRectStartdate	
	public FRectEndDate

	public sub fuserwithdrawlist()			'회원탈퇴현황
		dim sqlstr, i
		
		sqlstr = "select convert(varchar(10),wdrawDate,121) as wdrawDate"
		sqlstr = sqlstr & " ,sum(wdrawCount) as withdrowtotalcount"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '' then wdrawCount end) as wdrawReason"		
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '01' then wdrawCount end) as wdrawReason_01"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '02' then wdrawCount end) as wdrawReason_02"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '03' then wdrawCount end) as wdrawReason_03"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '04' then wdrawCount end) as wdrawReason_04"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '05' then wdrawCount end) as wdrawReason_05"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '06' then wdrawCount end) as wdrawReason_06"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '07' then wdrawCount end) as wdrawReason_07"
		sqlstr = sqlstr & " ,sum(case when wdrawSex = '남' then wdrawCount end) as mancount"
		sqlstr = sqlstr & " ,sum(case when wdrawSex = '여' then wdrawCount end) as girlcount"
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_withdraw_log"
		sqlstr = sqlstr & " where 1=1"
		
		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),wdrawDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if			
		
		sqlstr = sqlstr & " group by convert(varchar(10),wdrawDate,121)"
		sqlstr = sqlstr & " order by wdrawDate"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				set FItemList(i) = new cuserwithdrawoneitem 			'클래스를 넣고
	
					FItemList(i).fwdrawDate = db3_rsget("wdrawDate")
					FItemList(i).fwithdrowtotalcount = db3_rsget("withdrowtotalcount")
					FItemList(i).fwdrawReason = db3_rsget("wdrawReason")
					FItemList(i).fwdrawReason_01 = db3_rsget("wdrawReason_01")
					FItemList(i).fwdrawReason_02 = db3_rsget("wdrawReason_02")				
					FItemList(i).fwdrawReason_03 = db3_rsget("wdrawReason_03")
					FItemList(i).fwdrawReason_04 = db3_rsget("wdrawReason_04")
					FItemList(i).fwdrawReason_05 = db3_rsget("wdrawReason_05")						
					FItemList(i).fwdrawReason_06 = db3_rsget("wdrawReason_06")
					FItemList(i).fwdrawReason_07 = db3_rsget("wdrawReason_07")
					FItemList(i).fmancount = db3_rsget("mancount")
					FItemList(i).fgirlcount = db3_rsget("girlcount")

				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub	

	public sub fuserwithdraw_sexgraph()			'회원탈퇴현황(성별 그래프용)
		dim sqlstr, i
		
		sqlstr = "select"
		sqlstr = sqlstr & " wdrawSex,sum(wdrawCount) as wdrawCount"
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_withdraw_log"
		sqlstr = sqlstr & " where 1=1"
		
		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),wdrawDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if			
		
		sqlstr = sqlstr & " group by wdrawSex"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				set FItemList(i) = new cuserwithdrawoneitem 			'클래스를 넣고
	
					FItemList(i).fwdrawSex = db3_rsget("wdrawSex")
					FItemList(i).fwdrawCount = db3_rsget("wdrawCount")
					
				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub		

	public sub fuserwithdraw_areagraph()			'회원탈퇴현황(사유 그래프용)
		dim sqlstr, i
		
		sqlstr = "select"
		sqlstr = sqlstr & " wdrawReason,sum(wdrawCount) as wdrawCount"
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_withdraw_log"
		sqlstr = sqlstr & " where 1=1"
		
		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),wdrawDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if			
		
		sqlstr = sqlstr & " group by wdrawReason"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				set FItemList(i) = new cuserwithdrawoneitem 			'클래스를 넣고
	
					FItemList(i).fwdrawReason = db3_rsget("wdrawReason")
					FItemList(i).fwdrawCount = db3_rsget("wdrawCount")
					
				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub	

    Private Sub Class_Initialize()
	end sub

	Private Sub Class_Terminate()
	End Sub	
end class
%>