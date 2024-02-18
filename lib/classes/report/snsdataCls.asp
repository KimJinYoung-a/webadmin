<%
'###########################################################
' Description :  sns 회원가입 데이터 클래스
' History : 2017-06-30 유태욱 생성
'###########################################################
Class CSnsItem
    public Fdate

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CSnsContents
    public FRectjoindate		''회원가입 기간(년도별-yy, 월별-mm, 주차별-ww, 일별dd)
    public FRectjoinmm		''월
    public FRectjoinww		''분기
    public FRectjoinyy		''년도
	public tendb


	''---------------------------------------------------------------------------------
	public Function GetSnsjoinList()
		dim sqlstr, i, sqlselect, sqlweek, sqldatename, sqlpivotsum, sqlwhere, sqlcasewhen
		if FRectjoinww <> "" then
			if FRectjoinww = "1" then
				sqlweek = 1
			elseif FRectjoinww = "2" then
				sqlweek = 14
			elseif FRectjoinww = "3" then
				sqlweek = 27
			elseif FRectjoinww = "4" then
				sqlweek = 40
			else
				sqlweek = 1
			end if
		end if
'response.write FRectjoindate&"/"&FRectjoinww&"/"&sqlweek
'response.end
		if FRectjoindate <> "" then
			if FRectjoindate = "dd" then
				sqldatename = "일자별합계"
				sqlselect = " ,isnull([1],0) as '1일',isnull([2],0) as '2일',isnull([3],0) as '3일',isnull([4],0) as '4일',isnull([5],0) as '5일',isnull([6],0) as '6일',isnull([7],0) as '7일',isnull([8],0) as '8일',isnull([9],0) as '9일',isnull([10],0) as '10일',isnull([11],0) as '11일',isnull([12],0) as '12일',isnull([13],0) as '13일',isnull([14],0) as '14일',isnull([15],0) as '15일',isnull([16],0) as '16일',isnull([17],0) as '17일',isnull([18],0) as '18일',isnull([19],0) as '19일',isnull([20],0) as '20일',isnull([21],0) as '21일',isnull([22],0) as '22일',isnull([23],0) as '23일',isnull([24],0) as '24일',isnull([25],0) as '25일',isnull([26],0) as '26일',isnull([27],0) as '27일',isnull([28],0) as '28일',isnull([29],0) as '29일',isnull([30],0) as '30일',isnull([31],0) as '31일' "
				sqlselect = sqlselect & " ,isnull([1],0)+isnull([2],0)+isnull([3],0)+isnull([4],0)+isnull([5],0)+isnull([6],0)+isnull([7],0)+isnull([8],0)+isnull([9],0)+isnull([10],0)+isnull([11],0)+isnull([12],0)+isnull([13],0)+isnull([14],0)+isnull([15],0)+isnull([16],0)+isnull([17],0)+isnull([18],0)+isnull([19],0)+isnull([20],0)+isnull([21],0)+isnull([22],0)+isnull([23],0)+isnull([24],0)+isnull([25],0)+isnull([26],0)+isnull([27],0)+isnull([28],0)+isnull([29],0)+isnull([30],0)+isnull([31],0) as total "
				sqlpivotsum = "	sum(cnt) for day in ([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26],[27],[28],[29],[30],[31]) "
				sqlwhere = "		and datepart(yyyy,regdate)='"&FRectjoinyy&"' and datepart(mm,regdate)='"&FRectjoinmm&"' "
				sqlcasewhen = "datepart(dd,regdate) as day"
			elseif FRectjoindate = "ww" then
				sqldatename = "주차별합계"
				sqlselect = " ,isnull(["&sqlweek&"],0) as '"&sqlweek&"주차',isnull(["&sqlweek+1&"],0) as '"&sqlweek+1&"주차',isnull(["&sqlweek+2&"],0) as '"&sqlweek+2&"주차',isnull(["&sqlweek+3&"],0) as '"&sqlweek+3&"주차',isnull(["&sqlweek+4&"],0) as '"&sqlweek+4&"주차',isnull(["&sqlweek+5&"],0) as '"&sqlweek+5&"주차',isnull(["&sqlweek+6&"],0) as '"&sqlweek+6&"주차',isnull(["&sqlweek+7&"],0) as '"&sqlweek+7&"주차',isnull(["&sqlweek+8&"],0) as '"&sqlweek+8&"주차',isnull(["&sqlweek+9&"],0) as '"&sqlweek+9&"주차',isnull(["&sqlweek+10&"],0) as '"&sqlweek+10&"주차',isnull(["&sqlweek+11&"],0) as '"&sqlweek+11&"주차',isnull(["&sqlweek+12&"],0) as '"&sqlweek+12&"주차' "
				sqlselect = sqlselect & " ,isnull(["&sqlweek&"],0)+isnull(["&sqlweek+1&"],0)+isnull(["&sqlweek+2&"],0)+isnull(["&sqlweek+3&"],0)+isnull(["&sqlweek+4&"],0)+isnull(["&sqlweek+5&"],0)+isnull(["&sqlweek+6&"],0)+isnull(["&sqlweek+7&"],0)+isnull(["&sqlweek+8&"],0)+isnull(["&sqlweek+9&"],0)+isnull(["&sqlweek+10&"],0)+isnull(["&sqlweek+11&"],0)+isnull(["&sqlweek+12&"],0) as total "
				sqlpivotsum = "	sum(cnt) for week in (["&sqlweek&"],["&sqlweek+1&"],["&sqlweek+2&"],["&sqlweek+3&"],["&sqlweek+4&"],["&sqlweek+5&"],["&sqlweek+6&"],["&sqlweek+7&"],["&sqlweek+8&"],["&sqlweek+9&"],["&sqlweek+10&"],["&sqlweek+11&"],["&sqlweek+12&"]) "
				sqlwhere = "		and datepart(yyyy,regdate)='"&FRectjoinyy&"' and datepart(ww,regdate) >= "&sqlweek&" and datepart(ww,regdate) <= "&sqlweek+12&" "
				sqlcasewhen = "datepart(ww,regdate) as week"
			elseif FRectjoindate = "mm" then
				sqldatename = "월별합계"
				sqlselect = " ,isnull([1],0) as '1월',isnull([2],0) as '2월',isnull([3],0) as '3월',isnull([4],0) as '4월',isnull([5],0) as '5월',isnull([6],0) as '6월',isnull([7],0) as '7월',isnull([8],0) as '8월',isnull([9],0) as '9월',isnull([10],0) as '10월',isnull([11],0) as '11월',isnull([12],0) as '12월' "
				sqlselect = sqlselect & " ,isnull([1],0)+isnull([2],0)+isnull([3],0)+isnull([4],0)+isnull([5],0)+isnull([6],0)+isnull([7],0)+isnull([8],0)+isnull([9],0)+isnull([10],0)+isnull([11],0)+isnull([12],0) as total "
				sqlpivotsum = "	sum(cnt) for month in ([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12] ) "
				sqlwhere = "		and datepart(yyyy,regdate)='"&FRectjoinyy&"' "	
				sqlcasewhen = "datepart(mm,regdate) as month"
			elseif FRectjoindate = "yy" then
				sqldatename = "년도별합계"
				sqlselect = " ,isnull(["&(FRectjoinyy-3)&"],0) as '"&(FRectjoinyy-3)&"년',isnull(["&(FRectjoinyy-2)&"],0) as '"&(FRectjoinyy-2)&"년',isnull(["&(FRectjoinyy-1)&"],0) as '"&(FRectjoinyy-1)&"년',isnull(["&(FRectjoinyy)&"],0) as '"&(FRectjoinyy)&"년' "
				sqlselect = sqlselect & " ,isnull(["&(FRectjoinyy-3)&"],0)+isnull(["&(FRectjoinyy-2)&"],0)+isnull(["&(FRectjoinyy-1)&"],0)+isnull(["&(FRectjoinyy)&"],0) as total "
				sqlpivotsum = "	sum(cnt) for year in (["&(FRectjoinyy-3)&"],["&(FRectjoinyy-2)&"],["&(FRectjoinyy-1)&"],["&(FRectjoinyy)&"] ) "
				sqlwhere = "		and datepart(yyyy,regdate)>='"&(FRectjoinyy-3)&"' and datepart(yyyy,regdate)<='"&FRectjoinyy&"' "	''지정년에서 최근 4년간
				sqlcasewhen = "datepart(yyyy,regdate) as year"
			end if
		end if

		sqlstr = "SELECT evtid as [구분] " & vbcrlf
		sqlstr = sqlstr & sqlselect & vbcrlf
		sqlstr = sqlstr & "	FROM  " & vbcrlf
		sqlstr = sqlstr & "	( " & vbcrlf
		sqlstr = sqlstr & " 	select CASE WHEN GROUPING(eventid) = 0 THEN eventid ELSE '"&sqldatename&"' END as 'evtid', "&sqlcasewhen&", COUNT(*) as cnt " & vbcrlf
		sqlstr = sqlstr & "		from "& tendb &"db_user.dbo.tbl_user_n P " & vbcrlf
		sqlstr = sqlstr & "			join "& tendb &"db_user.dbo.tbl_logindata S  " & vbcrlf
		sqlstr = sqlstr & "				on P.userid=S.userid and userdiv='05'  " & vbcrlf
		sqlstr = sqlstr & "		where convert(varchar(10),regdate,120) >= '2017-06-14' " & vbcrlf	'6월 14일부터 sns로그인 시작(네이버먼저)
		sqlstr = sqlstr & 		sqlwhere & vbcrlf
		sqlstr = sqlstr & "			and eventid in ('PC','PC_nv','PC_ka','PC_fb','PC_gl','PC_ap','MO','MO_nv','MO_ka','MO_fb','MO_gl','MO_ap','AP','AP_nv','AP_ka','AP_fb','AP_gl','AP_ap') " & vbcrlf
		sqlstr = sqlstr & "		group by regdate, eventid WITH ROLLUP " & vbcrlf
		sqlstr = sqlstr & "	) as a " & vbcrlf
		sqlstr = sqlstr & " pivot " & vbcrlf
		sqlstr = sqlstr & " ( " & vbcrlf
		sqlstr = sqlstr & 	sqlpivotsum & vbcrlf
		sqlstr = sqlstr & " ) as tp "
		sqlstr = sqlstr & " ORDER BY [구분] ASC "
'response.write sqlstr&"<br>"
'response.end
		db3_rsget.open sqlstr,db3_dbget,1
			IF not db3_rsget.EOF THEN
				GetSnsjoinList = db3_rsget.getRows() 
			END IF	
		db3_rsget.close
	end Function


    Private Sub Class_Initialize()
		IF application("Svr_Info")="Dev" THEN
			tendb = "tendb."
		end if
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class
%>

