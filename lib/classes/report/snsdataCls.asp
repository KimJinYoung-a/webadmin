<%
'###########################################################
' Description :  sns ȸ������ ������ Ŭ����
' History : 2017-06-30 ���¿� ����
'###########################################################
Class CSnsItem
    public Fdate

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CSnsContents
    public FRectjoindate		''ȸ������ �Ⱓ(�⵵��-yy, ����-mm, ������-ww, �Ϻ�dd)
    public FRectjoinmm		''��
    public FRectjoinww		''�б�
    public FRectjoinyy		''�⵵
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
				sqldatename = "���ں��հ�"
				sqlselect = " ,isnull([1],0) as '1��',isnull([2],0) as '2��',isnull([3],0) as '3��',isnull([4],0) as '4��',isnull([5],0) as '5��',isnull([6],0) as '6��',isnull([7],0) as '7��',isnull([8],0) as '8��',isnull([9],0) as '9��',isnull([10],0) as '10��',isnull([11],0) as '11��',isnull([12],0) as '12��',isnull([13],0) as '13��',isnull([14],0) as '14��',isnull([15],0) as '15��',isnull([16],0) as '16��',isnull([17],0) as '17��',isnull([18],0) as '18��',isnull([19],0) as '19��',isnull([20],0) as '20��',isnull([21],0) as '21��',isnull([22],0) as '22��',isnull([23],0) as '23��',isnull([24],0) as '24��',isnull([25],0) as '25��',isnull([26],0) as '26��',isnull([27],0) as '27��',isnull([28],0) as '28��',isnull([29],0) as '29��',isnull([30],0) as '30��',isnull([31],0) as '31��' "
				sqlselect = sqlselect & " ,isnull([1],0)+isnull([2],0)+isnull([3],0)+isnull([4],0)+isnull([5],0)+isnull([6],0)+isnull([7],0)+isnull([8],0)+isnull([9],0)+isnull([10],0)+isnull([11],0)+isnull([12],0)+isnull([13],0)+isnull([14],0)+isnull([15],0)+isnull([16],0)+isnull([17],0)+isnull([18],0)+isnull([19],0)+isnull([20],0)+isnull([21],0)+isnull([22],0)+isnull([23],0)+isnull([24],0)+isnull([25],0)+isnull([26],0)+isnull([27],0)+isnull([28],0)+isnull([29],0)+isnull([30],0)+isnull([31],0) as total "
				sqlpivotsum = "	sum(cnt) for day in ([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26],[27],[28],[29],[30],[31]) "
				sqlwhere = "		and datepart(yyyy,regdate)='"&FRectjoinyy&"' and datepart(mm,regdate)='"&FRectjoinmm&"' "
				sqlcasewhen = "datepart(dd,regdate) as day"
			elseif FRectjoindate = "ww" then
				sqldatename = "�������հ�"
				sqlselect = " ,isnull(["&sqlweek&"],0) as '"&sqlweek&"����',isnull(["&sqlweek+1&"],0) as '"&sqlweek+1&"����',isnull(["&sqlweek+2&"],0) as '"&sqlweek+2&"����',isnull(["&sqlweek+3&"],0) as '"&sqlweek+3&"����',isnull(["&sqlweek+4&"],0) as '"&sqlweek+4&"����',isnull(["&sqlweek+5&"],0) as '"&sqlweek+5&"����',isnull(["&sqlweek+6&"],0) as '"&sqlweek+6&"����',isnull(["&sqlweek+7&"],0) as '"&sqlweek+7&"����',isnull(["&sqlweek+8&"],0) as '"&sqlweek+8&"����',isnull(["&sqlweek+9&"],0) as '"&sqlweek+9&"����',isnull(["&sqlweek+10&"],0) as '"&sqlweek+10&"����',isnull(["&sqlweek+11&"],0) as '"&sqlweek+11&"����',isnull(["&sqlweek+12&"],0) as '"&sqlweek+12&"����' "
				sqlselect = sqlselect & " ,isnull(["&sqlweek&"],0)+isnull(["&sqlweek+1&"],0)+isnull(["&sqlweek+2&"],0)+isnull(["&sqlweek+3&"],0)+isnull(["&sqlweek+4&"],0)+isnull(["&sqlweek+5&"],0)+isnull(["&sqlweek+6&"],0)+isnull(["&sqlweek+7&"],0)+isnull(["&sqlweek+8&"],0)+isnull(["&sqlweek+9&"],0)+isnull(["&sqlweek+10&"],0)+isnull(["&sqlweek+11&"],0)+isnull(["&sqlweek+12&"],0) as total "
				sqlpivotsum = "	sum(cnt) for week in (["&sqlweek&"],["&sqlweek+1&"],["&sqlweek+2&"],["&sqlweek+3&"],["&sqlweek+4&"],["&sqlweek+5&"],["&sqlweek+6&"],["&sqlweek+7&"],["&sqlweek+8&"],["&sqlweek+9&"],["&sqlweek+10&"],["&sqlweek+11&"],["&sqlweek+12&"]) "
				sqlwhere = "		and datepart(yyyy,regdate)='"&FRectjoinyy&"' and datepart(ww,regdate) >= "&sqlweek&" and datepart(ww,regdate) <= "&sqlweek+12&" "
				sqlcasewhen = "datepart(ww,regdate) as week"
			elseif FRectjoindate = "mm" then
				sqldatename = "�����հ�"
				sqlselect = " ,isnull([1],0) as '1��',isnull([2],0) as '2��',isnull([3],0) as '3��',isnull([4],0) as '4��',isnull([5],0) as '5��',isnull([6],0) as '6��',isnull([7],0) as '7��',isnull([8],0) as '8��',isnull([9],0) as '9��',isnull([10],0) as '10��',isnull([11],0) as '11��',isnull([12],0) as '12��' "
				sqlselect = sqlselect & " ,isnull([1],0)+isnull([2],0)+isnull([3],0)+isnull([4],0)+isnull([5],0)+isnull([6],0)+isnull([7],0)+isnull([8],0)+isnull([9],0)+isnull([10],0)+isnull([11],0)+isnull([12],0) as total "
				sqlpivotsum = "	sum(cnt) for month in ([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12] ) "
				sqlwhere = "		and datepart(yyyy,regdate)='"&FRectjoinyy&"' "	
				sqlcasewhen = "datepart(mm,regdate) as month"
			elseif FRectjoindate = "yy" then
				sqldatename = "�⵵���հ�"
				sqlselect = " ,isnull(["&(FRectjoinyy-3)&"],0) as '"&(FRectjoinyy-3)&"��',isnull(["&(FRectjoinyy-2)&"],0) as '"&(FRectjoinyy-2)&"��',isnull(["&(FRectjoinyy-1)&"],0) as '"&(FRectjoinyy-1)&"��',isnull(["&(FRectjoinyy)&"],0) as '"&(FRectjoinyy)&"��' "
				sqlselect = sqlselect & " ,isnull(["&(FRectjoinyy-3)&"],0)+isnull(["&(FRectjoinyy-2)&"],0)+isnull(["&(FRectjoinyy-1)&"],0)+isnull(["&(FRectjoinyy)&"],0) as total "
				sqlpivotsum = "	sum(cnt) for year in (["&(FRectjoinyy-3)&"],["&(FRectjoinyy-2)&"],["&(FRectjoinyy-1)&"],["&(FRectjoinyy)&"] ) "
				sqlwhere = "		and datepart(yyyy,regdate)>='"&(FRectjoinyy-3)&"' and datepart(yyyy,regdate)<='"&FRectjoinyy&"' "	''�����⿡�� �ֱ� 4�Ⱓ
				sqlcasewhen = "datepart(yyyy,regdate) as year"
			end if
		end if

		sqlstr = "SELECT evtid as [����] " & vbcrlf
		sqlstr = sqlstr & sqlselect & vbcrlf
		sqlstr = sqlstr & "	FROM  " & vbcrlf
		sqlstr = sqlstr & "	( " & vbcrlf
		sqlstr = sqlstr & " 	select CASE WHEN GROUPING(eventid) = 0 THEN eventid ELSE '"&sqldatename&"' END as 'evtid', "&sqlcasewhen&", COUNT(*) as cnt " & vbcrlf
		sqlstr = sqlstr & "		from "& tendb &"db_user.dbo.tbl_user_n P " & vbcrlf
		sqlstr = sqlstr & "			join "& tendb &"db_user.dbo.tbl_logindata S  " & vbcrlf
		sqlstr = sqlstr & "				on P.userid=S.userid and userdiv='05'  " & vbcrlf
		sqlstr = sqlstr & "		where convert(varchar(10),regdate,120) >= '2017-06-14' " & vbcrlf	'6�� 14�Ϻ��� sns�α��� ����(���̹�����)
		sqlstr = sqlstr & 		sqlwhere & vbcrlf
		sqlstr = sqlstr & "			and eventid in ('PC','PC_nv','PC_ka','PC_fb','PC_gl','PC_ap','MO','MO_nv','MO_ka','MO_fb','MO_gl','MO_ap','AP','AP_nv','AP_ka','AP_fb','AP_gl','AP_ap') " & vbcrlf
		sqlstr = sqlstr & "		group by regdate, eventid WITH ROLLUP " & vbcrlf
		sqlstr = sqlstr & "	) as a " & vbcrlf
		sqlstr = sqlstr & " pivot " & vbcrlf
		sqlstr = sqlstr & " ( " & vbcrlf
		sqlstr = sqlstr & 	sqlpivotsum & vbcrlf
		sqlstr = sqlstr & " ) as tp "
		sqlstr = sqlstr & " ORDER BY [����] ASC "
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

