<%
'########################################################
' 2008년 01월 29일 한용민 개발
' 2011.07.14 허진원 수정; 20, 30대 전후반으로 구분
'########################################################
%>
<%
Class cuserjoinoneitem		'회원가입현황
	public fjoinDate		'가입일
	public fjoinSex			'연령별
	public fjoinAreaSido	'지역
	public fjoinAreaGugun	'상세지역
	public fjoinAge			'나이
	public fjoinPath		'가입경로
	public fjoinsexcount	'연령별별 고객수
	public fjoinAreaSidocount	'지역별 고객수
	public fjoinPathCount	'채널별 고객수
	public fjoinPath_count
	Public fjoinPathCountArr

	public ftotal_count
	public ftotal_0_9_count
	public ftotal_10_14_count
	public ftotal_15_19_count
	public ftotal_10_19_count
	public ftotal_20_29_count
	public ftotal_20_24_count
	public ftotal_25_29_count
	public ftotal_30_39_count
	public ftotal_30_34_count
	public ftotal_35_39_count
	public ftotal_40_44_count
	public ftotal_45_49_count
	public ftotal_40_49_count
	public ftotal_50_59_count
	public ftotal_60_69_count
	public ftotal_70_79_count
	public ftotal_80_89_count
	public ftotal_90_99_count
	public ftotal_50_count
	public ftotal_100_count
	public ftotal_etc_count

	public fsexman_total_count
	public fsexman_0_9_count
	public fsexman_10_14_count
	public fsexman_15_19_count
	public fsexman_10_19_count
	public fsexman_20_29_count
	public fsexman_20_24_count
	public fsexman_25_29_count
	public fsexman_30_39_count
	public fsexman_30_34_count
	public fsexman_35_39_count
	public fsexman_40_44_count
	public fsexman_45_49_count
	public fsexman_40_49_count
	public fsexman_50_59_count
	public fsexman_60_69_count
	public fsexman_70_79_count
	public fsexman_80_89_count
	public fsexman_90_99_count
	public fsexman_50_count
	public fsexman_100_count
	public fsexman_etc_count

	public fsexgirl_total_count
	public fsexgirl_0_9_count
	public fsexgirl_10_14_count
	public fsexgirl_15_19_count
	public fsexgirl_10_19_count
	public fsexgirl_20_29_count
	public fsexgirl_20_24_count
	public fsexgirl_25_29_count
	public fsexgirl_30_39_count
	public fsexgirl_30_34_count
	public fsexgirl_35_39_count
	public fsexgirl_40_44_count
	public fsexgirl_45_49_count
	public fsexgirl_40_49_count
	public fsexgirl_50_59_count
	public fsexgirl_60_69_count
	public fsexgirl_70_79_count
	public fsexgirl_80_89_count
	public fsexgirl_90_99_count
	public fsexgirl_50_count
	public fsexgirl_100_count
	public fsexgirl_etc_count


    Private Sub Class_Initialize()
	end sub

	Private Sub Class_Terminate()
	End Sub
end Class

class cuserjoinlist
	public FItemList()
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public frectjoinSex
	public frectjoinAreaSido
	public FRectStartdate
	public FRectEndDate
	public frectjoinPath


	public sub fuserjoinlist()			'회원가입현황(연령별)
		dim sqlstr, i
		sqlstr = "select"
		sqlstr = sqlstr & " sum(joinCount) as total_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 0 and 9 then joinCount Else 0 end)) as total_0_9_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 10 and 14 then joinCount Else 0 end)) as total_10_14_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 15 and 19 then joinCount Else 0 end)) as total_15_19_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 20 and 24 then joinCount Else 0 end)) as total_20_24_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 25 and 29 then joinCount Else 0 end)) as total_25_29_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 30 and 34 then joinCount Else 0 end)) as total_30_34_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 35 and 39 then joinCount Else 0 end)) as total_35_39_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 40 and 44 then joinCount Else 0 end)) as total_40_44_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 45 and 49 then joinCount Else 0 end)) as total_45_49_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge between 50 and 998 then joinCount Else 0 end)) as total_50_count"
'		sqlstr = sqlstr & " ,(sum(case when joinAge between 50 and 59 then joinCount Else 0 end)) as total_50_59_count"
'		sqlstr = sqlstr & " ,(sum(case when joinAge between 60 and 69 then joinCount Else 0 end)) as total_60_69_count"
'		sqlstr = sqlstr & " ,(sum(case when joinAge between 70 and 79 then joinCount Else 0 end)) as total_70_79_count"
'		sqlstr = sqlstr & " ,(sum(case when joinAge between 80 and 89 then joinCount Else 0 end)) as total_80_89_count"
'		sqlstr = sqlstr & " ,(sum(case when joinAge between 90 and 99 then joinCount Else 0 end)) as total_90_99_count"
		sqlstr = sqlstr & " ,(sum(case when joinAge=999 then joinCount Else 0 end)) as total_etc_count"

		sqlstr = sqlstr & " ,sum(case when joinSex ='남' then joinCount Else 0 end) as sexman_total_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 0 and 9 then joinCount Else 0 end)) as sexman_0_9_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 10 and 14 then joinCount Else 0 end)) as sexman_10_14_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 15 and 19 then joinCount Else 0 end)) as sexman_15_19_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 20 and 24 then joinCount Else 0 end)) as sexman_20_24_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 25 and 29 then joinCount Else 0 end)) as sexman_25_29_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 30 and 34 then joinCount Else 0 end)) as sexman_30_34_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 35 and 39 then joinCount Else 0 end)) as sexman_35_39_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 40 and 44 then joinCount Else 0 end)) as sexman_40_44_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 45 and 49 then joinCount Else 0 end)) as sexman_45_49_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 50 and 998 then joinCount Else 0 end)) as sexman_50_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 50 and 59 then joinCount Else 0 end)) as sexman_50_59_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 60 and 69 then joinCount Else 0 end)) as sexman_60_69_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 70 and 79 then joinCount Else 0 end)) as sexman_70_79_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 80 and 89 then joinCount Else 0 end)) as sexman_80_89_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge between 90 and 99 then joinCount Else 0 end)) as sexman_90_99_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='남' and joinAge=999 then joinCount Else 0 end)) as sexman_etc_count"

		sqlstr = sqlstr & " ,sum(case when joinSex ='여' then joinCount Else 0 end) as sexgirl_total_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 0 and 9 then joinCount Else 0 end)) as sexgirl_0_9_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 10 and 14 then joinCount Else 0 end)) as sexgirl_10_14_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 15 and 19 then joinCount Else 0 end)) as sexgirl_15_19_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 20 and 24 then joinCount Else 0 end)) as sexgirl_20_24_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 25 and 29 then joinCount Else 0 end)) as sexgirl_25_29_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 30 and 34 then joinCount Else 0 end)) as sexgirl_30_34_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 35 and 39 then joinCount Else 0 end)) as sexgirl_35_39_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 40 and 44 then joinCount Else 0 end)) as sexgirl_40_44_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 45 and 49 then joinCount Else 0 end)) as sexgirl_45_49_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 50 and 998 then joinCount Else 0 end)) as sexgirl_50_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 50 and 59 then joinCount Else 0 end)) as sexgirl_50_59_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 60 and 69 then joinCount Else 0 end)) as sexgirl_60_69_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 70 and 79 then joinCount Else 0 end)) as sexgirl_70_79_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 80 and 89 then joinCount Else 0 end)) as sexgirl_80_89_count"
'		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge between 90 and 99 then joinCount Else 0 end)) as sexgirl_90_99_count"
		sqlstr = sqlstr & " ,(sum(case when joinSex ='여' and joinAge=999 then joinCount Else 0 end)) as sexgirl_etc_count"

		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_join_log"
		sqlstr = sqlstr & " where 1=1"

		if frectjoinAreaSido <> "" then
			sqlstr = sqlstr & " and joinAreaSido = '" & frectjoinAreaSido & "'"
		end if

		if frectjoinPath <> "" then
			sqlstr = sqlstr & " and (case when joinPath='' then '10X10' else joinPath end) = '" & frectjoinPath & "'"
		end if

		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),joinDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"
		end if

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"

		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0

		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				set FItemList(i) = new cuserjoinoneitem 			'클래스를 넣고

				FItemList(i).ftotal_count = db3_rsget("total_count")
				FItemList(i).ftotal_0_9_count = db3_rsget("total_0_9_count")
				FItemList(i).ftotal_10_14_count = db3_rsget("total_10_14_count")
				FItemList(i).ftotal_15_19_count = db3_rsget("total_15_19_count")
				FItemList(i).ftotal_20_24_count = db3_rsget("total_20_24_count")
				FItemList(i).ftotal_25_29_count = db3_rsget("total_25_29_count")
				FItemList(i).ftotal_30_34_count = db3_rsget("total_30_34_count")
				FItemList(i).ftotal_35_39_count = db3_rsget("total_35_39_count")
				FItemList(i).ftotal_40_44_count = db3_rsget("total_40_44_count")
				FItemList(i).ftotal_45_49_count = db3_rsget("total_45_49_count")
'				FItemList(i).ftotal_50_59_count = db3_rsget("total_50_59_count")
'				FItemList(i).ftotal_60_69_count = db3_rsget("total_60_69_count")
'				FItemList(i).ftotal_70_79_count = db3_rsget("total_70_79_count")
'				FItemList(i).ftotal_80_89_count = db3_rsget("total_80_89_count")
'				FItemList(i).ftotal_90_99_count = db3_rsget("total_90_99_count")
				FItemList(i).ftotal_50_count = db3_rsget("total_50_count")
				FItemList(i).ftotal_etc_count = db3_rsget("total_etc_count")

				FItemList(i).fsexman_total_count = db3_rsget("sexman_total_count")
				FItemList(i).fsexman_0_9_count = db3_rsget("sexman_0_9_count")
				FItemList(i).fsexman_10_14_count = db3_rsget("sexman_10_14_count")
				FItemList(i).fsexman_15_19_count = db3_rsget("sexman_15_19_count")
				FItemList(i).fsexman_20_24_count = db3_rsget("sexman_20_24_count")
				FItemList(i).fsexman_25_29_count = db3_rsget("sexman_25_29_count")
				FItemList(i).fsexman_30_34_count = db3_rsget("sexman_30_34_count")
				FItemList(i).fsexman_35_39_count = db3_rsget("sexman_35_39_count")
				FItemList(i).fsexman_40_44_count = db3_rsget("sexman_40_44_count")
				FItemList(i).fsexman_45_49_count = db3_rsget("sexman_45_49_count")
'				FItemList(i).fsexman_50_59_count = db3_rsget("sexman_50_59_count")
'				FItemList(i).fsexman_60_69_count = db3_rsget("sexman_60_69_count")
'				FItemList(i).fsexman_70_79_count = db3_rsget("sexman_70_79_count")
'				FItemList(i).fsexman_80_89_count = db3_rsget("sexman_80_89_count")
'				FItemList(i).fsexman_90_99_count = db3_rsget("sexman_90_99_count")
				FItemList(i).fsexman_50_count = db3_rsget("sexman_50_count")
				FItemList(i).fsexman_etc_count = db3_rsget("sexman_etc_count")

				FItemList(i).fsexgirl_total_count = db3_rsget("sexgirl_total_count")
				FItemList(i).fsexgirl_0_9_count = db3_rsget("sexgirl_0_9_count")
				FItemList(i).fsexgirl_10_14_count = db3_rsget("sexgirl_10_14_count")
				FItemList(i).fsexgirl_15_19_count = db3_rsget("sexgirl_15_19_count")
				FItemList(i).fsexgirl_20_24_count = db3_rsget("sexgirl_20_24_count")
				FItemList(i).fsexgirl_25_29_count = db3_rsget("sexgirl_25_29_count")
				FItemList(i).fsexgirl_30_34_count = db3_rsget("sexgirl_30_34_count")
				FItemList(i).fsexgirl_35_39_count = db3_rsget("sexgirl_35_39_count")
				FItemList(i).fsexgirl_40_44_count = db3_rsget("sexgirl_40_44_count")
				FItemList(i).fsexgirl_45_49_count = db3_rsget("sexgirl_45_49_count")
'				FItemList(i).fsexgirl_50_59_count = db3_rsget("sexgirl_50_59_count")
'				FItemList(i).fsexgirl_60_69_count = db3_rsget("sexgirl_60_69_count")
'				FItemList(i).fsexgirl_70_79_count = db3_rsget("sexgirl_70_79_count")
'				FItemList(i).fsexgirl_80_89_count = db3_rsget("sexgirl_80_89_count")
'				FItemList(i).fsexgirl_90_99_count = db3_rsget("sexgirl_90_99_count")
				FItemList(i).fsexgirl_50_count = db3_rsget("sexgirl_50_count")
				FItemList(i).fsexgirl_etc_count = db3_rsget("sexgirl_etc_count")

				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub

	public sub fuserjoinarealist()			'회원가입현황(지역별)
		dim sqlstr, i
		sqlstr = "select"
		sqlstr = sqlstr & " joinSex,joinAreaSido,sum(joincount) as joinAreaSido_count"

		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_join_log"
		sqlstr = sqlstr & " where 1=1"

		if frectjoinSex <> "" then
			sqlstr = sqlstr & " and joinSex = '"& frectjoinSex &"'"
		end if

		if frectjoinPath <> "" then
			sqlstr = sqlstr & " and (case when joinPath='' then '10X10' else joinPath end) = '" & frectjoinPath & "'"
		end if

		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),joinDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"
		end if

		sqlstr = sqlstr & " group by joinSex,joinAreaSido"
		sqlstr = sqlstr & " order by joinAreaSido asc"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"

		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0

		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				set FItemList(i) = new cuserjoinoneitem 			'클래스를 넣고

				FItemList(i).fjoinSex = db3_rsget("joinSex")
				FItemList(i).fjoinAreaSido = db3_rsget("joinAreaSido")
				FItemList(i).fjoinAreaSidocount = db3_rsget("joinAreaSido_count")
				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub

	public sub fuserjoin_sex()		'회원가입현황 그래프용(연령별)
		dim sqlstr, i
		sqlstr = "select joinSex,sum(joincount) as joinsexcount"
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_join_log"
		sqlstr = sqlstr & " where 1=1"

		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),joinDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"
		end if

		if frectjoinAreaSido <> "" then
			sqlstr = sqlstr & " and joinAreaSido = '" & frectjoinAreaSido & "'"
		end if

		if frectjoinPath <> "" then
			sqlstr = sqlstr & " and (case when joinPath='' then '10X10' else joinPath end) = '" & frectjoinPath & "'"
		end if

		sqlstr = sqlstr & " group by joinSex"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"

		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0

		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				set FItemList(i) = new cuserjoinoneitem 			'클래스를 넣고

					FItemList(i).fjoinsexcount = db3_rsget("joinsexcount")
					FItemList(i).fjoinSex = db3_rsget("joinSex")

				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub

	public sub fuserjoin_area()			'회원가입현황 그래프용(지역별)
		dim sqlstr, i
		sqlstr = "select joinAreaSido,sum(joincount) as joinAreaSidocount"
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_join_log"
		sqlstr = sqlstr & " where 1=1"

		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),joinDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"
		end if

		if frectjoinSex <> "" then
			sqlstr = sqlstr & " and joinSex = '"& frectjoinSex &"'"
		end if

		if frectjoinPath <> "" then
			sqlstr = sqlstr & " and (case when joinPath='' then '10X10' else joinPath end) = '" & frectjoinPath & "'"
		end if

		sqlstr = sqlstr & " group by joinAreaSido"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"

		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0

		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				set FItemList(i) = new cuserjoinoneitem 			'클래스를 넣고

					FItemList(i).fjoinAreaSido = db3_rsget("joinAreaSido")
					FItemList(i).fjoinAreaSidocount = db3_rsget("joinAreaSidocount")

				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end Sub

	' 회원가입현황(채널별)
	' admin/userjoin/userjoin.asp
	public sub fuserjoinchannellist()
		dim sqlstr, i, j, sqlsearch

		if frectjoinSex <> "" then
			sqlsearch = sqlsearch & " and l.joinSex = '"& frectjoinSex &"'"
		end if
		if frectjoinPath <> "" then
			sqlsearch = sqlsearch & " and (case when l.joinPath='' then '10X10' else l.joinPath end) = '" & frectjoinPath & "'"
		end if
		if FRectStartdate<>"" and FRectEndDate<>"" then
			sqlsearch = sqlsearch & " and convert(varchar(10),l.joinDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"
		end if

		sqlStr = "select count(*) as cnt"
		sqlstr = sqlstr & " from (" & vbcrlf
		sqlstr = sqlstr & " 	select" & vbcrlf
		sqlstr = sqlstr & " 	convert(varchar(10), l.joindate, 121) as joindate" & vbcrlf
		sqlstr = sqlstr & " 	, (case when l.joinPath = '' then '10x10' else joinPath end) as joinPath" & vbcrlf
		sqlstr = sqlstr & " 	from db_datamart.dbo.tbl_user_join_log l" & vbcrlf
		sqlstr = sqlstr & " 	where 1=1 " & sqlsearch
		sqlstr = sqlstr & " 	group by convert(varchar(10), l.joindate, 121)" & vbcrlf
		sqlstr = sqlstr & " 	, (case when l.joinPath = '' then '10x10' else joinPath end)" & vbcrlf
		sqlstr = sqlstr & " ) as t" & vbcrlf

		'response.write sqlstr &"<br>"
		db3_rsget.Open sqlstr,db3_dbget,1
			FTotalCount = db3_rsget("cnt")
		db3_rsget.Close
		
		if FTotalCount < 1 then exit sub

        sqlstr = "select top "& Cstr(FPageSize * FCurrPage)
		sqlstr = sqlstr & " convert(varchar(10), l.joindate, 121) as joindate" & vbcrlf
		sqlstr = sqlstr & " , sum(joincount) as joinPath_count" & vbcrlf
		sqlstr = sqlstr & " , (case when Left(l.joinPath,6) = 'mobile' then 2 else 1 end) as ordBy" & vbcrlf
		sqlstr = sqlstr & " , (case when l.joinPath = '' then '10x10' else joinPath end) as joinPath" & vbcrlf
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_join_log l" & vbcrlf
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " group by convert(varchar(10), l.joindate, 121)" & vbcrlf
		sqlstr = sqlstr & " , (case when Left(l.joinPath,6) = 'mobile' then 2 else 1 end)" & vbcrlf
		sqlstr = sqlstr & " , (case when l.joinPath = '' then '10x10' else joinPath end)" & vbcrlf
		sqlstr = sqlstr & " order by joindate desc, ordBy asc, joinPath asc" & vbcrlf

		'response.write sqlStr &"<br>"
		db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,1

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
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.EOF
				set FItemList(i) = new cuserjoinoneitem

				FItemList(i).fjoindate = db3_rsget("joindate")
				FItemList(i).fjoinPath_count = db3_rsget("joinPath_count")
				FItemList(i).fjoinPath = db3_rsget("joinPath")

				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	end Sub

	' 회원가입현황 그래프용(채널별)
	' admin/userjoin/userjoin.asp
	public sub fuserjoin_channel()
		dim sqlstr, i, sqlsearch

		if frectjoinSex <> "" then
			sqlsearch = sqlsearch & " and l.joinSex = '"& frectjoinSex &"'"
		end if
		if frectjoinPath <> "" then
			sqlsearch = sqlsearch & " and (case when l.joinPath='' then '10X10' else l.joinPath end) = '" & frectjoinPath & "'"
		end if
		if FRectStartdate<>"" and FRectEndDate<>"" then
			sqlsearch = sqlsearch & " and convert(varchar(10),l.joinDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"
		end if

        sqlstr = "select" & vbcrlf
		sqlstr = sqlstr & " (case when l.joinPath = '' then '10x10' else l.joinPath end) as joinPath" & vbcrlf
		sqlstr = sqlstr & " , sum(l.joincount) as joinPathCount" & vbcrlf
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_join_log l" & vbcrlf
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " group by (case when joinPath = '' then '10x10' else joinPath End)" & vbcrlf
		sqlstr = sqlstr & " order by joinPathCount desc, (case when l.joinPath = '' then '10x10' else l.joinPath end) asc" & vbcrlf

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"

		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0

		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				set FItemList(i) = new cuserjoinoneitem 			'클래스를 넣고

					FItemList(i).fjoinPath = db3_rsget("joinPath")
					FItemList(i).fjoinPathcount = db3_rsget("joinPathCount")

				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
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

    Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	end sub
	Private Sub Class_Terminate()
	End Sub
end class
%>