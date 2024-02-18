<%
'###########################################################
' Description : 데이터분석 영업이슈 클래스
' History : 2016.01.29 한용민 생성
'###########################################################

class cdataanalysis_salesissue_oneitem
	public fsalesidx
	public fdepartment_id
	public fstartdate
	public fenddate
	public ftitle
	public fcomment
	public freguserid
	public fregdate
	public fisusing
	public fdepartmentNameFull
	public fusername

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class

Class cdataanalysis_salesissue
	Public FItemList()
	public foneitem
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	public FPageCount

	public ftendb
	public frectisusing
	public frecttitle
	public frectsalesidx
	public frectdepartment_id
	public frectstartdate
	public frectenddate

	'//admin/dataanalysis/manager/salesissue/salesissue_edit.asp
	public sub getdataanalysis_salesissue_oneitem()
        dim sqlStr, sqlsearch

		if frectsalesidx<>"" then
			sqlsearch = sqlsearch & " and s.salesidx = "& frectsalesidx &"" + vbcrlf
		end if

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " s.salesidx, s.department_id, convert(varchar(19),s.startdate,121) as startdate, convert(varchar(19),s.enddate,121) as enddate" + vbcrlf
		sqlStr = sqlStr & " , s.title, s.comment, s.reguserid, s.regdate, s.isusing" + vbcrlf
		sqlStr = sqlStr & " , t.username" + vbcrlf
		sqlStr = sqlStr & " from db_analyze.dbo.tbl_analysis_salesissue s" + vbcrlf

		IF application("Svr_Info")="Dev" THEN
			sqlStr = sqlStr & " left join TENDB.db_partner.dbo.tbl_user_tenbyten t" + vbcrlf
			sqlStr = sqlStr & " 	on s.reguserid=t.userid" + vbcrlf
		else
			sqlStr = sqlStr & " left join db_analyze_data_raw.dbo.tbl_user_tenbyten t" + vbcrlf
			sqlStr = sqlStr & " 	on s.reguserid=t.userid" + vbcrlf
		end if

		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsAnalget.Open SqlStr, dbanalget, 1
        FResultCount = rsAnalget.RecordCount
        FtotalCount = rsAnalget.RecordCount

        set FOneItem = new cdataanalysis_salesissue_oneitem

        if Not rsAnalget.Eof then

			FOneItem.fsalesidx = rsAnalget("salesidx")
			FOneItem.fdepartment_id = rsAnalget("department_id")
			FOneItem.fstartdate = rsAnalget("startdate")
			FOneItem.fenddate = rsAnalget("enddate")
			FOneItem.ftitle = db2html(rsAnalget("title"))
			FOneItem.fcomment = db2html(rsAnalget("comment"))
			FOneItem.freguserid = rsAnalget("reguserid")
			FOneItem.fregdate = rsAnalget("regdate")
			FOneItem.fisusing = rsAnalget("isusing")
			FOneItem.fusername = db2html(rsAnalget("username"))

        end if
        rsAnalget.Close
    end Sub

	'//admin/dataanalysis/salesissue/salesissue.asp
	public sub getdataanalysis_salesissue_list()
		dim sqlStr,i, sqlsearch, sqldb

		if frectisusing<>"" then
			sqlsearch = sqlsearch & " and s.isusing = '"& frectisusing &"'" + vbcrlf
		end if
		if frecttitle<>"" then
			sqlsearch = sqlsearch & " and s.title like '%"& frecttitle &"%'" + vbcrlf
		end if
		if frectdepartment_id<>"" then
			sqlsearch = sqlsearch & " and s.department_id = "& frectdepartment_id &"" + vbcrlf
		end if
		if frectstartdate<>"" and frectenddate<>"" then
			'/날짜단위 : YYYY
			if len(frectstartdate)=4 then
				sqlsearch = sqlsearch & " and (" + vbcrlf
				sqlsearch = sqlsearch & " 	('"& frectstartdate &"' between convert(varchar(4),s.startdate,121) and convert(varchar(4),s.enddate,121))" + vbcrlf
				sqlsearch = sqlsearch & " 	or ('"& frectenddate &"' between convert(varchar(4),s.startdate,121) and convert(varchar(4),s.enddate,121))" + vbcrlf
				sqlsearch = sqlsearch & " )" + vbcrlf

			'/날짜단위 : YYYY-MM
			elseif len(frectstartdate)=7 then
				sqlsearch = sqlsearch & " and (" + vbcrlf
				sqlsearch = sqlsearch & " 	('"& frectstartdate &"' between convert(varchar(7),s.startdate,121) and convert(varchar(7),s.enddate,121))" + vbcrlf
				sqlsearch = sqlsearch & " 	or ('"& frectenddate &"' between convert(varchar(7),s.startdate,121) and convert(varchar(7),s.enddate,121))" + vbcrlf
				sqlsearch = sqlsearch & " )" + vbcrlf

			'/날짜단위 : YYYY-MM-DD
			else
				sqlsearch = sqlsearch & " and (" + vbcrlf
				sqlsearch = sqlsearch & " 	('"& frectstartdate &"' between s.startdate and s.enddate)" + vbcrlf
				sqlsearch = sqlsearch & " 	or ('"& frectenddate &"' between s.startdate and s.enddate)" + vbcrlf
				sqlsearch = sqlsearch & " )" + vbcrlf
			end if
		end if

		sqldb = sqldb & " from db_analyze.dbo.tbl_analysis_salesissue s" + vbcrlf
		IF application("Svr_Info")="Dev" THEN
			sqldb = sqldb & " left join TENDB.db_partner.dbo.tbl_user_tenbyten t" + vbcrlf
			sqldb = sqldb & " 	on s.reguserid=t.userid" + vbcrlf
			sqldb = sqldb & " left join TENDB.db_partner.dbo.vw_user_department dv" + vbcrlf
			sqldb = sqldb & " 	on s.department_id=dv.cid" + vbcrlf
			sqldb = sqldb & " 	and dv.useYN = 'Y'" + vbcrlf
		else
			sqldb = sqldb & " left join db_analyze_data_raw.dbo.tbl_user_tenbyten t" + vbcrlf
			sqldb = sqldb & " 	on s.reguserid=t.userid" + vbcrlf
			sqldb = sqldb & " left join db_analyze_data_raw.dbo.vw_user_department dv" + vbcrlf
			sqldb = sqldb & " 	on s.department_id=dv.cid" + vbcrlf
			sqldb = sqldb & " 	and dv.useYN = 'Y'" + vbcrlf
		end if

		sqlStr = "select count(s.salesidx) as cnt" + vbcrlf
		sqlStr = sqlStr & sqldb
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr & "<br>"
		rsAnalget.Open sqlStr,dbanalget,1
			FTotalCount = rsAnalget("cnt")
		rsAnalget.Close

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " s.salesidx, s.department_id, s.startdate, s.enddate, s.title, s.comment, s.reguserid, s.regdate, s.isusing" + vbcrlf
		sqlStr = sqlStr & " , t.username, dv.departmentNameFull" + vbcrlf
		sqlStr = sqlStr & sqldb
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by s.salesidx Desc" + vbcrlf

		'response.write sqlStr & "<br>"
		rsAnalget.pagesize = FPageSize
		rsAnalget.Open sqlStr,dbanalget,1

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
		if  not rsAnalget.EOF  then
			rsAnalget.absolutepage = FCurrPage
			do until rsAnalget.EOF
				set FItemList(i) = new cdataanalysis_salesissue_oneitem

				FItemList(i).fsalesidx = rsAnalget("salesidx")
				FItemList(i).fdepartment_id = rsAnalget("department_id")
				FItemList(i).fstartdate = rsAnalget("startdate")
				FItemList(i).fenddate = rsAnalget("enddate")
				FItemList(i).ftitle = db2html(rsAnalget("title"))
				FItemList(i).fcomment = db2html(rsAnalget("comment"))
				FItemList(i).freguserid = rsAnalget("reguserid")
				FItemList(i).fregdate = rsAnalget("regdate")
				FItemList(i).fisusing = rsAnalget("isusing")
				FItemList(i).fusername = db2html(rsAnalget("username"))
				FItemList(i).fdepartmentNameFull = db2html(rsAnalget("departmentNameFull"))

				rsAnalget.movenext
				i=i+1
			loop
		end if
		rsAnalget.Close
	end sub

	'//admin/dataanalysis/salesissue/salesissue.asp		'/admin/dataanalysis/mkt.asp
	public sub getdataanalysis_salesissue_top()
		dim sqlStr,i, sqlsearch, sqldb

		if frectisusing<>"" then
			sqlsearch = sqlsearch & " and s.isusing = '"& frectisusing &"'" + vbcrlf
		end if
		if frecttitle<>"" then
			sqlsearch = sqlsearch & " and s.title like '%"& frecttitle &"%'" + vbcrlf
		end if
		if frectdepartment_id<>"" then
			sqlsearch = sqlsearch & " and s.department_id = "& frectdepartment_id &"" + vbcrlf
		end if
		if frectstartdate<>"" and frectenddate<>"" then
			sqlsearch = sqlsearch & " and (" + vbcrlf
			sqlsearch = sqlsearch & " 	('"& frectstartdate &"' between s.startdate and s.enddate)" + vbcrlf
			sqlsearch = sqlsearch & " 	or ('"& frectenddate &"' between s.startdate and s.enddate)" + vbcrlf
			sqlsearch = sqlsearch & " )" + vbcrlf
		end if

		sqldb = sqldb & " from db_analyze.dbo.tbl_analysis_salesissue s" + vbcrlf

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " s.salesidx, s.department_id, s.startdate, s.enddate, s.title, s.comment, s.reguserid, s.regdate, s.isusing" + vbcrlf
		sqlStr = sqlStr & sqldb
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by s.salesidx Desc" + vbcrlf

		'response.write sqlStr & "<br>"
		rsAnalget.pagesize = FPageSize
		rsAnalget.Open sqlStr,dbanalget,1

		FTotalCount = rsAnalget.RecordCount
		FResultCount = rsAnalget.RecordCount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsAnalget.EOF  then
			rsAnalget.absolutepage = FCurrPage
			do until rsAnalget.EOF
				set FItemList(i) = new cdataanalysis_salesissue_oneitem

				FItemList(i).fsalesidx = rsAnalget("salesidx")
				FItemList(i).fdepartment_id = rsAnalget("department_id")
				FItemList(i).fstartdate = rsAnalget("startdate")
				FItemList(i).fenddate = rsAnalget("enddate")
				FItemList(i).ftitle = db2html(rsAnalget("title"))
				FItemList(i).fcomment = db2html(rsAnalget("comment"))
				FItemList(i).freguserid = rsAnalget("reguserid")
				FItemList(i).fregdate = rsAnalget("regdate")
				FItemList(i).fisusing = rsAnalget("isusing")

				rsAnalget.movenext
				i=i+1
			loop
		end if
		rsAnalget.Close
	end sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function
	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function
	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FScrollCount = 10

		IF application("Svr_Info")="Dev" THEN
			ftendb="TENDB."
		end if
	End Sub
	Private Sub Class_Terminate()
    End Sub
End Class
%>