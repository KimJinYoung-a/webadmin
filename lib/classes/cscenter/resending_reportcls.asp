<%
Class clsCSList
    public FID
    public Fdivcd
    public Fgubun01
    public Fgubun02

    public Fdivcd_Name
    public Fgubun01_Name
    public Fgubun02_Name

    public Forderserial
    public Ftitle

    public FRegdate
    public Freguserid
    public Ffinishdate
	public Ffinishuser

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

class CReportMasterItemList
    public FDivcd
    public FDivcdName
    public Fgubun01
    public Fgubun02
    public Fgubun01Name
    public Fgubun02Name
 	public Fcount

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CReportMaster
	public FMasterItemList()
	public FRectStart
	public FRectEnd
	public FRectDivcd

	public FRectFinishUser
    public FRectRegUserID
    public FRectRegStart
    public FRectRegEnd
	public FRectJumunSite

    public FItemList()
	public FOneItem

    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		redim  FMasterItemList(0)
		FResultCount = 0
		redim  FItemList(0)
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public function getCSReport()
		Dim strSql
		strSql = " db_datamart.dbo.sp_Ten_CS_Report_AS ('" & FRectStart & "','" & FRectEnd & "','" & FRectDivcd & "','" & FRectJumunSite & "')"

		'response.write strSql & "<br>"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

		Dim rs
		If Not db3_rsget.EOF Then
			rs = db3_rsget.getRows()
		End If
		db3_rsget.close

		getCSReport = rs

	End Function


    public Sub getCsListView2(ByVal finishDate1, ByVal finishDate2, ByVal divCd, ByVal gubun01, ByVal gubun02)
        dim sqlStr, i

		Dim paramInfo
		Dim strSQL, sqlColumn, sqlTable, sqlWhere, sqlOrder, sqlGroup	' 쿼리문 변수 선언

		sqlWhere = " and deleteyn='N' "
		sqlWhere = sqlWhere + " and currState='B007' "

        if (finishDate1<>"" and not(isnull(finishDate1))) then
            sqlWhere = sqlWhere + " and a.finishdate>= ? "
			Call redimParam(paramInfo, "@finishDate1"		, adVarchar	, adParamInput	, 10		, finishDate1)
        end if

        if (finishDate2<>"" and not(isnull(finishDate2))) then
            sqlWhere = sqlWhere + " and a.finishdate< ? "
			Call redimParam(paramInfo, "@finishDate2"		, adVarchar	, adParamInput	, 10		, dateAdd("d",1,finishDate2))
        end if

        if (FRectRegStart<>"" and not(isnull(FRectRegStart))) then
            sqlWhere = sqlWhere + " and a.regdate>= ? "
			Call redimParam(paramInfo, "@regdate1"		, adVarchar	, adParamInput	, 10		, FRectRegStart)
        end if

        if (FRectRegEnd<>"" and not(isnull(FRectRegEnd))) then
            sqlWhere = sqlWhere + " and a.regdate< ? "
			Call redimParam(paramInfo, "@regdate2"		, adVarchar	, adParamInput	, 10		, dateAdd("d",1,FRectRegEnd))
        end if

        if (divCd<>"") then
            if (divCd="") then
                sqlWhere = sqlWhere + " and a.divcd in ('A000','A001','A002')"
            else
                sqlWhere = sqlWhere + " and a.divcd= ? "
				Call redimParam(paramInfo, "@divCd"		, adVarchar	, adParamInput	, 4		, divCd)
            end if
        end if

        if (gubun01<>"") then
            sqlWhere = sqlWhere + " and a.gubun01 = ? "
			Call redimParam(paramInfo, "@gubun01"		, adVarchar	, adParamInput	, 4		, gubun01)
        else
            sqlWhere = sqlWhere + " and a.gubun01 IN ('C004','C005','C006','C007') "
        end if
        if (gubun02<>"") then
            sqlWhere = sqlWhere + " and a.gubun02 = ? "
			Call redimParam(paramInfo, "@gubun02"		, adVarchar	, adParamInput	, 4		, gubun02)
        end if

        if (FRectFinishUser<>"") then
            sqlWhere = sqlWhere + " and a.finishuser = ? "
			Call redimParam(paramInfo, "@finishuser"		, adVarchar	, adParamInput	, 32		, FRectFinishUser)
        end if

        if (FRectRegUserID<>"") then
            sqlWhere = sqlWhere + " and a.writeuser = ? "
			Call redimParam(paramInfo, "@writeuser"		, adVarchar	, adParamInput	, 32		, FRectRegUserID)
        end if

		' 쿼리문 조합용 변수 설정
		sqlColumn	= " a.id, a.orderserial, a.finishdate, a.title, a.finishuser "
        sqlColumn	= sqlColumn & " , db_cs.dbo.uf_getCodeName('Z001', a.divCd) divcd_Name "
        sqlColumn	= sqlColumn & " , db_cs.dbo.uf_getCodeName('Z020', a.gubun01) gubun01_Name "
        sqlColumn	= sqlColumn & " , db_cs.dbo.uf_getCodeName(a.gubun01, a.gubun02) gubun02_Name "
		''sqlColumn	= sqlColumn & " , IsNull(m.sitename, '10x10') as extsitename "
        sqlColumn	= sqlColumn & " , a.regdate "
        sqlColumn	= sqlColumn & " , a.writeuser as reguserid "
        sqlTable	= " from db_cs.dbo.tbl_new_as_list a"
		''sqlTable	= " Left Join [db_order].[dbo].tbl_order_master m on A.orderserial=m.orderserial "
        sqlOrder	= " order by a.id"

		strSQL = makeQuery("", sqlTable, sqlWhere, sqlOrder, "", "", "")	' 카운트 쿼리
		Call RecordSQL(strSQL, paramInfo)

		If Not rsget.EOF Then
			FTotalCount = rsget(0)
		End If
		rsget.Close

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		strSQL = makeQuery(sqlColumn, sqlTable, sqlWhere, sqlOrder, FCurrPage, FPageSize, "")	' 페이징 쿼리
		Call RecordSQL(strSQL, paramInfo)


		i=0
		if  not rsget.EOF  then
			do until rsget.eof

				redim preserve FItemList(i)
			    set FItemList(i) = new clsCSList
			    FItemList(i).FID            = rsget("id")
			    FItemList(i).Fdivcd_Name    = rsget("divcd_Name")
			    FItemList(i).Fgubun01_Name  = rsget("gubun01_Name")
			    FItemList(i).Fgubun02_Name  = rsget("gubun02_Name")

			    FItemList(i).Forderserial  = rsget("orderserial")
			    FItemList(i).Ftitle			= rsget("title")
			    FItemList(i).Ffinishdate    = rsget("finishdate")
				FItemList(i).Ffinishuser    = rsget("finishuser")
                FItemList(i).Fregdate    	= rsget("regdate")
                FItemList(i).Freguserid    	= rsget("reguserid")

			    i=i+1
				rsget.moveNext
			loop
		end If
		FREsultCount = i
		rsget.Close
    end Sub


    public Sub getCsListView(ByVal finishDate1, ByVal finishDate2, ByVal divCd, ByVal gubun01, ByVal gubun02)
        dim sqlStr, i

		Dim strSearch

		strSearch = " and deleteyn='N' "
		strSearch = strSearch + " and currState='B007' "

        if (finishDate1<>"") then
            strSearch = strSearch + " and a.finishdate>='" + finishDate1 + "'"
        end if

        if (finishDate2<>"") then
            strSearch = strSearch + " and a.finishdate<'" & dateAdd("d",1,finishDate2) & "'"
        end if

        if (divCd<>"") then
            if (divCd="") then
                strSearch = strSearch + " and a.divcd in ('A000','A001','A002')"
            else
                strSearch = strSearch + " and a.divcd='" + divCd + "'"
            end if
        end if

        if (gubun01<>"") then
            strSearch = strSearch + " and a.gubun01 = '" + gubun01 + "'"
        else
            strSearch = strSearch + " and a.gubun01 IN ('C004','C005','C006','C007') "
        end if
        if (gubun02<>"") then
            strSearch = strSearch + " and a.gubun02 = '" + gubun02 + "'"
        end if

        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr + " from db_cs.dbo.tbl_new_as_list a"
        sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + strSearch


        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

        '' too slow;;
        sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " a.id, a.orderserial, a.finishdate, a.title "
        sqlStr = sqlStr + " , db_cs.dbo.uf_getCodeName('Z001', a.divCd) divcd_Name "
        sqlStr = sqlStr + " , db_cs.dbo.uf_getCodeName('Z020', a.gubun01) gubun01_Name "
        sqlStr = sqlStr + " , db_cs.dbo.uf_getCodeName(a.gubun01, a.gubun02) gubun02_Name "
        sqlStr = sqlStr + " from db_cs.dbo.tbl_new_as_list a"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + strSearch

        sqlStr = sqlStr + " order by a.id"

        'rw sqlStr

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    rsget.absolutepage = FCurrPage
			do until rsget.eof

			    set FItemList(i) = new clsCSList
			    FItemList(i).FID            = rsget("id")
			    FItemList(i).Fdivcd_Name    = rsget("divcd_Name")
			    FItemList(i).Fgubun01_Name  = rsget("gubun01_Name")
			    FItemList(i).Fgubun02_Name  = rsget("gubun02_Name")

			    FItemList(i).Forderserial  = rsget("orderserial")
			    FItemList(i).Ftitle			= rsget("title")
			    FItemList(i).Ffinishdate    = rsget("finishdate")
			    i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	public Sub SearchReport()
        dim sql,i


		sql = "select  a.divcd, c1.comm_name as divcdName, count(a.id) as count "
		sql = sql + " from [db_cs].[dbo].tbl_new_as_list a" + vbcrlf
		sql = sql + "   left join [db_cs].[dbo].tbl_cs_comm_code c1" + vbcrlf
        sql = sql + "   on a.divcd=c1.comm_cd" + vbcrlf
		sql = sql + " where a.regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
		sql = sql + " and a.regdate < '" + Cstr(FRectEnd) + "'" + vbcrlf
		sql = sql + " and a.deleteyn = 'N'" + vbcrlf
		sql = sql + " group by a.divcd, c1.comm_name" + vbcrlf
		sql = sql + " order by a.divcd"

		rsget.Open sql,dbget,1

		FResultCount = rsget.recordcount
		redim preserve FMasterItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    do until rsget.EOF
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fdivcd       = rsget("divcd")
				FMasterItemList(i).FdivcdName   = db2html(rsget("divcdName"))
				FMasterItemList(i).Fcount       = rsget("count")

				rsget.movenext
				i=i+1
		    loop
		end if
		rsget.close
    end Sub

	public Sub SearchReportByGubun()

	    dim sql,i


		sql = "select  a.gubun01, c1.comm_name as gubun01Name, a.gubun02, c2.comm_name as gubun02Name, count(a.divcd) as count "
		sql = sql + " from [db_cs].[dbo].tbl_new_as_list a" + vbcrlf
		sql = sql + "   left join [db_cs].[dbo].tbl_cs_comm_code c1" + vbcrlf
        sql = sql + "   on a.gubun01=c1.comm_cd" + vbcrlf
        sql = sql + "   left join [db_cs].[dbo].tbl_cs_comm_code c2" + vbcrlf
        sql = sql + "   on a.gubun02=c2.comm_cd" + vbcrlf
		sql = sql + " where a.regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
		sql = sql + " and a.regdate < '" + Cstr(FRectEnd) + "'" + vbcrlf
		sql = sql + " and a.divcd = '" & FRectDivcd & "'" + vbcrlf
		sql = sql + " and a.deleteyn = 'N'" + vbcrlf
		sql = sql + " group by a.gubun01, c1.comm_name, a.gubun02, c2.comm_name" + vbcrlf
		sql = sql + " order by a.gubun01, a.gubun02"

		rsget.Open sql,dbget,1

		FResultCount = rsget.recordcount
		redim preserve FMasterItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    do until rsget.EOF
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fgubun01 = rsget("gubun01")
				FMasterItemList(i).Fgubun02 = rsget("gubun02")
				FMasterItemList(i).Fgubun01name = db2html(rsget("gubun01Name"))
				FMasterItemList(i).Fgubun02name = db2html(rsget("gubun02Name"))
				FMasterItemList(i).Fcount = rsget("count")

				rsget.movenext
				i=i+1
		    loop
		end if
		rsget.close

	end sub

	public Sub SearchObaesongReport()

	    dim sql,i

'		sql = "select count(divcd) as count from [db_cs].[dbo].tbl_as_list" + vbcrlf
'		sql = sql + " where deleteyn = 'N'" + vbcrlf
'		sql = sql + " and regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
'		sql = sql + " and regdate < '" + Cstr(FRectEnd) + "'" + vbcrlf
'		sql = sql + " and causedetail in ('오발송','상품불량','상품파손','상품누락')" + vbcrlf
'
'		rsget.Open sql,dbget,1
'
'		if  not rsget.EOF  then
'			Ftotalcount = rsget("count")
'		end if
'		rsget.close


		sql = "select  a.gubun01, c1.comm_name as gubun01Name, a.gubun02, c2.comm_name as gubun02Name, count(a.divcd) as count "
		sql = sql + " from [db_cs].[dbo].tbl_new_as_list a" + vbcrlf
		sql = sql + "   left join [db_cs].[dbo].tbl_cs_comm_code c1" + vbcrlf
        sql = sql + "   on a.gubun01=c1.comm_cd" + vbcrlf
        sql = sql + "   left join [db_cs].[dbo].tbl_cs_comm_code c2" + vbcrlf
        sql = sql + "   on a.gubun02=c2.comm_cd" + vbcrlf
		sql = sql + " where a.regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
		sql = sql + " and a.regdate < '" + Cstr(FRectEnd) + "'" + vbcrlf
		sql = sql + " and a.gubun01 in ('C006')" + vbcrlf
		sql = sql + " and a.deleteyn = 'N'" + vbcrlf
		sql = sql + " group by a.gubun01, c1.comm_name, a.gubun02, c2.comm_name" + vbcrlf
		sql = sql + " order by a.gubun01, a.gubun02"

		rsget.Open sql,dbget,1

		FResultCount = rsget.recordcount
		redim preserve FMasterItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    do until rsget.EOF
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fgubun01 = rsget("gubun01")
				FMasterItemList(i).Fgubun02 = rsget("gubun02")
				FMasterItemList(i).Fgubun01name = db2html(rsget("gubun01Name"))
				FMasterItemList(i).Fgubun02name = db2html(rsget("gubun02Name"))
				FMasterItemList(i).Fcount = rsget("count")

				rsget.movenext
				i=i+1
		    loop
		end if
		rsget.close

	end sub

end class

%>
