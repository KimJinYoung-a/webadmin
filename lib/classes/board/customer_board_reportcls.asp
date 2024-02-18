<%
'###########################################################
' Description : 1:1 상담
' History : 2015.05.27 이상구 생성
'			2016.03.25 한용민 수정
'###########################################################
%>
<%
class CReportMasterItemList
    public Fqadiv
 	public Fcount
    public FqadivName

    public function GetQadivName()
        if Fqadiv="00" then
            GetQadivName = "배송문의"
        elseif Fqadiv="01" then
            GetQadivName = "주문문의"
        elseif Fqadiv="02" then
            GetQadivName = "상품문의"
        elseif Fqadiv="03" then
            GetQadivName = "재고문의"
        elseif Fqadiv="04" then
            GetQadivName = "취소문의"
        elseif Fqadiv="05" then
            GetQadivName = "환불문의"
        elseif Fqadiv="06" then
            GetQadivName = "교환문의"
        elseif Fqadiv="07" then
            GetQadivName = "As문의"
        elseif Fqadiv="08" then
            GetQadivName = "이벤트문의"
        elseif Fqadiv="09" then
            GetQadivName = "증빙서류문의"
        elseif Fqadiv="10" then
            GetQadivName = "시스템문의"
        elseif Fqadiv="11" then
            GetQadivName = "회원제도문의"
        elseif Fqadiv="12" then
            GetQadivName = "개인정보관련"
        elseif Fqadiv="13" then
            GetQadivName = "당첨문의"
        elseif Fqadiv="14" then
            GetQadivName = "반품문의"
        elseif Fqadiv="15" then
            GetQadivName = "입금문의"
        elseif Fqadiv="16" then
            GetQadivName = "오프라인문의"
        elseif Fqadiv="17" then
            GetQadivName = "쿠폰/마일리지문의"
        elseif Fqadiv="18" then
            GetQadivName = "결제방법문의"

        elseif Fqadiv="20" then
            GetQadivName = "기타문의"
        else
            GetQadivName = Fqadiv
        end if
    end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CReportEvalItem
    public Freplydate
	public Freplyuser
	public FtotCnt
	public FtotEvalCnt
	public FevalCnt5
	public FevalCnt4
	public FevalCnt3
	public FevalCnt2
	public FevalCnt1
	public FnoEvalCnt
	public FevalSum
	public fd0
	public fd1
	public fd2
	public fd3
	public fd4

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CReportMaster
	public FMasterItemList()

    public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
	public FPageCount

	public FRectStart
	public FRectEnd
	public FRectReplyUser
	public FRectGroupByReplyUser
	public fTENDB
	public FRectSiteGubun
	public FRectuserlevel
	public FRectsitename

	public Sub SearchReport()
	    dim sql,i

'		sql = "select count(qadiv) as count from [db_cs].[dbo].tbl_myqna" + vbcrlf
'		sql = sql + " where regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
'		sql = sql + " and regdate < '" + Cstr(FRectEnd) + "'"
'
'		rsget.Open sql,dbget,1
'
'		if  not rsget.EOF  then
'			Ftotalcount = rsget("count")
'		end if
'		rsget.close

		sql = "select c.qadiv, c.qadivname, count(q.id) as count "
		sql = sql + " from [db_cs].[dbo].tbl_myqna q" + vbcrlf
		sql = sql + "   left join [db_cs].[dbo].tbl_myqna_comm_code c"
		sql = sql + "   on q.qadiv=c.qadiv"
		sql = sql + " where q.regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
		sql = sql + " and q.regdate < '" + Cstr(FRectEnd) + "'" + vbcrlf
		sql = sql + " and q.isusing='Y'"
		sql = sql + " and q.dispyn='Y'"
		sql = sql + " group by all c.qadiv, c.qadivname" + vbcrlf
		sql = sql + " order by c.qadiv asc"

		rsget.Open sql,dbget,1

		FResultCount = rsget.recordcount
		redim preserve FMasterItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    do until rsget.EOF
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fqadiv = rsget("qadiv")
				FMasterItemList(i).Fcount = rsget("count")
				FMasterItemList(i).FqadivName = rsget("qadivname")
				rsget.movenext
				i=i+1
		    loop
		end if
		rsget.close
	end sub

	public function getQnaDivName()
		Dim i, strSql

		strSql = " SELECT qaDivName, qaDiv FROM db_datamart.dbo.tbl_myqna_comm_code "
		strSql = strSql & " ORDER BY qaDiv	" & vbCrLf

		db3_rsget.open strSql,db3_dbget,1

		Dim rs
		If Not db3_rsget.EOF Then
			rs = db3_rsget.getRows()
		End If
		db3_rsget.close

		getQnaDivName = rs
	end function

	'/문의분야 카운트	'/2016.03.28 한용민 생성
	public function getQnaDivcount()
		Dim i, strSql, tmpcnt
		tmpcnt=0

		strSql = "select" & vbcrlf
		strSql = strSql & " count(c.comm_cd) as cnt" & vbcrlf
		strSql = strSql & " from "& fTENDB &"[db_cs].[dbo].[tbl_cs_comm_code] c" & vbcrlf
		strSql = strSql & " where comm_isdel='N'" & vbcrlf
		'strSql = strSql & " and dispyn='Y'" & vbcrlf
		strSql = strSql & " and left(comm_group,3)='D00'" & vbcrlf

		'response.write strSql & "<br>"
		db3_rsget.Open strSql,db3_dbget,1

		if not db3_rsget.EOF then
			tmpcnt = db3_rsget("cnt")
		End If
		db3_rsget.close

		getQnaDivcount = tmpcnt
	end function

	'/문의분야 DB화 시킴	'/2016.03.28 한용민 생성
	public function getQnaReport(ByVal userID)
		Dim i, strSql

        fTENDB = "[TENDB]."  ''2018/03/09 추가. db_cs 복원중 오류
		if userID<>"" then
			strSql = "SELECT" & vbcrlf
			strSql = strSql & " b.yyyymmdd, right(a.comm_cd,2) as qaDiv, isNull(c.cnt,0) cnt" & vbcrlf
			strSql = strSql & " , SUBSTRING(a.comm_name, CHARINDEX('!@#',a.comm_name)+3,255) as qaDivName, a.dispyn" & vbcrlf
			strSql = strSql & " FROM "& fTENDB &"[db_cs].[dbo].[tbl_cs_comm_code] a" & vbcrlf
			strSql = strSql & " INNER JOIN (" & vbcrlf
			strSql = strSql & " 	SELECT yyyymmdd FROM db_datamart.dbo.tbl_cs_daily_qna_summary" & vbcrlf
			strSql = strSql & " 	WHERE yyyymmdd between '" & FRectStart & "' and '" & FRectEnd & "'" & vbcrlf
			strSql = strSql & " 	AND userID = '" & userID & "'" & vbcrlf
			strSql = strSql & " 	GROUP BY yyyymmdd" & vbcrlf
			strSql = strSql & " ) b ON 1=1" & vbcrlf
			strSql = strSql & " LEFT OUTER JOIN (" & vbcrlf
			strSql = strSql & " 	SELECT yyyymmdd, qnaDiv" & vbcrlf
            if (FRectSiteGubun = "10x10") then
                strSql = strSql & " 	, Sum(cnt10x10) cnt" & vbcrlf
            elseif (FRectSiteGubun = "extall") then
                strSql = strSql & " 	, Sum(cntExtAll) cnt" & vbcrlf
            else
                strSql = strSql & " 	, Sum(cnt) cnt" & vbcrlf
            end if
			strSql = strSql & " 	FROM db_datamart.dbo.tbl_cs_daily_qna_summary" & vbcrlf
			strSql = strSql & " 	WHERE yyyymmdd between '" & FRectStart & "' and '" & FRectEnd & "'" & vbcrlf
			strSql = strSql & " 	AND userID = '" & userID & "'" & vbcrlf
			if (FRectSiteGubun = "10x10") then
				strSql = strSql & " 	AND cnt10x10 > 0 " & vbcrlf
			elseif (FRectSiteGubun = "extall") then
				strSql = strSql & " 	AND cntExtAll > 0 " & vbcrlf
			end if
			strSql = strSql & " 	GROUP BY yyyymmdd, qnaDiv" & vbcrlf
			strSql = strSql & " ) c" & vbcrlf
			strSql = strSql & " 	ON right(a.comm_cd,2) = c.qnaDiv AND b.yyyymmdd = c.yyyymmdd" & vbcrlf
			strSql = strSql & "where comm_isdel='N'" & vbcrlf
			'strSql = strSql & "and dispyn='Y'" & vbcrlf
			strSql = strSql & "and left(comm_group,3)='D00'" & vbcrlf
			strSql = strSql & "ORDER BY b.yyyymmdd asc, a.sortno asc, qaDiv asc"
		else
			strSql = strSql & "SELECT" & vbcrlf
			strSql = strSql & "b.userID, right(a.comm_cd,2) as qaDiv, isNull(c.cnt,0) cnt" & vbcrlf
			strSql = strSql & ", SUBSTRING(a.comm_name, CHARINDEX('!@#',a.comm_name)+3,255) as qaDivName, a.dispyn" & vbcrlf
			strSql = strSql & "FROM "& fTENDB &"[db_cs].[dbo].[tbl_cs_comm_code] a" & vbcrlf
			strSql = strSql & "INNER JOIN (" & vbcrlf
			strSql = strSql & "		SELECT userID FROM db_datamart.dbo.tbl_cs_daily_qna_summary" & vbcrlf
			strSql = strSql & "		WHERE yyyymmdd between '" & FRectStart & "' and '" & FRectEnd & "'" & vbcrlf
			strSql = strSql & "		GROUP BY userID" & vbcrlf
			strSql = strSql & ") b ON 1=1" & vbcrlf
			strSql = strSql & "LEFT OUTER JOIN (" & vbcrlf
			strSql = strSql & "		SELECT userID, qnaDiv" & vbcrlf
            if (FRectSiteGubun = "10x10") then
                strSql = strSql & " 	, Sum(cnt10x10) cnt" & vbcrlf
            elseif (FRectSiteGubun = "extall") then
                strSql = strSql & " 	, Sum(cntExtAll) cnt" & vbcrlf
            else
                strSql = strSql & " 	, Sum(cnt) cnt" & vbcrlf
            end if
			strSql = strSql & "		FROM db_datamart.dbo.tbl_cs_daily_qna_summary" & vbcrlf
			strSql = strSql & "		WHERE yyyymmdd between '" & FRectStart & "' and '" & FRectEnd & "'" & vbcrlf
			if (FRectSiteGubun = "10x10") then
				strSql = strSql & " 	AND cnt10x10 > 0 " & vbcrlf
			elseif (FRectSiteGubun = "extall") then
				strSql = strSql & " 	AND cntExtAll > 0 " & vbcrlf
			end if
			strSql = strSql & "		GROUP BY userID, qnaDiv" & vbcrlf
			strSql = strSql & ") c" & vbcrlf
			strSql = strSql & "		ON right(a.comm_cd,2) = c.qnaDiv AND b.userID = c.userID" & vbcrlf
			strSql = strSql & "where comm_isdel='N'" & vbcrlf
			'strSql = strSql & "and dispyn='Y'" & vbcrlf
			strSql = strSql & "and left(comm_group,3)='D00'" & vbcrlf
			strSql = strSql & "ORDER BY b.userID asc, a.sortno asc, qaDiv asc"
		end if

		'response.write strSql & "<br>"
		db3_rsget.Open strSql,db3_dbget,1

		Dim rs
		If Not db3_rsget.EOF Then
			rs = db3_rsget.getRows()
		End If
		db3_rsget.close

		getQnaReport = rs
	end function

	'/문의분야, 통계에서만 따로 쓰는 테이블로 통계냄.	'/사용안함
	public function getQnaDivReport(ByVal userID)
		Dim i, strSql

		strSql = " db_datamart.dbo.sp_Ten_CS_Report_Qna ('" & FRectStart & "','" & FRectEnd & "','" & userID & "')"

		'response.write strSql & "<br>"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

		Dim rs
		If Not db3_rsget.EOF Then
			rs = db3_rsget.getRows()
		End If
		db3_rsget.close

		getQnaDivReport = rs
	end function

	' /cscenter/board/customer_board_eval_report.asp
	public function getQnaEvalReport()
	    dim i,sqlStr, addSqlStr

		addSqlStr = ""

		addSqlStr = addSqlStr + " 1 = 1 "
		addSqlStr = addSqlStr + " and q.replydate is not NULL "
		addSqlStr = addSqlStr + " and q.replydate >= '" + CStr(FRectStart) + "' "
		addSqlStr = addSqlStr + " and q.replydate < '" + CStr(FRectEnd) + "' "

		if (FRectReplyUser <> "") then
			addSqlStr = addSqlStr + " and q.replyuser = '" + CStr(replyuser) + "' "
		end if

		if FRectuserlevel<>"" then
			addSqlStr = addSqlStr & " and q.userlevel='"& FRectuserlevel &"'"
		end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	convert(varchar(10), q.replydate, 127) as replydate "

		if (FRectGroupByReplyUser = "Y") then
			sqlStr = sqlStr + " 	, replyuser "
		else
			sqlStr = sqlStr + " 	, '' as replyuser "
		end if

		sqlStr = sqlStr + " 	, count(*) as totCnt "
		sqlStr = sqlStr + " 	, sum(case when IsNull(q.EvalPoint, 0) <> 0 then 1 else 0 end) as totEvalCnt "
		sqlStr = sqlStr + " 	, sum(case when IsNull(q.EvalPoint, 0) = 5 then 1 else 0 end) as evalCnt5 "
		sqlStr = sqlStr + " 	, sum(case when IsNull(q.EvalPoint, 0) = 4 then 1 else 0 end) as evalCnt4 "
		sqlStr = sqlStr + " 	, sum(case when IsNull(q.EvalPoint, 0) = 3 then 1 else 0 end) as evalCnt3 "
		sqlStr = sqlStr + " 	, sum(case when IsNull(q.EvalPoint, 0) = 2 then 1 else 0 end) as evalCnt2 "
		sqlStr = sqlStr + " 	, sum(case when IsNull(q.EvalPoint, 0) = 1 then 1 else 0 end) as evalCnt1 "
		sqlStr = sqlStr + " 	, sum(case when IsNull(q.EvalPoint, 0) = 0 then 1 else 0 end) as noEvalCnt "
		sqlStr = sqlStr + " 	, sum(IsNull(q.EvalPoint, 0)) as evalSum "
		sqlStr = sqlStr + " from db_cs.dbo.tbl_myqna q with (nolock)"
		sqlStr = sqlStr + " where "

		sqlStr = sqlStr + addSqlStr

		sqlStr = sqlStr + " group by "
		sqlStr = sqlStr + " 	convert(varchar(10), q.replydate, 127) "
		if (FRectGroupByReplyUser = "Y") then
			sqlStr = sqlStr + " 	, replyuser "
		end if

		sqlStr = sqlStr + " order by "
		if (FRectGroupByReplyUser = "Y") then
			sqlStr = sqlStr + " 	replyuser, "
		end if
		sqlStr = sqlStr + " 	convert(varchar(10), q.replydate, 127) "

		'response.write sqlStr & "<Br>"
	    rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CReportEvalItem

				FItemList(i).Freplydate			= rsget("replydate")
				FItemList(i).Freplyuser			= rsget("replyuser")
				FItemList(i).FtotCnt			= rsget("totCnt")
				FItemList(i).FtotEvalCnt		= rsget("totEvalCnt")
				FItemList(i).FevalCnt5			= rsget("evalCnt5")
				FItemList(i).FevalCnt4			= rsget("evalCnt4")
				FItemList(i).FevalCnt3			= rsget("evalCnt3")
				FItemList(i).FevalCnt2			= rsget("evalCnt2")
				FItemList(i).FevalCnt1			= rsget("evalCnt1")
				FItemList(i).FnoEvalCnt			= rsget("noEvalCnt")
				FItemList(i).FevalSum			= rsget("evalSum")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	' /cscenter/board/customer_board_sameday_report.asp
	public function getsameday_report()
	    dim i,sqlStr, addSqlStr

		if FRectStart<>"" and FRectEnd<>"" then
			addSqlStr = addSqlStr & " and q.replydate>='"& FRectStart &"'"
			addSqlStr = addSqlStr & " and q.replydate<'"& FRectEnd &"'"
		end if
		if FRectuserlevel<>"" then
			addSqlStr = addSqlStr & " and q.userlevel='"& FRectuserlevel &"'"
		end if
		if (FRectReplyUser <> "") then
			addSqlStr = addSqlStr & " and q.replyuser = '" & CStr(replyuser) & "' "
		end if
		if (FRectsitename = "10x10") then
			addSqlStr = addSqlStr & " and isnull(q.sitename,'10x10')='10x10'"
		elseif (FRectsitename = "10x10not") then
			addSqlStr = addSqlStr & " and isnull(q.sitename,'10x10')<>'10x10'"
		end if

		sqlStr = "select top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " t.replydate"
		if (FRectGroupByReplyUser = "Y") then
			sqlStr = sqlStr & " , t.replyuser"	
		end if
		sqlStr = sqlStr & " , sum(case when t.workdays=0 then 1 else 0 end) as d0"
		sqlStr = sqlStr & " , sum(case when t.workdays=1 then 1 else 0 end) as d1"
		sqlStr = sqlStr & " , sum(case when t.workdays=2 then 1 else 0 end) as d2"
		sqlStr = sqlStr & " , sum(case when t.workdays=3 then 1 else 0 end) as d3"
		sqlStr = sqlStr & " , sum(case when t.workdays>=4 then 1 else 0 end) as d4"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & " 	select q.id, q.replyuser, convert(nvarchar(10),q.regdate,121) as regdate"
		sqlStr = sqlStr & " 	, convert(nvarchar(10),q.replydate,121) as replydate"
		IF application("Svr_Info")="Dev" THEN
			sqlStr = sqlStr & " 	, [db_datamart].[dbo].[fn_Ten_WorkDays](convert(nvarchar(10),dateadd(hour,7,q.regdate),121),convert(nvarchar(10),dateadd(hour,7,q.replydate),121)) as workdays"
		else
			sqlStr = sqlStr & " 	, [db_sitemaster].[dbo].[fn_Ten_WorkDays](convert(nvarchar(10),dateadd(hour,7,q.regdate),121),convert(nvarchar(10),dateadd(hour,7,q.replydate),121)) as workdays"
		end if
		sqlStr = sqlStr & " 	from "&fTENDB&"db_cs.dbo.tbl_myqna q with (nolock)"
		sqlStr = sqlStr & " 	where q.replyuser is not null"
		sqlStr = sqlStr & " 	and q.isusing='Y' " & addSqlStr
		sqlStr = sqlStr & " ) as t"
		sqlStr = sqlStr & " group by t.replydate"
		if (FRectGroupByReplyUser = "Y") then
			sqlStr = sqlStr & " , t.replyuser"		
		end if
		if (FRectGroupByReplyUser = "Y") then
			sqlStr = sqlStr & " order by t.replyuser asc, t.replydate asc"
		else
			sqlStr = sqlStr & " order by t.replydate asc"
		end if

		'response.write sqlStr & "<Br>"
	    db3_rsget.pagesize = FPageSize
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CReportEvalItem

				if (FRectGroupByReplyUser = "Y") then
					FItemList(i).freplyuser			= db3_rsget("replyuser")
				end if

				FItemList(i).freplydate			= db3_rsget("replydate")
				FItemList(i).fd0			= db3_rsget("d0")
				FItemList(i).fd1			= db3_rsget("d1")
				FItemList(i).fd2		= db3_rsget("d2")
				FItemList(i).fd3			= db3_rsget("d3")
				FItemList(i).fd4			= db3_rsget("d4")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		redim  FMasterItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage = 0

		IF application("Svr_Info")="Dev" THEN
			fTENDB="TENDB."
		end if
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class

%>
