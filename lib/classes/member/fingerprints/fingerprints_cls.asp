<%
'###########################################################
' Description : 지문인식 근태관리
' Hieditor : 2011.03.22 한용민 생성
'###########################################################

class cfingerprints_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fidx
	public fplaceid
	public fempno
	public fYYYYMMDD
	public finoutType
	public finoutTime
	public fposIdx
	public fposDate
	public fregdate
	public flasteditupdate
	public fplaceiname
	public fvalidpart
	public fusername
	public finoutTypeName
	public fisusing
	public flastedituserid
	public fInTime
	public fOutTime
	public fworkmin
	public fexmin
	public freoutCNT
	public freinCNT
	public fpart_sn
	public fpart_name
	public FdepartmentNameFull
	public fuserid
end class

class cfingerprints_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public foneitem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public frectplaceid
	public FrectSDate
	public FrectEDate
	public frectidx
	public frectempno
	public frectpart_sn
	public FSearchType
	public FSearchText

	public Fdepartment_id
	public Finc_subdepartment

	'/common/member/fingerprints/fingerprints_poscode.asp
    public Sub fposcode_oneitem()
        dim SqlStr

        SqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " placeid ,placeiname ,validpart , isusing" + vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_user_inouttime_place" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
        SqlStr = SqlStr + " and placeid=" + CStr(FRectplaceid)

        'response.write SqlStr &"<br>" 
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        ftotalcount = rsget.RecordCount

        set FOneItem = new cfingerprints_item
        if Not rsget.Eof then

            FOneItem.fisusing = rsget("isusing")
            FOneItem.fplaceid = rsget("placeid")
            FOneItem.fplaceiname = db2html(rsget("placeiname"))
            FOneItem.fvalidpart	= rsget("validpart")

        end if
        rsget.close
    end Sub

	'/common/member/fingerprints/fingerprints_poscode.asp
	public sub fposcode_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_user_inouttime_place" + vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " placeid ,placeiname ,validpart ,isusing" + vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_user_inouttime_place" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cfingerprints_item

				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fplaceid = rsget("placeid")
				FItemList(i).fplaceiname = db2html(rsget("placeiname"))
				FItemList(i).fvalidpart = rsget("validpart")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/common/member/fingerprints/fingerprints_inouttime_sum.asp
	public sub ffingerprints_sum()
		dim sqlStr,i

		sqlStr = " exec db_partner.[dbo].[sp_Ten_user_inoutTimeLogSummary] '"&frectempno&"','"&frectplaceid&"','"&FrectSDate&"','"&FrectEDate&"','"&Frectpart_sn&"'"

		'Response.write sqlStr &"<br>"
		'response.end

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget

		fresultcount = rsget.recordcount
		ftotalcount = rsget.recordcount
		redim preserve FItemList(fresultcount)

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set fitemlist(i) = new cfingerprints_item

		            fitemlist(i).fyyyymmdd = rsget("yyyymmdd")
		            fitemlist(i).fempno = rsget("empno")
		            fitemlist(i).fusername = rsget("username")
		            fitemlist(i).fInTime = rsget("InTime")
		            fitemlist(i).fOutTime = rsget("OutTime")
		            fitemlist(i).fworkmin = rsget("workmin")
		            fitemlist(i).fexmin = rsget("exmin")
		            fitemlist(i).freoutCNT = rsget("reoutCNT")
		            fitemlist(i).freinCNT = rsget("reinCNT")
		            fitemlist(i).fpart_sn = rsget("part_sn")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/common/member/fingerprints/fingerprints_inouttime_edit.asp
    public Sub ffingerprints_item()
        dim sqlStr ,sqlsearch

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx="&frectidx&""
		end if

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " l.idx,l.placeid,l.empno,l.YYYYMMDD,l.inoutType,l.inoutTime,l.posIdx,l.posDate"
		sqlStr = sqlStr & " ,l.regdate,l.lasteditupdate,l.isusing"
		sqlStr = sqlStr & " , p.userid, P.username, p.part_sn"
		sqlStr = sqlStr & " from db_partner.dbo.tbl_user_inouttime_log l with (nolock)"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten P with (nolock)"
		sqlStr = sqlStr & " 	on L.empno=P.empno"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        FResultCount = rsget.RecordCount
        ftotalcount = rsget.RecordCount

        set FOneItem = new cfingerprints_item
        if Not rsget.Eof then
    		FOneItem.fidx = rsget("idx")
    		FOneItem.fplaceid = rsget("placeid")
    		FOneItem.fempno = rsget("empno")
    		FOneItem.fYYYYMMDD = rsget("YYYYMMDD")
    		FOneItem.finoutType = rsget("inoutType")
    		FOneItem.finoutTime = rsget("inoutTime")
    		FOneItem.fposIdx = rsget("posIdx")
    		FOneItem.fposDate = rsget("posDate")
    		FOneItem.fregdate = rsget("regdate")
    		FOneItem.flasteditupdate = rsget("lasteditupdate")
    		FOneItem.fisusing = rsget("isusing")
			FOneItem.fuserid = rsget("userid")
        end if
        rsget.Close
    end Sub

	'/common/member/fingerprints/fingerprints_inouttime.asp
	public sub ffingerprints_list()
		dim sqlStr,i , sqlsearch

		if frectpart_sn <> "" and frectpart_sn <> "1" then
			sqlsearch = sqlsearch & " and p.part_sn="&frectpart_sn&""
		end if

		if FrectSDate <> "" and FrectEDate <> "" then
			sqlsearch = sqlsearch & " and L.yyyymmdd between '"&FrectSDate&"' and  '"&FrectEDate&"'"
		end if

		if FSearchType <> "" and FSearchText <> "" then
			if FSearchType = "1" then
				sqlsearch = sqlsearch & " and L.empno = '"&FSearchText&"'"
			elseif FSearchType = "2" then
				sqlsearch = sqlsearch & " and p.username = '"&FSearchText&"'"
			end if
		end if

		if (Fdepartment_id <> "") then
			if (Finc_subdepartment = "N") then
				sqlsearch = sqlsearch & " AND p.department_id = '" & Fdepartment_id & "' "
			else
				sqlsearch = sqlsearch & " AND (IsNull(dv.cid1, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid2, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid3, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid4, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid5, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid6, -1) = '" & Fdepartment_id & "') "
			end if
		end if

		'총 갯수 구하기
		sqlStr = "select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg" + vbcrlf
		'sqlStr = "select count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_user_inouttime_log L with (nolock)"
	    sqlStr = sqlStr & " Join db_partner.dbo.tbl_user_inouttime_place a with (nolock)"
	    sqlStr = sqlStr & " 	on l.placeid = a.placeid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten P with (nolock)"
		sqlStr = sqlStr & " 	on L.empno=P.empno"
		sqlStr = sqlStr & " 	AND p.isUsing = 1 "		'and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0))
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partInfo pi with (nolock)"
		sqlStr = sqlStr & " 	on p.part_sn = pi.part_sn"
		sqlStr = sqlStr & " left join db_partner.dbo.vw_user_department dv with (nolock)" & vbCrLf
		sqlStr = sqlStr & " on " & vbCrLf
		sqlStr = sqlStr & " 	1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 	and p.department_id = dv.cid " & vbCrLf
		sqlStr = sqlStr & " 	and dv.useYN = 'Y' " & vbCrLf
		sqlStr = sqlStr & " where l.isusing='Y' and a.isusing='Y'" & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
	    sqlStr = sqlStr & " L.idx, L.empno, L.placeid, L.inoutTime ,l.lasteditupdate"    '', L.yyyymmdd
		sqlStr = sqlStr & " ,(CASE WHEN L.inoutType=0 THEN '출근'  "
		sqlStr = sqlStr & "   WHEN L.inoutType=1 THEN '퇴근' 	 "
		sqlStr = sqlStr & " 	WHEN L.inoutType=2 THEN '외출'  "
		sqlStr = sqlStr & " 	WHEN L.inoutType=3 THEN '복귀'  "
		sqlStr = sqlStr & " 	ELSE '' END) as	inoutTypeName	 "
		sqlStr = sqlStr & "  ,l.lastedituserid, P.username,a.placeiname ,pi.part_name ,p.part_sn, isNull(dv.departmentNameFull,'') AS departmentNameFull "
		sqlStr = sqlStr & " from db_partner.dbo.tbl_user_inouttime_log L with (nolock)"
	    sqlStr = sqlStr & " Join db_partner.dbo.tbl_user_inouttime_place a with (nolock)"
	    sqlStr = sqlStr & " 	on l.placeid = a.placeid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten P with (nolock)"
		sqlStr = sqlStr & " 	on L.empno=P.empno"
		sqlStr = sqlStr & " 	AND p.isUsing = 1 "		'and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0))
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partInfo pi with (nolock)"
		sqlStr = sqlStr & " 	on p.part_sn = pi.part_sn"
		sqlStr = sqlStr & " left join db_partner.dbo.vw_user_department dv with (nolock)" & vbCrLf
		sqlStr = sqlStr & " on " & vbCrLf
		sqlStr = sqlStr & " 	1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 	and p.department_id = dv.cid " & vbCrLf
		sqlStr = sqlStr & " 	and dv.useYN = 'Y' " & vbCrLf
		sqlStr = sqlStr & " where l.isusing='Y' and a.isusing='Y'" & sqlsearch
		sqlStr = sqlStr & " order by L.inoutTime desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cfingerprints_item

				FItemList(i).fpart_sn = rsget("part_sn")
				FItemList(i).fpart_name = rsget("part_name")
				FItemList(i).fplaceid = rsget("placeid")
				FItemList(i).flastedituserid = rsget("lastedituserid")
				FItemList(i).flasteditupdate = rsget("lasteditupdate")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fempno = rsget("empno")
				FItemList(i).fusername = db2html(rsget("username"))
				FItemList(i).finoutTypeName = rsget("inoutTypeName")
				FItemList(i).finoutTime = rsget("inoutTime")
				FItemList(i).fplaceiname = db2html(rsget("placeiname"))

				FItemList(i).FdepartmentNameFull		= rsget("departmentNameFull")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
end Class

function Drawplacegubun(selectBoxName,selectedId,changeFlag)
	dim tmp_str,query1
	%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
	 <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
	<%
	query1 = " select placeid ,placeiname ,validpart"
	query1 = query1 & " from db_partner.dbo.tbl_user_inouttime_place"
	query1 = query1 & " where isusing='Y'"

	'response.write query1 &"<Br>"
	rsget.Open query1,dbget,1

	if not rsget.EOF  then
	   do until rsget.EOF
	       if Lcase(selectedId) = Lcase(rsget("placeid")) then
	           tmp_str = " selected"
	       end if
	       response.write("<option value='"&rsget("placeid")&"' "&tmp_str&">" + db2html(rsget("placeiname")) + "</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   loop
	end if
	rsget.close
	response.write("</select>")
end function

function DrawinoutType(selectBoxName,selectedId,changeFlag)
	dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
		<option value='0' <%if selectedId="0" then response.write " selected"%> >출근</option>
		<option value='1' <%if selectedId="1" then response.write " selected"%> >퇴근</option>
		<option value='2' <%if selectedId="2" then response.write " selected"%> >외출</option>
		<option value='3' <%if selectedId="3" then response.write " selected"%> >복귀</option>
	</select>
<%
end function

function Drawisusing(selectBoxName,selectedId,changeFlag)
	dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
		<option value='Y' <%if selectedId="Y" then response.write " selected"%> >Y</option>
		<option value='N' <%if selectedId="N" then response.write " selected"%> >N</option>
	</select>
<%
end function
%>
