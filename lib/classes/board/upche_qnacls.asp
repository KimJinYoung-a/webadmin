<%
'####################################################
' Description : [업체]게시판 클래스
' History : 2015.05.27 이상구 생성
'		  :	2016.01.13 한용민 수정
'####################################################

class CUpcheQnASubItem
	public Fidx
	public Fgubun
	public Fuserid
	public Fusername
	public Ftitle
	public Fcontents
	public Fregdate
	public Freplyn
	public Freplyuser
	public Fmasterid
	public Fisusing
	public Freplydate
	public Fworker

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

    ''2009 구분 변경 (파트 위주로)
	function GubunName()
		if Fgubun = "01" then
			GubunName = "배송문의"
		elseif Fgubun = "02" then
			GubunName = "반품문의"
		elseif Fgubun = "03" then
			GubunName = "교환문의"
		elseif Fgubun = "04" then
			GubunName = "정산문의"
		elseif Fgubun = "05" then
			GubunName = "입고문의"
		elseif Fgubun = "06" then
			GubunName = "재고문의"
		elseif Fgubun = "07" then
			GubunName = "상품등록문의"
		elseif Fgubun = "08" then
			GubunName = "이벤트시행문의"
		elseif Fgubun = "20" then
			GubunName = "기타문의"
		end if
	end function

	function UpcheGubun()
		if Fmasterid = "01" then
			UpcheGubun = "업체"
		elseif Fmasterid = "02" then
			UpcheGubun = "제휴몰"
		end if
	end function
end Class

Class CUpcheQnA
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectGubun
	public FRectRelpy
	public FRectUsing
	public FRectUserid
	public FRectSearchKey
	public FRectSearchString
	public FRectSelDate
	public FRectSDate
	public FRectEDate
	public FRectIsRecenct
	public FWorkerGubun
	public FRectSortBy
	public Frectworkergubuntype

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'/admin/board/upche_qna_board_list.asp		'/2016.01.13 한용민 생성
	public Sub getqnalist()
		dim sql, addSql, sortSql, i

		if FRectUsing="N" then
			addSql = " A.isusing = 'N' "
		else
			addSql = " A.isusing = 'Y' "
		end if

		if FRectGubun <> "" then
			addSql = addSql + " and A.gubun = '" + FRectGubun + "'"
		end if

		if FRectRelpy = "N" then
			addSql = addSql + " and A.replyuser is null"
		elseif FRectRelpy = "Y" then
			addSql = addSql + " and A.replyuser is Not null"
		end if

		if FRectUserid <> "" then
			addSql = addSql + " and A.userid = '" + FRectUserid + "'"
		end if

		if FRectSearchString<>"" then
            if (FRectSearchKey = "replyuser") or (FRectSearchKey = "userid") then
                addSql = addSql + " and A." & FRectSearchKey & " = '" + html2db(FRectSearchString) + "'"
            else
                addSql = addSql + " and A." & FRectSearchKey & " like '%" + html2db(FRectSearchString) + "%'"
            end if
		end if

		If Frectworkergubuntype <> "" Then
			if FWorkerGubun<>"" then
				if Frectworkergubuntype="MY" then
					addSql = addSql + " and A.workerid = '" + FWorkerGubun + "'"
				elseif Frectworkergubuntype="SELECTID" then
					addSql = addSql + " and A.workerid = '" + FWorkerGubun + "'"
				elseif Frectworkergubuntype="SELECTNAME" then
					addSql = addSql + " and B.username = '" + FWorkerGubun + "'"
				end if
			end if
		End If

		if FRectSelDate<>"A" then
			if FRectSDate<>"" then addSql = addSql + " and A.regdate >= '" + FRectSDate + "'"
			if FRectEDate<>"" then addSql = addSql + " and A.regdate <= '" + FRectEDate + " 23:59:59'"
		else
			if FRectSDate<>"" then addSql = addSql + " and A.replydate >= '" + FRectSDate + "'"
			if FRectEDate<>"" then addSql = addSql + " and A.replydate <= '" + FRectEDate + " 23:59:59'"
		end if

		if FRectIsRecenct="Y" then
			addSql = addSql & " and A.regdate>DATEADD(month,-6,getdate()) "
		end if

		Select Case FRectSortBy
			Case "rd"
				sortSql = "A.idx desc"			'최근등록순
			Case "ra"
				sortSql = "A.idx asc"			'오랜등록순
			Case "ad"
				sortSql = "A.replydate desc"	'최근답변순
			Case Else
				sortSql = "A.idx desc"
		End Select

		sql = "select count(A.idx) as cnt, CEILING(CAST(Count(A.idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sql = sql + " from [db_board].[dbo].tbl_upche_qna AS A with(noLock)"
		sql = sql + " Left JOIN db_partner.dbo.tbl_user_tenbyten AS B with(noLock)"
		sql = sql + "  	ON A.workerid = B.userid"
		sql = sql + "  	and B.isusing=1 and isnull(B.userid,'')<>''"
		sql = sql + " where " + addSql

		'response.write sql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sql = " select "
		sql = sql + " A.idx, A.masterid, A.gubun, A.userid, A.username, A.title, A.regdate, A.replyuser, A.replydate"
		sql = sql + " , isnull(A.replyuser,'') as replyn, A.isusing, isNull(B.username,'') AS worker "
		sql = sql + " from [db_board].[dbo].tbl_upche_qna AS A with(noLock)"
		sql = sql + " Left JOIN db_partner.dbo.tbl_user_tenbyten AS B with(noLock)"
		sql = sql + "  	ON A.workerid = B.userid and isnull(B.userid,'')<>''"
		sql = sql + "  	and B.isusing=1"
		sql = sql + " where " + addSql
		sql = sql + " order by " & sortSql
		sql = sql + " OFFSET " & (FCurrPage-1)*FPageSize & " ROWS FETCH NEXT " & FPageSize & " ROWS ONLY "

		'response.write sql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CUpcheQnASubItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fmasterid	= rsget("masterid")
				FItemList(i).Fgubun		= rsget("gubun")
				FItemList(i).Fuserid	= rsget("userid")
				FItemList(i).Fusername	= db2html(rsget("username"))
				FItemList(i).Ftitle		=  db2html(rsget("title"))
				FItemList(i).Fregdate	= rsget("regdate")
				FItemList(i).Freplyuser	= rsget("replyuser")
				FItemList(i).Freplyn	= rsget("replyn")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).Freplydate	= rsget("replydate")

				If rsget("worker") = "" Then
					FItemList(i).Fworker	= "&nbsp;"
				Else
					FItemList(i).Fworker	= rsget("worker")
				End If
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub list()
		dim sql, addSql, i

		if FRectUsing="N" then
			addSql = " A.isusing = 'N' "
		else
			addSql = " A.isusing = 'Y' "
		end if

		if FRectGubun <> "" then
			addSql = addSql + " and A.gubun = '" + FRectGubun + "'"
		end if

		if FRectRelpy = "N" then
			addSql = addSql + " and A.replyuser is null"
		elseif FRectRelpy = "Y" then
			addSql = addSql + " and A.replyuser is Not null"
		end if

		if FRectUserid <> "" then
			addSql = addSql + " and A.userid = '" + FRectUserid + "'"
		end if

		if FRectSearchString<>"" then
			addSql = addSql + " and A." & SearchKey & " like '%" + html2db(FRectSearchString) + "%'"
		end if

		If FWorkerGubun <> "" Then
			addSql = addSql + " and A.workerid = '" + FWorkerGubun + "'"
		End If

		if FRectSelDate<>"A" then
			if FRectSDate<>"" then addSql = addSql + " and A.regdate >= '" + FRectSDate + "'"
			if FRectEDate<>"" then addSql = addSql + " and A.regdate <= '" + FRectEDate + " 23:59:59'"
		else
			if FRectSDate<>"" then addSql = addSql + " and A.replydate >= '" + FRectSDate + "'"
			if FRectEDate<>"" then addSql = addSql + " and A.replydate <= '" + FRectEDate + " 23:59:59'"
		end if

		If FRectGubun <> "" Then
			addSql = addSql + " and A.gubun = '" & FRectGubun & "'"
		End If

		sql = "select count(A.idx) as cnt, CEILING(CAST(Count(A.idx) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sql = sql + " from [db_board].[dbo].tbl_upche_qna AS A with (nolock)"
		sql = sql + " where " + addSql

		'response.write sql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료 2014-07-31 김진영 / 메모리 오류 수정
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sql = " select top " + CStr(FPageSize*FCurrPage) + ""
		sql = sql + " A.idx, A.masterid, A.gubun, A.userid, A.username, A.title, A.regdate, A.replyuser, A.replydate"
		sql = sql + " , isnull(A.replyuser,'') as replyn, A.isusing, isNull(B.username,'') AS worker "
		sql = sql + " from [db_board].[dbo].tbl_upche_qna AS A with (nolock)"
		sql = sql + " Left JOIN db_partner.dbo.tbl_user_tenbyten AS B with (nolock)"
		sql = sql + "  	ON A.workerid = B.userid  "
		sql = sql + " where " + addSql
		sql = sql + " order by A.regdate desc "

		'response.write sql & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CUpcheQnASubItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fmasterid	= rsget("masterid")
				FItemList(i).Fgubun		= rsget("gubun")
				FItemList(i).Fuserid	= rsget("userid")
				FItemList(i).Fusername	= db2html(rsget("username"))
				FItemList(i).Ftitle		=  db2html(rsget("title"))
				FItemList(i).Fregdate	= rsget("regdate")
				FItemList(i).Freplyuser	= rsget("replyuser")
				FItemList(i).Freplyn	= rsget("replyn")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).Freplydate	= rsget("replydate")

				If rsget("worker") = "" Then
					FItemList(i).Fworker	= "&nbsp;"
				Else
					FItemList(i).Fworker	= rsget("worker")
				End If
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

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

class CUpcheQnADetail
	public Fidx
	public Fgubun
	public Fuserid
	public Fusername
	public Ftitle
	public Fcontents
	public Fregdate
	public Freplyuser
	public Freplytitle
	public Freplycontents
	public Fisusing
	public FRectIdx
	public Freplyn
	public Freplydate
	public Fworkerid
	public FTeam
	public FRectRelpy
	public femail

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Sub read()
		dim sql, i, addSql

		if FRectRelpy = "N" then
			addSql = addSql & " and isnull(q.replyuser,'')=''"
		elseif FRectRelpy = "Y" then
			addSql = addSql & " and isnull(q.replyuser,'')<>''"
		end if
		if FRectIdx <> "" then
			addSql = addSql & " and q.idx=" + Cstr(FRectIdx)
		end if

		sql = " select top 1" & VbCrlf
		sql = sql & " q.idx,q.masterid,q.gubun,q.userid,q.username,q.title,q.contents,q.regdate,q.replyuser,isnull(q.replyuser,'') as replyn" & VbCrlf
		sql = sql & " ,q.replytitle,q.replycontents, q.isusing, q.replydate, q.workerid, p.email" & VbCrlf
		sql = sql & " from [db_board].[dbo].tbl_upche_qna q with (readuncommitted)"
		sql = sql & " left join db_partner.dbo.tbl_partner p with (readuncommitted)"
		sql = sql & " 	on q.userid=p.id"
		sql = sql & " where 1=1 " & addSql

		'response.write sql & "<Br>"
        rsget.CursorLocation = adUseClient
    	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		if  not rsget.EOF  then
			Fidx			= rsget("idx")
			Fgubun			= rsget("gubun")
			Fuserid			= rsget("userid")
			Fusername		= db2html(rsget("username"))
			Ftitle			= db2html(rsget("title"))
			Fcontents		= db2html(rsget("contents"))
			Fregdate		= rsget("regdate")
			Freplyuser		= rsget("replyuser")
			Freplyn			= rsget("replyn")
			Freplytitle		= db2html(rsget("replytitle"))
			Freplycontents	= db2html(rsget("replycontents"))
			Fisusing		= rsget("isusing")
			Freplydate		= rsget("replydate")
			Fworkerid		= rsget("workerid")
			femail		= db2html(rsget("email"))
		end if
		rsget.close
	end sub

	Public Function reply(byval idx, replytitle, replycontents, replyuser)
        dim sql, i

        sql = "update [db_board].[dbo].tbl_upche_qna " + VbCrlf
        sql = sql + " set replytitle = '" + html2db(replytitle) + "'," + VbCrlf
		sql = sql + " replycontents = '" + html2db(replycontents) + "'," + VbCrlf
		sql = sql + " replyuser = '" + replyuser + "'," + VbCrlf
		sql = sql + " replydate = getdate()" + VbCrlf
        sql = sql + " where (idx = " + idx + ") "
        rsget.CursorLocation = adUseClient
    	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
	end Function

	Public Function changeworker(byval idx, workerid)
        dim sql, i

        sql = "update [db_board].[dbo].tbl_upche_qna " + VbCrlf
        sql = sql + " set workerid = '" + workerid + "'" + VbCrlf
        sql = sql + " where (idx = " + idx + ") "
        rsget.CursorLocation = adUseClient
    	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
	end Function

	Public Function write(byval masterid,gubun,title,contents,userid,username,workerid)
        dim sql, i

        sql = " insert into [db_board].[dbo].tbl_upche_qna(masterid,gubun,userid,username,title,contents,workerid) "
        sql = sql + " values('" + masterid + "', '" + gubun + "', convert(varchar(32),'" + userid + "'), convert(varchar(32),'" + html2db(username) + "'), convert(varchar(128),'" + html2db(Replace(title, "'", "")) + "'),'" + html2db(Replace(contents, "'", "")) + "','" + workerid + "') "
		''response.write sql
		''dbget.close()	:	response.End
        rsget.CursorLocation = adUseClient
    	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
	end Function

	Public Function modify(byval idx,gubun,title,contents,workerid)
        dim sql, i

        sql = "update [db_board].[dbo].tbl_upche_qna" + VbCrlf
		sql = sql + " set gubun = '" + gubun + "'," + VbCrlf
		sql = sql + " title = '" + html2db(title) + "'," + VbCrlf
		sql = sql + " contents = '" + html2db(contents) + "', " + VbCrlf
		sql = sql + " workerid = '" + workerid + "' " + VbCrlf
        sql = sql + " where idx = " + Cstr(idx)
        rsget.CursorLocation = adUseClient
    	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
	end Function

	Public Function del(byval idx)
        dim sql, i

        sql = "update [db_board].[dbo].tbl_upche_qna" + VbCrlf
		sql = sql + " set isusing = 'N' " + VbCrlf
        sql = sql + " where (idx = " + idx + ") "
        rsget.CursorLocation = adUseClient
    	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
	end Function

	'####### 직원리스트 #######
	public Function fnGetMemberList
		Dim strSql, addSql
		If FTeam = "11" Then
			addSql = ""
			addSql = addSql & " and ((C.posit_sn<=8) OR C.posit_sn in ('12', '13')) "
		Else
			'addSql = addSql & " and C.posit_sn<=8 "	' 주석처리. MD팀에서 계약직분들이 안보인다는 문의옴.
		End If

		strSql = "	SELECT A.id, B.departmentname, C.posit_name, D.username as company_name, D.department_id, D.mywork " & _
				"		FROM [db_partner].[dbo].tbl_partner AS A " & _
				"		INNER JOIN [db_partner].[dbo].tbl_positInfo AS C ON A.posit_sn = C.posit_sn " & _
				"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS D ON A.id = D.userid " & _
				"		INNER JOIN [db_partner].[dbo].tbl_user_department AS B ON D.department_id = B.cid " & _
				"	WHERE A.isusing = 'Y' AND A.userdiv < 999 AND A.id <> '' AND Left(A.id,10) <> 'streetshop' " & _
				"		and d.isUsing=1 and (d.statediv ='Y' or (d.statediv ='N' and datediff(dd,d.retireday,getdate())<=0)) " &_
				"		AND (B.cid IN(" & FTeam & ")  or B.pid in (" & FTeam & ") )"  & addSql & _
				"	ORDER BY D.department_id ASC, A.posit_sn ASC, A.regdate ASC "
				'### AND A.id != 'yanan716' 휴직자임. CS 민연희.

        'response.write strSql & "<Br>"
		rsget.CursorLocation = adUseClient
    	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			fnGetMemberList = rsget.getRows()
		END IF
		rsget.close
	End Function
end Class

'####### 직원 이름구하기 #######
Function fnGetMemberName(id)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 D.username FROM [db_partner].[dbo].tbl_partner as A "
	strSql = strSql & " INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS D ON A.id = D.userid "
	strSql = strSql & " WHERE A.id = '" & id & "'  "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	IF not rsget.EOF THEN
		fnGetMemberName = rsget("username")
	Else
		fnGetMemberName = ""
	END IF
	rsget.close
End Function
%>
