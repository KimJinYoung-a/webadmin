<%

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

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

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
	public FRectUserid
	public FRectSearchKey
	public FRectSearchString

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

	public Sub list()
		dim sql, addSql, i

		'// 조건절
		if FRectGubun <> "" then
			addSql = addSql + " and gubun = '" + FRectGubun + "'"
		end if

		if FRectRelpy = "N" then
			addSql = addSql + " and replyuser is null"
		elseif FRectRelpy = "Y" then
			addSql = addSql + " and replyuser is Not null"
		end if

		if FRectUserid <> "" then
			addSql = addSql + " and userid = '" + FRectUserid + "'"
		end if

		if FRectSearchString<>"" then
			addSql = addSql + " and " & SearchKey & " like '%" + FRectSearchString + "%'"
		end if

'#######################################################
'총데이터
'#######################################################
		sql = "select count(idx) as cnt "
		sql = sql + " from [db_board].[10x10].tbl_upche_qna"
		sql = sql + " where isusing = 'Y' " + addSql
		rsget.Open sql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


'#######################################################
'데이터
'#######################################################

		sql = " select top " + CStr(FPageSize*FCurrPage) + " idx,masterid,gubun,userid,username, title, regdate,replyuser,isnull(replyuser,'') as replyn"
		sql = sql + " from [db_board].[10x10].tbl_upche_qna"
		sql = sql + " where isusing = 'Y' " + addSql
		sql = sql + " order by regdate desc "

'response.write sql
		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

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

				FItemList(i).Fidx          = rsget("idx")
				FItemList(i).Fmasterid          = rsget("masterid")
				FItemList(i).Fgubun          = rsget("gubun")
				FItemList(i).Fuserid   = rsget("userid")
				FItemList(i).Fusername   = db2html(rsget("username"))
				FItemList(i).Ftitle   =  db2html(rsget("title"))
				FItemList(i).Fregdate      = rsget("regdate")
				FItemList(i).Freplyuser      = rsget("replyuser")
				FItemList(i).Freplyn      = rsget("replyn")

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
	public FRectIdx
	public Freplyn

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub read()
		dim sql, i

		sql = " select top 1 idx,masterid,gubun,userid,username,title,contents,regdate,"
		sql = sql + " replyuser,isnull(replyuser,'') as replyn,replytitle,replycontents"
		sql = sql + " from [db_board].[10x10].tbl_upche_qna"
		sql = sql + " where isusing = 'Y'"
		if FRectIdx <> "" then
		sql = sql + " and  idx=" + Cstr(FRectIdx)
		end if

		rsget.Open sql, dbget, 1

		if  not rsget.EOF  then
			Fidx          = rsget("idx")
			Fgubun          = rsget("gubun")
			Fuserid   = rsget("userid")
			Fusername   = db2html(rsget("username"))
			Ftitle   = db2html(rsget("title"))
			Fcontents   = db2html(rsget("contents"))
			Fregdate      = rsget("regdate")
			Freplyuser      = rsget("replyuser")
			Freplyn      = rsget("replyn")
			Freplytitle      = db2html(rsget("replytitle"))
			Freplycontents      = db2html(rsget("replycontents"))
		end if
		rsget.close
	end sub

	Public Function reply(byval idx, replytitle, replycontents, replyuser)
                dim sql, i

                sql = "update [db_board].[10x10].tbl_upche_qna " + VbCrlf
                sql = sql + " set replytitle = '" + replytitle + "'," + VbCrlf
				sql = sql + " replycontents = '" + replycontents + "'," + VbCrlf
				sql = sql + " replyuser = '" + replyuser + "'" + VbCrlf
                sql = sql + " where (idx = " + idx + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function write(byval masterid,gubun,title,contents,userid,username)
                dim sql, i

                sql = " insert into [db_board].[10x10].tbl_upche_qna(masterid,gubun,userid,username,title,contents) "
                sql = sql + " values('" + masterid + "', '" + gubun + "', '" + userid + "', '" + username + "', '" + title + "','" + contents + "') "

'				dbget.close()	:	response.End
                rsget.Open sql, dbget, 1
	end Function

	Public Function modify(byval idx,gubun,title,contents)
                dim sql, i

                sql = "update [db_board].[10x10].tbl_upche_qna" + VbCrlf
				sql = sql + " set gubun = '" + gubun + "'," + VbCrlf
				sql = sql + " title = '" + title + "'," + VbCrlf
				sql = sql + " contents = '" + contents + "' " + VbCrlf
                sql = sql + " where idx = " + Cstr(idx)
                rsget.Open sql, dbget, 1
	end Function

	Public Function del(byval idx)
                dim sql, i

                sql = "update [db_board].[10x10].tbl_upche_qna" + VbCrlf
				sql = sql + " set isusing = 'N' " + VbCrlf
                sql = sql + " where (idx = " + idx + ") "
                rsget.Open sql, dbget, 1
	end Function

end Class

%>