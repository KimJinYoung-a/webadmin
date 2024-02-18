<%
Class CEnjoyItem
	public FdenjSn
	public Fmakerid
	public FsmallImage
	public FlistImage
	public FintroImage
	public FbestImage
	public Fv2mainimage
	public Fsubject
	public FvideoSn
	public FcmtCnt
	public Fisusing
	public Fregdate
	public Feventday

	public Fbrandname

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CEnjoy
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectEnSN
	public FRectMaker
	public FRectUsing

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function GetDiaryEnjoyList()
		dim sqlStr, addSql, i

		'추가 조건절
		if FRectEnSN<>"" then
			addSql = addSql & " and denjSn='" & FRectEnSN & "'"
		end if
		if FRectMaker<>"" then
			addSql = addSql & " and d.makerid='" & FRectMaker & "'"
		end if
		if FRectUsing<>"" then
			addSql = addSql & " and d.isUsing='" & FRectUsing & "'"
		end if

		'카운트
		sqlStr = "select count(denjSn) as cnt"
		sqlStr = sqlStr + " from [db_diary2010].[dbo].tbl_diary_enjoy as d" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " denjSn, d.makerid, d.smallImage, d.listImage, d.introImage, d.bestImage, d.subject, videoSn, d.cmtCnt, d.isusing, d.regdate, d.v2mainimage, d.eventday " + vbcrlf
		sqlStr = sqlStr + " 	,c.socname_kor " + vbcrlf
		sqlStr = sqlStr + " from [db_diary2010].[dbo].tbl_diary_enjoy as d " + vbcrlf
		sqlStr = sqlStr + " 	Join [db_user].[dbo].tbl_user_c as c " + vbcrlf
		sqlStr = sqlStr + " 		on d.makerid = c.userid " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " order by denjSn desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CEnjoyItem

				FItemList(i).FdenjSn		= rsget("denjSn")
				FItemList(i).Fmakerid		= rsget("makerid")
				FItemList(i).FsmallImage	= rsget("smallImage")
				FItemList(i).FlistImage		= rsget("listImage")
				FItemList(i).FintroImage	= rsget("introImage")
				FItemList(i).FbestImage		= rsget("bestImage")
				FItemList(i).Fsubject		= rsget("subject")
				FItemList(i).FvideoSn		= rsget("videoSn")
				FItemList(i).FcmtCnt		= rsget("cmtCnt")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Fbrandname		= rsget("socname_kor")
				FItemList(i).Fv2mainimage	= rsget("v2mainimage")
				FItemList(i).Feventday		= rsget("eventday")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

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
%>