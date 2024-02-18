<%
Class CVideoItem
	public FvideoSn
	public FvideoDiv
	public FvideoTitle
	public FvideoWidth
	public FvideoHeight
	public FvideoFile
	public FvideoThumb
	public Fisusing
	public Fregdate
	public FstartDate
	public FendDate
	public Flinkgubun
	public Flinkinfo

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CVideo
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectVSN
	public FRectDiv
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

	public Function GetVideoList()
		dim sqlStr, addSql, i

		'추가 조건절
		if FRectVSN<>"" then
			addSql = addSql & " and videoSn='" & FRectVSN & "'"
		end if
		if FRectDiv<>"" then
			addSql = addSql & " and videoDiv='" & FRectDiv & "'"
		end if
		if FRectUsing<>"" then
			addSql = addSql & " and isUsing='" & FRectUsing & "'"
		end if

		'카운트
		sqlStr = "select count(videoSn) as cnt"
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_VideoInfo as t1" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " videoSn, videoDiv, videoTitle, videoWidth, videoHeight, videoFile, videoThumb, isusing, regdate " + vbcrlf
		sqlStr = sqlStr + " , startDate, endDate, linkgubun, linkinfo" + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_VideoInfo " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " order by videoSn desc "

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
				set FItemList(i) = new CVideoItem

				FItemList(i).FvideoSn		= rsget("videoSn")
				FItemList(i).FvideoDiv		= rsget("videoDiv")
				FItemList(i).FvideoTitle	= db2html(rsget("videoTitle"))
				FItemList(i).FvideoWidth	= rsget("videoWidth")
				FItemList(i).FvideoHeight	= rsget("videoHeight")
				FItemList(i).FvideoFile		= rsget("videoFile")
				FItemList(i).FvideoThumb	= rsget("videoThumb")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FstartDate		= rsget("startDate")
				FItemList(i).FendDate		= rsget("endDate")
				FItemList(i).Flinkgubun		= rsget("linkgubun")
				FItemList(i).Flinkinfo		= rsget("linkinfo")

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

Function drawVDivSelect(fnm,svl)
	Dim strRet
	strRet = "<select name='" & fnm & "' class='select' id='" & fnm & "'>" & vbCrLf &_
			" <option value=''>구분</option>" & vbCrLf

	strRet = strRet  & "<option value='dia'"
	if svl="dia" then strRet = strRet  & " selected"
	strRet = strRet  & ">다이어리</option>" & vbCrLf

	strRet = strRet  & "<option value='prd'"
	if svl="prd" then strRet = strRet  & " selected"
	strRet = strRet  & ">상품</option>" & vbCrLf

	strRet = strRet  & "<option value='evt'"
	if svl="evt" then strRet = strRet  & " selected"
	strRet = strRet  & ">이벤트</option>" & vbCrLf
	
	strRet = strRet  & "<option value='fin'"
	if svl="fin" then strRet = strRet  & " selected"
	strRet = strRet  & ">디자인핑거스</option>" & vbCrLf

	strRet = strRet  & "<option value='mov'"
	if svl="mov" then strRet = strRet  & " selected"
	strRet = strRet  & ">PC메인동영상</option>" & vbCrLf

	strRet = strRet  & "</select>"

	drawVDivSelect = strRet
end Function

Function getVDivName(vdv)
	Select Case vdv
		Case "dia"
			getVDivName = "다이어리"
		Case "prd"
			getVDivName = "상품"
		Case "evt"
			getVDivName = "이벤트"
		Case "fin"
			getVDivName = "디자인핑거스"
		Case "mov"
			getVDivName = "PC메인동영상"
	end Select
End Function
%>