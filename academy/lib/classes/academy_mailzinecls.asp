<%

class CUploadMaster

  public Fregdate
  public Fimg1
  public Fimg2
  public Fimgmap1
  public Fimgmap2
  public Fisusing
  public Fcode1
  public Fcode2
  public Fcode3
  public Ftitle

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub MailzineUpload(byval ImageMap,regdate)

		dim sqlStr

		'###########################################################################
		'상품 데이터 입력
		'###########################################################################

		sqlStr = "insert into [db_academy].[dbo].tbl_academy_mailzine" + vbcrlf
		sqlStr = sqlStr & " (regdate,imgmap)" + vbcrlf
		sqlStr = sqlStr & " values(" + vbcrlf
		sqlStr = sqlStr & "'" & regdate & "'," + vbcrlf
		sqlStr = sqlStr & "'" & html2db(ImageMap) & "')"
'response.write	sqlStr
'dbget.close()	:	response.End
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end sub

	public sub MailzineDetail(byval idx)

		dim sqlStr,code

		'###########################################################################
		'상품 데이터
		'###########################################################################

		sqlStr = "select title,regdate,img1,img2,imgmap1,imgmap2,isusing" + vbcrlf
		sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_mailzine" + vbcrlf
		sqlStr = sqlStr & " where idx = " + idx

		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if  not rsACADEMYget.EOF  then
			Ftitle = db2html(rsACADEMYget("title"))
			Fregdate = rsACADEMYget("regdate")
			Fimg1 = rsACADEMYget("img1")
			Fimg2 = rsACADEMYget("img2")
			Fimgmap1 = db2html(rsACADEMYget("imgmap1"))
			Fimgmap2 = db2html(rsACADEMYget("imgmap2"))
			Fisusing = rsACADEMYget("isusing")
            code = split(Fregdate,".")
			Fcode1 = code(0)
			Fcode2 = code(1)
			Fcode3 = 	code(2)
		end if

		rsACADEMYget.Close
	end sub

	public sub MailzineModify(byval idx,ImageMap,regdate,display)

		dim sqlStr

		'###########################################################################
		'상품 데이터 입력
		'###########################################################################

		sqlStr = "update [db_academy].[dbo].tbl_academy_mailzine" + vbcrlf
		sqlStr = sqlStr & " set regdate = '" & regdate & "'," + vbcrlf
		sqlStr = sqlStr & " imgmap = '" + html2db(ImageMap) + "'," + vbcrlf
		sqlStr = sqlStr & " display = '" + display + "'"
		sqlStr = sqlStr & " where idx = " + idx
'response.write	sqlStr
'dbget.close()	:	response.End
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end sub

end Class

class CMailzineListSubItem

	public Fidx
	public Fregdate
	public Ftitle
	public Fisusing

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CMailzineList
	public FItemList()
	public FTotalCount
	public FResultCount
	public FRectDesignerID
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FPCount

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub MailzineList()
		dim sqlStr,i
		'###########################################################################
		'상품 총 갯수 구하기
		'###########################################################################
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_mailzine" + vbcrlf

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		'###########################################################################
		'상품 데이터
		'###########################################################################

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,title,regdate,isusing" + vbcrlf
		sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_mailzine" + vbcrlf
		sqlStr = sqlStr & " order by regdate Desc" + vbcrlf


		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPCount = FCurrPage - 1

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new CMailzineListSubItem
				FItemList(i).Fidx = rsACADEMYget("idx")
				FItemList(i).Ftitle = db2html(rsACADEMYget("title"))
			    FItemList(i).Fregdate = rsACADEMYget("regdate")
				FItemList(i).Fisusing = rsACADEMYget("isusing")
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

%>