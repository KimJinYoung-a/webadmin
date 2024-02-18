<%

class CUploadMaster

  public Fidx
  public Ftitle
  public Fnews
  public Fimg1
  public Fimg2
  public Fimg3
  public Fimg4
  public Fimg5
  public Fimg6
  public Fimg7
  public Furl1
  public Furl2
  public Furl3
  public Furl4
  public Fisusing
  public Fregdate
  public Fbrand
  public Fsendmailer

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub MailzineDetail(byval idx)

		dim sqlStr,code
		
		'###########################################################################
		'상품 데이터
		'###########################################################################

		sqlStr = "select title,news,img1,img2,img3,img4,img5,img6,img7,url1,url2,url3,url4,brand,sendmailer,isusing,regdate" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shopmaster_mail" + vbcrlf
		sqlStr = sqlStr & " where idx = " + idx

		rsget.Open sqlStr,dbget,1

		if  not rsget.EOF  then
			Ftitle = db2html(rsget("title"))
			Fnews = db2html(rsget("news"))
			Fimg1 = rsget("img1")
			Fimg2 = rsget("img2")
			Fimg3 = rsget("img3")
			Fimg4 = rsget("img4")
			Fimg5 = rsget("img5")
			Fimg6 = rsget("img6")
			Fimg7 = rsget("img7")
			Furl1 = db2html(rsget("url1"))
			Furl2 = db2html(rsget("url2"))
			Furl3 = db2html(rsget("url3"))
			Furl4 = db2html(rsget("url4"))
			Fbrand = db2html(rsget("brand"))
			Fsendmailer = db2html(rsget("sendmailer"))
			Fisusing = rsget("isusing")
			Fregdate = rsget("regdate")
		end if

		rsget.Close
	end sub

	public sub MailzineView(byval idx)

		dim sqlStr,code
		
		'###########################################################################
		'상품 데이터
		'###########################################################################

		sqlStr = "select title,news,img1,img2,img3,img4,img5,img6,img7,url1,url2,url3,url4,brand,sendmailer,isusing,regdate" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shopmaster_mail" + vbcrlf
		sqlStr = sqlStr & " where idx = " + idx

		rsget.Open sqlStr,dbget,1

		if  not rsget.EOF  then
			Ftitle = db2html(rsget("title"))
			Fnews = db2html(rsget("news"))
			Fimg1 = "http://imgstatic.10x10.co.kr/offshopmailzine/" + Cstr(left(rsget("regdate"),4)) + "/" + rsget("img1")
			Fimg2 = "http://imgstatic.10x10.co.kr/offshopmailzine/" + Cstr(left(rsget("regdate"),4)) + "/" + rsget("img2")
			Fimg3 = "http://imgstatic.10x10.co.kr/offshopmailzine/" + Cstr(left(rsget("regdate"),4)) + "/" + rsget("img3")
			Fimg4 = "http://imgstatic.10x10.co.kr/offshopmailzine/" + Cstr(left(rsget("regdate"),4)) + "/" + rsget("img4")
			Fimg5 = "http://imgstatic.10x10.co.kr/offshopmailzine/" + Cstr(left(rsget("regdate"),4)) + "/" + rsget("img5")
			Fimg6 = "http://imgstatic.10x10.co.kr/offshopmailzine/" + Cstr(left(rsget("regdate"),4)) + "/" + rsget("img6")
			Fimg7 = "http://imgstatic.10x10.co.kr/offshopmailzine/" + Cstr(left(rsget("regdate"),4)) + "/" + rsget("img7")
			Furl1 = db2html(rsget("url1"))
			Furl2 = db2html(rsget("url2"))
			Furl3 = db2html(rsget("url3"))
			Furl4 = db2html(rsget("url4"))
			Fbrand = db2html(rsget("brand"))
			Fsendmailer = db2html(rsget("sendmailer"))
			Fisusing = rsget("isusing")
			Fregdate = rsget("regdate")
		end if

		rsget.Close
	end sub

end Class

class CMailzineListSubItem

	public Fidx
	public Fregdate
	public Ftitle
	public Fisusing
   public Fmailyn

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
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shopmaster_mail" + vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		'###########################################################################
		'상품 데이터
		'###########################################################################

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,title,regdate,mailyn,isusing" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shopmaster_mail" + vbcrlf
		sqlStr = sqlStr & " order by regdate Desc" + vbcrlf


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

		FPCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CMailzineListSubItem
				FItemList(i).Fidx = rsget("idx")
				FItemList(i).Ftitle = rsget("title")
			   FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Fmailyn = rsget("mailyn")
				FItemList(i).Fisusing = rsget("isusing")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
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