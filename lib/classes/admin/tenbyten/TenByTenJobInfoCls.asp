<%
Class CTenByTenJobItem
	public Fjob_sn
	public Fjob_name
	public Fjob_isDel

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class


Class CTenByTenJob
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectJob_sn
	public FRectsearchKey
	public FRectsearchString

	Private Sub Class_Initialize()
		redim  FitemList(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'##### 직책 목록 #####
	public Sub GetList()
		dim SQL, AddSQL, i, strTemp

		'// 검색어 쿼리 //
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " Where " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		'// 개수 파악 //
		SQL =	"Select count(job_sn), CEILING(CAST(Count(job_sn) AS FLOAT)/" & FPageSize & ") " &_
				"From db_partner.dbo.tbl_JobInfo " & AddSQL
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'// 목록 접수 //
		SQL =	"Select top " & CStr(FPageSize*FCurrPage) & " * " &_
				"From db_partner.dbo.tbl_JobInfo " & AddSQL
		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTenByTenJobItem

				FItemList(i).Fjob_sn		= rsget("job_sn")
				FItemList(i).Fjob_name	= rsget("job_name")
				FItemList(i).Fjob_isDel	= rsget("job_isDel")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function


	public Sub GetInfo()
		dim SQL

		'// 목록 접수 //
		SQL =	"Select * " &_
				"From db_partner.dbo.tbl_JobInfo " &_
				"Where job_sn=" & FRectJob_sn
		rsget.Open SQL,dbget,1

		if Not(rsget.EOF or rsget.BOF) then

			FResultCount = 1
			redim preserve FItemList(1)
			set FItemList(1) = new CTenByTenJobItem

			FItemList(1).Fjob_name	= rsget("job_name")
			FItemList(1).Fjob_isDel	= rsget("job_isDel")
		else
			FResultCount = 0
		end if

		rsget.Close

	end Sub
end Class
%>