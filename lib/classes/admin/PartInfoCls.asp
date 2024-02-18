<%
Class CPartItem
	public Fpart_sn
	public Fpart_name
	public Fpart_sort
	public Fpart_isDel

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class


Class CPart
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectPart_sn
	public FRectsearchKey
	public FRectsearchString
	public FRectorderBy
	public FRectGroupBy

	public FRectPartName
	public FRectPartSortingNumber
	public FRectPartNumber

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

	'##### 부서 목록 접수 ##### 
	public Sub GetPartList()
		dim SQL, AddSQL, i, strTemp

		'// 검색어 쿼리 //
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " Where " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		'// 개수 파악 //
		SQL =	"Select count(part_sn), CEILING(CAST(Count(part_sn) AS FLOAT)/" & FPageSize & ") " &_
				"From db_partner.dbo.tbl_partInfo " & AddSQL
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'// 목록 접수 //
		SQL =	"Select top " & CStr(FPageSize*FCurrPage) & " * " &_
				"From db_partner.dbo.tbl_partInfo " & AddSQL &_
				"Order by part_sort "
		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPartItem

				FItemList(i).Fpart_sn		= rsget("part_sn")
				FItemList(i).Fpart_name		= rsget("part_name")
				FItemList(i).Fpart_sort		= rsget("part_sort")
				FItemList(i).Fpart_isDel	= rsget("part_isDel")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	'##### 부서 목록 접수 ##### 
	public Sub GetPartInfoList()
		dim SQL, AddSQL, i, strTemp , OrderBySQL , GroupBySQL

		if FRectorderBy <> "" then 
			OrderBySQL = " ORDER BY " & FRectorderBy
		end if 

		'// 검색어 쿼리 //
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " Where " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		'// 개수 파악 //
		SQL =	"Select count(part_sn) " &_
				"From db_partner.dbo.tbl_partInfo WITH(NOLOCK) " & AddSQL
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
		rsget.Close

		if FPageSize > 0 then 
		'// 목록 접수 //
			SQL =	"Select top " & CStr(FPageSize*FCurrPage) & " * " &_
					"From db_partner.dbo.tbl_partInfo WITH(NOLOCK) " & AddSQL & OrderBySQL
		else 
			SQL =	"Select * " &_
					"From db_partner.dbo.tbl_partInfo WITH(NOLOCK) " & AddSQL & OrderBySQL
		end if 

		if FPageSize > 0 then
			rsget.pagesize = FPageSize
		end if 
		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPartItem

				FItemList(i).Fpart_sn		= rsget("part_sn")
				FItemList(i).Fpart_name		= rsget("part_name")
				FItemList(i).Fpart_sort		= rsget("part_sort")
				FItemList(i).Fpart_isDel	= rsget("part_isDel")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	public Sub PostPartInfo()
		dim strSQL
		if FRectPartName = "" or FRectPartSortingNumber = "" then
			exit sub
		end if

		strSQL = "Insert into db_partner.dbo.tbl_partInfo " & vbCrLf
		strSQL = strSQL & " (part_name, part_sort, part_isDel) values " & vbCrLf
		strSQL = strSQL & " ('" & FRectPartName & "'" & vbCrLf
		strSQL = strSQL & " ," & FRectPartSortingNumber & vbCrLf
		strSQL = strSQL & " ,'N')"

		rsget.Open strSQL , dbget , 1
	end Sub

	public Sub PutPartInfo()
		dim strSQL
		if FRectPartNumber = "" then
			exit sub
		end if

		strSQL = "Update db_partner.dbo.tbl_partInfo Set " & vbCrLf
		strSQL = strSQL & "	part_name = '" & FRectPartName & "' " & vbCrLf
		strSQL = strSQL & "	,part_sort = " & FRectPartSortingNumber & " " & vbCrLf
		strSQL = strSQL & " Where part_sn = " & FRectPartNumber

		rsget.Open strSQL , dbget , 1
	end Sub

	public Sub DeletePartInfo()
		dim strSQL
		if FRectPartNumber = "" then
			exit sub
		end if

		strSQL = "IF EXISTS(SELECT * FROM db_partner.dbo.tbl_partinfo WITH(NOLOCK) WHERE part_sn = '"& FRectPartNumber &"' and part_isDel = 'N') " & vbCrLf
		strSQL = strSQL & " BEGIN " & vbCrLf
		strSQL = strSQL & " UPDATE db_partner.dbo.tbl_partinfo SET part_isDel = 'Y' WHERE part_sn = '"& FRectPartNumber &"' " & vbCrLf
		strSQL = strSQL & "	END " & vbCrLf
		strSQL = strSQL & " ELSE " & vbCrLf
		strSQL = strSQL & " BEGIN " & vbCrLf
		strSQL = strSQL & " UPDATE db_partner.dbo.tbl_partinfo SET part_isDel = 'N' WHERE part_sn = '"& FRectPartNumber &"' " & vbCrLf
		strSQL = strSQL & "	END "

		rsget.Open strSQL , dbget , 1
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


	'##### 부서 내용 접수 #####
	public Sub GetPartInfo()
		dim SQL

		'// 목록 접수 //
		SQL =	"Select * " &_
				"From db_partner.dbo.tbl_partInfo " &_
				"Where part_sn=" & FRectpart_sn
		rsget.Open SQL,dbget,1

		if Not(rsget.EOF or rsget.BOF) then

			FResultCount = 1
			redim preserve FItemList(1)
			set FItemList(1) = new CPartItem

			FItemList(1).Fpart_name		= rsget("part_name")
			FItemList(1).Fpart_sort		= rsget("part_sort")
			FItemList(1).Fpart_isDel	= rsget("part_isDel")
		else
			FResultCount = 0
		end if

		rsget.Close

	end Sub
	
	''부서리스트 가져오기 - 운영비 팀관리에서 사용
	public Function fnGetPartInfoList
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_partInfo_getList]"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPartInfoList =  rsget.getRows()
		END IF
		rsget.close
	End Function
end Class
%>