<%
'###########################################################
' Description : [cs]공통코드관리 클래스
' Hieditor : 이상구 생성
'			 2023.08.28 한용민 수정(쿼리튜닝, 고객노출여부 추가)
'###########################################################

Class CAsGubunHelpItem
    public Fdiv_comm_name
    public Fdiv_comm_cd
    public Fstate_comm_cd
    public FinfoHtml
    
    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CCommCdItem
	public Fcomm_cd
	public Fcomm_name
	public Fcomm_group
	public Fgroup_name
	public Fcomm_isDel
    public Fcomm_color
	public Fsortno
	public fdispyn
    
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CCommCd
	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectCommCd
	public FRectGroupCd
	public FRectsearchKey
	public FRectsearchString
	public FRectisDel
	public FRectdispyn
	public FSortType

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
    
    public Sub GetCommHelpStr()
        dim sqlStr, i
        
        sqlStr = " select top 10 c.comm_name, h.* "
        sqlStr = sqlStr + " from db_cs.dbo.tbl_cs_comm_code c with (nolock)"
        sqlStr = sqlStr + " left join db_cs.dbo.tbl_cs_comm_div_info h with (nolock)"
        sqlStr = sqlStr + " 	on c.comm_cd=h.div_comm_cd"
        sqlStr = sqlStr + " where c.comm_cd='" + FRectCommCd + "'"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CAsGubunHelpItem

				FItemList(i).Fdiv_comm_name = db2html(rsget("comm_name"))
                FItemList(i).Fdiv_comm_cd   = rsget("div_comm_cd")
                FItemList(i).Fstate_comm_cd = rsget("state_comm_cd")
                FItemList(i).FinfoHtml      = db2html(rsget("infoHtml"))

                
				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.Close
    end Sub

	public Sub GetCommList()
		dim SQL, AddSQL, sortSQL, i

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and t1." & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectGroupCd<>"" then
			AddSQL = AddSQL & " and t1.comm_group='" & FRectGroupCd & "' "
		end if

		if FRectisDel<>"" then
			AddSQL = AddSQL & " and t1.comm_isDel='" & FRectisDel & "' "
		end if
		if FRectdispyn<>"" then
			AddSQL = AddSQL & " and t1.dispyn='" & FRectdispyn & "' "
		end if

		Select Case FSortType
			Case "ca"
				sortSQL = " ORDER BY t1.comm_cd ASC "
			Case "cd"
				sortSQL = " ORDER BY t1.comm_cd DESC "
			Case "sa"
				sortSQL = " ORDER BY t1.sortno ASC, t1.comm_cd ASC "
			Case "sd"
				sortSQL = " ORDER BY t1.sortno DESC, t1.comm_cd ASC "
			Case "ga"
				sortSQL = " ORDER BY t1.comm_group ASC, t1.comm_cd ASC "
			Case "gd"
				sortSQL = " ORDER BY t1.comm_group DESC, t1.comm_cd ASC "
			Case else
				sortSQL = " ORDER BY t1.comm_cd ASC "
		End Select

		SQL = "Select count(t1.comm_cd), CEILING(CAST(Count(t1.comm_cd) AS FLOAT)/" & FPageSize & ")"
		SQL = SQL & " From db_cs.dbo.tbl_cs_comm_code as t1 with (nolock)"
		SQL = SQL & " where 1=1 " & AddSQL

		'response.write SQL & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		SQL = "Select top " + CStr(FPageSize*FCurrPage)
		SQL = SQL & " t1.comm_cd, t1.comm_name, t1.comm_group, t1.comm_color"
		SQL = SQL & " ,Case t1.comm_isDel When 'N' Then '<font color=darkblue>사용</font>' When 'Y' Then '<font color=darkred>삭제</font>' End comm_isDel"
		SQL = SQL & " ,(Select t2.comm_name From db_cs.dbo.tbl_cs_comm_code as t2 with(noLock) Where t2.comm_cd=t1.comm_group) as group_name"
		SQL = SQL & " , t1.sortno, t1.dispyn"
		SQL = SQL & " From db_cs.dbo.tbl_cs_comm_code as t1 with(noLock)"
		SQL = SQL & " where 1=1 " & AddSQL & sortSQL

		'response.write SQL & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCommCdItem

				FItemList(i).Fcomm_cd		= rsget("comm_cd")
				FItemList(i).Fcomm_name		= rsget("comm_name")
				FItemList(i).Fcomm_group	= rsget("comm_group")
				FItemList(i).Fgroup_name	= rsget("group_name")
				FItemList(i).Fcomm_isDel    = rsget("comm_isDel")
                FItemList(i).Fcomm_color    = rsget("comm_color")
				FItemList(i).Fsortno    = rsget("sortno")
				FItemList(i).fdispyn    = rsget("dispyn")
                
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

	public Sub GetCommRead()
		dim SQL

		SQL ="Select top " + CStr(FPageSize*FCurrPage)
		SQL = SQL & " t1.comm_cd, t1.comm_name, t1.comm_group, t1.comm_color"
		SQL = SQL & " ,Case t1.comm_isDel When 'N' Then '사용' When 'Y' Then '삭제' End comm_isDel"
		SQL = SQL & " ,(Select t2.comm_name From db_cs.dbo.tbl_cs_comm_code as t2 with (nolock) Where t2.comm_cd=t1.comm_group) as group_name"
		SQL = SQL & " ,t1.sortno, t1.dispyn"
		SQL = SQL & " From db_cs.dbo.tbl_cs_comm_code as t1 with (nolock)"
		SQL = SQL & " where t1.comm_cd = '" & FRectCommCd & "'"

		'response.write SQL & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		redim FItemList(0)

		if Not(rsget.EOF or rsget.BOF) then
			set FItemList(0) = new CCommCdItem

			FItemList(0).FComm_Cd		= rsget("Comm_Cd")
			FItemList(0).Fcomm_Name		= rsget("comm_Name")
			FItemList(0).Fcomm_group	= rsget("comm_group")
			FItemList(0).Fgroup_Name	= rsget("group_Name")
			FItemList(0).Fcomm_isDel	= rsget("comm_isDel")
            FItemList(0).Fcomm_color    = rsget("comm_color")
			FItemList(0).Fsortno	    = rsget("sortno")
			FItemList(0).fdispyn	    = rsget("dispyn")
		end if
		rsget.close
	end sub

	'// 그룹 옵션 생성 //
	function optGroupCd(nowCd)
		dim SQL, strOpt

		SQL = "Select comm_cd, comm_name From db_cs.dbo.tbl_cs_comm_code with (nolock) Where comm_isDel='N' and comm_group='Z999' "

		'response.write SQL & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
		if Not(rsget.EOF or rsget.BOF) then
			Do Until rsget.EOF
				strOpt = strOpt & "<option value='" & rsget("comm_cd") & "' "

				if nowCd=rsget("comm_cd") then strOpt = strOpt & "selected"

				strOpt = strOpt & " >" & rsget("comm_name") & "</option>"
				rsget.MoveNext
			Loop
		end if
		rsget.Close

		optGroupCd = strOpt
	end function
end Class
%>