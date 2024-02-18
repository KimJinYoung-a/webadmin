<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2016.06.30 한용민 수정
'###########################################################

Class CCSCenterBoardUserItem
    public Findexno
    public Fuserid
    public Fvacationyn
	public fvvipone2oneyn
	public Fvipone2oneyn
    public Fone2oneyn
    public Fmichulgoyn
    public Fstockoutyn
    public Freturnyn
    public Fuseyn
    public Flastupdate

    Private Sub Class_Initialize()
    End Sub
    Private Sub Class_Terminate()
    End Sub
end Class

Class CCSCenterBoardUser
    public FItemList()
    public FOneItem
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public Sub GetCSCenterBoardUserList()
        dim i,sqlStr

        sqlStr = "select top 300"
        sqlStr = sqlStr & " u.indexno, u.userid, 'N' as vacationyn, u.vipone2oneyn, u.vvipone2oneyn, u.one2oneyn, u.michulgoyn"
        sqlStr = sqlStr & " , u.stockoutyn, u.returnyn, u.useyn, u.lastupdate "
        sqlStr = sqlStr & " from db_cs.dbo.tbl_cs_board_user u "
        sqlStr = sqlStr & " order by indexno "

        'response.write sqlStr & "<br>"
        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount
        FTotalCount = FResultCount

        redim preserve FItemList(FResultCount)
        if  not rsget.EOF  then
            i = 0
            do until rsget.eof
				set FItemList(i) = new CCSCenterBoardUserItem
					FItemList(i).Findexno        	= rsget("indexno")
					FItemList(i).Fuserid            = rsget("userid")
					FItemList(i).Fvacationyn        = rsget("vacationyn")
					FItemList(i).Fvipone2oneyn      = rsget("vipone2oneyn")
					FItemList(i).fvvipone2oneyn      = rsget("vvipone2oneyn")
					FItemList(i).Fone2oneyn         = rsget("one2oneyn")
					FItemList(i).Fmichulgoyn        = rsget("michulgoyn")
					FItemList(i).Fstockoutyn        = rsget("stockoutyn")
					FItemList(i).Freturnyn        	= rsget("returnyn")
					FItemList(i).Fuseyn             = rsget("useyn")
					FItemList(i).Flastupdate        = rsget("lastupdate")
				rsget.MoveNext
				i = i + 1
            loop
        end if
        rsget.close
    end sub

    Private Sub Class_Initialize()
        FCurrPage       = 1
        FPageSize       = 20
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
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
%>
