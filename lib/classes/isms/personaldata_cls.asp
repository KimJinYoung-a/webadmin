<%
'####################################################
' Description : 개인정보 문서 파기 관리 클래스
' History : 2018.09.12 한용민 생성
'####################################################

class Cpersonaldata_item
	public Fidx
	public Flogtype
	public Fqryuserid
	public Frefip
	public Fscrname
	public FqryStr
	public FqryMethod
	public Fregdate
	public FdownFileGubun
	public FdownFilemenupos
	public FdownFileDelYN
	public FdownFileDelDate
	public FdownFileconfirmYN
	public FdownFileconfirmDelDate
    public fmenuname

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class Cpersonaldata
	public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectStartdate
    public FRectEnddate
    public frectlogtype
    public frectdownFileGubun
    public FRectqryuserid
    public FRectdownFileDelYN
    public FRectdownFileconfirmYN

    ' /partner/isms/personaldata_list.asp	' /admin/isms/personaldata_list.asp
	public function GetpersonaldataList()
		dim sqlStr,i , sqlsearch

        if FRectdownFileDelYN<>"" then
            sqlsearch = sqlsearch & " and l.downFileDelYN = '" & FRectdownFileDelYN & "'" & vbcrlf
        end if
        if FRectdownFileconfirmYN<>"" then
            sqlsearch = sqlsearch & " and l.downFileconfirmYN = '" & FRectdownFileconfirmYN & "'" & vbcrlf
        end if
        if FRectqryuserid<>"" then
            sqlsearch = sqlsearch & " and l.qryuserid = '" & FRectqryuserid & "'" & vbcrlf
        end if
        if frectdownFileGubun<>"" then
            sqlsearch = sqlsearch & " and isnull(l.downFileGubun,'') <> ''" & vbcrlf
        end if
        if frectlogtype<>"" then
            sqlsearch = sqlsearch & " and l.logtype = '" & frectlogtype & "'" & vbcrlf
        end if
        if FRectStartdate<>"" and FRectEnddate<>"" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch & " and l.regdate >= '" & FRectStartdate & "'" & vbcrlf
			end if
			if FRectEnddate<>"" then
				sqlsearch = sqlsearch & " and l.regdate < '" & FRectEnddate & "'" & vbcrlf
			end if
        end if

		sqlStr = "select count(l.idx) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_log.dbo.tbl_ChkAllowIpLog l with (nolock)" & vbcrlf
        sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner_menu m with (nolock)" & vbcrlf
        sqlStr = sqlStr & "     on l.downFilemenupos = m.id" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" & vbcrlf
		sqlStr = sqlStr & " l.idx, l.logtype, l.qryuserid, l.refip, l.scrname, l.qryStr, l.qryMethod, l.regdate, l.downFileGubun, l.downFilemenupos" & vbcrlf
		sqlStr = sqlStr & " , l.downFileDelYN, l.downFileDelDate, l.downFileconfirmYN, l.downFileconfirmDelDate" & vbcrlf
        sqlStr = sqlStr & " , m.menuname" & vbcrlf
		sqlStr = sqlStr & " from db_log.dbo.tbl_ChkAllowIpLog l with (nolock)" & vbcrlf
        sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner_menu m with (nolock)" & vbcrlf
        sqlStr = sqlStr & "     on l.downFilemenupos = m.id" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx desc" & vbcrlf

		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Cpersonaldata_item

				FItemList(i).fidx = rsget("idx")
				FItemList(i).flogtype = rsget("logtype")
				FItemList(i).fqryuserid = rsget("qryuserid")
				FItemList(i).frefip = rsget("refip")
				FItemList(i).fscrname = rsget("scrname")
				FItemList(i).fqryStr = rsget("qryStr")
				FItemList(i).fqryMethod = rsget("qryMethod")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fdownFileGubun = rsget("downFileGubun")
				FItemList(i).fdownFilemenupos = rsget("downFilemenupos")
				FItemList(i).fdownFileDelYN = rsget("downFileDelYN")
				FItemList(i).fdownFileDelDate = rsget("downFileDelDate")
				FItemList(i).fdownFileconfirmYN = rsget("downFileconfirmYN")
				FItemList(i).fdownFileconfirmDelDate = rsget("downFileconfirmDelDate")
                FItemList(i).fmenuname = rsget("menuname")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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