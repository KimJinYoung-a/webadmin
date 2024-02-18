<%
'###########################################################
' Description : 운송장전송주소오류관리
' Hieditor : 2022.06.27 한용민 생성
'###########################################################

Class cSongJangSendErrorOneItem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fidx
    public fSiteSEQ
    public fDIV_CD
    public fdivname
    public fSONGJANGNO
    public fGUBUNCD
    public fORDERSERIAL
    public fISUPLOADED
    public fnm
    public fTEL_NO
    public fHP_NO
    public fZIP_NO
    public fADDR
    public fADDR_ETC
    public fREGDATE
    public fonlinereqzipcode
    public fonlinereqzipaddr
    public fonlinereqaddress
end class

class cSongJangSendError
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

    public frectidx
	public FrectSongJangGubun
    public Frectsiteseq
    public FRectSongjangDiv
    public Frectgubuncd
	public tendb

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		IF application("Svr_Info")<>"Dev" THEN
			tendb="tendb."
		end if
	End Sub
	Private Sub Class_Terminate()
	End Sub

    public Sub GetSongJangSendErrorOne()
        dim sqlStr, sqlsearch

        if FrectSongJangGubun="" or isnull(FrectSongJangGubun) then exit sub

        if Frectgubuncd <> "" then
			if Frectgubuncd = "etc" then
				sqlsearch = sqlsearch & " and l.gubuncd <> '00'"
			else
            	sqlsearch = sqlsearch & " and l.gubuncd = '"& Frectgubuncd &"'"
			end if
        end if
        if FRectSongjangDiv <> "" then
            sqlsearch = sqlsearch & " and l.DIV_CD = '"& FRectSongjangDiv &"'"
        end if
        if frectsiteseq <> "" then
            sqlsearch = sqlsearch & " and l.SiteSEQ = "& frectsiteseq &""
        end if
        if frectidx <> "" then
            sqlsearch = sqlsearch & " and l.idx = "& frectidx &""
        end if

		sqlStr = "select top 1"
		sqlStr = sqlStr & " l.idx, l.SiteSEQ, l.DIV_CD"
		sqlStr = sqlStr & " , (select divname from "& tendb &"db_order.[dbo].tbl_songjang_div with (nolock)"
		sqlStr = sqlStr & " 	where divcd=DIV_CD and isusing='Y') as divname"
		sqlStr = sqlStr & " , l.SONGJANGNO, l.GUBUNCD, l.ORDERSERIAL, l.ISUPLOADED"
		sqlStr = sqlStr & " , l.nm,l.TEL_NO, l.HP_NO"
        sqlStr = sqlStr & " , replace(replace(replace(replace(replace(l.ZIP_NO,char(9),''),char(10),''),char(13),''),'""',''),'''','') as ZIP_NO"
        sqlStr = sqlStr & " , replace(replace(replace(replace(replace(l.ADDR,char(9),''),char(10),''),char(13),''),'""',''),'''','') as ADDR"
        sqlStr = sqlStr & " , replace(replace(replace(replace(replace(l.ADDR_ETC,char(9),''),char(10),''),char(13),''),'""',''),'''','') as ADDR_ETC"
        sqlStr = sqlStr & " , l.REGDATE"
        'sqlStr = sqlStr & " , replace(replace(replace(replace(replace(m.reqzipcode,char(9),''),char(10),''),char(13),''),'""',''),'''','') as onlinereqzipcode"
        'sqlStr = sqlStr & " , replace(replace(replace(replace(replace(m.reqzipaddr,char(9),''),char(10),''),char(13),''),'""',''),'''','') as onlinereqzipaddr"
        'sqlStr = sqlStr & " , replace(replace(replace(replace(replace(m.reqaddress,char(9),''),char(10),''),char(13),''),'""',''),'''','') as onlinereqaddress"

        if FrectSongJangGubun = "GENERAL" then
		    sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_Logistics_songjang_log] as l with (nolock)"
        elseif FrectSongJangGubun = "RETURN" then
            sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_Logistics_songjang_log_return] as l with (nolock)"
        end if

        'sqlStr = sqlStr & " left join db_order.dbo.tbl_order_master m with (nolock)"
        'sqlStr = sqlStr & "     on l.ORDERSERIAL=m.ORDERSERIAL"

		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly
        FResultCount = rsget_Logistics.RecordCount
        FTotalCount = rsget_Logistics.RecordCount
        set FOneItem = new cSongJangSendErrorOneItem

        if Not rsget_Logistics.Eof then
			FOneItem.fidx = rsget_Logistics("idx")
            FOneItem.fSiteSEQ = rsget_Logistics("SiteSEQ")
            FOneItem.fDIV_CD = rsget_Logistics("DIV_CD")
            FOneItem.fdivname = rsget_Logistics("divname")
            FOneItem.fSONGJANGNO = rsget_Logistics("SONGJANGNO")
            FOneItem.fGUBUNCD = rsget_Logistics("GUBUNCD")
            FOneItem.fORDERSERIAL = rsget_Logistics("ORDERSERIAL")
            FOneItem.fISUPLOADED = rsget_Logistics("ISUPLOADED")
            FOneItem.fnm = rsget_Logistics("nm")
            FOneItem.fTEL_NO = rsget_Logistics("TEL_NO")
            FOneItem.fHP_NO = rsget_Logistics("HP_NO")
            FOneItem.fZIP_NO = rsget_Logistics("ZIP_NO")
            FOneItem.fADDR = rsget_Logistics("ADDR")
            FOneItem.fADDR_ETC = rsget_Logistics("ADDR_ETC")
            FOneItem.fREGDATE = rsget_Logistics("REGDATE")
            'FOneItem.fonlinereqzipcode = rsget_Logistics("onlinereqzipcode")
            'FOneItem.fonlinereqzipaddr = rsget_Logistics("onlinereqzipaddr")
            'FOneItem.fonlinereqaddress = rsget_Logistics("onlinereqaddress")
        end if
        rsget_Logistics.Close
    end Sub

    ' /admin/logics/songjang/SongJangSendErrorList.asp
	public sub GetSongJangSendErrorList()
		dim sqlStr,i, sqlsearch

        if FrectSongJangGubun="" or isnull(FrectSongJangGubun) then exit sub

        if Frectgubuncd <> "" then
			if Frectgubuncd = "etc" then
				sqlsearch = sqlsearch & " and l.gubuncd <> '00'"
			else
            	sqlsearch = sqlsearch & " and l.gubuncd = '"& Frectgubuncd &"'"
			end if
        end if
        if FRectSongjangDiv <> "" then
            sqlsearch = sqlsearch & " and l.DIV_CD = '"& FRectSongjangDiv &"'"
        end if
        if frectsiteseq <> "" then
            sqlsearch = sqlsearch & " and l.SiteSEQ = "& frectsiteseq &""
        end if
        if frectidx <> "" then
            sqlsearch = sqlsearch & " and l.idx = "& frectidx &""
        end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " l.idx, l.SiteSEQ, l.DIV_CD"
		sqlStr = sqlStr & " , (select divname from "& tendb &"db_order.[dbo].tbl_songjang_div with (nolock)"
		sqlStr = sqlStr & " 	where divcd=DIV_CD and isusing='Y') as divname"
		sqlStr = sqlStr & " , l.SONGJANGNO, l.GUBUNCD, l.ORDERSERIAL, l.ISUPLOADED"
		sqlStr = sqlStr & " , l.nm,l.TEL_NO, l.HP_NO"
        sqlStr = sqlStr & " , replace(replace(replace(replace(replace(l.ZIP_NO,char(9),''),char(10),''),char(13),''),'""',''),'''','') as ZIP_NO"
        sqlStr = sqlStr & " , replace(replace(replace(replace(replace(l.ADDR,char(9),''),char(10),''),char(13),''),'""',''),'''','') as ADDR"
        sqlStr = sqlStr & " , replace(replace(replace(replace(replace(l.ADDR_ETC,char(9),''),char(10),''),char(13),''),'""',''),'''','') as ADDR_ETC"
        sqlStr = sqlStr & " , l.REGDATE"

        if FrectSongJangGubun = "GENERAL" then
		    sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_Logistics_songjang_log] as l with (nolock)"
        elseif FrectSongJangGubun = "RETURN" then
            sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_Logistics_songjang_log_return] as l with (nolock)"
        end if

		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " and l.ISUPLOADED='X'"
		sqlStr = sqlStr & " order by l.idx desc"

		'response.write sqlStr &"<br>"
		rsget_Logistics.pagesize = FPageSize
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

		FResultCount=rsget_Logistics.recordcount
		FtotalCount=rsget_Logistics.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget_Logistics.EOF  then
			rsget_Logistics.absolutepage = FCurrPage
			do until rsget_Logistics.EOF
				set FItemList(i) = new cSongJangSendErrorOneItem

				FItemList(i).fidx = rsget_Logistics("idx")
				FItemList(i).fSiteSEQ = rsget_Logistics("SiteSEQ")
				FItemList(i).fDIV_CD = rsget_Logistics("DIV_CD")
				FItemList(i).fdivname = rsget_Logistics("divname")
				FItemList(i).fSONGJANGNO = rsget_Logistics("SONGJANGNO")
				FItemList(i).fGUBUNCD = rsget_Logistics("GUBUNCD")
				FItemList(i).fORDERSERIAL = rsget_Logistics("ORDERSERIAL")
				FItemList(i).fISUPLOADED = rsget_Logistics("ISUPLOADED")
				FItemList(i).fnm = rsget_Logistics("nm")
				FItemList(i).fTEL_NO = rsget_Logistics("TEL_NO")
				FItemList(i).fHP_NO = rsget_Logistics("HP_NO")
				FItemList(i).fZIP_NO = rsget_Logistics("ZIP_NO")
                FItemList(i).fADDR = rsget_Logistics("ADDR")
                FItemList(i).fADDR_ETC = rsget_Logistics("ADDR_ETC")
                FItemList(i).fREGDATE = rsget_Logistics("REGDATE")

				rsget_Logistics.movenext
				i=i+1
			loop
		end if
		rsget_Logistics.Close
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

' 송장구분      ' 2022.06.27 한용민 생성
sub drawSelectBoxSongJangGubun(gubuncdName, gubuncdValue, chplg)
%>
    <select class="select" name="<%= gubuncdName %>" <%= chplg %> >
        <option value="">선택</option>
        <option value="GENERAL" <% if gubuncdValue="GENERAL" then response.write "selected" %> >일반송장</option>
        <option value="RETURN" <% if gubuncdValue="RETURN" then response.write "selected" %> >반품송장</option>
    </select>
<%
End sub

' 송장구분      ' 2022.06.27 한용민 생성
function getSongJangGubun(SongJangGubun)
    dim tmpSongJangGubun

    if SongJangGubun="GENERAL" then
        tmpSongJangGubun="일반송장"
    elseif SongJangGubun="RETURN" then
        tmpSongJangGubun="반품송장"
    else
        tmpSongJangGubun=""
    end if

    getSongJangGubun=tmpSongJangGubun
End function

' 물류센터 출고구분      '/2021.04.14 한용민 생성
function getgubuncdname(gubuncd)
    dim tmpgubuncdname

    if gubuncd="00" then
        tmpgubuncdname="온라인출고"
    elseif gubuncd="etc" then
        tmpgubuncdname="기타출고"
    else
        tmpgubuncdname=""
    end if

    getgubuncdname=tmpgubuncdname
End function

' 물류센터 출고구분      '/2021.04.15 한용민 생성
sub drawSelectBoxgubuncd(gubuncdName, gubuncdValue, chplg)
%>
    <select class="select" name="<%= gubuncdName %>" <%= chplg %> >
        <option value="">선택</option>
        <option value="00" <% if gubuncdValue="00" then response.write "selected" %> >온라인출고</option>
        <option value="etc" <% if gubuncdValue="etc" then response.write "selected" %> >기타출고</option>
    </select>
<%
End sub
%>
