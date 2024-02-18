<%
'###########################################################
' Description : 송장대역관리
' Hieditor : 2021.04.14 한용민 생성
'###########################################################

Class cinvoice_band_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fiidx
	public fsiteseq
	public fgubuncd
	public fstartsongjangno
	public fendsongjangno
	public fstartrealsongjangno
	public fendrealsongjangno
	public fremainsongjangcount
	public fbasicsongjangyn
	public fisusing
	public fregdate
	public flastupdate
	public freguserid
	public flastuserid
	public fSONGJANGNO
	public fREALSONGJANGNO
	public fORDERSERIAL
    public Fsongjangdiv
    public Fdivname
end class

class cinvoice_band_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

    public frectsiteseq
    public frectisusing
	public frectiidx
	public Frectgubuncd
    public FRectSongjangDiv

	public fendrealsongjangno
	public fcurrentbasicsongjangidx
	public tendb

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		fendrealsongjangno =0
		IF application("Svr_Info")<>"Dev" THEN
			tendb="tendb."
		end if
	End Sub
	Private Sub Class_Terminate()
	End Sub

    ' /admin/logics/invoice_band.asp
	public sub finvoice_band()
		dim sqlStr,i, sqlsearch

        if frectsiteseq <> "" then
            sqlsearch = sqlsearch & " and b.siteseq = "& frectsiteseq &""
        end if
        if frectisusing <> "" then
            sqlsearch = sqlsearch & " and b.isusing = '"& frectisusing &"'"
        end if
        if Frectgubuncd <> "" then
            sqlsearch = sqlsearch & " and b.gubuncd = '"& Frectgubuncd &"'"
        end if
        if FRectSongjangDiv <> "" then
            sqlsearch = sqlsearch & " and b.songjangdiv = '"& FRectSongjangDiv &"'"
        end if

		sqlStr = "select count(iidx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_invoice_band] b with (nolock)" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget_Logistics("cnt")
		rsget_Logistics.Close

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " b.iidx, b.siteseq, b.gubuncd, b.startsongjangno, b.endsongjangno, b.startrealsongjangno, b.endrealsongjangno"
		sqlStr = sqlStr & " , b.remainsongjangcount, b.basicsongjangyn, b.isusing, b.regdate, b.lastupdate, b.reguserid, b.lastuserid, b.songjangdiv "
		sqlStr = sqlStr & " , (select divname from "& tendb &"db_order.[dbo].tbl_songjang_div with (nolock) where divcd=b.songjangdiv and isusing='Y') as divname"
		sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_invoice_band] b with (nolock)" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by iidx Desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget_Logistics.pagesize = FPageSize
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget_Logistics.EOF  then
			rsget_Logistics.absolutepage = FCurrPage
			do until rsget_Logistics.EOF
				set FItemList(i) = new cinvoice_band_oneitem

				FItemList(i).fiidx = rsget_Logistics("iidx")
				FItemList(i).fsiteseq = rsget_Logistics("siteseq")
				FItemList(i).fgubuncd = rsget_Logistics("gubuncd")
				FItemList(i).fstartsongjangno = rsget_Logistics("startsongjangno")
				FItemList(i).fendsongjangno = rsget_Logistics("endsongjangno")
				FItemList(i).fstartrealsongjangno = rsget_Logistics("startrealsongjangno")
				FItemList(i).fendrealsongjangno = rsget_Logistics("endrealsongjangno")
				FItemList(i).fremainsongjangcount = rsget_Logistics("remainsongjangcount")
				FItemList(i).fbasicsongjangyn = rsget_Logistics("basicsongjangyn")
				FItemList(i).fisusing = rsget_Logistics("isusing")
				FItemList(i).fregdate = rsget_Logistics("regdate")
				FItemList(i).flastupdate = rsget_Logistics("lastupdate")
				FItemList(i).freguserid = rsget_Logistics("reguserid")
				FItemList(i).flastuserid = rsget_Logistics("lastuserid")
                FItemList(i).fsongjangdiv = rsget_Logistics("songjangdiv")
				FItemList(i).fdivname = rsget_Logistics("divname")

				if FItemList(i).fbasicsongjangyn="Y" then
					fendrealsongjangno=FItemList(i).fendrealsongjangno
					fcurrentbasicsongjangidx=FItemList(i).fiidx
				end if

				rsget_Logistics.movenext
				i=i+1
			loop
		end if
		rsget_Logistics.Close
	end sub

    ' /admin/logics/invoice_band_reg.asp
    public Sub finvoice_band_one()
        dim sqlStr, sqlsearch

		if frectiidx="" or isnull(frectiidx) then exit Sub

        if frectiidx <> "" then
            sqlsearch = sqlsearch & " and iidx = "& frectiidx &""
        end if

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " b.iidx, b.siteseq, b.gubuncd, b.startsongjangno, b.endsongjangno, b.startrealsongjangno, b.endrealsongjangno"
		sqlStr = sqlStr & " , b.remainsongjangcount, b.basicsongjangyn, b.isusing, b.regdate, b.lastupdate, b.reguserid, b.lastuserid, b.songjangdiv "
		sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_invoice_band] b with (nolock)" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly
        FResultCount = rsget_Logistics.RecordCount
        FTotalCount = rsget_Logistics.RecordCount
        set FOneItem = new cinvoice_band_oneitem

        if Not rsget_Logistics.Eof then
			FOneItem.fiidx = rsget_Logistics("iidx")
			FOneItem.fsiteseq = rsget_Logistics("siteseq")
			FOneItem.fgubuncd = rsget_Logistics("gubuncd")
			FOneItem.fstartsongjangno = rsget_Logistics("startsongjangno")
			FOneItem.fendsongjangno = rsget_Logistics("endsongjangno")
			FOneItem.fstartrealsongjangno = rsget_Logistics("startrealsongjangno")
			FOneItem.fendrealsongjangno = rsget_Logistics("endrealsongjangno")
			FOneItem.fremainsongjangcount = rsget_Logistics("remainsongjangcount")
			FOneItem.fbasicsongjangyn = rsget_Logistics("basicsongjangyn")
			FOneItem.fisusing = rsget_Logistics("isusing")
			FOneItem.fregdate = rsget_Logistics("regdate")
			FOneItem.flastupdate = rsget_Logistics("lastupdate")
			FOneItem.freguserid = rsget_Logistics("reguserid")
			FOneItem.flastuserid = rsget_Logistics("lastuserid")
            FOneItem.Fsongjangdiv = rsget_Logistics("songjangdiv")
        end if
        rsget_Logistics.Close
    end Sub

    ' /admin/logics/invoice_band.asp
	public sub finvoice_band_log()
		dim sqlStr,i, sqlsearch

        if frectsiteseq <> "" then
            sqlsearch = sqlsearch & " and siteseq = "& frectsiteseq &""
        end if
        if Frectgubuncd <> "" then
			if Frectgubuncd = "etc" then
				sqlsearch = sqlsearch & " and gubuncd <> '00'"
			else
            	sqlsearch = sqlsearch & " and gubuncd = '"& Frectgubuncd &"'"
			end if
        end if

        if FRectSongjangDiv <> "" then
            sqlsearch = sqlsearch & " and div_cd = '"& FRectSongjangDiv &"'"
        end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " l.SONGJANGNO, l.REALSONGJANGNO, l.ORDERSERIAL"
		sqlStr = sqlStr & " from db_aLogistics.dbo.tbl_Logistics_songjang_log l with (nolock)" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by l.idx Desc" + vbcrlf

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
				set FItemList(i) = new cinvoice_band_oneitem

				FItemList(i).fSONGJANGNO = rsget_Logistics("SONGJANGNO")
				FItemList(i).fREALSONGJANGNO = rsget_Logistics("REALSONGJANGNO")
				FItemList(i).fORDERSERIAL = rsget_Logistics("ORDERSERIAL")

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
