<%

class CTplJungsanMasterItem
	public Fidx
    public Ftplcompanyid
    public Fyyyymm
    public Ftitle
    public Fst_totalcash
    public Fio_totalcash
    public Fet_totalcash
    public Fregdate
    public Fcancelyn
    public Ffinishflag
    public Fipkumdate
    public Ftaxregdate
    public Ftaxinputdate
    public Ftaxtype
    public Fdifferencekey
    public Ftaxlinkidx
    public Fneotaxno
    public Fbankingupflag
    public Fgroupid
    public FitemvatYn

    Public Fcompany_name
    public Fjungsan_hp
    public Fjungsan_email
    public Fjungsan_gubun
    public Fjungsan_date

	public function GetSimpleTaxtypeName()
		if Ftaxtype="01" then
			GetSimpleTaxtypeName = "과세"
		elseif Ftaxtype="02" then
			GetSimpleTaxtypeName = "면세"
		elseif Ftaxtype="03" then
			GetSimpleTaxtypeName = "원천" '''"간이"
		end if
	end function

	public function GetTaxtypeNameColor()
		if Ftaxtype="01" then
			GetTaxtypeNameColor = "#000000"
		elseif Ftaxtype="02" then
			GetTaxtypeNameColor = "#FF3333"
		elseif Ftaxtype="03" then
			GetTaxtypeNameColor = "#3333FF"
		end if
	end function

	public function GetStateName()
		if Ffinishflag="0" then
			GetStateName = "수정중"
		elseif Ffinishflag="1" then
		    IF isNULL(FISSU_SEQNO) or (FISSU_SEQNO="") then
		        GetStateName = "업체확인대기"
		    ELSE
			    GetStateName = "역발행등록중"
			ENd IF
		elseif Ffinishflag="2" then
		    IF isNULL(FISSU_SEQNO) or (FISSU_SEQNO="") then
		        GetStateName = "업체확인완료"
		    ELSE
			    GetStateName = "업체역발행대기"
			ENd IF
		elseif Ffinishflag="3" then
			GetStateName = "정산확정"
		elseif Ffinishflag="7" then
			GetStateName = "입금완료"
		else

		end if
	end function

	public function GetStateColor()
		if Ffinishflag="0" then
			GetStateColor = "#000000"
		elseif Ffinishflag="1" then
		    IF isNULL(FISSU_SEQNO) or (FISSU_SEQNO="") then
		        GetStateColor = "#448888"
		    ELSE
			    GetStateColor = "#884488"
		    ENd IF
		elseif Ffinishflag="2" then
		    IF isNULL(FISSU_SEQNO) or (FISSU_SEQNO="") then
			    GetStateColor = "#0000FF"
			ELSE
			    GetStateColor = "#AA44AA"
		    ENd IF
		elseif Ffinishflag="3" then
			GetStateColor = "#0000FF"
		elseif Ffinishflag="7" then
			GetStateColor = "#FF0000"
		else

		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CTplJungsanDetailItem
	public Fidx
    public Fmasteridx
    public Fgubuncd
    public Fgubunname
    public Fgubundetailname
    public Ftypename
    public Funitprice
    public Favgcbm
    public Fcurrcbm
    public Fprevcbm
    public Fitemno
    public FtotPrice
    public Fmastercode
    public Fcomment

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CTplJungsanCbmItem
	public Fidx
    public Fmasteridx
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public Fbarcode
    public Fitemname
    public Fitemoptionname
    public Fitemno
    public FcbmX
    public FcbmY
    public FcbmZ

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CTplJungsanEtcItem
    ''idx, masteridx, gubuncd, gubunname, gubundetailname, typename, unitprice, itemno, totPrice, mastercode, comment
	public Fidx
    public Fmasteridx
    public Fgubuncd
    public Fgubunname
    public Fgubundetailname
    public Ftypename
    public Funitprice
    public Fitemno
    public FtotPrice
    public Fmastercode
    public Fcomment

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CTplJungsanGubunDetailItem
    public Fgubuncd
    public Fgubunname
    public Fgubundetailname
    public Ftypename
    public Funitprice

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CTplJungsan
    public FItemList()
    public FOneItem
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectGubun
    public FRectYYYYMM
    public FRectTplCompanyID
    public FRectCancelYN
    public FRectIdx
    public FRectMasterIdx

	public Sub GetTPLJungsanMasterList()
		dim i,sqlStr, addSql

		addSql = ""

		if (FRectIdx <> "") then
			addSql = addSql & " and m.idx = '" & FRectIdx & "'" & vbcrlf
		end if

		if (FRectYYYYMM <> "") then
			addSql = addSql & " and m.yyyymm = '" & FRectYYYYMM & "'" & vbcrlf
		end if

		if (FRectTplCompanyID <> "") then
			addSql = addSql & " and m.tplcompanyid = '" & FRectTplCompanyID & "'" & vbcrlf
		end if

		if (FRectCancelYN <> "") then
			addSql = addSql & " and m.cancelyn = '" & FRectCancelYN & "'" & vbcrlf
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_tpl_jungsan_master] m " & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & addSql
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.Open sqlStr,dbget_TPL,1
			FTotalCount = rsget_TPL("cnt")
			FTotalPage = rsget_TPL("totPg")
		rsget_TPL.Close


		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If


        sqlStr = " select top " & CStr(FPageSize*FCurrPage) & vbcrlf
        sqlStr = sqlStr & " m.* " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_tpl_jungsan_master] m " & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by m.idx desc " & vbcrlf
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget_TPL.EOF Then
			rsget_TPL.absolutepage = FCurrPage
			Do until rsget_TPL.EOF
				Set FItemList(i) = new CTplJungsanMasterItem

                	''idx, tplcompanyid, yyyymm, title, st_totalcash, io_totalcash, et_totalcash, regdate, cancelyn, finishflag, ipkumdate, taxregdate, taxinputdate, taxtype, differencekey, taxlinkidx, neotaxno, bankingupflag, groupid, itemvatYn
                    ''company_name, jungsan_hp, jungsan_email

					FItemList(i).Fidx				= rsget_TPL("idx")
                    FItemList(i).Ftplcompanyid		= rsget_TPL("tplcompanyid")
                    FItemList(i).Fyyyymm			= rsget_TPL("yyyymm")
                    FItemList(i).Ftitle				= rsget_TPL("title")
                    FItemList(i).Fst_totalcash		= rsget_TPL("st_totalcash")
                    FItemList(i).Fio_totalcash		= rsget_TPL("io_totalcash")
                    FItemList(i).Fet_totalcash		= rsget_TPL("et_totalcash")
                    FItemList(i).Fregdate			= rsget_TPL("regdate")
                    FItemList(i).Fcancelyn			= rsget_TPL("cancelyn")
                    FItemList(i).Ffinishflag		= rsget_TPL("finishflag")
                    FItemList(i).Fipkumdate			= rsget_TPL("ipkumdate")
                    FItemList(i).Ftaxregdate		= rsget_TPL("taxregdate")
                    FItemList(i).Ftaxinputdate		= rsget_TPL("taxinputdate")
                    FItemList(i).Ftaxtype			= rsget_TPL("taxtype")
                    FItemList(i).Fdifferencekey		= rsget_TPL("differencekey")
                    FItemList(i).Ftaxlinkidx		= rsget_TPL("taxlinkidx")
                    FItemList(i).Fneotaxno			= rsget_TPL("neotaxno")
                    FItemList(i).Fbankingupflag		= rsget_TPL("bankingupflag")
                    FItemList(i).Fgroupid			= rsget_TPL("groupid")
                    FItemList(i).FitemvatYn			= rsget_TPL("itemvatYn")

                    FItemList(i).Fcompany_name		= db2html(rsget_TPL("company_name"))
                    FItemList(i).Fjungsan_hp		= db2html(rsget_TPL("jungsan_hp"))
                    FItemList(i).Fjungsan_email		= db2html(rsget_TPL("jungsan_email"))

	            rsget_TPL.MoveNext
				i = i + 1
			Loop
        End If
        rsget_TPL.close
	end sub

	public Sub GetTplJungsanDetailList()
		dim i,sqlStr, addSql

		addSql = ""

		if (FRectMasterIdx <> "") then
			addSql = addSql & " and d.masteridx = '" & FRectMasterIdx & "'" & vbcrlf
		end if

        if (FRectGubun <> "") then
			addSql = addSql & " and d.gubuncd = '" & FRectGubun & "'" & vbcrlf
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_tpl_jungsan_detail] d " & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & addSql
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.Open sqlStr,dbget_TPL,1
			FTotalCount = rsget_TPL("cnt")
			FTotalPage = rsget_TPL("totPg")
		rsget_TPL.Close


		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

        sqlStr = " select top " & CStr(FPageSize*FCurrPage) & vbcrlf
        sqlStr = sqlStr & " d.* " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_tpl_jungsan_detail] d " & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by d.idx " & vbcrlf
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget_TPL.EOF Then
			rsget_TPL.absolutepage = FCurrPage
			Do until rsget_TPL.EOF
				Set FItemList(i) = new CTplJungsanDetailItem

					FItemList(i).Fidx					= rsget_TPL("idx")
                    FItemList(i).Fmasteridx				= rsget_TPL("masteridx")
                    FItemList(i).Fgubuncd				= rsget_TPL("gubuncd")
                    FItemList(i).Fgubunname				= rsget_TPL("gubunname")
                    FItemList(i).Fgubundetailname		= rsget_TPL("gubundetailname")
                    FItemList(i).Ftypename				= rsget_TPL("typename")
                    FItemList(i).Funitprice				= rsget_TPL("unitprice")
                    FItemList(i).Favgcbm				= rsget_TPL("avgcbm")
                    FItemList(i).Fcurrcbm				= rsget_TPL("currcbm")
                    FItemList(i).Fprevcbm				= rsget_TPL("prevcbm")
                    FItemList(i).Fitemno				= rsget_TPL("itemno")
                    FItemList(i).FtotPrice				= rsget_TPL("totPrice")
                    FItemList(i).Fmastercode			= rsget_TPL("mastercode")
                    FItemList(i).Fcomment				= rsget_TPL("comment")

	            rsget_TPL.MoveNext
				i = i + 1
			Loop
        End If
        rsget_TPL.close
    end sub

	public Sub GetTplJungsanCbmList()
		dim i,sqlStr, addSql

		addSql = ""

		if (FRectMasterIdx <> "") then
			addSql = addSql & " and m.idx = '" & FRectMasterIdx & "'" & vbcrlf
		end if

        if (FRectTplCompanyID <> "") then
			addSql = addSql & " and m.tplcompanyid = '" & FRectTplCompanyID & "'" & vbcrlf
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & vbCrLf
        sqlStr = sqlStr & " from " & vbCrLf
        sqlStr = sqlStr & " 	[db_threepl].[dbo].[tbl_tpl_jungsan_master] m " & vbCrLf
        sqlStr = sqlStr & " 	join [db_threepl].[dbo].[tbl_tpl_jungsan_cbm] c " & vbCrLf
        sqlStr = sqlStr & " 	on " & vbCrLf
        sqlStr = sqlStr & " 		m.idx = c.masteridx " & vbCrLf
        sqlStr = sqlStr & " where " & vbCrLf
        sqlStr = sqlStr & " 	1 = 1 " & vbCrLf
		sqlStr = sqlStr & addSql
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.Open sqlStr,dbget_TPL,1
			FTotalCount = rsget_TPL("cnt")
			FTotalPage = rsget_TPL("totPg")
		rsget_TPL.Close


		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

        sqlStr = " select top " & CStr(FPageSize*FCurrPage) & vbcrlf
        sqlStr = sqlStr & " c.* " & vbcrlf
        sqlStr = sqlStr & " from " & vbCrLf
        sqlStr = sqlStr & " 	[db_threepl].[dbo].[tbl_tpl_jungsan_master] m " & vbCrLf
        sqlStr = sqlStr & " 	join [db_threepl].[dbo].[tbl_tpl_jungsan_cbm] c " & vbCrLf
        sqlStr = sqlStr & " 	on " & vbCrLf
        sqlStr = sqlStr & " 		m.idx = c.masteridx " & vbCrLf
        sqlStr = sqlStr & " where " & vbCrLf
        sqlStr = sqlStr & " 	1 = 1 " & vbCrLf
		sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by c.idx " & vbcrlf
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget_TPL.EOF Then
			rsget_TPL.absolutepage = FCurrPage
			Do until rsget_TPL.EOF
				Set FItemList(i) = new CTplJungsanCbmItem

					FItemList(i).Fidx					= rsget_TPL("idx")
                    FItemList(i).Fmasteridx				= rsget_TPL("masteridx")
                    FItemList(i).Fitemgubun				= rsget_TPL("itemgubun")
                    FItemList(i).Fitemid				= rsget_TPL("itemid")
                    FItemList(i).Fitemoption			= rsget_TPL("itemoption")
                    FItemList(i).Fbarcode				= rsget_TPL("barcode")
                    FItemList(i).Fitemname				= rsget_TPL("itemname")
                    FItemList(i).Fitemoptionname		= rsget_TPL("itemoptionname")
                    FItemList(i).Fitemno				= rsget_TPL("itemno")
                    FItemList(i).FcbmX					= rsget_TPL("cbmX")
                    FItemList(i).FcbmY					= rsget_TPL("cbmY")
                    FItemList(i).FcbmZ					= rsget_TPL("cbmZ")

	            rsget_TPL.MoveNext
				i = i + 1
			Loop
        End If
        rsget_TPL.close
    end sub

	public Sub GetTplJungsanEtcList()
		dim i,sqlStr, addSql

		addSql = ""

		if (FRectMasterIdx <> "") then
			addSql = addSql & " and m.idx = '" & FRectMasterIdx & "'" & vbcrlf
		end if

        if (FRectTplCompanyID <> "") then
			addSql = addSql & " and m.tplcompanyid = '" & FRectTplCompanyID & "'" & vbcrlf
		end if

        if (FRectGubun <> "") then
			addSql = addSql & " and c.gubuncd = '" & FRectGubun & "'" & vbcrlf
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & vbCrLf
        sqlStr = sqlStr & " from " & vbCrLf
        sqlStr = sqlStr & " 	[db_threepl].[dbo].[tbl_tpl_jungsan_master] m " & vbCrLf
        sqlStr = sqlStr & " 	join [db_threepl].[dbo].[tbl_tpl_jungsan_etc] c " & vbCrLf
        sqlStr = sqlStr & " 	on " & vbCrLf
        sqlStr = sqlStr & " 		m.idx = c.masteridx " & vbCrLf
        sqlStr = sqlStr & " where " & vbCrLf
        sqlStr = sqlStr & " 	1 = 1 " & vbCrLf
		sqlStr = sqlStr & addSql
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.Open sqlStr,dbget_TPL,1
			FTotalCount = rsget_TPL("cnt")
			FTotalPage = rsget_TPL("totPg")
		rsget_TPL.Close


		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

        sqlStr = " select top " & CStr(FPageSize*FCurrPage) & vbcrlf
        sqlStr = sqlStr & " c.* " & vbcrlf
        sqlStr = sqlStr & " from " & vbCrLf
        sqlStr = sqlStr & " 	[db_threepl].[dbo].[tbl_tpl_jungsan_master] m " & vbCrLf
        sqlStr = sqlStr & " 	join [db_threepl].[dbo].[tbl_tpl_jungsan_etc] c " & vbCrLf
        sqlStr = sqlStr & " 	on " & vbCrLf
        sqlStr = sqlStr & " 		m.idx = c.masteridx " & vbCrLf
        sqlStr = sqlStr & " where " & vbCrLf
        sqlStr = sqlStr & " 	1 = 1 " & vbCrLf
		sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by c.idx " & vbcrlf
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget_TPL.EOF Then
			rsget_TPL.absolutepage = FCurrPage
			Do until rsget_TPL.EOF
				Set FItemList(i) = new CTplJungsanEtcItem

					FItemList(i).Fidx					= rsget_TPL("idx")
                    FItemList(i).Fmasteridx				= rsget_TPL("masteridx")
                    FItemList(i).Fgubuncd				= rsget_TPL("gubuncd")
                    FItemList(i).Fgubunname				= rsget_TPL("gubunname")
                    FItemList(i).Fgubundetailname		= rsget_TPL("gubundetailname")
                    FItemList(i).Ftypename				= rsget_TPL("typename")
                    FItemList(i).Funitprice				= rsget_TPL("unitprice")
                    FItemList(i).Fitemno				= rsget_TPL("itemno")
                    FItemList(i).FtotPrice				= rsget_TPL("totPrice")
                    FItemList(i).Fmastercode			= rsget_TPL("mastercode")
                    FItemList(i).Fcomment				= rsget_TPL("comment")

	            rsget_TPL.MoveNext
				i = i + 1
			Loop
        End If
        rsget_TPL.close
    end sub

	public Sub GetTplJungsanGubunDetailList()
		dim i,sqlStr, addSql

		addSql = ""

        if (FRectGubun <> "") then
			addSql = addSql & " and c1.comm_cd = '" & FRectGubun & "'" & vbcrlf
		end if

        sqlStr = " select c1.comm_cd as gubuncd, c1.comm_name gubunname, c2.comm_name gubundetailname, IsNull(c3.comm_name, '') typename, IsNull(c3.comm_price, c2.comm_price) unitprice "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c1 "
        sqlStr = sqlStr + " 	left join [db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c2 on c1.comm_cd = c2.comm_group "
        sqlStr = sqlStr + " 	left join [db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c3 on c2.comm_cd = c3.comm_group "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and c1.comm_isDel = 'N' "
        sqlStr = sqlStr + " 	and c1.dispyn = 'Y' "
        sqlStr = sqlStr + " 	and c1.comm_group = 'gubun' "
        sqlStr = sqlStr + addSql
        sqlStr = sqlStr + " order by c1.sortno, c2.sortno, c3.sortno "
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.Open sqlStr,dbget_TPL,1
		FResultCount = rsget_TPL.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget_TPL.EOF Then
			rsget_TPL.absolutepage = FCurrPage
			Do until rsget_TPL.EOF
				Set FItemList(i) = new CTplJungsanGubunDetailItem

                    FItemList(i).Fgubuncd				= rsget_TPL("gubuncd")
                    FItemList(i).Fgubunname				= rsget_TPL("gubunname")
                    FItemList(i).Fgubundetailname		= rsget_TPL("gubundetailname")
                    FItemList(i).Ftypename				= rsget_TPL("typename")
                    FItemList(i).Funitprice				= rsget_TPL("unitprice")

	            rsget_TPL.MoveNext
				i = i + 1
			Loop
        End If
        rsget_TPL.close
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
