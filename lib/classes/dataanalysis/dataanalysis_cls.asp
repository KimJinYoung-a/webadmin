<%
'###########################################################
' Description : 데이터분석 클래스
' History : 2016.01.29 한용민 생성
'###########################################################

class cdataanalysis_oneWishItem
	public FItemID
	public FlistImage

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class

class cdataanalysis_oneitem
	public Fmainidx
	public Fkind
	public Fmeasure
	public Fmeasurename
	public Fapiurl
	public Fdimensiongubun
	public Fpretypegubun
	public Fshchannelgubun
	public Fshmakeridgubun
	public fshdategubun
	public fshdateunit
	public fshdatetermgubun
	public Fcomment
	public Fisusing
	public Fregdate
	public Fgubun
	public Fgubunkey
	public Fgubunname
	public Fsortno
	public Fchartidx
	public Fchanneltype
	public Fcharttype
	public Fposition
	public Fpositionpretypegubun
	public Fpositionpretype
	public Foption1
	public Foption2
	public fordtypegubun
	public fchartsortno
	public fgroupcd
	public fgroupname
	public fgroupsortno
	public fyyyy
	public fmm
	public fmaechul
	public fprofit

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class

Class cdataanalysis
	Public FItemList()
	public foneitem
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	public FPageCount

	public ftendb
	public fDBDATAMART
	public fdefault_measure
	public fdefault_dimensiongubun
	public fdefault_dimension
	public fdefault_pretypegubun
	public fdefault_apiurl
	public fdefault_shchannelgubun
	public fdefault_shmakeridgubun
	public fdefault_channeltype

	public fpurposemaechul
	public fpurposeprofit
	public fcurrentmaechul
	public fcurrentprofit
	public fbeforemaechul
	public fbeforeprofit

	public frectkind
	public frectisusing
	public frectmainidx
	public frectgroupcd
	public frectyyyy
	public frectmm
	public frectstartdate
	public frectenddate

	public FRectItemID
	public FRectCateCode

	'//admin/dataanalysis/md_purpose.asp
	Public Sub Getpurposelist()
		Dim sqlStr, i, addsql

		if frectyyyy="" or frectmm="" then exit Sub

		sqlStr = "select TOP " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " t.gubun, t.sortno, t.yyyy, t.mm, t.maechul, t.profit" & vbCrLf
		sqlStr = sqlStr & " from (" & vbCrLf
		sqlStr = sqlStr & " 	select" & vbCrLf
		sqlStr = sqlStr & " 	'purpose' as gubun, '1' as sortno, p.yyyy, p.mm, sum(p.targetmoney) as maechul, sum(p.profitmoney) as profit" & vbCrLf
		sqlStr = sqlStr & " 	from "& ftendb &"db_partner.dbo.tbl_mdmenu_purpose p" & vbCrLf
		sqlStr = sqlStr & " 	where p.gubun='ON'" & vbCrLf
		sqlStr = sqlStr & " 	and p.yyyy='"& frectyyyy &"' and p.mm='"& frectmm &"'" & vbCrLf
		sqlStr = sqlStr & " 	group by p.yyyy, p.mm" & vbCrLf
		sqlStr = sqlStr & " 	union all" & vbCrLf
		sqlStr = sqlStr & " 	select" & vbCrLf
		sqlStr = sqlStr & " 	'currentmaechul' as gubun, '2' as sortno, yyyy, mm, sum(c.itemcost) as maechul, sum(c.maechulprofit) as profit" & vbCrLf
		sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_mdMenu_maechul c" & vbCrLf
		sqlStr = sqlStr & " 	where c.gubun='ON'" & vbCrLf
		sqlStr = sqlStr & " 	and c.yyyy='"& frectyyyy &"' and c.mm='"& frectmm &"'" & vbCrLf
		sqlStr = sqlStr & " 	group by c.yyyy, c.mm" & vbCrLf
		sqlStr = sqlStr & " 	union all" & vbCrLf
		sqlStr = sqlStr & " 	select" & vbCrLf
		sqlStr = sqlStr & " 	'beforemaechul' as gubun, '3' as sortno, yyyy, mm, sum(b.itemcost) as maechul, sum(b.maechulprofit) as profit" & vbCrLf
		sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_mdMenu_maechul b" & vbCrLf
		sqlStr = sqlStr & " 	where b.gubun='ON'" & vbCrLf
		sqlStr = sqlStr & " 	and b.yyyy='"& frectyyyy-1 &"' and b.mm='"& frectmm &"'" & vbCrLf
		sqlStr = sqlStr & " 	group by b.yyyy, b.mm" & vbCrLf
		sqlStr = sqlStr & " ) as t" & vbCrLf
		sqlStr = sqlStr & " order by sortno asc" & vbCrLf

		'response.write sqlStr & "<br>"
		db3_rsget.open sqlStr , db3_dbget, 1
		FResultCount = db3_rsget.RecordCount
		ftotalcount = db3_rsget.RecordCount

		i=0
		if  not db3_rsget.EOF  then
			redim preserve FItemList(FResultCount)

			do until db3_rsget.eof
				set FItemList(i) = new cdataanalysis_oneitem
					FItemList(i).fgubun = db3_rsget("gubun")
					FItemList(i).fsortno = db3_rsget("sortno")
					FItemList(i).fyyyy = db3_rsget("yyyy")
					FItemList(i).fmm = db3_rsget("mm")
					FItemList(i).fmaechul = db3_rsget("maechul")
					FItemList(i).fprofit = db3_rsget("profit")

					if FItemList(i).fgubun="purpose" then
						fpurposemaechul=db3_rsget("maechul")
						fpurposeprofit=db3_rsget("profit")
					end if
					if FItemList(i).fgubun="currentmaechul" then
						fcurrentmaechul=db3_rsget("maechul")
						fcurrentprofit=db3_rsget("profit")
					end if
					if FItemList(i).fgubun="beforemaechul" then
						fbeforemaechul=db3_rsget("maechul")
						fbeforeprofit=db3_rsget("profit")
					end if
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
	End Sub

	'//admin/dataanalysis/md_maechul_ajax.asp
	Public Sub Getmaechullist()
		Dim sqlStr, i, addsql

		if frectstartdate="" or frectenddate="" then exit Sub

		'///////////////// 목표 가져옴 ///////////////////////////////
		sqlStr = "select" & vbCrLf
		sqlStr = sqlStr & " 'purpose' as gubun, '1' as sortno, sum(p.targetmoney) as maechul, sum(p.profitmoney) as profit" & vbCrLf
		sqlStr = sqlStr & " into #tmppurpose" & vbCrLf
		sqlStr = sqlStr & " from "& fDBDATAMART &"db_partner.dbo.tbl_mdmenu_purpose p" & vbCrLf
		sqlStr = sqlStr & " where p.gubun='ON'" & vbCrLf
		sqlStr = sqlStr & " and p.yyyy + '-' + p.mm >= '"& left(frectstartdate,7) &"'" & vbCrLf
		sqlStr = sqlStr & " and p.yyyy + '-' + p.mm <= '"& left(frectenddate,7) &"'" & vbCrLf

		'response.write sqlStr & "<br>"
		dbanalget.Execute sqlStr
		'///////////////// 목표 가져옴 ///////////////////////////////

		'///////////////// 매출 가져옴 ///////////////////////////////
		sqlStr = "SELECT" & vbCrLf
		sqlStr = sqlStr & " 'currentmaechul' as gubun, '2' as sortno" & vbCrLf
		sqlStr = sqlStr & " , isNull(sum(d.itemcost*d.itemno),0) AS maechul" & vbCrLf
		sqlStr = sqlStr & " , isNull(sum(d.itemcost*d.itemno) - sum(d.buycash*d.itemno),0) as profit" & vbCrLf
		sqlStr = sqlStr & " into #tmpcurrentmaechul" & vbCrLf
		sqlStr = sqlStr & " FROM [db_analyze_data_raw].[dbo].[tbl_order_master] as m" & vbCrLf
		sqlStr = sqlStr & " INNER JOIN [db_analyze_data_raw].[dbo].[tbl_order_detail] as d" & vbCrLf
		sqlStr = sqlStr & " 	ON m.orderserial = d.orderserial" & vbCrLf
		'sqlStr = sqlStr & " left join db_analyze_data_raw.dbo.tbl_partner p2" & vbCrLf
		'sqlStr = sqlStr & " 	on m.sitename=p2.id" & vbCrLf
		sqlStr = sqlStr & " WHERE d.beasongdate >='"& frectstartdate &"'" & vbCrLf
		sqlStr = sqlStr & " and d.beasongdate <'"& dateadd("d", +1, frectenddate) &"'" & vbCrLf
		sqlStr = sqlStr & " AND m.ipkumdiv>3" & vbCrLf
		sqlStr = sqlStr & " AND m.cancelyn='N'" & vbCrLf
		sqlStr = sqlStr & " AND d.cancelyn<>'Y'" & vbCrLf
		sqlStr = sqlStr & " AND d.itemid not in (0, 100)" & vbCrLf
		'sqlStr = sqlStr & " and isNULL(p2.tplcompanyid,'')=''" & vbCrLf
		sqlStr = sqlStr & " AND m.beadaldiv<>90" & vbCrLf  '3pl 제외

		'response.write sqlStr & "<br>"
		dbanalget.Execute sqlStr
		'///////////////// 매출 가져옴 ///////////////////////////////

		'///////////////// 전년대비 매출 가져옴 ///////////////////////////////
		sqlStr = "SELECT" & vbCrLf
		sqlStr = sqlStr & " 'beforemaechul' as gubun, '3' as sortno" & vbCrLf
		sqlStr = sqlStr & " , isNull(sum(d.itemcost*d.itemno),0) AS maechul" & vbCrLf
		sqlStr = sqlStr & " , isNull(sum(d.itemcost*d.itemno) - sum(d.buycash*d.itemno),0) as profit" & vbCrLf
		sqlStr = sqlStr & " into #tmpbeforemaechul" & vbCrLf
		sqlStr = sqlStr & " FROM [db_analyze_data_raw].[dbo].[tbl_order_master] as m" & vbCrLf
		sqlStr = sqlStr & " INNER JOIN [db_analyze_data_raw].[dbo].[tbl_order_detail] as d" & vbCrLf
		sqlStr = sqlStr & " 	ON m.orderserial = d.orderserial" & vbCrLf
		'sqlStr = sqlStr & " left join db_analyze_data_raw.dbo.tbl_partner p2" & vbCrLf
		'sqlStr = sqlStr & " 	on m.sitename=p2.id" & vbCrLf
		sqlStr = sqlStr & " WHERE d.beasongdate >='"& dateadd("yyyy", -1, frectstartdate) &"'" & vbCrLf
		sqlStr = sqlStr & " and d.beasongdate <'"& dateadd("d", +1, dateadd("yyyy", -1, frectenddate)) &"'" & vbCrLf
		sqlStr = sqlStr & " AND m.ipkumdiv>3" & vbCrLf
		sqlStr = sqlStr & " AND m.cancelyn='N'" & vbCrLf
		sqlStr = sqlStr & " AND d.cancelyn<>'Y'" & vbCrLf
		sqlStr = sqlStr & " AND d.itemid not in (0, 100)" & vbCrLf
		'sqlStr = sqlStr & " and isNULL(p2.tplcompanyid,'')=''" & vbCrLf
		sqlStr = sqlStr & " AND m.beadaldiv<>90" & vbCrLf  '3pl 제외

		'response.write sqlStr & "<br>"
		dbanalget.Execute sqlStr
		'///////////////// 전년대비 매출 가져옴 ///////////////////////////////

		sqlStr = "select TOP " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " t.gubun, t.sortno, t.maechul, t.profit" & vbCrLf
		sqlStr = sqlStr & " from (" & vbCrLf
		sqlStr = sqlStr & " 	select * from #tmppurpose" & vbCrLf
		sqlStr = sqlStr & " 	union all" & vbCrLf
		sqlStr = sqlStr & " 	select * from #tmpcurrentmaechul" & vbCrLf
		sqlStr = sqlStr & " 	union all" & vbCrLf
		sqlStr = sqlStr & " 	select * from #tmpbeforemaechul" & vbCrLf
		sqlStr = sqlStr & " ) as t" & vbCrLf
		sqlStr = sqlStr & " order by sortno asc" & vbCrLf

		'response.write sqlStr & "<br>"
		rsAnalget.open sqlStr , dbAnalget, 1
		FResultCount = rsAnalget.RecordCount
		ftotalcount = rsAnalget.RecordCount

		i=0
		if  not rsAnalget.EOF  then
			redim preserve FItemList(FResultCount)

			do until rsAnalget.eof
				set FItemList(i) = new cdataanalysis_oneitem
					FItemList(i).fgubun = rsAnalget("gubun")
					FItemList(i).fsortno = rsAnalget("sortno")
					FItemList(i).fmaechul = rsAnalget("maechul")
					FItemList(i).fprofit = rsAnalget("profit")

					if FItemList(i).fgubun="purpose" then
						fpurposemaechul=rsAnalget("maechul")
						fpurposeprofit=rsAnalget("profit")
					end if
					if FItemList(i).fgubun="currentmaechul" then
						fcurrentmaechul=rsAnalget("maechul")
						fcurrentprofit=rsAnalget("profit")
					end if
					if FItemList(i).fgubun="beforemaechul" then
						fbeforemaechul=rsAnalget("maechul")
						fbeforeprofit=rsAnalget("profit")
					end if
				i=i+1
				rsAnalget.moveNext
			loop
		end if
		rsAnalget.Close
	End Sub

	'//admin/dataanalysis/mkt.asp
	Public Sub Getdataanalysis_maingroup_list()
		Dim sqlStr, i, addsql

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " gr.groupcd, g.gubunname as groupname, gr.groupsortno" & vbCrLf
		sqlStr = sqlStr & " ,m.mainidx, m.kind, m.measure, m.measurename, m.apiurl, m.dimensiongubun, m.pretypegubun, m.shchannelgubun" & vbCrLf
		sqlStr = sqlStr & " , m.shmakeridgubun, m.shdateunit, m.shdategubun, m.shdatetermgubun, m.comment, m.isusing, m.regdate, m.ordtypegubun" & vbCrLf
		sqlStr = sqlStr & " from db_analyze.dbo.tbl_analysis_main m" & vbCrLf
		sqlStr = sqlStr & " join db_analyze.dbo.tbl_analysis_group gr" & vbCrLf
		sqlStr = sqlStr & " 	on m.mainidx=gr.mainidx" & vbCrLf
		sqlStr = sqlStr & " left join db_analyze.dbo.tbl_analysis_gubun g" & vbCrLf
		sqlStr = sqlStr & " 	on gr.groupcd=g.gubunkey" & vbCrLf
		sqlStr = sqlStr & " 	and g.gubun='groupcd'" & vbCrLf
		sqlStr = sqlStr & " where 1=1"

		If frectgroupcd <> "" Then
			sqlStr = sqlStr & " AND gr.groupcd = '" & html2db(frectgroupcd) & "'" & vbCrLf
		End IF
		If frectisusing <> "" Then
			sqlStr = sqlStr & " AND m.isusing = '" & frectisusing & "'" & vbCrLf
		End IF

		sqlStr = sqlStr & " order by gr.groupsortno asc"

		'response.write sqlStr & "<br>"
		rsAnalget.pagesize = FPageSize
		rsAnalget.Open sqlStr,dbanalget,1
		FResultCount = rsAnalget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsAnalget.EOF Then
			rsAnalget.absolutepage = FCurrPage
			Do until rsAnalget.EOF
				Set FItemList(i) = new cdataanalysis_oneitem

					FItemList(i).fmainidx = rsAnalget("mainidx")
					FItemList(i).fkind = db2html(rsAnalget("kind"))
					FItemList(i).fmeasure = db2html(rsAnalget("measure"))
					FItemList(i).fmeasurename = db2html(rsAnalget("measurename"))
					FItemList(i).fapiurl = db2html(rsAnalget("apiurl"))
					FItemList(i).fdimensiongubun = rsAnalget("dimensiongubun")
					FItemList(i).fpretypegubun = rsAnalget("pretypegubun")
					FItemList(i).fshchannelgubun = rsAnalget("shchannelgubun")
					FItemList(i).fshmakeridgubun = rsAnalget("shmakeridgubun")
					FItemList(i).fshdateunit = rsAnalget("shdateunit")
					FItemList(i).fshdategubun = rsAnalget("shdategubun")
					FItemList(i).fshdatetermgubun = rsAnalget("shdatetermgubun")
					FItemList(i).fordtypegubun = db2html(rsAnalget("ordtypegubun"))
					FItemList(i).fcomment = db2html(rsAnalget("comment"))
					FItemList(i).fisusing = rsAnalget("isusing")
					FItemList(i).fregdate = rsAnalget("regdate")
					FItemList(i).fgroupcd = db2html(rsAnalget("groupcd"))
					FItemList(i).fgroupname = db2html(rsAnalget("groupname"))
					FItemList(i).fgroupsortno = rsAnalget("groupsortno")

				i = i + 1
				rsAnalget.moveNext
			Loop
		End If
		rsAnalget.Close
	End Sub

	'// /admin/itemmaster/wishCollection.asp
	Public Sub Getdataanalysis_wish_list()
		Dim sqlStr, i, addsql

		sqlStr = "select top " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " w.itemidA, i.listimage "
		sqlStr = sqlStr & " from "
		sqlStr = sqlStr & " [db_analyze].[dbo].[tbl_buy_together_wish_SUM] w "
		sqlStr = sqlStr & " join [db_analyze_data_raw].[dbo].[tbl_item] i on w.itemidA = i.itemid "
		sqlStr = sqlStr & " join [db_analyze].[dbo].[tbl_buy_together_wish_title] t on w.itemidA = t.itemidA "
		sqlStr = sqlStr & " join [db_analyze_data_raw].[dbo].[tbl_display_cate_item] c on w.itemidA = c.itemid and c.isDefault = 'Y' "
		sqlStr = sqlStr & " where 1 = 1 "
		''sqlStr = sqlStr & " and itemCnt*1.5 <= sumCnt and itemCnt >= 8 "
		sqlStr = sqlStr & " and itemCnt >= 2 "
		if FRectCateCode <> "" then
			if (FRectCateCode = "case") then
				sqlStr = sqlStr & " and Left(convert(varchar,c.catecode), 6) in ('102101', '102102')"
			else
				sqlStr = sqlStr & " and Left(convert(varchar,c.catecode), " & Len(FRectCateCode) & ") = '" & FRectCateCode & "'"
			end if
		end if
		sqlStr = sqlStr & " order by DateDiff(day, w.regdate, getdate()), itemcnt desc, itemidA desc "

		'response.write sqlStr & "<br>"
		rsAnalget.pagesize = FPageSize
		rsAnalget.Open sqlStr,dbanalget,1
		FResultCount = rsAnalget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsAnalget.EOF Then
			rsAnalget.absolutepage = FCurrPage
			Do until rsAnalget.EOF
				Set FItemList(i) = new cdataanalysis_oneWishItem

				FItemList(i).FItemID = rsAnalget("itemidA")
				FItemList(i).FlistImage = rsAnalget("listimage")

				if ((Not IsNULL(FItemList(i).FlistImage)) and (FItemList(i).FlistImage<>"")) then FItemList(i).FlistImage    = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FlistImage

				i = i + 1
				rsAnalget.moveNext
			Loop
		End If
		rsAnalget.Close
	End Sub

	'// /admin/itemmaster/wishCollection.asp
	Public Sub Getdataanalysis_wish_detail()
		Dim sqlStr, i, addsql

		sqlStr = "select top " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " w.itemidB, i.listimage "
		sqlStr = sqlStr & " from "
		sqlStr = sqlStr & " [db_analyze].[dbo].[tbl_buy_together_wish] w "
		sqlStr = sqlStr & " join [db_analyze_data_raw].[dbo].[tbl_item] i on w.itemidB = i.itemid "
		sqlStr = sqlStr & " where w.itemidA = " & FRectItemID
		sqlStr = sqlStr & " order by w.rnk "

		'response.write sqlStr & "<br>"
		rsAnalget.pagesize = FPageSize
		rsAnalget.Open sqlStr,dbanalget,1
		FResultCount = rsAnalget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsAnalget.EOF Then
			rsAnalget.absolutepage = FCurrPage
			Do until rsAnalget.EOF
				Set FItemList(i) = new cdataanalysis_oneWishItem

				FItemList(i).FItemID = rsAnalget("itemidB")
				FItemList(i).FlistImage = rsAnalget("listimage")

				if ((Not IsNULL(FItemList(i).FlistImage)) and (FItemList(i).FlistImage<>"")) then FItemList(i).FlistImage    = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FlistImage

				i = i + 1
				rsAnalget.moveNext
			Loop
		End If
		rsAnalget.Close
	End Sub

	'//admin/dataanalysis/mkt.asp	'//admin/dataanalysis/manager/chart_edit.asp
	Public Sub Getdataanalysis_chart_list()
		Dim sqlStr, i, addsql

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " c.chartidx, c.mainidx, c.channeltype, c.charttype, c.position, c.positionpretype" & vbCrLf
		sqlStr = sqlStr & " , c.option1, c.option2, c.isusing, c.chartsortno, c.regdate" & vbCrLf
		sqlStr = sqlStr & " from db_analyze.dbo.tbl_analysis_chart c" & vbCrLf
		sqlStr = sqlStr & " left join db_analyze.dbo.tbl_analysis_gubun g" & vbcrlf
		sqlStr = sqlStr & " 	on g.gubun='channeltype'" & vbcrlf
		sqlStr = sqlStr & " 	and c.channeltype=g.gubunkey" & vbcrlf
		sqlStr = sqlStr & " where 1=1"

		If frectmainidx <> "" Then
			sqlStr = sqlStr & " AND c.mainidx = " & frectmainidx & "" & vbCrLf
		End IF
		If frectisusing <> "" Then
			sqlStr = sqlStr & " AND c.isusing = '" & frectisusing & "'" & vbCrLf
		End IF

		sqlStr = sqlStr & " order by c.chartsortno asc, g.sortno asc" & vbCrLf

		'response.write sqlStr & "<br>"
		rsAnalget.pagesize = FPageSize
		rsAnalget.Open sqlStr,dbanalget,1
		FResultCount = rsAnalget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsAnalget.EOF Then
			rsAnalget.absolutepage = FCurrPage
			Do until rsAnalget.EOF
				Set FItemList(i) = new cdataanalysis_oneitem

					FItemList(i).fchartidx = rsAnalget("chartidx")
					FItemList(i).fmainidx = rsAnalget("mainidx")
					FItemList(i).fchanneltype = rsAnalget("channeltype")
					FItemList(i).fcharttype = rsAnalget("charttype")
					FItemList(i).fposition = db2html(rsAnalget("position"))
					FItemList(i).fpositionpretype = db2html(rsAnalget("positionpretype"))
					FItemList(i).foption1 = db2html(rsAnalget("option1"))
					FItemList(i).foption2 = db2html(rsAnalget("option2"))
					FItemList(i).fisusing = rsAnalget("isusing")
					FItemList(i).fchartsortno = rsAnalget("chartsortno")
					FItemList(i).fregdate = rsAnalget("regdate")

				i = i + 1
				rsAnalget.moveNext
			Loop
		End If
		rsAnalget.Close
	End Sub

	'//admin/dataanalysis/mkt.asp
	public sub getdefaultsetting()
	    dim i, tmp_measure, tmp_dimensiongubun, tmp_pretypegubun

	    if (isEmpty(cdata)) then Exit sub
	    if (cdata is Nothing) then Exit sub
	    if isArray(cdata.FItemList) then
	        for i=LBound(cdata.FItemList) to UBound(cdata.FItemList)
	            if Not (isEmpty(cdata.FItemList(i))) then
	    			if Not (cdata.FItemList(i) is Nothing) then
	    				'/초기값을 받아옴
						if i = 0 then
							tmp_measure=cdata.FItemList(i).fmeasure
							tmp_dimensiongubun=cdata.FItemList(i).fdimensiongubun
							tmp_pretypegubun=cdata.FItemList(i).fpretypegubun
							exit for
						end if
	    			end if
	    		end if
			next
			'response.write tmp_measure & "<Br>"
			cdata.fdefault_measure=tmp_measure
			cdata.fdefault_dimensiongubun=tmp_dimensiongubun
			if tmp_dimensiongubun="1" then
				cdata.fdefault_dimension="date"
			elseif tmp_dimensiongubun="2" then
				cdata.fdefault_dimension="date"
			end if
			cdata.fdefault_pretypegubun=tmp_pretypegubun
	    end if
	end sub

	'//admin/dataanalysis/mkt.asp
	public sub drawlayout_shdate(selectBoxName, selectedId, chplg, shdategubun)
	    dim i, tmp_str
%>
		<select name="<%= selectBoxName %>" <%=chplg%>>
			<% if shdategubun=1 then %>
				<option value='regdate' <% if selectedId="regdate" then response.write " selected" %>>주문일</option>
			<% elseif shdategubun=2 then %>
				<option value='regdate' <% if selectedId="regdate" then response.write " selected" %>>주문일</option>
				<option value='ipkumdate' <% if selectedId="ipkumdate" then response.write " selected" %>>결제일</option>
			<% elseif shdategubun=3 then %>
				<option value='regdate' <% if selectedId="regdate" then response.write " selected" %>>주문일</option>
				<option value='ipkumdate' <% if selectedId="ipkumdate" then response.write " selected" %>>결제일</option>
				<option value='beasongdate' <% if selectedId="beasongdate" then response.write " selected" %>>출고일</option>
			<% else %>
				<option value='defaultdate' <% if selectedId="defaultdate" then response.write " selected" %>>기간</option>
			<% end if %>
		</select>
<%
	end sub

	'//admin/dataanalysis/mkt.asp
	public sub drawlayout_shmakerid(BoxName, selectedId, chplg, shmakeridgubun)
		if shmakeridgubun="" or isnull(shmakeridgubun) or shmakeridgubun="0" then exit sub
%>
		브랜드 :
		<% if shmakeridgubun="1" then %>
			<input type='text' name='<%= BoxName %>' value='<%= selectedId %>' size="20">
			<input type='button' class='button' value='IDSearch' <%=chplg%>>
		<% end if %>
<%
	end sub

	'//admin/dataanalysis/mkt.asp
	public sub drawlayout_ordtype(selectBoxName, selectedId, chplg, allyn, ordtypegubun)
		if ordtypegubun="" or isnull(ordtypegubun) or ordtypegubun="0" then  exit sub
%>
		정렬 :
		<% if ordtypegubun="1" then %>
			<select name="<%=selectBoxName%>" <%=chplg%>>
				<% if allyn="Y" then %>
					<option value='' <% if selectedId="" then response.write " selected" %>>선택하세요</option>
				<% end if %>

				<option value='asc' <% if selectedId="asc" then response.write " selected" %>>오름차순</option>
				<option value='desc' <% if selectedId="desc" then response.write " selected" %>>내림차순</option>
			</select>
		<% elseif ordtypegubun="2" then %>
			<select name="<%=selectBoxName%>" <%=chplg%>>
				<option value='categubun' <% if selectedId="categubun" then response.write " selected" %>>카테고리구분순</option>
				<option value='maechul' <% if selectedId="maechul" then response.write " selected" %>>매출순</option>
				<option value='attain' <% if selectedId="attain" then response.write " selected" %>>매출달성순</option>
			</select>
		<% end if %>
<%
	end sub

	'//admin/dataanalysis/mkt.asp
	public sub drawlayout_shchannel(selectBoxName, selectedId, chplg, allyn, shchannelgubun)
		if shchannelgubun="" or isnull(shchannelgubun) or shchannelgubun="0" then exit sub
%>
		채널구분 :
		<% if shchannelgubun="1" then %>
			<select name="<%=selectBoxName%>" <%=chplg%>>
				<% if allyn="Y" then %>
					<option value='' <% if selectedId="" then response.write " selected" %>>선택하세요</option>
				<% end if %>

				<option value='WEB' <% if selectedId="WEB" then response.write " selected" %>>WEB</option>
				<option value='MOB' <% if selectedId="MOB" then response.write " selected" %>>MOB</option>
				<option value='APP' <% if selectedId="APP" then response.write " selected" %>>APP</option>
				<option value='OUT' <% if selectedId="OUT" then response.write " selected" %>>제휴몰</option>

				<option value='MOBONL' <% if selectedId="MOBONL" then response.write " selected" %>>MOB_제휴제외</option>
				<option value='MOBLNK' <% if selectedId="MOBLNK" then response.write " selected" %>>MOB_제휴</option>
				<option value='APPONL' <% if selectedId="APPONL" then response.write " selected" %>>APP_제휴제외</option>
				<option value='APPLNK' <% if selectedId="APPLNK" then response.write " selected" %>>APP_제휴</option>

				<!-- option value='3PL' <% if selectedId="3PL" then response.write " selected" %>>3PL</option-->
			</select>
		<% end if %>
<%
	end sub

	'//admin/dataanalysis/mkt.asp
	public sub drawlayout_dimension(selectBoxName, selectedId, chplg, allyn, dimensiongubun)
	    dim i, tmp_str
		if dimensiongubun="" or isnull(dimensiongubun) or dimensiongubun="0" then exit sub
%>
		일정 :
		<% if dimensiongubun="1" then %>
			<input type='radio' name='<%= selectBoxName %>' value='date' <% if selectedId="date" then response.write " checked" %> <%= chplg %>>일</option>
			<input type='radio' name='<%= selectBoxName %>' value='datehour' <% if selectedId="datehour" then response.write " checked" %> <%= chplg %>>시간</option>
			<input type='radio' name='<%= selectBoxName %>' value='week' <% if selectedId="week" then response.write " checked" %> <%= chplg %>>주</option>
			<input type='radio' name='<%= selectBoxName %>' value='month' <% if selectedId="month" then response.write " checked" %> <%= chplg %>>월</option>
			<input type='radio' name='<%= selectBoxName %>' value='year' <% if selectedId="year" then response.write " checked" %> <%= chplg %>>년</option>
			<input type='radio' name='<%= selectBoxName %>' value='weekday' <% if selectedId="weekday" then response.write " checked" %> <%= chplg %>>요일</option>
		<% elseif dimensiongubun="2" then %>
			<input type='radio' name='<%= selectBoxName %>' value='date' <% if selectedId="date" then response.write " checked" %> <%= chplg %>>일</option>
			<input type='radio' name='<%= selectBoxName %>' value='week' <% if selectedId="week" then response.write " checked" %> <%= chplg %>>주</option>
			<input type='radio' name='<%= selectBoxName %>' value='month' <% if selectedId="month" then response.write " checked" %> <%= chplg %>>월</option>
			<input type='radio' name='<%= selectBoxName %>' value='year' <% if selectedId="year" then response.write " checked" %> <%= chplg %>>년</option>
			<input type='radio' name='<%= selectBoxName %>' value='weekday' <% if selectedId="weekday" then response.write " checked" %> <%= chplg %>>요일</option>
		<% end if %>
<%
	end sub

	'//admin/dataanalysis/mkt.asp
	public sub drawlayout_pretype(selectBoxName, selectedId, chplg, allyn, pretypegubun, dimension, pretypeuse, pretypeusechplg)
	    dim i, tmp_str
		if pretypegubun="" or isnull(pretypegubun) or pretypegubun="0" then exit sub
		if dimension="year" then exit sub
%>
		<input type='checkbox' name='pretypeuse' value='on' <%= chkiif(pretypeuse<>"", "checked", "") %> <%= pretypeusechplg %>>
		비교 :
		<% if pretypegubun="1" then %>
			<%
			'/일 단위
			if dimension="date" then
			%>
				<input type='radio' name='<%= selectBoxName %>' value='pweek' <% if selectedId="pweek" then response.write " checked" %> <%= chplg %>>전주</option>
				<input type='radio' name='<%= selectBoxName %>' value='pmonth' <% if selectedId="pmonth" then response.write " checked" %> <%= chplg %>>전월</option>
				<input type='radio' name='<%= selectBoxName %>' value='pyear' <% if selectedId="pyear" then response.write " checked" %> <%= chplg %>>전년</option>
				<input type='radio' name='<%= selectBoxName %>' value='pyearwd' <% if selectedId="pyearwd" then response.write " checked" %> <%= chplg %>>전년동요일</option>
			<%
			'/시간 단위
			elseif dimension="datehour" then
			%>
				<input type='radio' name='<%= selectBoxName %>' value='pday' <% if selectedId="pday" then response.write " checked" %> <%= chplg %>>전일</option>
				<input type='radio' name='<%= selectBoxName %>' value='pweek' <% if selectedId="pweek" then response.write " checked" %> <%= chplg %>>전주</option>
				<input type='radio' name='<%= selectBoxName %>' value='pmonth' <% if selectedId="pmonth" then response.write " checked" %> <%= chplg %>>전월</option>
				<input type='radio' name='<%= selectBoxName %>' value='pyear' <% if selectedId="pyear" then response.write " checked" %> <%= chplg %>>전년</option>
				<input type='radio' name='<%= selectBoxName %>' value='pyearwd' <% if selectedId="pyearwd" then response.write " checked" %> <%= chplg %>>전년동요일</option>
			<%
			'/주 단위
			elseif dimension="week" then
			%>
				<input type='radio' name='<%= selectBoxName %>' value='pmonth' <% if selectedId="pmonth" then response.write " checked" %> <%= chplg %>>전월</option>
				<input type='radio' name='<%= selectBoxName %>' value='pyear' <% if selectedId="pyear" then response.write " checked" %> <%= chplg %>>전년</option>
				<input type='radio' name='<%= selectBoxName %>' value='pyearwd' <% if selectedId="pyearwd" then response.write " checked" %> <%= chplg %>>전년동요일</option>
			<%
			'/월 단위
			elseif dimension="month" then
			%>
				<input type='radio' name='<%= selectBoxName %>' value='pyear' <% if selectedId="pyear" then response.write " checked" %> <%= chplg %>>전년</option>
				<input type='radio' name='<%= selectBoxName %>' value='pyearwd' <% if selectedId="pyearwd" then response.write " checked" %> <%= chplg %>>전년동요일</option>
			<%
			'/년 단위
			elseif dimension="year" then
			%>
			<%
			'/요일 단위
			elseif dimension="weekday" then
			%>
				<input type='radio' name='<%= selectBoxName %>' value='pweek' <% if selectedId="pweek" then response.write " checked" %> <%= chplg %>>전주</option>
				<input type='radio' name='<%= selectBoxName %>' value='pmonth' <% if selectedId="pmonth" then response.write " checked" %> <%= chplg %>>전월</option>
				<input type='radio' name='<%= selectBoxName %>' value='pyear' <% if selectedId="pyear" then response.write " checked" %> <%= chplg %>>전년</option>
				<input type='radio' name='<%= selectBoxName %>' value='pyearwd' <% if selectedId="pyearwd" then response.write " checked" %> <%= chplg %>>전년동요일</option>
			<% end if %>
		<% end if %>
<%
	end sub

	'//admin/dataanalysis/mkt.asp
	public sub drawlayout_measure(selectBoxName, selectedId, chplg, allyn)
	    dim i, tmp_str

	    if (isEmpty(cdata)) then Exit sub
	    if (cdata is Nothing) then Exit sub
	    if isArray(cdata.FItemList) then
	    	response.write "측정값 : <select name="& selectBoxName &" "& chplg &">" & vbcrlf
	    	if allyn="Y" then
				if selectedId = "" then
					tmp_str = " selected"
				end if
	    		response.write "<option value='' "& tmp_str &">선택하세요</option>" & vbcrlf
	    		tmp_str = ""
	    	end if
	        for i=LBound(cdata.FItemList) to UBound(cdata.FItemList)
	            if Not (isEmpty(cdata.FItemList(i))) then
	    			if Not (cdata.FItemList(i) is Nothing) then
						if selectedId = cdata.FItemList(i).fmeasure then
							tmp_str = " selected"
						end if
						response.write "<option value='"& cdata.FItemList(i).fmeasure &"' "&tmp_str&">"& cdata.FItemList(i).fmeasurename &"</option>" & vbcrlf
						tmp_str = ""
	    			end if
	    		end if
			next
			response.write("</select>")
	    end if
	end sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function
	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function
	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FScrollCount = 10

		IF application("Svr_Info")="Dev" THEN
			ftendb = "tendb."
			fDBDATAMART = "tendb."
		else
			fDBDATAMART = "DBDATAMART."
		end if

		fdefault_apiurl="http://wapi.10x10.co.kr/anal/getque.asp"
	End Sub
	Private Sub Class_Terminate()
    End Sub
End Class

'//admin/dataanalysis/mkt.asp
function drawSelectBoxdataanalysischart(selectBoxName, selectedId, chplg, isusing, allyn, mainidx)
	dim tmp_str,query1

	response.write "채널 : "

	query1 = " select" & vbcrlf
	query1 = query1 & " c.channeltype, g.gubunname, g.sortno" & vbcrlf
	query1 = query1 & " from db_analyze.dbo.tbl_analysis_chart c" & vbcrlf
	query1 = query1 & " left join db_analyze.dbo.tbl_analysis_gubun g" & vbcrlf
	query1 = query1 & " 	on g.gubun='channeltype'" & vbcrlf
	query1 = query1 & " 	and c.channeltype=g.gubunkey" & vbcrlf
	query1 = query1 & " where 1=1" & vbcrlf

	If mainidx <> "" Then
		query1 = query1 & " AND c.mainidx = " & mainidx & "" & vbCrLf
	End IF
	If isusing <> "" Then
		query1 = query1 & " AND c.isusing = '" & isusing & "'" & vbCrLf
	End IF

	query1 = query1 & " group by c.channeltype, g.gubunname, g.sortno" & vbCrLf
	query1 = query1 & " order by g.sortno asc"

	rsAnalget.Open query1,dbAnalget,1

	if  not rsAnalget.EOF  then
	rsAnalget.Movefirst

	do until rsAnalget.EOF
%>
		<input type='radio' name='<%= selectBoxName %>' value='<%= db2html(rsAnalget("channeltype")) %>' <% if selectedId=db2html(rsAnalget("channeltype")) then response.write " checked" %> <%= chplg %>><%= db2html(rsAnalget("gubunname")) %></option>
<%
		rsAnalget.MoveNext
	loop
	end if
	rsAnalget.close
end function

'//admin/dataanalysis/mkt.asp	'//admin/dataanalysis/manager/chart_edit.asp
function drawSelectBoxdataanalysisgubun(selectBoxName, selectedId, chplg, allyn, gubun)
	dim tmp_str,query1

	query1 = " select" & vbcrlf
	query1 = query1 & " g.gubunkey, g.gubunname" & vbcrlf
	query1 = query1 & " from db_analyze.dbo.tbl_analysis_gubun g" & vbcrlf
	query1 = query1 & " where 1=1" & vbcrlf

	if gubun <> "" then
		query1 = query1 & " and g.gubun='"& html2db(gubun) &"'" & vbcrlf
	end if

	query1 = query1 & " group by g.gubunkey, g.gubunname"

	rsAnalget.Open query1,dbAnalget,1
%>
	<select name="<%=selectBoxName%>" <%=chplg%>>
		<% if allyn="Y" then %>
			<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>
		<% end if %>
<%
	if  not rsAnalget.EOF  then
	rsAnalget.Movefirst

	do until rsAnalget.EOF
		if selectedId = html2db(rsAnalget("gubunkey")) then
			tmp_str = " selected"
		end if
		response.write "<option value='"& db2html(rsAnalget("gubunkey")) &"' "&tmp_str&">"& db2html(rsAnalget("gubunname")) &"</option>" & vbcrlf
		tmp_str = ""
		rsAnalget.MoveNext
	loop
	end if
	rsAnalget.close
	response.write("</select>")
end function

'//admin/dataanalysis/mkt.asp
function getdataanalysisgubun(tmpval, gubun)
	dim tmp_str,query1

	query1 = " select" & vbcrlf
	query1 = query1 & " g.gubunkey, g.gubunname" & vbcrlf
	query1 = query1 & " from db_analyze.dbo.tbl_analysis_gubun g" & vbcrlf
	query1 = query1 & " where 1=1" & vbcrlf

	if gubun <> "" then
		query1 = query1 & " and g.gubun='"& html2db(gubun) &"'" & vbcrlf
	end if

	query1 = query1 & " group by g.gubunkey, g.gubunname"

	rsAnalget.Open query1,dbAnalget,1

	if  not rsAnalget.EOF  then
	rsAnalget.Movefirst

	do until rsAnalget.EOF
		if tmpval=db2html(rsAnalget("gubunkey")) then
			tmp_str = db2html(rsAnalget("gubunname"))
		end if

		rsAnalget.MoveNext
	loop
	end if
	rsAnalget.close

	getdataanalysisgubun=tmp_str
end function

'//admin/dataanalysis/mkt.asp
function drawSelectBoxdataanalysisgroup(selectBoxName, selectedId, chplg, allyn, groupcd)
	dim tmp_str,query1

	query1 = "select" & vbcrlf
	query1 = query1 & " gr.groupcd, g.gubunname as groupname, g.sortno" & vbcrlf
	query1 = query1 & " from db_analyze.dbo.tbl_analysis_group gr" & vbcrlf
	query1 = query1 & " join db_analyze.dbo.tbl_analysis_main m" & vbcrlf
	query1 = query1 & " 	on gr.mainidx=m.mainidx" & vbcrlf
	query1 = query1 & " 	and m.isusing='Y'" & vbcrlf
	query1 = query1 & " left join db_analyze.dbo.tbl_analysis_gubun g" & vbCrLf
	query1 = query1 & " 	on gr.groupcd=g.gubunkey" & vbCrLf
	query1 = query1 & " 	and g.gubun='groupcd'" & vbCrLf
	query1 = query1 & " where 1=1" & vbcrlf

	if groupcd <> "" then
		query1 = query1 & " and gr.groupcd='"& groupcd &"'" & vbcrlf
	end if

	query1 = query1 & " group by gr.groupcd, g.gubunname, g.sortno" & vbcrlf
	query1 = query1 & " order by g.sortno asc" & vbcrlf

	'response.write query1 & "<Br>"
	rsAnalget.Open query1,dbAnalget,1
%>
	<select name="<%=selectBoxName%>" <%=chplg%>>
		<% if allyn="Y" then %>
			<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>
		<% end if %>
		<% if allyn="NEW" then %>
			<option value='NEWREG' <%if selectedId="NEWREG" then response.write " selected"%>>신규등록</option>
		<% end if %>
<%
	if  not rsAnalget.EOF  then
	rsAnalget.Movefirst

	do until rsAnalget.EOF
		if selectedId = html2db(rsAnalget("groupcd")) then
			tmp_str = " selected"
		end if
		response.write "<option value='"& db2html(rsAnalget("groupcd")) &"' "&tmp_str&">"& db2html(rsAnalget("groupname")) &"</option>" & vbcrlf
		tmp_str = ""
		rsAnalget.MoveNext
	loop
	end if
	rsAnalget.close
	response.write("</select>")
end function

function getgubunname(vgubun)
	dim tmpgubunname

	if vgubun="" then exit function

	if vgubun="purpose" then
		tmpgubunname="목표"
	elseif vgubun="currentmaechul" then
		tmpgubunname="실적"
	elseif vgubun="beforemaechul" then
		tmpgubunname="전년실적"
	end if

	getgubunname = tmpgubunname
end function

function getgrade(vgrade)
	dim tmpgrade

	if vgrade="" then exit function

	if vgrade>=100 and vgrade<101 then
		tmpgrade="<img src='/images/grade/grade_100.png'>"
	elseif vgrade>=101 then
		tmpgrade="<img src='/images/grade/grade_100UP.png'>"
	elseif vgrade<90 then
		tmpgrade="<img src='/images/grade/grade_90DOWN.png'>"
	end if

	getgrade = tmpgrade
end function
%>
