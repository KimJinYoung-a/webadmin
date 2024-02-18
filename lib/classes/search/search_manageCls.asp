<%
'//play class
Class CSearchMngItem
    public Fidx
    public Ftopidx
	public Ftitle
	public Fautotype
	public Furl_pc
	public Furl_m
	public Ficon
	public Fmemo
	public Fuseyn
	public Fsortno
	public Freguserid
	public Fregusername
	public Fregdate
	public Flastuserid
	public Flastusername
	public Flastdate
	public Fquicktype
	public Fquickname
	public Fviewgubun
	public Fsdate
	public Fedate
	public Fshhmmss
	public Fehhmmss
	public Fsubcopy
	public Fhtmlcont
	public Fbtnname
	public Fbtn_pclink
	public Fbtn_mlink
	public Fbggubun
	public Fbgcolor
	public Fbgimgpc
	public Fbgimgm
	public Fqimg_useyn
	public Fqimgpc
	public Fqimgm
	public Fbtn_color
	public Fbrandid
	public Fkeyword
	public Fbgimg
	public Fmaskingimg
	public Ftextinfouse
	public Ftextinfo1
	public Ftextinfo1url
	public Ftextinfo2
	public Ftextinfo2url
	public Fbgclass
    
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CSearchMng
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FKeywordArr
	public FUnitArr
	
	public FRectIdx
	public FRectDateGubun
	public FRectSDate
	public FRectEDate
	public FRectUseYN
	public FRectAutoType
	public FRectQuickType
	public FRectEndType
	public FRectSearchGubun
	public FRectSearchTxt
	public FRectOnlyUnitList
	
	
	public function fnAutoCompleteList()
        dim sqlStr, sqlsearch, i

		'sqlsearch = sqlsearch & " and a.autotype <> 'ky' "
		
		If FRectSDate <> "" Then
			sqlsearch = sqlsearch & " and a.regdate >= '" & FRectSDate & "' "
		End If

		If FRectEDate <> "" Then
			sqlsearch = sqlsearch & " and a.regdate <= '" & DateAdd("d",1,FRectEDate) & "' "
		End If
		
		If FRectAutoType <> "" Then
			sqlsearch = sqlsearch & " and a.autotype = '" & FRectAutoType & "' "
		End If
		
		If FRectUseYN <> "" Then
			sqlsearch = sqlsearch & " and a.useyn = '" & FRectUseYN & "' "
		End IF
		
		If FRectSearchTxt <> "" Then
			sqlsearch = sqlsearch & " and a.title = '" & FRectSearchTxt & "' "
		End If


		'// 결과수 카운트
		sqlStr = "select count(a.idx) as cnt, CEILING(CAST(Count(a.idx) AS FLOAT)/" & FPageSize & ") AS totPg"
        sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_autocomplete] as a"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
            FTotalCount = rsget("cnt")
            FTotalPage	= rsget("totPg")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " a.idx, a.autotype, a.title, a.url_pc, a.url_m, a.icon, a.memo, a.useyn, a.reguserid, a.regdate, a.lastupdateid, a.lastupdatedate "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = a.reguserid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as regusername "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = a.lastupdateid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as lastusername "
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_autocomplete] as a "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by a.idx DESC"

		'response.write sqlStr &"<Br>"
		'response.end
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

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
            do until rsget.EOF
                set FItemList(i) = new CSearchMngItem

					FItemList(i).Fidx				= rsget("idx")
					FItemList(i).Fautotype		= rsget("autotype")
					FItemList(i).Ftitle			= db2html(rsget("title"))
					FItemList(i).Furl_pc			= rsget("url_pc")
					FItemList(i).Furl_m			= rsget("url_m")
					FItemList(i).Ficon			= rsget("icon")
					FItemList(i).Fmemo			= db2html(rsget("memo"))
					FItemList(i).Fuseyn			= rsget("useyn")
					FItemList(i).Freguserid		= rsget("reguserid")
					FItemList(i).Fregusername	= rsget("regusername")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).Flastuserid		= rsget("lastupdateid")
					FItemList(i).Flastusername	= rsget("lastusername")
					FItemList(i).Flastdate		= rsget("lastupdatedate")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function
    
    
	public Sub sbAutoCompleteDetail()
		dim sqlStr, addsql
		
		If FRectIdx <> "" Then
			addsql = addsql & " and a.idx = '" & FRectIdx & "'"
		End If
		
		sqlStr = "select "
		sqlStr = sqlStr & " a.idx, a.autotype, a.title, a.url_pc, a.url_m, a.icon, a.memo, a.useyn, a.sortno, a.reguserid, a.regdate, a.lastupdateid, a.lastupdatedate "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = a.reguserid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as regusername "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = a.lastupdateid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as lastusername "
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_autocomplete] as a"
		sqlStr = sqlStr & " where 1=1 " & addsql
		'response.write sqlStr
		
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		FResultCount = rsget.RecordCount
	
		set FOneItem = new CSearchMngItem
	
		if Not rsget.Eof then
	
			FOneItem.Fidx				= rsget("idx")
			FOneItem.Fautotype		= rsget("autotype")
			FOneItem.Ftitle			= db2html(rsget("title"))
			FOneItem.Furl_pc			= rsget("url_pc")
			FOneItem.Furl_m			= rsget("url_m")
			FOneItem.Ficon			= rsget("icon")
			FOneItem.Fsortno			= rsget("sortno")
			If isNull(rsget("memo")) Then
				FOneItem.Fmemo = ""
			Else
				FOneItem.Fmemo			= db2html(rsget("memo"))
			End If
			FOneItem.Fuseyn			= rsget("useyn")
			FOneItem.Freguserid		= rsget("reguserid")
			FOneItem.Fregusername	= rsget("regusername")
			FOneItem.Fregdate			= rsget("regdate")
			FOneItem.Flastuserid		= rsget("lastupdateid")
			FOneItem.Flastusername	= rsget("lastusername")
			FOneItem.Flastdate		= rsget("lastupdatedate")

		end if
		rsget.Close
		
	end Sub
	
	
	public function fnQuickLinkList()
        dim sqlStr, sqlsearch, i
	
		If FRectDateGubun = "write" Then	'### 기간 작성일
			If FRectSDate <> "" Then
				sqlsearch = sqlsearch & " and q.regdate >= '" & FRectSDate & "' "
			End If

			If FRectEDate <> "" Then
				sqlsearch = sqlsearch & " and q.regdate <= '" & DateAdd("d",1,FRectEDate) & "' "
			End If
		ElseIf FRectDateGubun = "sdate" OR FRectDateGubun = "edate" Then	'### 기간 시작일, 종료일
			If FRectSDate <> "" Then
				sqlsearch = sqlsearch & " and q." & FRectDateGubun & " >= '" & FRectSDate & "' "
			End If

			If FRectEDate <> "" Then
				sqlsearch = sqlsearch & " and q." & FRectDateGubun & " <= '" & DateAdd("d",1,FRectEDate) & "' "
			End If
		End IF
	
		If FRectQuickType <> "" Then	'### 퀵링크 유형
			sqlsearch = sqlsearch & " and q.type = '" & FRectQuickType & "' "
		End If
		
		If FRectEndType <> "" Then	'### 종료 유형
			If FRectEndType = "now" Then
				sqlsearch = sqlsearch & " and q.viewgubun = 'period' and q.sdate <= getdate() and q.edate >= getdate() "
			ElseIf FRectEndType = "end" Then
				sqlsearch = sqlsearch & " and q.viewgubun = 'period' and q.edate < getdate() "
			End If
		End If
		
		If FRectUseYN <> "" Then	'### 사용 유형
			sqlsearch = sqlsearch & " and q.useyn = '" & FRectUseYN & "' "
		End IF

		If FRectSearchGubun <> "" Then
			If FRectSearchTxt <> "" Then
				If FRectSearchGubun = "k.keyword" Then
					sqlsearch = sqlsearch & " and q.idx in (select k.topidx from [db_sitemaster].[dbo].[tbl_search_keyword] as k where k.topgubun = 'q' and k.keyword = '" & FRectSearchTxt & "') "
				Else
					sqlsearch = sqlsearch & " and " & FRectSearchGubun & " = '" & FRectSearchTxt & "' "
				End If
			End If
		End IF


		'// 결과수 카운트
		sqlStr = "select count(q.idx) as cnt, CEILING(CAST(Count(q.idx) AS FLOAT)/" & FPageSize & ") AS totPg"
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_quicklink] as q "
		sqlStr = sqlStr & " left join [db_partner].[dbo].[tbl_user_tenbyten] as t" & vbcrlf
		sqlStr = sqlStr & " 	on t.userid = q.reguserid" & vbcrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "		and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr & " 	and t.userid <> ''" & vbcrlf
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
            FTotalCount = rsget("cnt")
            FTotalPage	= rsget("totPg")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " q.idx, q.type, q.name, q.viewgubun, q.sdate, q.edate, q.url_pc, q.url_m, q.useyn, t.username as regusername, q.lastupdatedate, "
		sqlStr = sqlStr & " STUFF(( "
		sqlStr = sqlStr & " 	SELECT ',' + k.keyword "
		sqlStr = sqlStr & " 	FROM [db_sitemaster].[dbo].[tbl_search_keyword] as k "
		sqlStr = sqlStr & " 	WHERE topidx = q.idx and k.topgubun = 'q' "
		sqlStr = sqlStr & " FOR XML PATH('') "
		sqlStr = sqlStr & " ), 1, 1, '') AS keyword "
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_quicklink] as q "
		sqlStr = sqlStr & " left join [db_partner].[dbo].[tbl_user_tenbyten] as t" & vbcrlf
		sqlStr = sqlStr & " 	on t.userid = q.reguserid" & vbcrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "		and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr & " 	and t.userid <> ''" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by q.lastupdatedate DESC"

		'response.write sqlStr &"<Br>"
		'response.end
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

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
            do until rsget.EOF
                set FItemList(i) = new CSearchMngItem

					FItemList(i).Fidx				= rsget("idx")
					FItemList(i).Fquicktype		= rsget("type")
					FItemList(i).Fquickname		= db2html(rsget("name"))
					FItemList(i).Fviewgubun		= rsget("viewgubun")
					FItemList(i).Fsdate			= rsget("sdate")
					FItemList(i).Fedate			= rsget("edate")
					FItemList(i).Furl_pc			= rsget("url_pc")
					FItemList(i).Furl_m			= rsget("url_m")
					FItemList(i).Fuseyn			= rsget("useyn")
					FItemList(i).Fregusername	= rsget("regusername")
					FItemList(i).Fregdate			= rsget("lastupdatedate")
					FItemList(i).Fkeyword			= rsget("keyword")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function
    
    
	public Sub sbQuickLinkDetail()
		dim sqlStr, addsql

		sqlStr = "select "
		sqlStr = sqlStr & " q.idx, q.type, q.name, q.brandid, q.subcopy, q.url_pc, q.url_m, q.viewgubun, q.sdate, q.edate, q.memo, q.useyn, "
		sqlStr = sqlStr & " q.reguserid, q.regdate, q.lastupdateid, q.lastupdatedate, q.htmlcont, q.btnname, q.btn_pclink, q.btn_mlink, "
		sqlStr = sqlStr & " q.bggubun, q.bgcolor, q.bgimgpc, q.bgimgm, q.qimg_useyn, q.qimgpc, q.qimgm, q.btn_color "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = q.reguserid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as regusername "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = q.lastupdateid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as lastusername "
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_quicklink] as q "
		sqlStr = sqlStr & " where q.idx = '" & FRectIdx & "'"
		'response.write sqlStr
		
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		FResultCount = rsget.RecordCount
	
		set FOneItem = new CSearchMngItem
	
		if Not rsget.Eof then
	
			FOneItem.Fidx				= rsget("idx")
			FOneItem.Fquicktype		= rsget("type")
			FOneItem.Fquickname		= db2html(rsget("name"))
			FOneItem.Fsubcopy			= db2html(rsget("subcopy"))
			FOneItem.Furl_pc			= rsget("url_pc")
			FOneItem.Furl_m			= rsget("url_m")
			FOneItem.Fviewgubun		= rsget("viewgubun")
			If Left(rsget("sdate"),4) = "1900" Then
				FOneItem.Fsdate = ""
			Else
				FOneItem.Fsdate			= Left(rsget("sdate"),10)
				FOneItem.Fshhmmss		= TwoNumber(Hour(rsget("sdate"))) & ":" & TwoNumber(Minute(rsget("sdate"))) & ":" & TwoNumber(Second(rsget("sdate")))
			End If
			If Left(rsget("edate"),4) = "1900" Then
				FOneItem.Fedate = ""
			Else
				FOneItem.Fedate			= Left(rsget("edate"),10)
				FOneItem.Fehhmmss		= TwoNumber(Hour(rsget("edate"))) & ":" & TwoNumber(Minute(rsget("edate"))) & ":" & TwoNumber(Second(rsget("edate")))
			End If
			If isNull(rsget("memo")) Then
				FOneItem.Fmemo = ""
			Else
				FOneItem.Fmemo			= db2html(rsget("memo"))
			End If
			FOneItem.Fuseyn			= rsget("useyn")
			FOneItem.Freguserid		= rsget("reguserid")
			FOneItem.Fregusername	= rsget("regusername")
			FOneItem.Fregdate			= rsget("regdate")
			FOneItem.Flastuserid		= rsget("lastupdateid")
			FOneItem.Flastusername	= rsget("lastusername")
			FOneItem.Flastdate		= rsget("lastupdatedate")
			If isNull(rsget("htmlcont")) Then
				FOneItem.Fhtmlcont = ""
			Else
				FOneItem.Fhtmlcont	= db2html(rsget("htmlcont"))
			End If
			FOneItem.Fbtnname			= db2html(rsget("btnname"))
			FOneItem.Fbtn_pclink		= rsget("btn_pclink")
			FOneItem.Fbtn_mlink		= rsget("btn_mlink")
			FOneItem.Fbggubun			= rsget("bggubun")
			FOneItem.Fbgcolor			= rsget("bgcolor")
			FOneItem.Fbgimgpc			= rsget("bgimgpc")
			FOneItem.Fbgimgm			= rsget("bgimgm")
			FOneItem.Fqimg_useyn		= rsget("qimg_useyn")
			FOneItem.Fqimgpc			= rsget("qimgpc")
			FOneItem.Fqimgm			= rsget("qimgm")
			FOneItem.Fbtn_color		= rsget("btn_color")
			FOneItem.Fbrandid			= rsget("brandid")

		end if
		rsget.Close
		
		sqlStr = "select keyword from [db_sitemaster].[dbo].[tbl_search_keyword] as k where k.topidx = '" & FRectIdx & "' and k.topgubun = 'q'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		If Not rsget.Eof Then
			FKeywordArr = rsget.getRows()
		End If
		rsget.Close
		
	end Sub
	
	
	public function fnMainManageList()
        dim sqlStr, sqlsearch, i
	
		If FRectDateGubun = "write" Then	'### 기간 작성일
			If FRectSDate <> "" Then
				sqlsearch = sqlsearch & " and m.regdate >= '" & FRectSDate & "' "
			End If

			If FRectEDate <> "" Then
				sqlsearch = sqlsearch & " and m.regdate <= '" & DateAdd("d",1,FRectEDate) & "' "
			End If
		ElseIf FRectDateGubun = "sdate" OR FRectDateGubun = "edate" Then	'### 기간 시작일, 종료일
			If FRectSDate <> "" Then
				sqlsearch = sqlsearch & " and m." & FRectDateGubun & " >= '" & FRectSDate & "' "
			End If

			If FRectEDate <> "" Then
				sqlsearch = sqlsearch & " and m." & FRectDateGubun & " <= '" & DateAdd("d",1,FRectEDate) & "' "
			End If
		End IF

		If FRectEndType <> "" Then	'### 종료 유형
			If FRectEndType = "now" Then
				sqlsearch = sqlsearch & " and m.viewgubun = 'period' and m.sdate <= getdate() and m.edate >= getdate() "
			ElseIf FRectEndType = "end" Then
				sqlsearch = sqlsearch & " and m.viewgubun = 'period' and m.edate < getdate() "
			End If
		End If
		
		If FRectUseYN <> "" Then	'### 사용 유형
			sqlsearch = sqlsearch & " and m.useyn = '" & FRectUseYN & "' "
		End IF

		If FRectSearchTxt <> "" Then
			sqlsearch = sqlsearch & " and t.username = '" & FRectSearchTxt & "' "
		End If


		'// 결과수 카운트
		sqlStr = "select count(m.idx) as cnt, CEILING(CAST(Count(m.idx) AS FLOAT)/" & FPageSize & ") AS totPg"
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_mainmanage] as m "
		sqlStr = sqlStr & " inner join [db_partner].[dbo].[tbl_user_tenbyten] as t" & vbcrlf
		sqlStr = sqlStr & " 	on t.userid = m.lastupdateid" & vbcrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "		and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr & " 	and t.userid <> '' "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
            FTotalCount = rsget("cnt")
            FTotalPage	= rsget("totPg")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " m.idx, m.viewgubun, m.sdate, m.edate, m.useyn, m.lastupdatedate "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = m.lastupdateid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as lastusername "
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_mainmanage] as m "
		sqlStr = sqlStr & " inner join [db_partner].[dbo].[tbl_user_tenbyten] as t" & vbcrlf
		sqlStr = sqlStr & " 	on t.userid = m.lastupdateid" & vbcrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "		and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr & " 	and t.userid <> '' "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by m.idx DESC"

'		response.write "<pre>" & sqlStr &"</pre>"
'		response.end
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

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
            do until rsget.EOF
                set FItemList(i) = new CSearchMngItem

					FItemList(i).Fidx				= rsget("idx")
					FItemList(i).Fviewgubun		= rsget("viewgubun")
					If Left(rsget("sdate"),4) = "1900" Then
						FItemList(i).Fsdate = ""
					Else
						FItemList(i).Fsdate		= rsget("sdate")
					End If
					If Left(rsget("edate"),4) = "1900" Then
						FItemList(i).Fedate = ""
					Else
						FItemList(i).Fedate		= rsget("edate")
					End If
					FItemList(i).Fuseyn			= rsget("useyn")
					FItemList(i).Flastdate		= rsget("lastupdatedate")
					FItemList(i).Flastusername	= rsget("lastusername")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function
    
    
	public Sub sbMainManageDetail()
		dim sqlStr, addsql

		sqlStr = "select "
		sqlStr = sqlStr & " m.idx, m.bggubun, m.bgcolor, m.bgimg, m.maskingimg, m.viewgubun, m.sdate, m.edate, m.useyn,  "
		sqlStr = sqlStr & " m.textinfouse, m.textinfo1, m.textinfo1url, m.textinfo2, m.textinfo2url, m.memo, "
		sqlStr = sqlStr & " m.reguserid, m.regdate, m.lastupdateid, m.lastupdatedate "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = m.reguserid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as regusername "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = m.lastupdateid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as lastusername "
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_mainmanage] as m "
		sqlStr = sqlStr & " where m.idx = '" & FRectIdx & "'"
		'response.write sqlStr
		
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		FResultCount = rsget.RecordCount
	
		set FOneItem = new CSearchMngItem
	
		if Not rsget.Eof then
	
			FOneItem.Fidx				= rsget("idx")
			FOneItem.Fbggubun			= rsget("bggubun")
			FOneItem.Fbgcolor			= rsget("bgcolor")
			FOneItem.Fbgimg			= rsget("bgimg")
			FOneItem.Fmaskingimg		= rsget("maskingimg")
			FOneItem.Fviewgubun		= rsget("viewgubun")
			If Left(rsget("sdate"),4) = "1900" Then
				FOneItem.Fsdate = ""
			Else
				FOneItem.Fsdate			= Left(rsget("sdate"),10)
				FOneItem.Fshhmmss		= TwoNumber(Hour(rsget("sdate"))) & ":" & TwoNumber(Minute(rsget("sdate"))) & ":" & TwoNumber(Second(rsget("sdate")))
			End If
			If Left(rsget("edate"),4) = "1900" Then
				FOneItem.Fedate = ""
			Else
				FOneItem.Fedate			= Left(rsget("edate"),10)
				FOneItem.Fehhmmss		= TwoNumber(Hour(rsget("edate"))) & ":" & TwoNumber(Minute(rsget("edate"))) & ":" & TwoNumber(Second(rsget("edate")))
			End If
			If isNull(rsget("memo")) Then
				FOneItem.Fmemo = ""
			Else
				FOneItem.Fmemo			= db2html(rsget("memo"))
			End If
			FOneItem.Ftextinfouse	= rsget("textinfouse")
			FOneItem.Ftextinfo1		= rsget("textinfo1")
			FOneItem.Ftextinfo1url	= rsget("textinfo1url")
			FOneItem.Ftextinfo2		= rsget("textinfo2")
			FOneItem.Ftextinfo2url	= rsget("textinfo2url")
			FOneItem.Fuseyn			= rsget("useyn")
			FOneItem.Freguserid		= rsget("reguserid")
			FOneItem.Fregusername	= rsget("regusername")
			FOneItem.Fregdate			= rsget("regdate")
			FOneItem.Flastuserid		= rsget("lastupdateid")
			FOneItem.Flastusername	= rsget("lastusername")
			FOneItem.Flastdate		= rsget("lastupdatedate")

		end if
		rsget.Close

	end Sub


	public function fnCuratorList()
        dim sqlStr, sqlsearch, i
	
		If FRectDateGubun = "write" Then	'### 기간 작성일
			If FRectSDate <> "" Then
				sqlsearch = sqlsearch & " and c.regdate >= '" & FRectSDate & "' "
			End If

			If FRectEDate <> "" Then
				sqlsearch = sqlsearch & " and c.regdate <= '" & DateAdd("d",1,FRectEDate) & "' "
			End If
		ElseIf FRectDateGubun = "sdate" OR FRectDateGubun = "edate" Then	'### 기간 시작일, 종료일
			If FRectSDate <> "" Then
				sqlsearch = sqlsearch & " and c." & FRectDateGubun & " >= '" & FRectSDate & "' "
			End If

			If FRectEDate <> "" Then
				sqlsearch = sqlsearch & " and c." & FRectDateGubun & " <= '" & DateAdd("d",1,FRectEDate) & "' "
			End If
		End IF

		If FRectEndType <> "" Then	'### 종료 유형
			If FRectEndType = "now" Then
				sqlsearch = sqlsearch & " and c.viewgubun = 'period' and c.sdate <= getdate() and c.edate >= getdate() "
			ElseIf FRectEndType = "end" Then
				sqlsearch = sqlsearch & " and c.viewgubun = 'period' and c.edate < getdate() "
			End If
		End If
		
		If FRectUseYN <> "" Then	'### 사용 유형
			sqlsearch = sqlsearch & " and c.useyn = '" & FRectUseYN & "' "
		End IF

		If FRectSearchGubun <> "" Then
			If FRectSearchTxt <> "" Then
				If FRectSearchGubun = "k.keyword" Then
					sqlsearch = sqlsearch & " and c.idx in (select k.topidx from [db_sitemaster].[dbo].[tbl_search_keyword] as k where k.topgubun = 'c' and k.keyword = '" & FRectSearchTxt & "') "
				Else
					sqlsearch = sqlsearch & " and " & FRectSearchGubun & " = '" & FRectSearchTxt & "' "
				End If
			End If
		End IF


		'// 결과수 카운트
		sqlStr = "select count(c.idx) as cnt, CEILING(CAST(Count(c.idx) AS FLOAT)/" & FPageSize & ") AS totPg"
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_curator] as c "
		sqlStr = sqlStr & " left join [db_partner].[dbo].[tbl_user_tenbyten] as t" & vbcrlf
		sqlStr = sqlStr & " 	on t.userid = c.lastupdateid" & vbcrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "		and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr & " 	and t.userid <> '' " & vbcrlf
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
            FTotalCount = rsget("cnt")
            FTotalPage	= rsget("totPg")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " c.idx, c.title, c.viewgubun, c.sdate, c.edate, c.useyn, c.lastupdatedate, t.username as lastusername, "
		sqlStr = sqlStr & " STUFF(( "
		sqlStr = sqlStr & " 	SELECT ',' + k.keyword "
		sqlStr = sqlStr & " 	FROM [db_sitemaster].[dbo].[tbl_search_keyword] as k "
		sqlStr = sqlStr & " 	WHERE topidx = c.idx and k.topgubun = 'c' "
		sqlStr = sqlStr & " FOR XML PATH('') "
		sqlStr = sqlStr & " ), 1, 1, '') AS keyword "
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_curator] as c "
		sqlStr = sqlStr & " left join [db_partner].[dbo].[tbl_user_tenbyten] as t" & vbcrlf
		sqlStr = sqlStr & " 	on t.userid = c.lastupdateid" & vbcrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "		and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr & " 	and t.userid <> '' " & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by c.idx DESC"

		'response.write sqlStr &"<Br>"
		'response.end
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

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
            do until rsget.EOF
                set FItemList(i) = new CSearchMngItem

					FItemList(i).Fidx				= rsget("idx")
					FItemList(i).Ftitle			= db2html(rsget("title"))
					FItemList(i).Fviewgubun		= rsget("viewgubun")
					FItemList(i).Fsdate			= rsget("sdate")
					FItemList(i).Fedate			= rsget("edate")
					FItemList(i).Fuseyn			= rsget("useyn")
					FItemList(i).Flastdate		= rsget("lastupdatedate")
					FItemList(i).Flastusername	= rsget("lastusername")
					FItemList(i).Fkeyword			= rsget("keyword")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function
    
    
	public Sub sbCuratorDetail()
		dim sqlStr, addsql

		If FRectOnlyUnitList <> "o" Then
			sqlStr = "select "
			sqlStr = sqlStr & " c.idx, c.title, c.viewgubun, c.sdate, c.edate, c.memo, c.useyn, c.bgclass, c.reguserid, c.regdate, c.lastupdateid, c.lastupdatedate "
			sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.reguserid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as regusername "
			sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.lastupdateid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as lastusername "
			sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_search_curator] as c "
			sqlStr = sqlStr & " where c.idx = '" & FRectIdx & "'"
			'response.write sqlStr
			
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			FResultCount = rsget.RecordCount
		
			set FOneItem = new CSearchMngItem
		
			if Not rsget.Eof then
		
				FOneItem.Fidx				= rsget("idx")
				FOneItem.Ftitle			= db2html(rsget("title"))
				FOneItem.Fviewgubun		= rsget("viewgubun")
				If Left(rsget("sdate"),4) = "1900" Then
					FOneItem.Fsdate = ""
				Else
					FOneItem.Fsdate		= Left(rsget("sdate"),10)
					FOneItem.Fshhmmss		= TwoNumber(Hour(rsget("sdate"))) & ":" & TwoNumber(Minute(rsget("sdate"))) & ":" & TwoNumber(Second(rsget("sdate")))
				End If
				If Left(rsget("edate"),4) = "1900" Then
					FOneItem.Fedate = ""
				Else
					FOneItem.Fedate		= Left(rsget("edate"),10)
					FOneItem.Fehhmmss		= TwoNumber(Hour(rsget("edate"))) & ":" & TwoNumber(Minute(rsget("edate"))) & ":" & TwoNumber(Second(rsget("edate")))
				End If
				If isNull(rsget("memo")) Then
					FOneItem.Fmemo = ""
				Else
					FOneItem.Fmemo			= db2html(rsget("memo"))
				End If
				FOneItem.Fuseyn			= rsget("useyn")
				FOneItem.Fbgclass		= rsget("bgclass")
				FOneItem.Freguserid		= rsget("reguserid")
				FOneItem.Fregusername	= rsget("regusername")
				FOneItem.Fregdate			= rsget("regdate")
				FOneItem.Flastuserid		= rsget("lastupdateid")
				FOneItem.Flastusername	= rsget("lastusername")
				FOneItem.Flastdate		= rsget("lastupdatedate")

			end if
			rsget.Close
			
			sqlStr = "select keyword from [db_sitemaster].[dbo].[tbl_search_keyword] as k where k.topidx = '" & FRectIdx & "' and k.topgubun = 'c'"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			If Not rsget.Eof Then
				FKeywordArr = rsget.getRows()
			End If
			rsget.Close
		End If
		
		
		sqlStr = ""
		sqlStr = sqlStr & "select e.evt_name, cu.gubun, cu.contentsidx, cu.sortno, e.evt_enddate "
		sqlStr = sqlStr & "	from [db_sitemaster].[dbo].[tbl_search_curator_unit] as cu "
		sqlStr = sqlStr & "	inner join [db_event].[dbo].[tbl_event] as e on cu.contentsidx = e.evt_code "
		sqlStr = sqlStr & "where cu.topidx = '" & FRectIdx & "' and cu.gubun = 'event' "
		sqlStr = sqlStr & "union all "
		sqlStr = sqlStr & "select i.itemname, cu.gubun, cu.contentsidx, cu.sortno, getdate() "
		sqlStr = sqlStr & "	from [db_sitemaster].[dbo].[tbl_search_curator_unit] as cu "
		sqlStr = sqlStr & "	inner join [db_item].[dbo].[tbl_item] as i on cu.contentsidx = i.itemid "
		sqlStr = sqlStr & "where cu.topidx = '" & FRectIdx & "' and cu.gubun = 'item' "
		sqlStr = sqlStr & "order by sortno asc"
		
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		If Not rsget.Eof Then
			FUnitArr = rsget.getRows()
		End If
		rsget.Close
		
	end Sub

	
    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 15
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

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


Function fnAutoCompleteTypeName(t)
	Dim vName
	SELECT CASE t
		Case "sc" : vName = "바로가기"
		Case "ca" : vName = "카테고리"
		Case "br" : vName = "브랜드"
		Case "ky" : vName = "키워드"
	END SELECT

	fnAutoCompleteTypeName = vName
End Function

Function fnAutoCompleteTypeSelect(t)
	Dim vBody
	vBody = "<option value=""sc"" " & CHKIIF(t="sc","selected","") & ">바로가기</option>"
	vBody = vBody & "<option value=""ca"" " & CHKIIF(t="ca","selected","") & ">카테고리</option>"
	vBody = vBody & "<option value=""br"" " & CHKIIF(t="br","selected","") & ">브랜드</option>"
	vBody = vBody & "<option value=""ky"" " & CHKIIF(t="ky","selected","") & ">키워드</option>"

	fnAutoCompleteTypeSelect = vBody
End Function

Function fnAutoCompleteIconName(i)
	Dim vName
	SELECT CASE i
		Case "none" : vName = "사용안함"
		Case "best" : vName = "베스트"
		Case "jump" : vName = "급상승 검색어"
	END SELECT

	fnAutoCompleteIconName = vName
End Function

Function fnIsExistValue(i,a,t)
	Dim vQuery, vIsExist
	
	vIsExist = True
	vQuery = "SELECT count(idx) FROM [db_sitemaster].[dbo].[tbl_search_autocomplete] as a "
	vQuery = vQuery & "WHERE useyn = 'y' AND idx <> '" & i & "' AND title = '" & t & "'"
	' AND autotype = '" & a & "'
	rsget.CursorLocation = adUseClient
	rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
	
	If rsget(0) > 0 Then
		vIsExist = True
	Else
		vIsExist = False
	End If
	rsget.close
	fnIsExistValue = vIsExist
End Function

Function fnQuickLinkTypeName(t)
	Dim vName
	SELECT CASE t
		Case "txt" : vName = "텍스트형"
		Case "nor" : vName = "기본형"
		Case "set" : vName = "설정형"
		Case "brd" : vName = "브랜드형"
		Case "cus" : vName = "커스텀형"
	END SELECT

	fnQuickLinkTypeName = vName
End Function

Function fnQuickLinkTypeSelect(t)
	Dim vBody
	vBody = "<option value=""txt"" " & CHKIIF(t="txt","selected","") & ">텍스트형</option>"
	vBody = vBody & "<option value=""nor"" " & CHKIIF(t="nor","selected","") & ">기본형</option>"
	vBody = vBody & "<option value=""set"" " & CHKIIF(t="set","selected","") & ">설정형</option>"
	vBody = vBody & "<option value=""brd"" " & CHKIIF(t="brd","selected","") & ">브랜드형</option>"
	vBody = vBody & "<option value=""cus"" " & CHKIIF(t="cus","selected","") & ">커스텀형</option>"

	fnQuickLinkTypeSelect = vBody
End Function

Function fnKeywordExistCheck(arr,gubun,topidx)
	Dim i, vQuery, j, arr2, vArr, vKwd, vCount, vResult
	arr2 = arr
	vArr = "," & arr & ","
	vResult = "0"
	vCount = 0

	'### arr 안에 같은게 있는지 체크.
	For i = LBound(Split(arr,",")) To UBound(Split(arr,","))
	
		For j = LBound(Split(arr2,",")) To UBound(Split(arr2,","))
			If (","&Trim(Split(arr,",")(i))&",") = (","&Trim(Split(arr2,",")(j))&",") Then
				vCount = vCount + 1
			End IF
		Next
		
		If vCount > 1 Then
			vResult = "1"
			Exit For
		Else
			vCount = 0
		End IF
		
	Next

	If vResult = "0" Then
		'### 전체 키워드로 검색. 중복 안됨.
		For i = LBound(Split(arr,",")) To UBound(Split(arr,","))
			vKwd = vKwd & "'" & Trim(Split(arr,",")(i)) & "',"
		Next
		vKwd = Left(vKwd, Len(vKwd)-1)
		
		'// 기존에는 tbl_search_keyword만 확인하여 중복되는 모든 keyword값이 있을경우 등록 안되게 막았지만,
		'// 2018년 2월 20일 수정 이후 에는 기간이 지났거나 사용하지 않는 키워드일경우엔 중복 등록 되게 변경함.
		vQuery = "SELECT count(topidx) FROM [db_sitemaster].[dbo].[tbl_search_keyword] as k "
		vQuery = vQuery & "inner join [db_sitemaster].dbo.[tbl_search_curator] as w on k.topidx = w.idx "
		vQuery = vQuery & "WHERE k.topgubun = '" & gubun & "' AND k.topidx <> '" & topidx & "' AND k.keyword in(" & vKwd & ")"
		vQuery = vQuery & "And w.useyn='y' And getdate() <= w.edate "
		rsget.CursorLocation = adUseClient
		rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
		If rsget(0) > 0 Then
			vResult = "2"
		Else
			vResult = "0"
		End If
		rsget.close
	End IF

	fnKeywordExistCheck = vResult
End Function
%>