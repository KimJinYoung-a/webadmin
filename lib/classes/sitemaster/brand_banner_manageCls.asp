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
	public Fcompany_name
	public Fsocname
	public Fsocname_kor

    
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
    public FRectMasterIDX
	public FRectBrandID
	
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
	
		If FRectUseYN <> "" Then	'### 사용 유형
			sqlsearch = sqlsearch & " and q.isusing = '" & FRectUseYN & "' "
		End IF

		If FRectSearchTxt <> "" Then
			sqlsearch = sqlsearch & " and q.name like '%" & FRectSearchTxt & "%'"
		End If

		If FRectEndType <> "" Then	'### 종료 유형
			If FRectEndType = "now" Then
				sqlsearch = sqlsearch & " and m.sdate <= getdate() and m.edate >= getdate() "
			ElseIf FRectEndType = "end" Then
				sqlsearch = sqlsearch & " and m.edate < getdate() "
			End If
		End If

		'// 결과수 카운트
		sqlStr = "select count(q.idx) as cnt, CEILING(CAST(Count(q.idx) AS FLOAT)/" & FPageSize & ") AS totPg"
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_brand_link_banner] as q "
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
		sqlStr = sqlStr & " q.idx, q.name, q.sdate, q.edate, q.url_pc,"
        sqlStr = sqlStr & " q.url_m, q.isusing, t.username as regusername, q.lastupdatedate"
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_brand_link_banner] as q "
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
					FItemList(i).Fquickname		= db2html(rsget("name"))
					FItemList(i).Fsdate			= rsget("sdate")
					FItemList(i).Fedate			= rsget("edate")
					FItemList(i).Furl_pc			= rsget("url_pc")
					FItemList(i).Furl_m			= rsget("url_m")
					FItemList(i).Fuseyn			= rsget("isusing")
					FItemList(i).Fregusername	= rsget("regusername")
					FItemList(i).Fregdate			= rsget("lastupdatedate")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	public Sub sbQuickLinkDetail()
		dim sqlStr, addsql

		sqlStr = "select "
		sqlStr = sqlStr & " q.idx, q.name, q.url_pc, q.url_m, q.sdate, q.edate, q.isusing, "
		sqlStr = sqlStr & " q.reguserid, q.regdate, q.lastupdateid, q.lastupdatedate, q.qimgpc, q.qimgm,"
		sqlStr = sqlStr & " (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = q.reguserid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as regusername "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = q.lastupdateid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as lastusername "
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_brand_link_banner] as q "
		sqlStr = sqlStr & " where q.idx = '" & FRectIdx & "'"
		'response.write sqlStr
		
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		FResultCount = rsget.RecordCount
	
		set FOneItem = new CSearchMngItem
	
		if Not rsget.Eof then
	
			FOneItem.Fidx				= rsget("idx")
			FOneItem.Fquickname		= db2html(rsget("name"))
			FOneItem.Furl_pc			= rsget("url_pc")
			FOneItem.Furl_m			= rsget("url_m")
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
			FOneItem.Fuseyn			= rsget("isusing")
			FOneItem.Freguserid		= rsget("reguserid")
			FOneItem.Fregusername	= rsget("regusername")
			FOneItem.Fregdate			= rsget("regdate")
			FOneItem.Flastuserid		= rsget("lastupdateid")
			FOneItem.Flastusername	= rsget("lastusername")
			FOneItem.Flastdate		= rsget("lastupdatedate")
			FOneItem.Fqimgpc			= rsget("qimgpc")
			FOneItem.Fqimgm			= rsget("qimgm")

		end if
		rsget.Close
	
	end Sub

	public function fnQuickLinkBrandList()
        dim sqlStr, sqlsearch, i
	
		If FRectSearchTxt <> "" Then	'### 사용 유형
            if FRectSearchGubun="brandid" then
			    sqlsearch = sqlsearch & " and q.brandid = '" & FRectSearchTxt & "' "
            elseif FRectSearchGubun="company_name" then
                sqlsearch = sqlsearch & " and p.company_name like '%" & FRectSearchTxt & "%'"
            elseif FRectSearchGubun="socname_kor" then
                sqlStr = sqlStr + " and c.socname_kor like '%" + FRectSearchTxt + "%'"
            end if
		End IF

		'// 결과수 카운트
		sqlStr = "select count(q.idx) as cnt, CEILING(CAST(Count(q.idx) AS FLOAT)/" & FPageSize & ") AS totPg"
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_brand_link_banner_brand_list] as q "
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c as c on q.brandid=c.userid" & vbcrlf
		sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner as p on q.brandid=p.id" & vbcrlf
        sqlStr = sqlStr & " where q.masteridx=" & FRectMasterIDX & vbcrlf
        sqlStr = sqlStr & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
            FTotalCount = rsget("cnt")
            FTotalPage	= rsget("totPg")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " q.idx, q.brandid, c.socname_kor, c.socname, p.company_name, q.regdate"
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_brand_link_banner_brand_list] as q "
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c as c on q.brandid=c.userid" & vbcrlf
		sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner as p on q.brandid=p.id" & vbcrlf
		sqlStr = sqlStr & " where q.masteridx=" & FRectMasterIDX & vbcrlf
		sqlStr = sqlStr & " and q.isusing='Y'" & vbcrlf
        sqlStr = sqlStr & sqlsearch
		sqlStr = sqlStr & " order by q.idx DESC"

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

					FItemList(i).Fidx = rsget("idx")
                    FItemList(i).Fbrandid = rsget("brandid")
					FItemList(i).Fcompany_name = db2html(rsget("company_name"))
					FItemList(i).Fsocname = db2html(rsget("socname"))
					FItemList(i).Fsocname_kor = db2html(rsget("socname_kor"))
					FItemList(i).Fregdate = rsget("regdate")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

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