<%
Class CBrandSellReportItem
	public Fuserid
	public Fuserdiv
	public Fmaeipdiv
	public Fdefaultmargine
	public Fsocname_kor
	public Fisusing
	public Fmduserid
	public Fregdate
	public Fitemcount

	public Fsellttl
	public Fbuyttl
	public Fmdusername

	public function GetUserDivName
		if Fuserdiv="02" then
			GetUserDivName = "디자인업체"
		elseif Fuserdiv="03" then
			GetUserDivName = "플라워업체"
		elseif Fuserdiv="04" then
			GetUserDivName = "패션업체"
		elseif Fuserdiv="05" then
			GetUserDivName = "쥬얼리업체"
		elseif Fuserdiv="06" then
			GetUserDivName = "케어업체"
		elseif Fuserdiv="07" then
			GetUserDivName = "애견업체"
		elseif Fuserdiv="08" then
			GetUserDivName = "보드게임"
		elseif Fuserdiv="13" then
			GetUserDivName = "여행몰업체"
		elseif Fuserdiv="14" then
			GetUserDivName = "강사"
		elseif Fuserdiv="20" then
			GetUserDivName = "텐바이텐소호"
		else
			GetUserDivName = Fuserdiv
		end if
	end function

	public function GetMaeipDivName
		if Fmaeipdiv="M" then
			GetMaeipDivName = "매입"
		elseif Fmaeipdiv="W" then
			GetMaeipDivName = "특정"
		elseif Fmaeipdiv="U" then
			GetMaeipDivName = "업체"
		else

		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CNewReport
	public FItemList()
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
    public FCurrPage

	public FRectFromDate
	public FRectToDate
	public FRectSearchType
	public FRectMakerid
	public FRectOrdType
	public FRectMdid

	public Sub GetNewBrandSellReport
		dim sqlStr, i, addsql
		If FRectMakerid <> "" Then
			addSql = addSql & " and c.userid='" & FRectMakerid & "'"
		End If

		If FRectMdid <> "" Then
			addSql = addSql & " and c.mduserid='" & FRectMdid & "'"
		End If

		sqlStr = "select top " + CStr(FPageSize) + " c.userid,c.userdiv,c.maeipdiv,c.defaultmargine,c.socname_kor,c.isusing,c.mduserid,c.regdate,c.itemcount"
		sqlStr = sqlStr + " ,IsNULL(T.sellttl,0) as sellttl, IsNULL(T.buyttl,0) as buyttl, y.username as mdusername"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select d.makerid,	sum(d.itemcost*d.itemno) as sellttl, sum(d.buycash*d.itemno) as buyttl"
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " 	 [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.regdate>='" + Cstr(FRectFromDate) + "'"
		sqlStr = sqlStr + " 	and m.regdate<'" + Cstr(FRectToDate) + "'"
		sqlStr = sqlStr + " 	and m.ipkumdiv>3"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.itemid<>0"
		sqlStr = sqlStr + " 	group by d.makerid"
		sqlStr = sqlStr + " ) T on c.userid=T.makerid"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten as y on c.mduserid = y.userid and isnull(c.mduserid, '') <> '' "
		sqlStr = sqlStr + " where c.userdiv<21" & addSql
		If FRectSearchType="N" Then
			sqlStr = sqlStr + " and datediff(d,c.regdate,getdate())<31"
		End If

'		If FRectSearchType="N" then
'			sqlStr = sqlStr + " order by T.sellttl/(datediff(d,c.regdate,getdate())+1) desc"
'		Else
'			sqlStr = sqlStr + " order by T.sellttl  desc"
'		End If

		If FRectSearchType="N" then
			If FRectOrdType="" then 
				sqlStr = sqlStr + " ORDER BY T.sellttl/(datediff(d,c.regdate,getdate())+1) DESC"
			ElseIf FRectOrdType="1" then
				sqlStr = sqlStr + " ORDER BY T.sellttl/(datediff(d,c.regdate,getdate())+1) DESC, c.regdate DESC "
			ElseIf FRectOrdType="2" then
				sqlStr = sqlStr + " ORDER BY T.sellttl/(datediff(d,c.regdate,getdate())+1) DESC, c.regdate ASC "
			ElseIf FRectOrdType="3" then
				sqlStr = sqlStr + " ORDER BY T.sellttl/(datediff(d,c.regdate,getdate())+1) DESC, T.sellttl DESC "
			ElseIf FRectOrdType="4" then
				sqlStr = sqlStr + " ORDER BY T.sellttl/(datediff(d,c.regdate,getdate())+1) DESC, T.sellttl ASC "
			End If
		Else
			If FRectOrdType="" then 
				sqlStr = sqlStr + " ORDER by T.sellttl DESC"
			ElseIf FRectOrdType="1" then
				sqlStr = sqlStr + " ORDER BY c.regdate DESC "
			ElseIf FRectOrdType="2" then
				sqlStr = sqlStr + " ORDER BY c.regdate ASC "
			ElseIf FRectOrdType="3" then
				sqlStr = sqlStr + " ORDER by T.sellttl DESC"
			ElseIf FRectOrdType="4" then
				sqlStr = sqlStr + " ORDER by T.sellttl ASC"
			End If
		End If
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(i) = new CBrandSellReportItem
			    FItemList(i).Fuserid        = rsget("userid")
				FItemList(i).Fuserdiv       = rsget("userdiv")
				FItemList(i).Fmaeipdiv      = rsget("maeipdiv")
				FItemList(i).Fdefaultmargine= rsget("defaultmargine")
				FItemList(i).Fsocname_kor   = db2html(rsget("socname_kor"))
				FItemList(i).Fisusing       = rsget("isusing")
				FItemList(i).Fmduserid      = rsget("mduserid")
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).Fitemcount		= rsget("itemcount")

				FItemList(i).Fsellttl       = rsget("sellttl")
				FItemList(i).Fbuyttl        = rsget("buyttl")
				FItemList(i).Fmdusername      = rsget("mdusername")

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close

	end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

Function fnGetMdlist(selectBoxName,selectedId)
	Dim tmp_str, strSql

%>
	<select name="<%=selectBoxName%>" class="select">
<%
	response.write("<option value='' selected>-선택-</option>")
		strSql = "	SELECT A.id, D.username " & _
			"		FROM [db_partner].[dbo].tbl_partner AS A " & _
			"		INNER JOIN [db_partner].[dbo].tbl_positInfo AS C ON A.posit_sn = C.posit_sn " & _
			"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS D ON A.id = D.userid " & _
			"	WHERE A.isusing = 'Y' AND A.userdiv < 999 AND A.id <> '' AND Left(A.id,10) <> 'streetshop' " & _
			"			AND A.part_sn in ('11', '21') " & _
			"			AND A.id != 'yanan716' " & _
			"			and ((C.posit_sn<=8) OR C.posit_sn in ('12', '13')) " & _
			"	ORDER BY A.part_sn ASC, A.posit_sn ASC, A.regdate ASC "
	rsget.Open strSql,dbget,1

	If not rsget.EOF Then
		Do Until rsget.EOF
			If rsget("id") = selectedId Then
				tmp_str = " selected"
			End If
			response.write("<option value='"&rsget("id")&"' "&tmp_str&">" + db2html(rsget("username")) + "</option>")
			tmp_str = ""
			rsget.MoveNext
		Loop
	End If
	rsget.close

	response.write("</select>")
End Function
%>