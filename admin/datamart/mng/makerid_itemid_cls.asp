<%
'### MIS MakeridItemidStatistic
Class CMISItem
    public Fidx

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMIS
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectIdx
	public FRectCompareKey
	public FRectSDate
	public FRectEDate
	public FItemTotalCount

    public function fnGetItemSellStdateByDispmdcatecode()
    	dim sqlStr
 
    	sqlStr = "EXEC [db_analyze_data_raw].[dbo].[sp_Ten_item_sellSTDate_by_dispcate1_v2] '" & FRectSDate & "','" & FRectEDate & "'"

   		'response.write sqlStr & "<Br>"
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
		IF not rsAnalget.EOF THEN
			fnGetItemSellStdateByDispmdcatecode = rsAnalget.getRows()
		End IF
    	rsAnalget.Close
    end Function

    public function fnGetItemSellStdateByDisp()
    	dim sqlStr
    	sqlStr = "EXEC [db_analyze_data_raw].[dbo].[sp_Ten_item_sellSTDate_by_dispcate1] '" & FRectSDate & "','" & FRectEDate & "'"
    	''response.write sqlStr
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
		IF not rsAnalget.EOF THEN
			fnGetItemSellStdateByDisp = rsAnalget.getRows()
		End IF
    	rsAnalget.Close

    end Function

    public function fnGetUserCitemSellStdateByDispmdcatecode()
    	dim sqlStr

    	sqlStr = "EXEC [db_analyze_data_raw].[dbo].[sp_Ten_userc_itemSellStdate_by_dispcate1_v2] '" & FRectSDate & "','" & FRectEDate & "'"

   		'response.write sqlStr & "<Br>"
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
		IF not rsAnalget.EOF THEN
			fnGetUserCitemSellStdateByDispmdcatecode = rsAnalget.getRows()
		End IF
    	rsAnalget.Close

    end Function

    public function fnGetUserCitemSellStdateByDisp()
    	dim sqlStr
    	sqlStr = "EXEC [db_analyze_data_raw].[dbo].[sp_Ten_userc_itemSellStdate_by_dispcate1] '" & FRectSDate & "','" & FRectEDate & "'"
    	''response.write sqlStr
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
		IF not rsAnalget.EOF THEN
			fnGetUserCitemSellStdateByDisp = rsAnalget.getRows()
		End IF
    	rsAnalget.Close

    end Function


    public function fnGetItemTotalCountByDisp()
    	dim sqlStr
    	sqlStr = "select "
		sqlStr = sqlStr & "	sum(d101cnt) as d101cnt, sum(d102cnt) as d102cnt, sum(d103cnt) as d103cnt, sum(d104cnt) as d104cnt, sum(d124cnt) as d124cnt, "
		sqlStr = sqlStr & "	sum(d121cnt) as d121cnt, sum(d122cnt) as d122cnt, sum(d120cnt) as d120cnt, sum(d112cnt) as d112cnt, "
		sqlStr = sqlStr & "	sum(d119cnt) as d119cnt, sum(d117cnt) as d117cnt, sum(d116cnt) as d116cnt, sum(d125cnt) as d125cnt, sum(d118cnt) as d118cnt, "
		sqlStr = sqlStr & "	sum(d115cnt) as d115cnt, sum(d110cnt) as d110cnt, sum(d000cnt) as d000cnt "
		sqlStr = sqlStr & "from "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "	select "
		sqlStr = sqlStr & "	case when dispcate1 = 101 then count(itemid) else 0 end as d101cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 102 then count(itemid) else 0 end as d102cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 103 then count(itemid) else 0 end as d103cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 104 then count(itemid) else 0 end as d104cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 124 then count(itemid) else 0 end as d124cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 121 then count(itemid) else 0 end as d121cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 122 then count(itemid) else 0 end as d122cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 120 then count(itemid) else 0 end as d120cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 112 then count(itemid) else 0 end as d112cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 119 then count(itemid) else 0 end as d119cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 117 then count(itemid) else 0 end as d117cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 116 then count(itemid) else 0 end as d116cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 125 then count(itemid) else 0 end as d125cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 118 then count(itemid) else 0 end as d118cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 115 then count(itemid) else 0 end as d115cnt, "
		sqlStr = sqlStr & "	case when dispcate1 = 110 then count(itemid) else 0 end as d110cnt, "
		sqlStr = sqlStr & "	case when dispcate1 is NULL then count(itemid) else 0 end as d000cnt "
		sqlStr = sqlStr & "	from [db_analyze_data_raw].[dbo].[tbl_item] as i with (nolock)"
		sqlStr = sqlStr & "	where i.sellSTDate is Not Null "
		sqlStr = sqlStr & "	group by i.dispcate1 "
		sqlStr = sqlStr & ") as A "
    	'response.write sqlStr
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
		IF not rsAnalget.EOF THEN
			FItemTotalCount = rsAnalget(0) + rsAnalget(1) + rsAnalget(2) + rsAnalget(3) + rsAnalget(4) + rsAnalget(5) + rsAnalget(6) + rsAnalget(7) + rsAnalget(8)
			FItemTotalCount = FItemTotalCount + rsAnalget(9) + rsAnalget(10) + rsAnalget(11) + rsAnalget(12) + rsAnalget(13) + rsAnalget(14) + rsAnalget(15) + rsAnalget(16)
			fnGetItemTotalCountByDisp = rsAnalget.getRows()
		End IF
    	rsAnalget.Close

    end Function

    public function fnGetItemTotalCountByDispmdcatecode()
    	dim sqlStr
    	sqlStr = "select "
		sqlStr = sqlStr & "	sum(d101cnt) as d101cnt, sum(d102cnt) as d102cnt, sum(d103cnt) as d103cnt, sum(d104cnt) as d104cnt, sum(d124cnt) as d124cnt, "
		sqlStr = sqlStr & "	sum(d121cnt) as d121cnt, sum(d122cnt) as d122cnt, sum(d120cnt) as d120cnt, sum(d112cnt) as d112cnt, "
		sqlStr = sqlStr & "	sum(d119cnt) as d119cnt, sum(d117cnt) as d117cnt, sum(d116cnt) as d116cnt, sum(d125cnt) as d125cnt, sum(d118cnt) as d118cnt, "
		sqlStr = sqlStr & "	sum(d115cnt) as d115cnt, sum(d110cnt) as d110cnt, sum(d000cnt) as d000cnt "
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & "		select "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 101 then count(itemid) else 0 end as d101cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 102 then count(itemid) else 0 end as d102cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 103 then count(itemid) else 0 end as d103cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 104 then count(itemid) else 0 end as d104cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 124 then count(itemid) else 0 end as d124cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 121 then count(itemid) else 0 end as d121cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 122 then count(itemid) else 0 end as d122cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 120 then count(itemid) else 0 end as d120cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 112 then count(itemid) else 0 end as d112cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 119 then count(itemid) else 0 end as d119cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 117 then count(itemid) else 0 end as d117cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 116 then count(itemid) else 0 end as d116cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 125 then count(itemid) else 0 end as d125cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 118 then count(itemid) else 0 end as d118cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 115 then count(itemid) else 0 end as d115cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode = 110 then count(itemid) else 0 end as d110cnt, "
		sqlStr = sqlStr & "		case when c.standardmdcatecode is NULL then count(itemid) else 0 end as d000cnt "
		sqlStr = sqlStr & "		from [db_analyze_data_raw].dbo.tbl_user_c c with (nolock)"
		sqlStr = sqlStr & "		join [db_analyze_data_raw].[dbo].[tbl_item] as i with (nolock)"
		sqlStr = sqlStr & "			on c.userid=i.makerid"
		sqlStr = sqlStr & "		where i.sellSTDate is Not Null "
		sqlStr = sqlStr & "		group by c.standardmdcatecode"
		sqlStr = sqlStr & " ) as A "

    	'response.write sqlStr & "<Br>"
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
		IF not rsAnalget.EOF THEN
			FItemTotalCount = rsAnalget(0) + rsAnalget(1) + rsAnalget(2) + rsAnalget(3) + rsAnalget(4) + rsAnalget(5) + rsAnalget(6) + rsAnalget(7) + rsAnalget(8)
			FItemTotalCount = FItemTotalCount + rsAnalget(9) + rsAnalget(10) + rsAnalget(11) + rsAnalget(12) + rsAnalget(13) + rsAnalget(14) + rsAnalget(15) + rsAnalget(16)
			fnGetItemTotalCountByDispmdcatecode = rsAnalget.getRows()
		End IF
    	rsAnalget.Close

    end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
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


Function fnCompareValue(v1, v2)
	Dim vValue

	If v1 = 0 AND v2 = 0 Then
		vValue = 0
	Else
		If v1 <> 0 AND v2 <> 0 Then
			vValue = (v1/v2) * 100
		ElseIf v1 <> 0 AND v2 = 0 Then
			vValue = v1 * 100
		ElseIf v1 = 0 AND v2 <> 0 Then
			vValue = -(v2*100)
		End If
	End If

	fnCompareValue = vValue
End Function


Function fnCompareUpDownValue(g, v1, v2)
	Dim vValue

	If g = "up" Then
		If Fix(Split(v1,"|")(0)) > Fix(Split(v2,"|")(0)) Then
			vValue = v1
		Else
			vValue = v2
		End If
	ElseIf g = "down" Then
		If Fix(Split(v1,"|")(0)) < Fix(Split(v2,"|")(0)) Then
			vValue = v1
		Else
			vValue = v2
		End If
	End If
'response.write v1 &",,,"& v2
	fnCompareUpDownValue = vValue
End Function


Function fnBGcolorCompare(a, b, gubun)
	Dim vTmp
	If a <> "" and b <> "" Then
		If gubun = "normal" Then
			If CDbl(a) > CDbl(b) Then
				vTmp = "bgred"
			ElseIf CDbl(a) < CDbl(b) Then
				vTmp = "bgblue"
			End If
		ElseIf gubun = "title" Then
			If CDbl(a) > CDbl(b) Then
				vTmp = "bgredtt"
			ElseIf CDbl(a) < CDbl(b) Then
				vTmp = "bgbluett"
			End If

			If vTmp = "" Then
				vTmp = "bggraytt"
			End If
		End If
	Else
		If gubun = "title" Then
			vTmp = "bggraytt"
		End If
	End If
	fnBGcolorCompare = vTmp
End Function

Function fnWeekNameReturn(w)
	Dim vWeek
	SELECT CASE w
		Case 1 : vWeek = "일"
		Case 2 : vWeek = "월"
		Case 3 : vWeek = "화"
		Case 4 : vWeek = "수"
		Case 5 : vWeek = "목"
		Case 6 : vWeek = "금"
		Case 7 : vWeek = "토"
	End SELECT
	fnWeekNameReturn = vWeek
End Function
%>
