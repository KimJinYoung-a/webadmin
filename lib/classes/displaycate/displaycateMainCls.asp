<%
class cDispCateMainOneItem
	public FCateCode
	public FDepth
	public FCateName
	public FCateName_E
	public FUseYN
	public FSortNo
	public FItemID
	public FIsDefault
	public FCDL
	public FCDM
	public FCDS
	public FItemName
	public FMakerID
	public FSmallImage
	public FIdx
	public Freguserid
	public Fregusername
	public Fworkcomment
	public Fregdate
	public Fstartdate
	public Ficon
	public Flinkurl
	public Fimgurl
	public Fsubcopy
	public Ftitle
	public Fcode
	public Ftype
	public Fpage
	public Fenddate
end Class

Class cDispCateMain
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	public FRectPage
	public FRectCateCode
	public FRectStartDate
	public FRectDepth
	public FRectCateName
	public FRectUseYN
	public FRectSortNo
	public FRectItemID
	public FRectIsDefault
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectMakerId
	public FRectItemName
	public FRectKeyword
	public FRectSellYN
	public FRectIsUsing
	public FRectDanjongyn
	public FRectLimityn
	public FRectSailYn
	public FRectDeliveryType
	public FRectSortDiv
	public FCateCode
	public FDepth
	public FCateName
	public FCateName_E
	public FCateFullName
	public FUseYN
	public FSortNo
	public FItemID
	public FIsDefault
	public FCDL
	public FCDM
	public FCDS
	public FItemName
	public FCateNameTitle
	public Fworkcomment
	public Freguserid
	public FRectIdx
	public Fidx

	
	Public Sub GetDispCateMainList()
		Dim sqlStr, i, addsql
		
		If FRectPage <> "" Then
			addsql = addsql & " AND m.page = '" & FRectPage & "' "
		End If
		
		If FRectStartDate <> "" Then
			addsql = addsql & " AND m.startdate = '" & FRectStartDate & "' "
		End IF

		sqlStr = "SELECT count(A.idx) AS cnt, CEILING(CAST(Count(A.idx) AS FLOAT)/5) AS totPg FROM ( " & vbCrLf
		sqlStr = sqlStr & "	SELECT m.idx " & vbCrLf
		sqlStr = sqlStr & " 	FROM [db_sitemaster].[dbo].[tbl_display_catemain] AS m " & vbCrLf
		sqlStr = sqlStr & " 	INNER JOIN [db_sitemaster].[dbo].[tbl_display_catemain_detail] AS d ON m.startdate = d.startdate AND m.catecode = d.catecode AND m.page = d.page " & vbCrLf
		sqlStr = sqlStr & " 	WHERE m.catecode = '" & FRectCateCode & "' " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " 	GROUP BY m.idx " & vbCrLf
		sqlStr = sqlStr & " ) AS A " & vbCrLf
'rw sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		sqlStr = "SELECT Top " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " 	m.idx, d.startdate, m.reguserid, m.regdate, d.page, "
		sqlStr = sqlStr & " 	(select username from db_partner.dbo.tbl_user_tenbyten where userid = m.reguserid) as regusername " & vbCrLf
		sqlStr = sqlStr & " FROM [db_sitemaster].[dbo].[tbl_display_catemain] AS m " & vbCrLf
		sqlStr = sqlStr & " INNER JOIN [db_sitemaster].[dbo].[tbl_display_catemain_detail] AS d ON m.startdate = d.startdate AND m.catecode = d.catecode AND m.page = d.page " & vbCrLf
		sqlStr = sqlStr & " 	WHERE m.catecode = '" & FRectCateCode & "' " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " 	GROUP BY m.idx, d.startdate, m.reguserid, m.regdate, d.page " & vbCrLf
		sqlStr = sqlStr & "ORDER BY m.idx DESC" & vbCrLf
'rw sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new cDispCateMainOneItem
					FItemList(i).FIdx 			= rsget("idx")
					FItemList(i).Fstartdate 	= rsget("startdate")
					FItemList(i).Freguserid		= rsget("reguserid")
					FItemList(i).Fregusername	= rsget("regusername")
					FItemList(i).Fregdate 		= rsget("regdate")
					FItemList(i).Fpage			= rsget("page")
					
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
		
	End Sub
	
	
	Public Sub GetDispCateMainComment()
	Dim sqlStr, i, addsql
		sqlStr = "SELECT m.idx, m.workcomment, m.reguserid FROM [db_sitemaster].[dbo].[tbl_display_catemain] AS m "
		sqlStr = sqlStr & " WHERE m.catecode = '" & FRectCateCode & "' AND m.page = '" & FRectPage & "' AND m.startdate = '" & FRectStartDate & "' "
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF THEN
			Fidx = rsget("idx")
			Fworkcomment = db2html(rsget("workcomment"))
			Freguserid = rsget("reguserid")
		END IF
		rsget.Close
	End Sub
	
	
	Public Function GetDispCateMainDetailList()
		Dim sqlStr, i, addsql

		sqlStr = "SELECT " & vbCrLf
		sqlStr = sqlStr & " 	d.idx, d.type, isNull(d.code,'') as code, isNull(d.title,'') as title, isNull(d.subcopy,'') as subcopy, "
		sqlStr = sqlStr & " 	isNull(d.imgurl,'') as imgurl, isNull(d.linkurl,'') as linkurl, isNull(d.icon,'') as icon, d.reguserid, "
		sqlStr = sqlStr & " 	(select username from db_partner.dbo.tbl_user_tenbyten where userid = d.reguserid) as regusername " & vbCrLf
		sqlStr = sqlStr & " FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] AS d " & vbCrLf
		sqlStr = sqlStr & " 	WHERE d.catecode = '" & FRectCateCode & "' AND d.page = '" & FRectPage & "' AND d.startdate = '" & FRectStartDate & "' " & vbCrLf
		sqlStr = sqlStr & "ORDER BY d.idx ASC" & vbCrLf
'rw sqlStr
		rsget.Open sqlStr,dbget,1
		IF not rsget.EOF THEN
			GetDispCateMainDetailList = rsget.getRows()
		END IF
		rsget.Close
		
	End Function
	
	
	Public Sub GetCateMainIssueList()
		Dim sqlStr, i, addsql
		
		If FRectIdx <> "" Then
			addsql = addsql & " and ci.idx = '" & FRectIdx & "' "
		End If
		
		sqlStr = sqlStr & "select count(ci.idx) AS cnt, CEILING(CAST(Count(ci.idx) AS FLOAT)/5) AS totPg " & vbCrLf
		sqlStr = sqlStr & "from [db_sitemaster].[dbo].[tbl_display_catemain_issue] as ci " & vbCrLf
		sqlStr = sqlStr & "WHERE ci.catecode = '" & FRectCateCode & "' " & addsql & " "
'rw sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		
		sqlStr = "SELECT Top " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " 	ci.idx, ci.imgurl, ci.linkurl, ci.title, ci.subcopy, convert(varchar(10),ci.startdate,120) as startdate, convert(varchar(10),ci.enddate,120) as enddate, ci.reguserid, ci.regdate, "
		sqlStr = sqlStr & " 	(select username from db_partner.dbo.tbl_user_tenbyten where userid = ci.reguserid) as regusername " & vbCrLf
		sqlStr = sqlStr & " FROM [db_sitemaster].[dbo].[tbl_display_catemain_issue] AS ci " & vbCrLf
		sqlStr = sqlStr & " 	WHERE ci.catecode = '" & FRectCateCode & "' " & addsql & " " & vbCrLf
		sqlStr = sqlStr & "ORDER BY ci.idx DESC" & vbCrLf
'rw sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new cDispCateMainOneItem
					FItemList(i).FIdx 			= rsget("idx")
					FItemList(i).Fimgurl		= rsget("imgurl")
					FItemList(i).Flinkurl		= rsget("linkurl")
					FItemList(i).Ftitle			= db2html(rsget("title"))
					FItemList(i).Fsubcopy		= db2html(rsget("subcopy"))
					FItemList(i).Fstartdate 	= rsget("startdate")
					FItemList(i).Fenddate 		= rsget("enddate")
					FItemList(i).Freguserid		= rsget("reguserid")
					FItemList(i).Fregusername	= rsget("regusername")
					FItemList(i).Fregdate 		= rsget("regdate")
					
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	
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
	End Sub

	Private Sub Class_Terminate()
    End Sub
End Class


'// 카테고리 Histoty 출력
Sub printCategoryHistory(code)
	dim strHistory, strLink, SQL, i, j
	j = (len(code)/3)
    
	'히스토리 기본
	strHistory = "HOME"

	'// 카테고리 이름 접수
	SQL = "SELECT ([db_item].[dbo].[getCateCodeFullDepthName]('" & code & "'))"
	rsget.Open SQL, dbget, 1

	if NOT(rsget.EOF or rsget.BOF) then
		for i = 1 to j
			strHistory = strHistory & "&nbsp;&gt;&nbsp;"
			if i = j then
				strHistory = strHistory & "<strong>" & Split(db2html(rsget(0)),"^^")(i-1) & "</strong>"
			else
				strHistory = strHistory & Split(db2html(rsget(0)),"^^")(i-1)
			end if
		next
	end if
	
	rsget.Close

	Response.Write strHistory
end Sub


Function fnGetEventIcon(issale, isgift, iscoupon, isOnlyTen, isoneplusone, isfreedelivery, isbookingsell, iscomment)
	Dim vIcon
	If isbookingsell = True Then
		fnGetEventIcon = "tagReserv"
		EXIT Function
	ElseIf issale = True Then
		fnGetEventIcon = "tagRed"
		EXIT Function
	ElseIf iscoupon = True Then
		fnGetEventIcon = "tagGreen"
		EXIT Function
	ElseIf isoneplusone = True Then
		fnGetEventIcon = "tagOneplus"
		EXIT Function
	ElseIf isgift = True Then
		fnGetEventIcon = "tagGift"
		EXIT Function
	ElseIf isOnlyTen = True Then
		fnGetEventIcon = "tagOnly"
		EXIT Function
	ElseIf isfreedelivery = True Then
		fnGetEventIcon = "tagFreeship"
		EXIT Function
	ElseIf iscomment = True Then
		fnGetEventIcon = "tagInvolve"
		EXIT Function
	Else
		fnGetEventIcon = "tagNew"
		EXIT Function
	End If
End Function

Function fnGetEventIconName(icon)
	SELECT Case icon
		Case "tagReserv" : fnGetEventIconName = "예약판매"
		Case "tagRed" : fnGetEventIconName = "SALE"
		Case "tagGreen" : fnGetEventIconName = "쿠폰"
		Case "tagOneplus" : fnGetEventIconName = "1+1"
		Case "tagGift" : fnGetEventIconName = "GIFT"
		Case "tagOnly" : fnGetEventIconName = "ONLY"
		Case "tagFreeship" : fnGetEventIconName = "무료배송"
		Case "tagInvolve" : fnGetEventIconName = "참여"
		Case "tagNew" : fnGetEventIconName = "NEW"
	END SELECT
End Function

Function chrbyte(str,chrlen,dot)

    Dim charat, wLen, cut_len, ext_chr, cblp

    if IsNULL(str) then Exit function

    for cblp=1 to len(str)
        charat=mid(str, cblp, 1)
        if asc(charat)>0 and asc(charat)<255 then
            wLen=wLen+1
        else
            wLen=wLen+2
        end if

        if wLen >= cint(chrlen) then
           cut_len = cblp
           exit for
        end if
    next

    if len(cut_len) = 0 then
        cut_len = len(str)
    end if

	if len(str)>cut_len and dot="Y" then
		ext_chr = "..."
	else
		ext_chr = ""
	end if

    chrbyte = Trim(left(str,cut_len)) & ext_chr

end function
%>