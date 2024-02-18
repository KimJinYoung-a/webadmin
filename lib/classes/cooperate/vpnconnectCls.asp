<%
class Cvpnconnect_item
	public Fidx
	public Fstime
	public Fetime
	public Fequip
	public Fuserid
	public Fusername
	public Frealip
	public Fassignip
	public Floginstate
	public Fconstate
	public Fwhycon
	public Fwhyuserid
	public Fwhyregdate
	public Freguserid
	public Fregdate
	public Fsign
	public Fsigndate


    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class


Class Cvpnconnect_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	Public FRectSTime1
	Public FRectSTime2
	Public FRectETime1
	Public FRectETime2
	Public FRectSDate
	Public FRectEDate
	Public FRectIdx
	Public FRectIsSign
	Public FRectUserID


	Public Sub sbVPNLogList
		Dim sqlStr, i, sqladd

		If FRectSTime1 <> "" Then
			sqladd = sqladd & " and convert(varchar(10),v.stime,120) >= '" & FRectSTime1 & "' "
		End If
		
		If FRectSTime2 <> "" Then
			sqladd = sqladd & " and convert(varchar(10),v.stime,120) <= '" & FRectSTime2 & "' "
		End If
		
		If FRectETime1 <> "" Then
			sqladd = sqladd & " and convert(varchar(10),v.etime,120) >= '" & FRectETime1 & "' "
		End If
		
		If FRectETime2 <> "" Then
			sqladd = sqladd & " and convert(varchar(10),v.etime,120) <= '" & FRectETime2 & "' "
		End If
		
		If FRectIsSign = "x" Then
			sqladd = sqladd & " and v.sign = '' "
		ElseIf FRectIsSign = "o" Then
			sqladd = sqladd & " and v.sign <> '' "
		End If
		
		If FRectUserID <> "" Then
			sqladd = sqladd & " and v.userid = '" & FRectUserID & "' "
		End If
        
        if (application("Svr_Info")	= "Dev") then
        
        else
            sqladd = sqladd & " and v.idx>1136"
        end if

		sqlStr = "SELECT COUNT(v.idx), CEILING(CAST(Count(v.idx) AS FLOAT)/" & FPageSize & ") AS totPg "
		sqlStr = sqlStr & " FROM [db_board].[dbo].[tbl_vpn_connect_log] as v "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		rsget.Open sqlStr,dbget
		IF not rsget.EOF THEN
			FTotalCount = rsget(0)
			FTotalPage = rsget(1)
		END IF
		rsget.close
		
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		IF FTotalCount > 0 THEN
			sqlStr = "SELECT Top "&CStr(FPageSize*FCurrPage)&" "
			sqlStr = sqlStr & "v.idx, v.stime, v.etime, v.equip, v.userid, v.realip, v.assignip, v.loginstate, v.constate, v.whycon, v.whyuserid, v.regdate, v.reguserid, v.regdate, v.sign, v.signdate, "
			sqlStr = sqlStr & "(select username from [db_partner].[dbo].[tbl_user_tenbyten] where userid = (case when v.userid = 'VPN_eastone' then 'icommang' when v.userid = 'VPN_thensi' then 'thensi7' else replace(v.userid,'VPN_','') end)) as username "
			sqlStr = sqlStr & " FROM [db_board].[dbo].[tbl_vpn_connect_log] as v "
			sqlStr = sqlStr & " WHERE 1=1 " & sqladd
			sqlStr = sqlStr & " ORDER BY v.stime DESC"
			'response.write sqlStr & "<Br>"
			rsget.pagesize = FPageSize
			rsget.Open sqlStr,dbget,1
			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
			Redim preserve FItemList(FResultCount)
			i = 0
			
			If  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				Do until rsget.eof
					Set FItemList(i) = new Cvpnconnect_item
	
						FItemList(i).Fidx			= rsget("idx")
						FItemList(i).Fstime			= rsget("stime")
						FItemList(i).Fetime			= rsget("etime")
						FItemList(i).Fequip			= rsget("equip")
						FItemList(i).Fuserid		= rsget("userid")
						FItemList(i).Fusername		= rsget("username")
						FItemList(i).Frealip		= rsget("realip")
						FItemList(i).Fassignip		= rsget("assignip")
						FItemList(i).Floginstate	= rsget("loginstate")
						FItemList(i).Fconstate		= rsget("constate")
						FItemList(i).Fwhycon		= rsget("whycon")
						FItemList(i).Fwhyuserid		= rsget("whyuserid")
						FItemList(i).Fwhyregdate	= rsget("regdate") 
						FItemList(i).Freguserid		= rsget("reguserid")
						FItemList(i).Fregdate		= rsget("regdate")
						FItemList(i).Fsign			= rsget("sign")
						FItemList(i).Fsigndate		= rsget("signdate")
	
					i = i + 1
					rsget.moveNext
				Loop
			End If
			rsget.Close
		End If
	End Sub
	
	
	Public Sub sbVPNLogView
		Dim sqlStr, i, sqladd
		sqlStr = "SELECT whycon FROM [db_board].[dbo].[tbl_vpn_connect_log] WHERE idx = '" & FRectIdx & "'"
		rsget.Open sqlStr,dbget,1
		If Not rsget.Eof Then
			Set FOneItem = new Cvpnconnect_item
			FOneItem.Fwhycon = rsget("whycon")
		End If
		rsget.Close
	End Sub

	

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
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