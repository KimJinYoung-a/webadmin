<%
'#######################################################
' Description : �ֹ�/��� ��Ƽ
' History	:  ������ ����
'              2022.01.19 �ѿ�� ����
'#######################################################

Class CNotiItem
	Public FItemid
	Public FNotignb
	Public FChkData
	Public FNoticnt
	Public FRegdate
	Public FLastupdate
	Public FNotistr
	Public FIsconfirmed
	Public FLastconfirmDT
	Public FLastconfirmUser
	public forderserial

	Public Function getNotignbStr
		Select Case FNotignb
			Case "11"		getNotignbStr = "�ǸŰ�"
			Case Else		getNotignbStr = FNotignb
		End Select
	End Function

	Public Function getConfirmedStr
		If FIsconfirmed = 0 Then
			getConfirmedStr = "<font color='BLUE'>Ȯ����</font>"
		Else
			getConfirmedStr = "<font color='RED'>Ȯ�οϷ�</font>"
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CNoti
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	public frectselect_type
	public farrlist
	Public FRectItemID
	Public FRectNotignb	
	Public FRectIsconfirmed

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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

	Public Sub getCheckNotiList
		Dim sqlStr, i, addSql

		'��ǰ�ڵ� �˻�
        If FRectItemid <> "" then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and itemid in (" + FRectItemid + ")"
            End If
        End If

		'���� �˻�
		If FRectNotignb <> "" Then
			addSql = addSql & " and notignb = '"&FRectNotignb&"' "
		End If

		'Ȯ�ο��� �˻�
		If FRectIsconfirmed <> "" Then
			addSql = addSql & " and isconfirmed = '"&FRectIsconfirmed&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_contents].[dbo].[tbl_check_noti_log] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " itemid, notignb, chkData, noticnt, regdate, lastupdate, notistr, isnull(isconfirmed, 0) as isconfirmed, lastconfirmDT, lastconfirmUser "
		sqlStr = sqlStr & " FROM [db_contents].[dbo].[tbl_check_noti_log] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY isconfirmed , lastupdate DESC, noticnt DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CNotiItem
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FNotignb			= rsget("notignb")
					FItemList(i).FChkData			= rsget("chkData")
					FItemList(i).FNoticnt			= rsget("noticnt")
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FLastupdate		= rsget("lastupdate")
					FItemList(i).FNotistr			= rsget("notistr")
					FItemList(i).FIsconfirmed		= rsget("isconfirmed")
					FItemList(i).FLastconfirmDT		= rsget("lastconfirmDT")
					FItemList(i).FLastconfirmUser	= rsget("lastconfirmUser")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	' /admin/ordermaster/chulgobogo.asp
	Public Sub getchulgobogo
		Dim sqlStr, i, addSql

		sqlStr = "exec db_order.dbo.usp_TEN_logics_chulgo_alarm_v2_admin '"& frectselect_type &"'"

		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		dbget.CommandTimeout = 60*10   ' 10��
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			farrlist			= rsget.getrows
		End If
		rsget.Close
	End Sub
End Class
%>