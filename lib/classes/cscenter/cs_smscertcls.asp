<%

Class CCSSMSCertItem
	''idx, userid, confDiv, usermail, usercell, smsCD, isConfirm, regdate, confDate, pFlag, evtFlag

	public Fidx
	public Fuserid
	public FconfDiv
	public Fusermail
	public Fusercell
	public FsmsCD
	public FisConfirm
	public Fregdate
	public FconfDate
	public FpFlag
	public FevtFlag

	public FREQDATE
	public FSENTDATE
	public FMSG

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end Class

Class CCSSMSCert
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectUserID
	public FRectUserMail
	public FRectUserCell
	public logicsdb

	public Sub GetCSSMSCertLogList()
		dim i, sqlStr, addSql

		if (FRectUserID <> "") then
			addSql = " 	and userid = '" & FRectUserID & "' "
		end if

		if (FRectUserMail <> "") then
			addSql = " 	and usermail = '" & FRectUserMail & "' "
		end if

		if (FRectUserCell <> "") then
			addSql = " 	and usercell = '" & FRectUserCell & "' "
		end if

		sqlStr = " select top " & (FCurrPage * FPageSize) & " idx, userid, confDiv, usermail, usercell, smsCD, isConfirm, regdate, confDate, pFlag, evtFlag "
		sqlStr = sqlStr + " from db_log.dbo.tbl_userConfirm "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "

		sqlStr = sqlStr + addSql

		sqlStr = sqlStr + " order by idx desc "
        'response.write sqlStr

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FResultCount = rsget.RecordCount

        redim preserve FItemList(FResultCount)
        if  not rsget.EOF  then
            i = 0
            do until rsget.eof
                set FItemList(i) = new CCSSMSCertItem

                FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fuserid        = rsget("userid")
				FItemList(i).FconfDiv       = rsget("confDiv")
				FItemList(i).Fusermail      = rsget("usermail")
				FItemList(i).Fusercell      = rsget("usercell")
				FItemList(i).FsmsCD         = rsget("smsCD")
				FItemList(i).FisConfirm     = rsget("isConfirm")
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).FconfDate      = rsget("confDate")
				FItemList(i).FpFlag         = rsget("pFlag")
				FItemList(i).FevtFlag       = rsget("evtFlag")

                rsget.MoveNext
                i = i + 1
            loop
        end if
        rsget.close

	end sub

	Public Sub GetCSKakaoLogList()
		Dim i, sqlStr, addSql

		If (FRectUserCell <> "") Then
			'addSql = " and replace(trim(PHONE),'-','') = '"& replace(trim(FRectUserCell),"-","") &"'"		부하때문에 안됨.인덱스를 안탐.
			addSql = " and PHONE = '"& trim(FRectUserCell) &"'"
		End If

		sqlStr = "SELECT TOP " & (FCurrPage * FPageSize)
		sqlStr = sqlStr & " MSGKEY AS idx , PHONE AS usercell, REQDATE, SENTDATE, MSG"
		sqlStr = sqlStr & " FROM "& logicsdb &"[db_kakaoMsg_v4].[dbo].[KKO_MSG_LOG] with (nolock)"
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & "	ORDER BY MSGKEY DESC"

        'response.write sqlStr & "<br>"
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly
        FResultCount = rsget_Logistics.RecordCount
        Redim preserve FItemList(FResultCount)
        If  not rsget_Logistics.EOF Then
            i = 0
            Do Until rsget_Logistics.eof
                SET FItemList(i) = new CCSSMSCertItem
					FItemList(i).Fidx		= rsget_Logistics("idx")
					FItemList(i).Fusercell	= rsget_Logistics("usercell")
					FItemList(i).FREQDATE	= rsget_Logistics("REQDATE")
					FItemList(i).FSENTDATE	= rsget_Logistics("SENTDATE")
					FItemList(i).FMSG      	= rsget_Logistics("MSG")
					rsget_Logistics.MoveNext
				i = i + 1
            Loop
        End If
        rsget_Logistics.close
	End Sub

	Public Sub GetCSKakaoLogList_cs()
		Dim i, sqlStr, addSql

		If (FRectUserCell <> "") Then
			'addSql = " and replace(trim(PHONE),'-','') = '"& replace(trim(FRectUserCell),"-","") &"'"		부하때문에 안됨.인덱스를 안탐.
			addSql = " and PHONE = '"& trim(FRectUserCell) &"'"
		End If

		sqlStr = "SELECT TOP " & (FCurrPage * FPageSize)
		sqlStr = sqlStr & " MSGKEY AS idx , PHONE AS usercell, REQDATE, SENTDATE, MSG"
		sqlStr = sqlStr & " FROM "& logicsdb &"[db_kakaoMsg_v4_cs].[dbo].[KKO_MSG_LOG] with (nolock)"
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & "	ORDER BY MSGKEY DESC"

        'response.write sqlStr & "<br>"
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly
        FResultCount = rsget_Logistics.RecordCount
        Redim preserve FItemList(FResultCount)
        If  not rsget_Logistics.EOF Then
            i = 0
            Do Until rsget_Logistics.eof
                SET FItemList(i) = new CCSSMSCertItem
					FItemList(i).Fidx		= rsget_Logistics("idx")
					FItemList(i).Fusercell	= rsget_Logistics("usercell")
					FItemList(i).FREQDATE	= rsget_Logistics("REQDATE")
					FItemList(i).FSENTDATE	= rsget_Logistics("SENTDATE")
					FItemList(i).FMSG      	= rsget_Logistics("MSG")
					rsget_Logistics.MoveNext
				i = i + 1
            Loop
        End If
        rsget_Logistics.close
	End Sub

	Public Sub GetCSKakaoLogList_mkt()
		Dim i, sqlStr, addSql

		If (FRectUserCell <> "") Then
			'addSql = " and replace(trim(PHONE),'-','') = '"& replace(trim(FRectUserCell),"-","") &"'"		부하때문에 안됨.인덱스를 안탐.
			addSql = " and PHONE = '"& trim(FRectUserCell) &"'"
		End If

		sqlStr = "SELECT TOP " & (FCurrPage * FPageSize)
		sqlStr = sqlStr & " MSGKEY AS idx , PHONE AS usercell, REQDATE, SENTDATE, MSG"
		sqlStr = sqlStr & " FROM "& logicsdb &"[db_kakaoMsg_v4_mkt].[dbo].[KKO_MSG_LOG] with (nolock)"
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & "	ORDER BY MSGKEY DESC"

        'response.write sqlStr & "<br>"
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly
        FResultCount = rsget_Logistics.RecordCount
        Redim preserve FItemList(FResultCount)
        If  not rsget_Logistics.EOF Then
            i = 0
            Do Until rsget_Logistics.eof
                SET FItemList(i) = new CCSSMSCertItem
					FItemList(i).Fidx		= rsget_Logistics("idx")
					FItemList(i).Fusercell	= rsget_Logistics("usercell")
					FItemList(i).FREQDATE	= rsget_Logistics("REQDATE")
					FItemList(i).FSENTDATE	= rsget_Logistics("SENTDATE")
					FItemList(i).FMSG      	= rsget_Logistics("MSG")
					rsget_Logistics.MoveNext
				i = i + 1
            Loop
        End If
        rsget_Logistics.close
	End Sub

    Private Sub Class_Initialize()
        FCurrPage       = 1
        FPageSize       = 20
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
		IF application("Svr_Info")="Dev" THEN
			logicsdb="LOGISTICSDB."
		end if
    End Sub

    Private Sub Class_Terminate()
		'
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
