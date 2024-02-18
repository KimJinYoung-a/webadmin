<%
'###########################################################
' Description : play class
' Hieditor : 이종화 생성
'			 2022.07.06 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################

Class CPlayItem
    public Fidx
    public Fmidx
    public Fdidx
    public Fvolnum
    public Ftitle
    public Ftitlestyle
    public Fmobgcolor
    public Fstartdate
    public Fstate
    public Fworktext
    public Fregdate
    public Flastupdate
    public FlastupdateID
    public Flastupdatename
    public FpartWDname
    public FpartWDID
    public FpartMKname
    public FpartMKID
    public FpartPBname
    public FpartPBID
    public Fcate
    public Fsubcopy
    public Fimgidx
    public Fimgurl
    public Flinkurl
    public Fentryvalue
    public Fdevice
    public Fuser
    public Fistagview
    public Ftagsdate
    public Ftagedate
    public Ftagannouncedate
    public Fkeyword
    
	public FCate1VideoURL
	public FCate1Type
	public FCate1Directer
	public FCate1LinkBanImg
	public FCate1LinkBanURL
	public FCate1CommTitle
	public FCate1Comment1
	public FCate1Comment2
	public FCate1Comment3
	public FCate1precomm1
	public FCate1precomm2
	public FCate1precomm3
	public FCate1VideoOrigin
	public FCate1RewardCopy
	public FCate3iconimg
	public FCate3EntryCont
	public FCate3EntrySDate
	public FCate3EntryEDate
	public FCate3AnnounDate
	public FCate3Notice
	public FCate3EntryMethod
	'// 2017.06.01 원승현 azit comma 스타일 추가
	public FCate31Directer
	public FCate41PCIsExec
	public FCate41PCExecFile
	public FCate41MoIsExec
	public FCate41MoExecFile
	public FCate41PCContent
	public FCate41MoContent
	public FCate42EntrySDate
	public FCate42EntryEDate
	public FCate42AnnounDate
	public FCate42WinnerTxt
	public FCate42WinnerValue
	public FCate42Notice
	public FCate42Entrycopy
	public FCate42Badgetag ''2017-07-17 유태욱 추가
	public FCate5Directer
	public FCate6VideoURL
	public FCate6BannSub
	public FCate6BannTitle
	public FCate6BannBtnTitle
	public FCate6BannBtnLink

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CPlay
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	public FDetailList
	public FPlayImageList
	public FPlayItemList
	public FPlayAzipList
	public FPlayThingThingWinlist
	public FPlayThingPCDownList
	public FPlayThingMoDownList

    public FRectIdx
    public FRectMIdx
    public FRectDIdx
    public FRectCate
    public FRectState
    public FRectVolnum
    public FRectTitle
    public FRectMKID
    public FRectWDID
    public FRectImgGubun


	public function fnPlayMasterList()
        dim sqlStr, sqlsearch, i

		If FRectState <> "" Then
			sqlsearch = sqlsearch & " and p.state = '" & FRectState & "' "
		End If

		If FRectVolnum <> "" Then
			sqlsearch = sqlsearch & " and p.volnum = '" & FRectVolnum & "' "
		End If
		
		If FRectMKID <> "" Then
			sqlsearch = sqlsearch & " and p.partMKid = '" & FRectMKID & "' "
		End If
		
		If FRectWDID <> "" Then
			sqlsearch = sqlsearch & " and p.partWDid = '" & FRectWDID & "' "
		End If


		'// 결과수 카운트
		sqlStr = "select count(p.midx) as cnt, CEILING(CAST(Count(p.midx) AS FLOAT)/" & FPageSize & ") AS totPg"
        sqlStr = sqlStr & " from [db_giftplus].[dbo].[tbl_play_master] as p with (nolock)"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
            FTotalPage	= rsget("totPg")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " p.midx, p.volnum, p.title, p.mo_bgcolor, p.startdate, p.state, p.regdate, p.lastupdate "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = p.partMKid and statediv = 'Y' and userid <> '') as partMKname "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = p.partWDid and statediv = 'Y' and userid <> '') as partWDname "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = p.partPBid and statediv = 'Y' and userid <> '') as partPBname "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten WHERE userid = p.lastupdateID and statediv = 'Y' and userid <> '') as lastupdatename "
		sqlStr = sqlStr & " from [db_giftplus].[dbo].[tbl_play_master] as p with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by p.volnum DESC, p.startdate DESC"

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
                set FItemList(i) = new CPlayItem

					FItemList(i).Fmidx			= rsget("midx")
					FItemList(i).Fvolnum			= rsget("volnum")
					FItemList(i).Ftitle			= db2html(rsget("title"))
					FItemList(i).Fmobgcolor		= rsget("mo_bgcolor")
					FItemList(i).Fstartdate		= rsget("startdate")
					FItemList(i).Fstate			= rsget("state")
					FItemList(i).Fregdate		= rsget("regdate")
					FItemList(i).Flastupdate	= rsget("lastupdate")
					FItemList(i).Flastupdatename	= rsget("lastupdatename")
					FItemList(i).FpartWDname	= rsget("partWDname")
					FItemList(i).FpartMKname	= rsget("partMKname")
					FItemList(i).FpartPBname	= rsget("partPBname")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function
    
	
	public Sub sbPlayMasterDetail()
		dim sqlStr, addsql
		
		If FRectMIdx <> "" Then
			addsql = addsql & " and m.midx = '" & FRectMIdx & "'"
		End If
		
		sqlStr = "select midx, volnum, title, mo_bgcolor, convert(varchar(10),startdate,120) as startdate, state, partWDid, partMKid, partPBid, "
		sqlStr = sqlStr & " worktext, regdate, lastupdate, lastupdateID "
		sqlStr = sqlStr & " from [db_giftplus].[dbo].[tbl_play_master] as m with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & addsql
		'response.write sqlStr
		
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		FResultCount = rsget.RecordCount
	
		set FOneItem = new CPlayItem
	
		if Not rsget.Eof then
	
			FOneItem.Fmidx			= rsget("midx")
			FOneItem.Fvolnum			= rsget("volnum")
			FOneItem.Ftitle			= db2html(rsget("title"))
			FOneItem.Fmobgcolor		= rsget("mo_bgcolor")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fstate			= rsget("state")
			FOneItem.Fworktext		= db2html(rsget("worktext"))
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Flastupdate		= rsget("lastupdate")
			FOneItem.FlastupdateID	= rsget("lastupdateID")
			FOneItem.FpartWDID		= rsget("partWDid")
			FOneItem.FpartMKID		= rsget("partMKid")
			FOneItem.FpartPBID		= rsget("partPBid")
	
		end if
		rsget.Close
		
		If FRectMIdx <> "" Then
			sqlStr = "select d.didx, d.cate, d.title, d.startdate, d.state, isNull(i.imgurl,'') as imgurl, isNull(i.linkurl,'') as linkurl, ([db_giftplus].[dbo].[getPlayCateName](d.cate)) as catename "
			sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten with (nolock) WHERE userid = d.partWDid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as partWDname "
			sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten with (nolock) WHERE userid = d.partPBid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as partPBname "
			sqlStr = sqlStr & " , d.viewcnt_w, d.viewcnt_m, d.viewcnt_a "
			sqlStr = sqlStr & " from [db_giftplus].[dbo].[tbl_play_detail] as d with (nolock)"
			sqlStr = sqlStr & " inner join [db_giftplus].[dbo].[tbl_play_image] as i with (nolock) on d.didx = i.didx "
			sqlStr = sqlStr & " left join [db_giftplus].[dbo].[tbl_play_cate] as c with (nolock) on d.cate = c.cate "
			sqlStr = sqlStr & "where d.midx = '" & FRectMIdx & "' and i.gubun = '" & FRectImgGubun & "' "
			sqlStr = sqlStr & "order by d.didx desc "
			'response.write sqlStr
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			
			if Not rsget.Eof then
				FDetailList = rsget.getRows()
			end if
			rsget.Close
		End If
	end Sub
	
	
	public Sub sbPlayCornerDetail()
		dim sqlStr, addsql

		sqlStr = "select * "
		sqlStr = sqlStr & " from [db_giftplus].[dbo].[tbl_play_detail] as d "
		If FRectCate = "1" Then
			sqlStr = sqlStr & " left join [db_giftplus].[dbo].[tbl_play_playlist] as p on d.didx = p.didx "

		End If
		sqlStr = sqlStr & "where d.didx = '" & FRectDIdx & "' "
		'response.write sqlStr
		
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		FResultCount = rsget.RecordCount
	
		set FOneItem = new CPlayItem
	
		if Not rsget.Eof then
	
			FOneItem.Fcate			= rsget("cate")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fstate			= rsget("state")
			FOneItem.Ftitle			= db2html(rsget("title"))
			FOneItem.Ftitlestyle		= db2html(rsget("titlestyle"))
			FOneItem.Fsubcopy			= db2html(rsget("subcopy"))
			FOneItem.Fmobgcolor		= rsget("mo_bgcolor")
			FOneItem.Fworktext		= db2html(rsget("worktext"))
			'FOneItem.Fregdate		= rsget("regdate")
			'FOneItem.Flastupdate		= rsget("lastupdate")
			'FOneItem.FlastupdateID	= rsget("lastupdateID")
			FOneItem.FpartWDID		= rsget("partWDid")
			FOneItem.FpartMKID		= rsget("partMKid")
			FOneItem.FpartPBID		= rsget("partPBid")
			FOneItem.Fistagview		= rsget("isTagView")
			FOneItem.Ftagsdate		= rsget("tag_sdate")
			FOneItem.Ftagedate		= rsget("tag_edate")
			FOneItem.Ftagannouncedate = rsget("tag_announcedate")
			FOneItem.Fkeyword			= db2html(rsget("keyword"))
			
			If FRectCate = "1" Then
				FOneItem.FCate1VideoURL		= rsget("videourl")
				FOneItem.FCate1Type			= rsget("type")
				FOneItem.FCate1Directer		= rsget("directer")
				FOneItem.FCate1CommTitle		= rsget("comm_title")
				FOneItem.FCate1Comment1		= rsget("comment1")
				FOneItem.FCate1Comment2		= rsget("comment2")
				FOneItem.FCate1Comment3		= rsget("comment3")
				FOneItem.FCate1precomm1		= rsget("precomm1")
				FOneItem.FCate1precomm2		= rsget("precomm2")
				FOneItem.FCate1precomm3		= rsget("precomm3")
				FOneItem.FCate1VideoOrigin	= rsget("videoorigin")
				FOneItem.FCate1RewardCopy	= rsget("rewardcopy")
			ElseIf FRectCate = "3" Then
				FOneItem.FCate3iconimg		= rsget("iconimg")
			ElseIf FRectCate = "41" Then
				FOneItem.FCate41PCIsExec		= rsget("pc_isExec")
				FOneItem.FCate41PCExecFile	= rsget("pc_execfile")
				FOneItem.FCate41MoIsExec		= rsget("mo_isExec")
				FOneItem.FCate41MoExecFile	= rsget("mo_execfile")
				FOneItem.FCate41PCContent	= db2html(rsget("pc_contents"))
				FOneItem.FCate41MoContent	= db2html(rsget("mo_contents"))
			End If
	
		end if
		rsget.Close
		
		If FRectCate = "3" Then '### azit
			sqlStr = "select groupnum, isNull(title,'') as imgurl, isNull(address,'') as linkurl, isNull(addrlink,'') as addrlink "
			sqlStr = sqlStr & "from [db_giftplus].[dbo].[tbl_play_azit] where didx = '" & FRectDIdx & "' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			if Not rsget.Eof then
				FPlayAzipList = rsget.getRows()
			end if
			rsget.Close
			
			
			sqlStr = "select * from [db_giftplus].[dbo].[tbl_play_azit_entry] where didx = '" & FRectDIdx & "' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			if Not rsget.Eof then
				FOneItem.FCate3EntryCont  = db2html(rsget("entry_content"))
				FOneItem.FCate3EntrySDate = rsget("entry_sdate")
				FOneItem.FCate3EntryEDate = rsget("entry_edate")
				FOneItem.FCate3AnnounDate = rsget("announce_date")
				FOneItem.FCate3Notice	 = db2html(rsget("notice"))
				FOneItem.FCate3EntryMethod = rsget("entry_method")
			end if
			rsget.Close
		'// 2017.06.01 원승현 azit comma 스타일 추가
		ElseIf FRectCate = "31" Then  '### AzitCOMMA
			sqlStr = "select * from [db_giftplus].[dbo].[tbl_play_azit_comma] where didx = '" & FRectDIdx & "' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			if Not rsget.Eof then
				FOneItem.FCate31Directer 	= rsget("directer")
			end if
			rsget.Close
		ElseIf FRectCate = "42" Then  '### THING thingthing
			sqlStr = "select * from [db_giftplus].[dbo].[tbl_play_thingthing] where didx = '" & FRectDIdx & "' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			if Not rsget.Eof then
				FOneItem.FCate42EntrySDate 	= rsget("entry_sdate")
				FOneItem.FCate42EntryEDate 	= rsget("entry_edate")
				FOneItem.FCate42AnnounDate 	= rsget("announce_date")
				FOneItem.FCate42WinnerTxt 	= rsget("winnertxt")
				FOneItem.FCate42WinnerValue = rsget("winnervalue")
				FOneItem.FCate42Notice		= rsget("notice")
				FOneItem.FCate42Entrycopy	= rsget("entrycopy")
				FOneItem.FCate42Badgetag	= rsget("badgetag")
			end if
			rsget.Close
			
			'### 차선 10개 리스트.
			sqlStr = "select isNull(winnervalue,'') as winnervalue "
			sqlStr = sqlStr & "from [db_giftplus].[dbo].[tbl_play_thingthing_winlist] where didx = '" & FRectDIdx & "' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			if Not rsget.Eof then
				FPlayThingThingWinlist = rsget.getRows()
			end if
			rsget.Close
		ElseIf FRectCate = "43" Then  '### THING 배경화면
			sqlStr = "select isNull(download,'') as download, isNull(link,'') as link "
			sqlStr = sqlStr & "from [db_giftplus].[dbo].[tbl_play_thingdownload] where didx = '" & FRectDIdx & "' and device = 'pc' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			if Not rsget.Eof then
				FPlayThingPCDownList = rsget.getRows()
			end if
			rsget.Close
			
			sqlStr = "select isNull(download,'') as download, isNull(link,'') as link "
			sqlStr = sqlStr & "from [db_giftplus].[dbo].[tbl_play_thingdownload] where didx = '" & FRectDIdx & "' and device = 'mo' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			if Not rsget.Eof then
				FPlayThingMoDownList = rsget.getRows()
			end if
			rsget.Close
		ElseIf FRectCate = "5" Then  '### COMMA
			sqlStr = "select * from [db_giftplus].[dbo].[tbl_play_comma] where didx = '" & FRectDIdx & "' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			if Not rsget.Eof then
				FOneItem.FCate5Directer 	= rsget("directer")
			end if
			rsget.Close
		ElseIf FRectCate = "6" Then  '### HOWHOW
			sqlStr = "select * from [db_giftplus].[dbo].[tbl_play_howhow] where didx = '" & FRectDIdx & "' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			if Not rsget.Eof then
				FOneItem.FCate6VideoURL 		= rsget("videourl")
				FOneItem.FCate6BannSub		= rsget("bannsub")
				FOneItem.FCate6BannTitle 	= rsget("banntitle")
				FOneItem.FCate6BannBtnTitle	= rsget("bannbtntitle")
				FOneItem.FCate6BannBtnLink	= rsget("bannbtnlink")
			end if
			rsget.Close
			
		End If

		sqlStr = "select idx, cate, gubun, isNull(imgurl,'') as imgurl, isNull(linkurl,'') as linkurl, imagecopy, isNull(sortno,'') as sortno, isNull(groupnum,'') as groupnum "
		sqlStr = sqlStr & "from [db_giftplus].[dbo].[tbl_play_image] where didx = '" & FRectDIdx & "' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		if Not rsget.Eof then
			FPlayImageList = rsget.getRows()
		end if
		rsget.Close
	end Sub
    
    
    public function fnGetCodeList()
    	dim sqlStr
    	
    	set FOneItem = new CPlayMoContentsItem
    	If FRectType <> "" Then
	    	sqlStr = "select * from db_sitemaster.dbo.tbl_play_mo_code where type = '" & FRectType & "'"
	    	rsget.Open sqlStr,dbget,1
	    	IF not rsget.EOF THEN
	    		FOneItem.Ftypename = rsget("typename")
	    		FOneItem.Fisusing = rsget("isusing")
	    	End IF
	    	rsget.Close
    	End If
    	
    	sqlStr = "select * from db_sitemaster.dbo.tbl_play_mo_code"
    	rsget.Open sqlStr,dbget,1
			IF not rsget.EOF THEN
				fnGetCodeList = rsget.getRows()
			End IF
    	rsget.Close
    end Function
    
    
    public function fnGetPlayThingThingUser()
    	dim sqlStr
		sqlStr = "select count(t.idx) as cnt, CEILING(CAST(Count(t.idx) AS FLOAT)/" & FPageSize & ") AS totPg "
        sqlStr = sqlStr & "from [db_giftplus].[dbo].[tbl_play_thingthing_entry] as t "
        sqlStr = sqlStr & "inner join db_user.dbo.tbl_logindata as l on t.userid = l.userid "
        sqlStr = sqlStr & "where t.didx = '" & FRectDIdx & "'"

		'response.write sqlStr &"<Br>"
		'response.end
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
            FTotalPage	= rsget("totPg")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " t.idx, t.userid, t.entryvalue, t.device, t.regdate, l.userlevel "
        sqlStr = sqlStr & "from [db_giftplus].[dbo].[tbl_play_thingthing_entry] as t "
        sqlStr = sqlStr & "inner join db_user.dbo.tbl_logindata as l on t.userid = l.userid "
		sqlStr = sqlStr & "where t.didx = '" & FRectDIdx & "'"
		sqlStr = sqlStr & "order by t.idx DESC"

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
                set FItemList(i) = new CPlayItem

					FItemList(i).Fidx			= rsget("idx")
					FItemList(i).Fentryvalue	= rsget("entryvalue")
					FItemList(i).Fdevice		= rsget("device")
					FItemList(i).Fregdate		= rsget("regdate")
					FItemList(i).Fuser		= rsget("userid") & "(<font color='"&getUserLevelColor(rsget("userlevel"))&"'>" & getUserLevelStr(rsget("userlevel")) & "</font>)"

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function
    
    
    public function fnGetPlayPlaylistUser()
    	dim sqlStr
		sqlStr = "select count(t.idx) as cnt, CEILING(CAST(Count(t.idx) AS FLOAT)/" & FPageSize & ") AS totPg "
        sqlStr = sqlStr & "from [db_giftplus].[dbo].[tbl_play_playlist_comment] as t "
        sqlStr = sqlStr & "inner join db_user.dbo.tbl_logindata as l on t.userid = l.userid "
        sqlStr = sqlStr & "where t.didx = '" & FRectDIdx & "'"

		'response.write sqlStr &"<Br>"
		'response.end
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
            FTotalPage	= rsget("totPg")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " t.idx, t.userid, t.comment1, t.comment2, t.comment3, t.device, t.regdate, l.userlevel "
        sqlStr = sqlStr & "from [db_giftplus].[dbo].[tbl_play_playlist_comment] as t "
        sqlStr = sqlStr & "inner join db_user.dbo.tbl_logindata as l on t.userid = l.userid "
		sqlStr = sqlStr & "where t.didx = '" & FRectDIdx & "'"
		sqlStr = sqlStr & "order by t.idx DESC"

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
                set FItemList(i) = new CPlayItem

					FItemList(i).Fidx			= rsget("idx")
					FItemList(i).FCate1Comment1 = rsget("comment1")
					FItemList(i).FCate1Comment2 = rsget("comment2")
					FItemList(i).FCate1Comment3 = rsget("comment3")
					FItemList(i).Fdevice		= rsget("device")
					FItemList(i).Fregdate		= rsget("regdate")
					FItemList(i).Fuser		= rsget("userid") & "(<font color='"&getUserLevelColor(rsget("userlevel"))&"'>" & getUserLevelStr(rsget("userlevel")) & "</font>)"

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
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


Function fnPlayImage(didx,ca,gb,gn,sno,v)
	dim sqlStr, vValue
	sqlStr = "select isNull(imgurl,'') as imgurl, isNull(linkurl,'') as linkurl, isNull(imagecopy,'') as imagecopy from [db_giftplus].[dbo].[tbl_play_image] "
	sqlStr = sqlStr & "where didx = '" & didx & "' and cate = '" & ca & "' and gubun = '" & gb & "' "
	If gn <> "" Then
		sqlStr = sqlStr & "and groupnum = '" & gn & "' "
	End If
	If sno <> "" Then
		sqlStr = sqlStr & "and sortno = '" & sno & "' "
	End If
	'response.write sqlStr &"<BR>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
	
	if not rsget.eof THEN
		If v = "i" Then
			vValue = rsget(0)
		ElseIf v = "l" Then
			vValue = rsget(1)
		ElseIf v = "c" Then
			vValue = rsget(2)
		End If
	end if
	rsget.Close
	fnPlayImage = vValue
end function


Function fnPlayImageSelect(arr,ca,gb,v)
'### 온니 1개인 경우. sortno 가 없는 경우.
	Dim i, vValue
	'select idx, cate, gubun, isNull(imgurl,'') as imgurl, isNull(linkurl,'') as linkurl, imagecopy, sortno 
	IF isArray(arr) THEN
		For i =0 To UBound(arr,2)
			If CStr(ca) = CStr(arr(1,i)) and CStr(gb) = CStr(arr(2,i)) Then
				If v = "i" Then
					vValue = arr(3,i)
				ElseIf v = "l" Then
					vValue = arr(4,i)
				ElseIf v = "c" Then
					vValue = db2html(arr(5,i))
				End If
				Exit For
			End IF
		Next
	End If
	fnPlayImageSelect = vValue
End Function


Function fnPlayImageSelectSortNo(arr,ca,gb,v,gn,sn)
'### gb가 여러개인 경우. sortno 가 지정 된 경우.
	Dim i, vValue
	'select idx, cate, gubun, isNull(imgurl,'') as imgurl, isNull(linkurl,'') as linkurl, imagecopy, sortno , groupnum
	IF isArray(arr) THEN
		For i =0 To UBound(arr,2)
			If CStr(ca) = CStr(arr(1,i)) and CStr(gb) = CStr(arr(2,i)) and CStr(gn) = CStr(arr(7,i)) and CStr(sn) = CStr(arr(6,i)) Then
				If v = "i" Then
					vValue = arr(3,i)
				ElseIf v = "l" Then
					vValue = arr(4,i)
				ElseIf v = "c" Then
					vValue = db2html(arr(5,i))
				ElseIf v = "x" Then
					vValue = arr(0,i)
				End If
				Exit For
			End IF
		Next
	End If
	fnPlayImageSelectSortNo = vValue
End Function


Function fnPlayAzitSelect(arr,gn,v)
	Dim i, vValue
	'select groupnum, isNull(title,'') as title, isNull(address,'') as address, isNull(addrlink,'') as addrlink
	IF isArray(arr) THEN
		For i =0 To UBound(arr,2)
			If CStr(gn) = CStr(arr(0,i)) Then
				vValue = arr(v,i)
				Exit For
			End IF
		Next
	End If
	fnPlayAzitSelect = vValue
End Function


Function fnPlayDownloadSelect(arr,v,ii)
	Dim i, vValue
	'select isNull(download,'') as download, isNull(link,'') as link
	IF isArray(arr) THEN
		For i =0 To UBound(arr,2)
			If CStr(i) = CStr(ii) Then
				vValue = arr(v,i)
				Exit For
			End IF
		Next
	End If
	fnPlayDownloadSelect = vValue
End Function


Function fnStateSelectBox(gubun,sstate)
	Dim sqlStr, vBody
	
	If gubun = "one" Then
		SELECT Case sstate
			Case "0" : vBody = "등록대기"
			Case "3" : vBody = "이미지등록요청"
			Case "4" : vBody = "코딩등록요청"
			Case "5" : vBody = "오픈요청"
			Case "7" : vBody = "오픈"
			Case "9" : vBody = "종료"
			Case Else : vBody = ""
		END SELECT
	Else
		vBody = vBody & "<option value="""" " & CHKIIF(sstate="","selected","") & "> - 선택 - </option>" & vbCrLf
		vBody = vBody & "<option value=""0"" " & CHKIIF(sstate="0","selected","") & ">등록대기</option>" & vbCrLf
		vBody = vBody & "<option value=""3"" " & CHKIIF(sstate="3","selected","") & ">이미지등록요청</option>" & vbCrLf
		vBody = vBody & "<option value=""4"" " & CHKIIF(sstate="4","selected","") & ">코딩등록요청</option>" & vbCrLf
		vBody = vBody & "<option value=""5"" " & CHKIIF(sstate="5","selected","") & ">오픈요청</option>" & vbCrLf
		vBody = vBody & "<option value=""7"" " & CHKIIF(sstate="7","selected","") & ">오픈</option>" & vbCrLf
		vBody = vBody & "<option value=""9"" " & CHKIIF(sstate="9","selected","") & ">종료</option>"
	End If

	fnStateSelectBox = vBody
End Function


Function fnTypeSelectBox(gubun,depth,ttype,isusing)
	Dim sqlStr, vBody, vCate
	vCate = ttype
	
	If depth = "1" AND Len(vCate) > 1 Then
		vCate = Left(vCate,1)
	End If
	
	sqlStr = "select cate, catename "
	sqlStr = sqlStr & " from [db_giftplus].dbo.tbl_play_cate"
	sqlStr = sqlStr & " where isusing = '" & isusing & "' "
	
	If depth <> "" Then
		sqlStr = sqlStr & " and Len(cate) = '" & depth & "'"
	End If
	
	If gubun = "one" Then
		sqlStr = sqlStr & " and cate = '" & vCate & "'"
	End If
	
	If gubun = "select" and vCate <> "" and depth > 1 Then
		sqlStr = sqlStr & " and Left(cate,1) = '" & Left(vCate,1) & "'"
	End If
	'response.write sqlStr
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
	
	If gubun = "select" Then
		vBody = vBody & "<option value="""" " & CHKIIF(vCate="","selected","") & "> - 선택 - </option>"
		
		if  not rsget.EOF  then
			do until rsget.EOF
				vBody = vBody & "<option value=""" & rsget("cate") & """"
				If CStr(rsget("cate")) = CStr(vCate) Then
					vBody = vBody & " selected"
				End If
				vBody = vBody & ">" & rsget("catename") & "</option>"
				
				rsget.movenext
			loop
		end if
		rsget.Close
	ElseIf gubun = "one" Then
		if  not rsget.EOF  then
		vBody = rsget("catename")
		end if
		rsget.Close
	End If
	
	fnTypeSelectBox = vBody
End Function


'/담당MD 리스트가져오기 (팀장 미만,직원 이상)
Sub sbGetpartid(ByVal selName, ByVal sIDValue, ByVal sScript,part_sn)
	Dim strSql, arrList, intLoop

	if part_sn = "" then exit sub

	strSql = " SELECT userid, username"
	strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "
	strSql = strSql & " WHERE part_sn IN("&part_sn&") and  posit_sn>='4' and  posit_sn<='12' and   isUsing=1  and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '' order by posit_sn, empno"

	'response.write strSql &"<Br>"
	rsget.Open strSql,dbget
	IF not rsget.eof THEN
	arrList = rsget.getRows()
	End IF
	rsget.close
%>
	<select name="<%=selName%>" <%=sScript%>>
	<option value="">선택</option>
	<%
	If isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
	%>
	<option value="<%=arrList(0,intLoop)%>" <%if arrList(0,intLoop) = sIDValue then %>selected<%end if%>><%=arrList(1,intLoop)%></option>
	<%
		Next
	End IF
	%>
	</select>
<%
End Sub


Function fnPlayItemList(didx)
	dim sqlStr, arr, i, t, iitem
	sqlStr = "select itemid from [db_giftplus].[dbo].[tbl_play_item] where didx = '" & didx & "' "
	'response.write sqlStr &"<BR>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
	
	if not rsget.eof THEN
		arr = rsget.getRows()
		For i =0 To UBound(arr,2)
			t = t & arr(0,i) & ","
		Next
		iitem = Trim(Left(t,(Len(t)-1)))
	end if
	rsget.Close
	fnPlayItemList = iitem
end function


Function fnPlayCateName(c)
	Dim vTmp
	SELECT CASE c
		Case "41" : vTmp = "THING. > thing"
		Case "42" : vTmp = "THING. > thingthing"
		Case "43" : vTmp = "THING. > 배경화면"
		Case "3" : vTmp = "TALK > AZIT&"
		'// 2017.06.01 원승현 azit comma 스타일 추가
		Case "31" : vTmp = "TALK > AZIT&COMMA"
		Case "1" : vTmp = "TALK > PLAYLIST♬"
		Case "21" : vTmp = "!NSPIRATION > DESIGN"
		Case "22" : vTmp = "!NSPIRATION > STYLE"
		Case "5" : vTmp = "!NSPIRATION > COMMA,"
		Case "6" : vTmp = "!NSPIRATION > HOWHOW?"
	END SELECT
	fnPlayCateName = vTmp
end function

'####### 이미지 구분값 #######
'
'	1 : 리스트이미지(직사각형)
'	2 : playlist 컨텐츠 이미지
'	3 : playlist 연결배너 PC 이미지
'	4 : inspiration design 컨텐츠 이미지
'	5 : inspiration style 컨텐츠 이미지
'	6 : azit Mo 컨텐츠 이미지
'	7 : azit 장소 이미지
'	8 : thingthing 롤링 이미지
'	9 : 배경화면 Mo 컨텐츠 이미지
'	10 : 배경화면 QR 이미지(저장된이미지링크만)
'	11 : 리스트이미지(정사각형)
'	12 : comma 컨텐츠 PC 상단 이미지
'	13 : comma 컨텐츠 Mo 상단 이미지
'	14 : comma 컨텐츠(에디터) 이미지
'	15 : comma 연결배너 PC 이미지
'	16 : comma 연결배너 Mo 이미지
'	17 : howhow 컨텐츠(에디터) 이미지
'	18 : playlist 연결배너 Mo 이미지
'	19 : azit PC 컨텐츠 이미지
'	20 : 배경화면 PC 컨텐츠 이미지
'	21 : thingthing 연결배너 PC 이미지
'	22 : thingthing 연결배너 Mo 이미지
'// 2017.06.01 원승현 azit comma 스타일 추가
'	23 : comma 컨텐츠 PC 상단 이미지
'	24 : comma 컨텐츠 Mo 상단 이미지
'	25 : comma 컨텐츠(에디터) 이미지
'	26 : comma 연결배너 PC 이미지
'	27 : comma 연결배너 Mo 이미지
'	28 : 검색리스트 배너 이미지
'############################
%>