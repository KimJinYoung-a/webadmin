<%
'###########################################################
' Description :  play class
' History : 2013.09.03 이종화 생성
'			2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<%
Class CPlayContentsItem
	public fstyle_html_m
    public Fidx
	Public Fidxsub
	public Flistimg
	Public Fmainimg
	public Fviewimg
	public Fviewtitle
	public Fviewtext
	Public Freservationdate
	Public Fstate
	Public Fviewno
	Public Forgimg
	Public Fworktext
	Public Fvideourl

	Public FSubtitle
	Public FIsusing
	Public FPPimg
	Public FRegdate

	Public Ftagname
	Public Ftagurl
	Public Ftagurl2
	Public Ftagurl3
	Public Ftagurl4

	Public Ftagcnt

	Public Fviewimg1
	Public Fviewimg2
	Public Fviewimg3
	Public Fviewimg4
	Public Fviewimg5

	Public Ftextimg
	Public FpartMDid
	Public FpartWDid
	Public FpartMKid
	public FpartMKname
	Public FpartMDname
	Public FpartWDname
	Public Fitemcnt

	Public Fitemcnt1
	Public Fitemcnt2
	Public Fitemcnt3
	Public Fitemcnt4
	Public Fitemcnt5

	public Fplaymainimg
	public Fbeforeimg
	public Fafterimg
	public Ftopbgimg
	Public FmainTopBGColor
	public Fsideltimg
	public Fsidertimg
	public FsubBGColor
	public Fviewcontents
	public Fviewthumbimg1
	public Fviewthumbimg2
	Public Fmyplayimg
	Public FvideourlM
	Public Fmo_contents '//모바일용 컨텐츠
	Public Fmo_exec_check '//모바일용 execute 체크 
	Public Fexec_check '//웹용 execute 체크 
	Public Fexec_filepath '//웹용 filepath execute 있을경우만 입력
	Public Fmo_idx '//모바일용 IDX


    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CPlayContents
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectIdx
	Public FRectsubidx
    public FRectgIdx
    public FRectPlaycate

	Public FRecttitle
	Public FRectstate

	Public FRPlaycate
	Public FRectTag

	Public FRectNo

	Public FRectpartMDid
	Public FRectpartWDid

	Public FRectIsusing
	Public FRectViewTitle

	'play ground 2013-09-16 이종화
	public Sub GetRowGroundMain()
	dim sqlStr
	sqlStr = "select * "
	sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_play_ground_main"
	sqlStr = sqlStr + " where gidx=" + CStr(FRectIdx)
	'response.write sqlStr
	rsget.Open SqlStr, dbget, 1
	FResultCount = rsget.RecordCount

	set FOneItem = new CPlayContentsItem

	if Not rsget.Eof then

		FOneItem.Fidx						= rsget("gidx")
		FOneItem.Flistimg					= rsget("titleimg")
		FOneItem.Fmainimg					= rsget("mainimg")
		FOneItem.Fviewtitle				= rsget("viewtitle")
		FOneItem.Freservationdate	= rsget("reservationdate")
		FOneItem.Fstate					= rsget("state")
		FOneItem.Fviewno				= rsget("viewno")
		FOneItem.Fworktext				= rsget("worktext")
		FOneItem.FpartMKid				= rsget("partMKid")
		FOneItem.FpartWDid			= rsget("partWDid")

	end if
	rsget.Close
	end Sub

	public function fnGetGroundMainList()
        dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and gidx = '"&FRectIdx&"'"
		end If

		if FRectNo <> "" then
			sqlsearch = sqlsearch & " and viewno = '"&FRectNo&"'"
		end if

		if FRecttitle <> "" then
			sqlsearch = sqlsearch & " and viewtitle like '%"&FRecttitle&"%'"
		end if

		If FRectstate <> "" THEN
			IF FRectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and state = 7 and convert(varchar(10),getdate(),120) <= reservationdate"
			ELSE
				sqlsearch  = sqlsearch & " and state = " &FRectstate & ""
			END IF
		End If


		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_ground_main"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " gidx , viewno , viewtitle , titleimg , reservationdate , state "
		sqlStr = sqlStr & " , (SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = partMKid and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0)) and p.userid <> '') as partMKname "
		sqlStr = sqlStr & " , (SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = partWDid and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0)) and p.userid <> '') as partWDname "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_ground_main "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by viewNo DESC, reservationdate desc"

		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new CPlayContentsItem

                FItemList(i).Fidx						= rsget("gidx")
				FItemList(i).Fviewno					= rsget("viewno")
                FItemList(i).Fviewtitle					= rsget("viewtitle")
				FItemList(i).Flistimg					= rsget("titleimg")
                FItemList(i).Freservationdate			= rsget("reservationdate")
                FItemList(i).Fstate						= rsget("state")
                FItemList(i).FpartWDname				= rsget("partWDname")
                FItemList(i).FpartMKname				= rsget("partMKname")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	'play ground Sub 2013-09-23 이종화
	public Sub GetRowGroundSub()
	dim sqlStr
	sqlStr = "select * "
	sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_play_ground_sub"
	sqlStr = sqlStr + " where gcidx=" + CStr(FRectIdx) + " and gidx=" + CStr(FRectgIdx)
	'response.write sqlStr
	rsget.Open SqlStr, dbget, 1
	FResultCount = rsget.RecordCount

	set FOneItem = new CPlayContentsItem

	if Not rsget.Eof then

		FOneItem.Fidx					= rsget("gcidx")
		FOneItem.Fidxsub				= rsget("gidx")
		FOneItem.Fviewno				= rsget("viewno")
		FOneItem.Fviewtitle				= rsget("viewtitle")
		FOneItem.Fstate					= rsget("state")
		FOneItem.Freservationdate		= rsget("reservationdate")
		FOneItem.FpartMKid				= rsget("partMKid")
		FOneItem.FpartWDid				= rsget("partWDid")
		FOneItem.Fworktext				= rsget("worktext")

		FOneItem.Fplaymainimg			= rsget("playmainimg")
		FOneItem.Fbeforeimg				= rsget("viewthumbimg1")
		FOneItem.Fafterimg				= rsget("viewthumbimg2")
		FOneItem.Ftopbgimg				= rsget("viewbgimg")
		FOneItem.Fsideltimg				= rsget("downsideimg1")

		FOneItem.FsubBGColor			= rsget("downbgcolor")
		FOneItem.FmainTopBGColor		= rsget("mainbgcolor")
		FOneItem.Fviewcontents			= rsget("viewcontents")

		FOneItem.Fmyplayimg				= rsget("myplayimg")
		FOneItem.Fmo_contents			= rsget("mo_contents")
		FOneItem.Fmo_exec_check			= rsget("mo_exec_check")
		FOneItem.Fexec_check			= rsget("exec_check")
		FOneItem.Fexec_filepath			= rsget("exec_filepath")

	end if
	rsget.Close
	end Sub

	public function fnGetGroundSubList()
        dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and gidx = '"& FRectIdx &"'"
		end If

		If FRectstate <> "" THEN
			IF FRectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and state = 7 and convert(varchar(10),getdate(),120) <= reservationdate"
			ELSE
				sqlsearch  = sqlsearch & " and state = " & FRectstate & ""
			END IF
		End If

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_ground_sub"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " gcidx , gidx , viewno , state , viewthumbimg1 , viewthumbimg2 , viewtitle , reservationdate "
		sqlStr = sqlStr & " , isnull((select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = gidx and t.playidxsub = gcidx  ),0) as tagcnt "
		sqlStr = sqlStr & " , (SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = partMKid and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0)) and p.userid <> '') as partMKname "
		sqlStr = sqlStr & " , (SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = partWDid and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0)) and p.userid <> '') as partWDname "
		sqlStr = sqlStr & " , isnull((select count(itemid) from db_sitemaster.dbo.tbl_play_ground_item as I where I.subidx = gcidx ),0) as itemcnt "
		sqlStr = sqlStr & " , (select top 1 idx from [db_sitemaster].[dbo].[tbl_play_mo] where type = 1 and contents_idx = gcidx) as mo_idx "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_ground_sub "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by viewNo DESC, reservationdate desc"

		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new CPlayContentsItem

				FItemList(i).Fidxsub					= rsget("gcidx")
                FItemList(i).Fidx						= rsget("gidx")
				FItemList(i).Fviewno					= rsget("viewno")
                FItemList(i).Fviewtitle					= rsget("viewtitle")
                FItemList(i).Freservationdate			= rsget("reservationdate")
                FItemList(i).Fstate						= rsget("state")
                FItemList(i).Ftagcnt					= rsget("tagcnt")
                FItemList(i).FpartMKname				= rsget("partMKname")
                FItemList(i).FpartWDname				= rsget("partWDname")
				FItemList(i).Fviewthumbimg1				= rsget("viewthumbimg1")
				FItemList(i).Fviewthumbimg2				= rsget("viewthumbimg2")
				FItemList(i).Fitemcnt					= rsget("itemcnt")
				FItemList(i).Fmo_idx					= rsget("mo_idx")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	'play picture diary 2013-09-03 이종화
	public Sub GetOneRowContent()
	dim sqlStr
	sqlStr = "select * "
	sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_play_picture_diary"
	sqlStr = sqlStr + " where pdidx=" + CStr(FRectIdx)
	rsget.Open SqlStr, dbget, 1
	FResultCount = rsget.RecordCount

	set FOneItem = new CPlayContentsItem

	if Not rsget.Eof then

		FOneItem.Fidx						= rsget("pdidx")
		FOneItem.Flistimg					= rsget("listimg")
		FOneItem.Fviewimg				= rsget("viewimg")
		FOneItem.Fviewtitle				= rsget("viewtitle")
		FOneItem.Fviewtext				= rsget("viewtext")
		FOneItem.Freservationdate	= rsget("reservationdate")
		FOneItem.Fstate					= rsget("state")
		FOneItem.Fviewno				= rsget("viewno")
		FOneItem.Forgimg				= rsget("orgimg")
		FOneItem.Fworktext				= rsget("worktext")

	end if
	rsget.Close
	end Sub

	'Play Tag
	public function GetRowTagContent()
		dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and playidx="& FRectIdx &""
		end If

		if FRectsubidx <> "" then
			sqlsearch = sqlsearch & " and playidxsub="& FRectsubidx &""
		end If
		
		if FRectPlaycate <> "" then
			sqlsearch = sqlsearch & " and playcate='"& FRectPlaycate &"'"
		end if

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_tag"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit function

		'// 본문 내용 접수
		sqlStr = "select "
		sqlStr = sqlStr & " tagname , tagurl , tagurl_mo , tagurl_appchk , tagurl_appurl "
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_tag"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by tagidx asc "

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
		redim preserve FItemList(FTotalCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CPlayContentsItem

				FItemList(i).Ftagname        = rsget("tagname")
				FItemList(i).Ftagurl         = rsget("tagurl")
				FItemList(i).Ftagurl2        = rsget("tagurl_mo")
				FItemList(i).Ftagurl3        = rsget("tagurl_appchk")
				FItemList(i).Ftagurl4        = rsget("tagurl_appurl")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
    end Function

	public function fnGetPictureDiaryList()
        dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and pdidx = '"&FRectIdx&"'"
		end If

		if FRectNo <> "" then
			sqlsearch = sqlsearch & " and viewno = '"&FRectNo&"'"
		end if

		if FRecttitle <> "" then
			sqlsearch = sqlsearch & " and viewtitle like '%"&FRecttitle&"%'"
		end if

		If FRectstate <> "" THEN
			IF FRectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and state = 7 and convert(varchar(10),getdate(),120) <= reservationdate"
			ELSE
				sqlsearch  = sqlsearch & " and  state = "&FRectstate & ""
			END IF
		End If

		If FRectTag <> "" Then
			If FRectTag = "Y" then
				sqlsearch  = sqlsearch & " and (select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = pdidx  ) > 0"
			Else
				sqlsearch  = sqlsearch & " and (select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = pdidx  ) = 0"
			End If
		End If

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_picture_diary"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " pdidx , viewtitle , listimg , reservationdate , state , viewno"
		sqlStr = sqlStr & " , isnull((select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = pdidx  ),0) as tagcnt "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_picture_diary"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by viewNo DESC, reservationdate desc"

		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new CPlayContentsItem

                FItemList(i).Fidx						= rsget("pdidx")
                FItemList(i).Fviewtitle				= rsget("viewtitle")
                FItemList(i).Flistimg					= rsget("listimg")
                FItemList(i).Freservationdate		= rsget("reservationdate")
                FItemList(i).Fstate						= rsget("state")
                FItemList(i).Fviewno					= rsget("viewno")
                FItemList(i).Ftagcnt					= rsget("tagcnt")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	'play style+ 2013-09-05 이종화
	'//admin/sitemaster/play/styleplus/popstyleplusEdit.asp
	public Sub GetOneRowStyleContent()
		dim sqlStr
	
		sqlStr = "select * "
		sqlStr = sqlStr + ", isnull((select count(itemid) from db_sitemaster.dbo.tbl_play_style_item as I where I.styleidx = "& CStr(FRectIdx) &" and viewidx = 1 ),0) as itemcnt1 "
		sqlStr = sqlStr + ", isnull((select count(itemid) from db_sitemaster.dbo.tbl_play_style_item as I where I.styleidx = "& CStr(FRectIdx) &" and viewidx = 2 ),0) as itemcnt2 "
		sqlStr = sqlStr + ", isnull((select count(itemid) from db_sitemaster.dbo.tbl_play_style_item as I where I.styleidx = "& CStr(FRectIdx) &" and viewidx = 3 ),0) as itemcnt3 "
		sqlStr = sqlStr + ", isnull((select count(itemid) from db_sitemaster.dbo.tbl_play_style_item as I where I.styleidx = "& CStr(FRectIdx) &" and viewidx = 4 ),0) as itemcnt4 "
		sqlStr = sqlStr + ", isnull((select count(itemid) from db_sitemaster.dbo.tbl_play_style_item as I where I.styleidx = "& CStr(FRectIdx) &" and viewidx = 5 ),0) as itemcnt5 "
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_play_style_list"
		sqlStr = sqlStr + " where styleidx=" + CStr(FRectIdx)
	
		'response.write sqlStr & "<Br>"
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
	
		set FOneItem = new CPlayContentsItem
	
		if Not rsget.Eof then
			FOneItem.fstyle_html_m		= db2html(rsget("style_html_m"))
			FOneItem.Fidx				= rsget("styleidx")
			FOneItem.Flistimg			= rsget("listimg")
			FOneItem.Fviewimg1			= rsget("viewimg1")
			FOneItem.Fviewimg2			= rsget("viewimg2")
			FOneItem.Fviewimg3			= rsget("viewimg3")
			FOneItem.Fviewimg4			= rsget("viewimg4")
			FOneItem.Fviewimg5			= rsget("viewimg5")
			FOneItem.Ftextimg			= rsget("textimg")
			FOneItem.Fviewtitle			= rsget("viewtitle")
			FOneItem.Freservationdate	= rsget("reservationdate")
			FOneItem.Fstate				= rsget("state")
			FOneItem.Fviewno			= rsget("viewno")
			FOneItem.Fworktext			= rsget("worktext")
			FOneItem.FpartMDid			= rsget("partMDid")
			FOneItem.FpartWDid			= rsget("partWDid")
			FOneItem.Fitemcnt1			= rsget("itemcnt1")
			FOneItem.Fitemcnt2			= rsget("itemcnt2")
			FOneItem.Fitemcnt3			= rsget("itemcnt3")
			FOneItem.Fitemcnt4			= rsget("itemcnt4")
			FOneItem.Fitemcnt5			= rsget("itemcnt5")
		end if
		rsget.Close
	end Sub

	public function fnGetStylePlusList()
        dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and styleidx = '"&FRectIdx&"'"
		end If

		if FRectNo <> "" then
			sqlsearch = sqlsearch & " and viewno = '"&FRectNo&"'"
		end if

		if FRecttitle <> "" then
			sqlsearch = sqlsearch & " and viewtitle like '%"&FRecttitle&"%'"
		end if

		If FRectstate <> "" THEN
			IF FRectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and state = 7 and convert(varchar(10),getdate(),120) <= reservationdate"
			ELSE
				sqlsearch  = sqlsearch & " and state = " &FRectstate & ""
			END IF
		End If

		If FRectTag <> "" Then
			If FRectTag = "Y" then
				sqlsearch  = sqlsearch & " and (select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = styleidx  ) > 0"
			Else
				sqlsearch  = sqlsearch & " and (select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = styleidx  ) = 0"
			End If
		End If

		If FRectpartMDid <> "" Then
			sqlsearch = sqlsearch & " and partMDid = '"& FRectpartMDid  &"'"
		End If

		If FRectpartWDid <> "" Then
			sqlsearch = sqlsearch & " and partwDid = '"& FRectpartWDid  &"'"
		End If

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_style_list"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " styleidx , viewno , listimg , viewtitle , reservationdate , state "
		sqlStr = sqlStr & " , isnull((select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = styleidx  ),0) as tagcnt "
		sqlStr = sqlStr & " , (SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = partMDid and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0)) and p.userid <> '') as partMDname "
		sqlStr = sqlStr & " , (SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = partWDid and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0)) and p.userid <> '') as partWDname "
		sqlStr = sqlStr & " ,  isnull((select count(itemid) from db_sitemaster.dbo.tbl_play_style_item as S where S.styleidx = l.styleidx),0) as itemcnt "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_style_list as l "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by viewNo DESC, reservationdate desc"

		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new CPlayContentsItem

                FItemList(i).Fidx						= rsget("styleidx")
				FItemList(i).Fviewno					= rsget("viewno")
                FItemList(i).Flistimg					= rsget("listimg")
                FItemList(i).Fviewtitle				= rsget("viewtitle")
                FItemList(i).Freservationdate		= rsget("reservationdate")
                FItemList(i).Fstate						= rsget("state")
                FItemList(i).Ftagcnt					= rsget("tagcnt")
                FItemList(i).FpartWDname			= rsget("partWDname")
                FItemList(i).FpartMDname			= rsget("partMDname")
                FItemList(i).Fitemcnt					= rsget("itemcnt")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	'play VideoClip 2013-09-13 이종화
	public Sub GetOneRowVideoClipContent()
	dim sqlStr
	sqlStr = "select * "
	sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_play_video_clip"
	sqlStr = sqlStr + " where vidx=" + CStr(FRectIdx)
	rsget.Open SqlStr, dbget, 1
	FResultCount = rsget.RecordCount

	set FOneItem = new CPlayContentsItem

	if Not rsget.Eof then

		FOneItem.Fidx						= rsget("vidx")
		FOneItem.Flistimg					= rsget("listimg")
		FOneItem.Fviewtitle				= rsget("viewtitle")
		FOneItem.Fviewtext				= rsget("viewtext")
		FOneItem.Freservationdate	= rsget("reservationdate")
		FOneItem.Fstate					= rsget("state")
		FOneItem.Fviewno				= rsget("viewno")
		FOneItem.Fworktext				= rsget("worktext")
		FOneItem.Fvideourl				= rsget("videourl")
		FOneItem.Fpartwdid				= rsget("partwdid")
		FOneItem.FvideourlM				= rsget("videourlM")

	end if
	rsget.Close
	end Sub

	public function fnGetVideoClipList()
        dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and vidx = '"&FRectIdx&"'"
		end If

		if FRectNo <> "" then
			sqlsearch = sqlsearch & " and viewno = '"&FRectNo&"'"
		end if

		if FRecttitle <> "" then
			sqlsearch = sqlsearch & " and viewtitle like '%"&FRecttitle&"%'"
		end if

		If FRectstate <> "" THEN
			IF FRectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and state = 7 and convert(varchar(10),getdate(),120) <= reservationdate"
			ELSE
				sqlsearch  = sqlsearch & " and state = " &FRectstate & ""
			END IF
		End If

		If FRectTag <> "" Then
			If FRectTag = "Y" then
				sqlsearch  = sqlsearch & " and (select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = vidx  ) > 0"
			Else
				sqlsearch  = sqlsearch & " and (select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = vidx  ) = 0"
			End If
		End If

		If FRectpartWDid <> "" Then
			sqlsearch = sqlsearch & " and partwDid = '"& FRectpartWDid  &"'"
		End If

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_video_clip"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " vidx , viewno , listimg , viewtitle , reservationdate , state , videourl"
		sqlStr = sqlStr & " , isnull((select count(*) from db_sitemaster.dbo.tbl_play_tag as t where t.playcate = "& FRPlaycate &" and t.playidx = vidx  ),0) as tagcnt "
		sqlStr = sqlStr & " , (SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = partWDid and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0)) and p.userid <> '') as partWDname "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_video_clip "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by viewNo DESC, reservationdate desc"

		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new CPlayContentsItem

                FItemList(i).Fidx						= rsget("vidx")
				FItemList(i).Fviewno					= rsget("viewno")
                FItemList(i).Flistimg					= rsget("listimg")
                FItemList(i).Fviewtitle				= rsget("viewtitle")
                FItemList(i).Freservationdate		= rsget("reservationdate")
                FItemList(i).Fstate						= rsget("state")
                FItemList(i).Ftagcnt					= rsget("tagcnt")
                FItemList(i).FpartWDname			= rsget("partWDname")
                FItemList(i).Fvideourl					= rsget("videourl")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	Public Sub sbGetphotopickList()
        dim sqlStr, sqlsearch, i

		If FRectIsusing <> "" then
			sqlsearch = sqlsearch & " and isusing = '"& FRectIsusing &"'"
		End If

		If FRectViewTitle <> "" THEN
			sqlsearch  = sqlsearch & " and viewtitle like '%"&FRectViewTitle&"%'"
		End If

		sqlStr = "SELECT count(*) as cnt"
        sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_play_photo_pick with (nolock)"
        sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
        rsget.Close

		If FTotalCount < 1 Then Exit Sub

        sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " idx, viewtitle, subtitle, isusing, PPimg, regdate "
        sqlStr = sqlStr & " ,isnull((SELECT count(*) FROM db_sitemaster.dbo.tbl_play_tag as t with (nolock) where t.playcate = "& FRPlaycate &" and t.playidx = idx),0) as tagcnt "
		sqlStr = sqlStr & " , isnull((select count(itemid) from db_sitemaster.dbo.tbl_play_photopick_item as I with (nolock) where I.subidx = A.idx ),0) as itemcnt "
        sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_play_photo_pick A with (nolock)"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " ORDER BY idx DESC"

        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        FTotalPage =  CInt(FTotalCount \ FPageSize)
		If (FTotalCount \ FPageSize) <> (FTotalCount/FPageSize) Then
			FTotalPage = FtotalPage + 1
		End If
		FResultCount = rsget.RecordCount - (FPageSize * (FCurrPage - 1))
        If (FResultCount < 1) then FResultCount = 0
        Redim preserve FItemList(FResultCount)

        i=0
        If not rsget.EOF Then
            rsget.absolutepage = FCurrPage
            Do until rsget.EOF
                Set FItemList(i) = new CPlayContentsItem
					FItemList(i).FIdx			= rsget("idx")
	                FItemList(i).FViewtitle		= rsget("viewtitle")
	                FItemList(i).FSubtitle		= rsget("subtitle")
	                FItemList(i).FIsusing		= rsget("isusing")
	                FItemList(i).FPPimg			= rsget("PPimg")
	                FItemList(i).FRegdate		= rsget("regdate")
	                FItemList(i).FTagCnt		= rsget("tagcnt")
	                FItemList(i).FitemCnt	= rsget("itemcnt")
                rsget.movenext
                i = i + 1
            Loop
        End If
        rsget.Close
    End Sub

	Public Sub GetPhotoPickOne()
		Dim sqlStr

		sqlStr = "SELECT TOP 1 idx, viewtitle, subtitle, isusing, PPimg, regdate , style_html_m "
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_play_photo_pick with (nolock)"
		sqlStr = sqlStr & " where idx=" & CStr(FRectIdx)

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		Set FOneItem = new CPlayContentsItem
		If Not rsget.EOF Then
			FOneItem.Fidx		= rsget("idx")
			FOneItem.FViewtitle	= rsget("viewtitle")
			FOneItem.FSubtitle	= rsget("subtitle")
			FOneItem.FIsusing	= rsget("isusing")
			FOneItem.FPPimg		= rsget("PPimg")
			FOneItem.FRegdate	= rsget("regdate")
			FOneItem.fstyle_html_m		= db2html(rsget("style_html_m"))
		End If
		rsget.Close
	End Sub

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


'//메인페이지 , 이벤트 공통함수		'/오픈예정 노출함 , 검색페이지용
function Draweventstate2(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value="" <%if selectedId="" then response.write " selected"%>>선택</option>
		<option value="0" <% if selectedId="0" then response.write "selected" %>>등록대기</option>
		<option value="3" <% if selectedId="3" then response.write "selected" %>>이미지등록요청</option>
		<option value="5" <% if selectedId="5" then response.write "selected" %>>오픈요청</option>
		<option value="6" <% if selectedId="6" then response.write "selected" %>>오픈예정</option>
		<option value="7" <% if selectedId="7" then response.write "selected" %>>오픈</option>
		<option value="9" <% if selectedId="9" then response.write "selected" %>>종료</option>
	</select>
<%
end Function

'//메인페이지 , 이벤트 모두 공통
function geteventstate(v)
	if v = "0" then
		geteventstate = "등록대기"
	elseif v = "3" then
		geteventstate = "이미지등록요청"
	elseif v = "5" then
		geteventstate = "오픈요청"
	elseif v = "6" then
		geteventstate = "오픈예정"
	elseif v = "7" then
		geteventstate = "오픈"
	elseif v = "9" then
		geteventstate = "종료"
	end if
end Function

'/담당MD 리스트가져오기 (팀장 미만,직원 이상)
Sub sbGetpartid(ByVal selName, ByVal sIDValue, ByVal sScript,part_sn)
	Dim strSql, arrList, intLoop

	if part_sn = "" then exit sub

	strSql = " SELECT userid, username"
	strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "
	strSql = strSql & " WHERE part_sn IN("&part_sn&") and  posit_sn>='4' and  posit_sn<='12' and   isUsing=1" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	strSql = strSql & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	strSql = strSql & " and userid <> ''" & vbcrlf
	strSql = strSql & " order by posit_sn, empno" & vbcrlf

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
%>