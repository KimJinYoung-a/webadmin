<%
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

	'### 공통 리스트이미지(직사각형) gubun = 1
	Call fnImageSave(vDidx, vCate, "1", vJikListImgURL, "", "", "0", "")
	
	'### 공통 리스트이미지(정사각형) gubun = 11
	Call fnImageSave(vDidx, vCate, "11", vJungListImgURL, "", "", "0", "")
	
	'### 공통 검색리스트배너이미지 gubun = 28
	Call fnImageSave(vDidx, vCate, "28", vSearchListImg, "", "", "0", "")
	
	If vCate = "1" Then '### playlist
		
		'### playlist : cate = 1
		Dim vCate1VideoURL, vCate1ImageURL, vCate1Type, vCate1Directer, vCate1PCLinkBanImg, vCate1PCLinkBanURL, vCate1MoLinkBanImg, vCate1MoLinkBanURL
		Dim vCate1CommTitle, vCate1Comment1, vCate1Comment2, vCate1Comment3, vCate1precomm1, vCate1precomm2, vCate1precomm3, vCate1VideoOrigin, vCate1RewardCopy
		vCate1Directer		= requestCheckVar(Request("cate1directer"),150)
		vCate1Type			= requestCheckVar(Request("cate1type"),1)
		vCate1VideoURL		= requestCheckVar(Request("cate1videourl"),190)
		vCate1ImageURL		= requestCheckVar(Request("cate1imageurl"),100)
		vCate1PCLinkBanImg	= requestCheckVar(Request("cate1pclinkbanimg"),100)
		vCate1PCLinkBanURL	= requestCheckVar(Request("cate1pclinkbanurl"),200)
		vCate1MoLinkBanImg	= requestCheckVar(Request("cate1molinkbanimg"),100)
		vCate1MoLinkBanURL	= requestCheckVar(Request("cate1molinkbanurl"),200)
		vCate1CommTitle		= requestCheckVar(Request("cate1commtitle"),150)
		vCate1Comment1		= requestCheckVar(Request("cate1comment1"),50)
		vCate1Comment2		= requestCheckVar(Request("cate1comment2"),50)
		vCate1Comment3		= requestCheckVar(Request("cate1comment3"),50)
		vCate1precomm1		= requestCheckVar(Request("cate1precomm1"),50)
		vCate1precomm2		= requestCheckVar(Request("cate1precomm2"),50)
		vCate1precomm3		= requestCheckVar(Request("cate1precomm3"),50)
		vCate1VideoOrigin	= requestCheckVar(Request("cate1videoorigin"),50)
		vCate1RewardCopy		= requestCheckVar(Request("cate1rewardcopy"),600)
		
		'### [db_giftplus].[dbo].[tbl_play_playlist] 저장
		vQuery = "IF EXISTS(select idx from [db_giftplus].[dbo].[tbl_play_playlist] where didx = '" & vDidx & "') "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	UPDATE [db_giftplus].[dbo].[tbl_play_playlist] SET "
		vQuery = vQuery & "		directer = '" & vCate1Directer & "' "
		vQuery = vQuery & "		,type = '" & vCate1Type & "' "
		vQuery = vQuery & "		,videourl = '" & vCate1VideoURL & "' "
		vQuery = vQuery & "		,comm_title = '" & vCate1CommTitle & "' "
		vQuery = vQuery & "		,comment1 = '" & vCate1Comment1 & "' "
		vQuery = vQuery & "		,comment2 = '" & vCate1Comment2 & "' "
		vQuery = vQuery & "		,comment3 = '" & vCate1Comment3 & "' "
		vQuery = vQuery & "		,precomm1 = '" & vCate1precomm1 & "' "
		vQuery = vQuery & "		,precomm2 = '" & vCate1precomm2 & "' "
		vQuery = vQuery & "		,precomm3 = '" & vCate1precomm3 & "' "
		vQuery = vQuery & "		,videoorigin = '" & vCate1VideoOrigin & "' "
		vQuery = vQuery & "		,rewardcopy = '" & vCate1RewardCopy & "' "
		vQuery = vQuery & "	where didx = '" & vDidx & "'"
		vQuery = vQuery & "END "
		vQuery = vQuery & "ELSE "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	INSERT INTO [db_giftplus].[dbo].[tbl_play_playlist](didx, directer, type, videourl, comm_title, comment1, comment2, comment3, precomm1, precomm2, precomm3, videoorigin, rewardcopy) VALUES("
		vQuery = vQuery & "'" & vDidx & "', '" & vCate1Directer & "', '" & vCate1Type & "', '" & vCate1VideoURL & "', '" & vCate1CommTitle & "' "
		vQuery = vQuery & ", '" & vCate1Comment1 & "','" & vCate1Comment2 & "', '" & vCate1Comment3 & "', '" & vCate1precomm1 & "','" & vCate1precomm2 & "', '" & vCate1precomm3 & "', '" & vCate1VideoOrigin & "', '" & vCate1RewardCopy & "')"
		vQuery = vQuery & "END "
		dbget.Execute vQuery
		
		'### 대표컨텐츠 이미지 gubun = 2
		Call fnImageSave(vDidx, vCate, "2", vCate1ImageURL, "", "", "0", "")
		
		'### playlist 연결배너 PC 이미지 gubun = 3
		Call fnImageSave(vDidx, vCate, "3", vCate1PCLinkBanImg, vCate1PCLinkBanURL, "", "0", "")
		
		'### playlist 연결배너 Mo 이미지 gubun = 18
		Call fnImageSave(vDidx, vCate, "18", vCate1MoLinkBanImg, vCate1MoLinkBanURL, "", "0", "")

	ElseIf vCate = "21" Then '### inspiration design
		
		'### inspiration design : cate = 21
		Dim vCate21ImageURL(5), vCate21Item
		vCate21Item			= Trim(requestCheckVar(Request("cate21item"),1000))


		'### inspiration design 컨텐츠 이미지 gubun = 4
		For l=1 To 5
			vCate21ImageURL(l) = Trim(requestCheckVar(Request("cate21img"&l&""),100))
		Next
		
		For l=1 To 5
			If vCate21ImageURL(l) <> "" Then
				Call fnImageSave(vDidx, vCate, "4", vCate21ImageURL(l), "", "", l, 0)
			End If
		Next
		
		If vCate21Item <> "" Then
			vQuery = "DELETE [db_giftplus].[dbo].[tbl_play_item] WHERE didx = '" & vDidx & "' "
			For i = LBound(Split(vCate21Item,",")) To UBound(Split(vCate21Item,","))
				vQuery = vQuery & "INSERT INTO [db_giftplus].[dbo].[tbl_play_item](didx, itemid) VALUES('" & vDidx & "', '" & Split(vCate21Item,",")(i) & "') "
			Next
			dbget.Execute vQuery
		End If
	ElseIf vCate = "22" Then '### inspiration style
		
		'### inspiration style : cate = 22
		Dim vCate22ImageURL(5), vCate22Item
		vCate22Item			= Trim(requestCheckVar(Request("cate22item"),1000))
		
		'### inspiration style 컨텐츠 이미지 gubun = 5
		For l=1 To 5
			vCate22ImageURL(l) = Trim(requestCheckVar(Request("cate22img"&l&""),100))
		Next
		
		For l=1 To 5
			If vCate22ImageURL(l) <> "" Then
				Call fnImageSave(vDidx, vCate, "5", vCate22ImageURL(l), "", "", l, 0)
			End If
		Next
		
		If vCate22Item <> "" Then
			vQuery = "DELETE [db_giftplus].[dbo].[tbl_play_item] WHERE didx = '" & vDidx & "' "
			For i = LBound(Split(vCate22Item,",")) To UBound(Split(vCate22Item,","))
				vQuery = vQuery & "INSERT INTO [db_giftplus].[dbo].[tbl_play_item](didx, itemid) VALUES('" & vDidx & "', '" & Split(vCate22Item,",")(i) & "') "
			Next
			dbget.Execute vQuery
		End If
	ElseIf vCate = "3" Then '### azit
		
		'### azit : cate = 3
		Dim vCate3PCImageURL, vCate3MoImageURL, vCate3Notice, vCate3Ptitle(4), vCate3Pjuso(4), vCate3Plink(4), vCate3PImg(20), vCate3PCopy(20), p
		Dim vCate3EntryCont, vCate3EntrySDate, vCate3EntryEDate, vCate3AnnounDate, vCate3EntryMethod

		tmp = 0
		vCate3PCImageURL		= requestCheckVar(Request("cate3pcimg"),100)
		vCate3MoImageURL		= requestCheckVar(Request("cate3moimg"),100)
		vCate3EntryCont		= html2db(Request("cate3entrycont"))
		vCate3EntrySDate 	= requestCheckVar(Request("cate3entrysdate"),10)
		vCate3EntryEDate 	= requestCheckVar(Request("cate3entryedate"),10)
		vCate3AnnounDate 	= requestCheckVar(Request("cate3announdate"),10)
		vCate3Notice			= html2db(Request("cate3notice"))
		vCate3EntryMethod	= requestCheckVar(Request("entry_method"),1)
		
	
		For p=1 To 4
			vCate3Ptitle(p)	= Trim(requestCheckVar(Request("cate3P"&p&"title"),150))
			vCate3Pjuso(p)	= Trim(requestCheckVar(Request("cate3P"&p&"juso"),150))
			vCate3Plink(p)	= Trim(requestCheckVar(Request("cate3P"&p&"link"),150))
			
			For l=1 To 5
				tmp = tmp + 1
				vCate3PImg(tmp)	= Trim(requestCheckVar(Request("cate3P"&p&"Img"&l&""),100))
				vCate3PCopy(tmp)	= html2db(Trim(Request("cate3P"&p&"copy"&l&"")))
			Next
		Next
		
		'### azit PC 컨텐츠 이미지 gubun = 19
		Call fnImageSave(vDidx, vCate, "19", vCate3PCImageURL, "", "", "0", "")
		
		'### azit Mo 컨텐츠 이미지 gubun = 6
		Call fnImageSave(vDidx, vCate, "6", vCate3MoImageURL, "", "", "0", "")
		
		tmp = 0
		For p=1 To 4
			If vCate3Ptitle(p) <> "" Then	'### 비어있으면 저장 안함.
				'### [db_giftplus].[dbo].[tbl_play_azit] 저장
				vQuery = "IF EXISTS(select idx from [db_giftplus].[dbo].[tbl_play_azit] where didx = '" & vDidx & "' and groupnum = '" & p & "') "
				vQuery = vQuery & "BEGIN "
				vQuery = vQuery & "	UPDATE [db_giftplus].[dbo].[tbl_play_azit] SET "
				vQuery = vQuery & "		title = '" & vCate3Ptitle(p) & "' "
				vQuery = vQuery & "		,address = '" & vCate3Pjuso(p) & "' "
				vQuery = vQuery & "		,addrlink = '" & vCate3Plink(p) & "' "
				vQuery = vQuery & "	where didx = '" & vDidx & "' and groupnum = '" & p & "'"
				vQuery = vQuery & "END "
				vQuery = vQuery & "ELSE "
				vQuery = vQuery & "BEGIN "
				vQuery = vQuery & "	INSERT INTO [db_giftplus].[dbo].[tbl_play_azit](didx, groupnum, title, address, addrlink) VALUES("
				vQuery = vQuery & "'" & vDidx & "', '" & p & "', '" & vCate3Ptitle(p) & "', '" & vCate3Pjuso(p) & "', '" & vCate3Plink(p) & "')"
				vQuery = vQuery & "END "
				dbget.Execute vQuery
			End If
			
			'### azit 장소 이미지 gubun = 7
			For l=1 To 5
				tmp = tmp + 1
				If vCate3PImg(tmp) <> "" Then
					Call fnImageSave(vDidx, vCate, "7", vCate3PImg(tmp), "", vCate3PCopy(tmp), l, p)
				End If
			Next
		Next
		
		'### [db_giftplus].[dbo].[tbl_play_azit_entry] 저장
		vQuery = "IF EXISTS(select didx from [db_giftplus].[dbo].[tbl_play_azit_entry] where didx = '" & vDidx & "') "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	UPDATE [db_giftplus].[dbo].[tbl_play_azit_entry] SET "
		vQuery = vQuery & "		entry_content = '" & vCate3EntryCont & "' "
		vQuery = vQuery & "		,entry_sdate = '" & vCate3EntrySDate & "' "
		vQuery = vQuery & "		,entry_edate = '" & vCate3EntryEDate & "' "
		vQuery = vQuery & "		,announce_date = '" & vCate3AnnounDate & "' "
		vQuery = vQuery & "		,notice = '" & vCate3Notice & "' "
		vQuery = vQuery & "		,entry_method = '" & vCate3EntryMethod & "' "
		vQuery = vQuery & "	where didx = '" & vDidx & "' "
		vQuery = vQuery & "END "
		vQuery = vQuery & "ELSE "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	INSERT INTO [db_giftplus].[dbo].[tbl_play_azit_entry](didx, entry_content, entry_sdate, entry_edate, announce_date, notice, entry_method) VALUES("
		vQuery = vQuery & "'" & vDidx & "', '" & vCate3EntryCont & "', '" & vCate3EntrySDate & "', '" & vCate3EntryEDate & "', '" & vCate3AnnounDate & "', '" & vCate3Notice & "', '" & vCate3EntryMethod & "')"
		vQuery = vQuery & "END "
		dbget.Execute vQuery

	'// 2017.06.01 원승현 azit comma 스타일 추가
	ElseIf vCate = "31" Then '### Azit COMMA
		Dim vCate31PCTopImageURL, vCate31MoTopImageURL, vCate31Directer, vCate31Img(5), vCate31Copy(5), vCate31PCLinkBanImg, vCate31PCLinkBanURL, vCate31MoLinkBanImg, vCate31MoLinkBanURL
		tmp = 0
		vCate31PCTopImageURL		= requestCheckVar(Request("cate31pctopimg"),100)
		vCate31MoTopImageURL		= requestCheckVar(Request("cate31motopimg"),100)
		vCate31Directer			= requestCheckVar(Request("cate31directer"),150)
		vCate31PCLinkBanImg		= requestCheckVar(Request("cate31pclinkbanimg"),100)
		vCate31PCLinkBanURL		= requestCheckVar(Request("cate31pclinkbanurl"),200)
		vCate31MoLinkBanImg		= requestCheckVar(Request("cate31molinkbanimg"),100)
		vCate31MoLinkBanURL		= requestCheckVar(Request("cate31molinkbanurl"),200)
		
		'### comma 컨텐츠 PC 상단 이미지 gubun = 23
		Call fnImageSave(vDidx, vCate, "23", vCate31PCTopImageURL, "", "", "0", "")
		
		'### comma 컨텐츠 Mo 상단 이미지 gubun = 24
		Call fnImageSave(vDidx, vCate, "24", vCate31MoTopImageURL, "", "", "0", "")
		
		'### [db_giftplus].[dbo].[tbl_play_azit_comma] 저장
		vQuery = "IF EXISTS(select didx from [db_giftplus].[dbo].[tbl_play_azit_comma] where didx = '" & vDidx & "') "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	UPDATE [db_giftplus].[dbo].[tbl_play_comma] SET directer = '" & vCate31Directer & "' where didx = '" & vDidx & "' "
		vQuery = vQuery & "END "
		vQuery = vQuery & "ELSE "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	INSERT INTO [db_giftplus].[dbo].[tbl_play_azit_comma](didx, directer) VALUES('" & vDidx & "', '" & vCate31Directer & "')"
		vQuery = vQuery & "END "
		dbget.Execute vQuery
		
		'### comma 연결배너 PC 이미지 gubun = 26
		Call fnImageSave(vDidx, vCate, "26", vCate31PCLinkBanImg, vCate31PCLinkBanURL, "", "0", "")
		
		'### comma 연결배너 Mo 이미지 gubun = 27
		Call fnImageSave(vDidx, vCate, "27", vCate31MoLinkBanImg, vCate31MoLinkBanURL, "", "0", "")
		
		'### comma 컨텐츠(에디터) 이미지 gubun = 25
		For l=1 To 5
			tmp = tmp + 1
			vCate31Img(tmp)	= Trim(requestCheckVar(Request("cate31Img"&l&""),100))
			vCate31Copy(tmp)	= html2db(Trim(Request("cate31copy"&l&"")))
		Next
		
		tmp = 0
		For l=1 To 5
			tmp = tmp + 1
			Call fnImageSave(vDidx, vCate, "25", vCate31Img(tmp), "", vCate31Copy(tmp), l, 0)
		Next

	ElseIf vCate = "42" Then '### THING thingthing
		
		'### THING thingthing : cate = 42
		Dim vCate42Img(3), vCate42EntrySDate, vCate42EntryEDate, vCate42AnnounDate, vCate42WinnerTxt, vCate42WinnerValue, vCate42Value(10), vCate42Notice, vCate42Item
		Dim vCate42EntryCopy, vCate42PCLinkBanImg, vCate42PCLinkBanURL, vCate42MoLinkBanImg, vCate42MoLinkBanURL, vCate42Badgetag
		tmp = 0
		vCate42EntrySDate 	= requestCheckVar(Request("cate42entrysdate"),10)
		vCate42EntryEDate 	= requestCheckVar(Request("cate42entryedate"),10)
		vCate42AnnounDate 	= requestCheckVar(Request("cate42announdate"),10)
		vCate42WinnerTxt 	= requestCheckVar(Request("cate42winnertxt"),150)
		vCate42WinnerValue 	= requestCheckVar(Request("cate42winnervalue"),100)
		vCate42Notice			= html2db(Request("cate42notice"))
		vCate42EntryCopy		= html2db(Request("cate42entrycopy"))
		vCate42PCLinkBanImg	= requestCheckVar(Request("cate42pclinkbanimg"),100)
		vCate42PCLinkBanURL	= requestCheckVar(Request("cate42pclinkbanurl"),200)
		vCate42MoLinkBanImg	= requestCheckVar(Request("cate42molinkbanimg"),100)
		vCate42MoLinkBanURL	= requestCheckVar(Request("cate42molinkbanurl"),200)
		vCate42Item			= Trim(requestCheckVar(Request("cate42item"),10))
		vCate42Badgetag		= requestCheckVar(Request("badgetag"),16)
		
		'### [db_giftplus].[dbo].[tbl_play_thingthing] 저장
		vQuery = "IF EXISTS(select didx from [db_giftplus].[dbo].[tbl_play_thingthing] where didx = '" & vDidx & "') "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	UPDATE [db_giftplus].[dbo].[tbl_play_thingthing] SET "
		vQuery = vQuery & "		entry_sdate = '" & vCate42EntrySDate & "' "
		vQuery = vQuery & "		,entry_edate = '" & vCate42EntryEDate & "' "
		vQuery = vQuery & "		,announce_date = '" & vCate42AnnounDate & "' "
		vQuery = vQuery & "		,winnertxt = '" & vCate42WinnerTxt & "' "
		vQuery = vQuery & "		,winnervalue = '" & vCate42WinnerValue & "' "
		vQuery = vQuery & "		,notice = '" & vCate42Notice & "' "
		vQuery = vQuery & "		,entrycopy = '" & vCate42EntryCopy & "' "
		vQuery = vQuery & "		,badgetag = '" & vCate42Badgetag & "' "
		vQuery = vQuery & "	where didx = '" & vDidx & "' "
		vQuery = vQuery & "END "
		vQuery = vQuery & "ELSE "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	INSERT INTO [db_giftplus].[dbo].[tbl_play_thingthing](didx, entry_sdate, entry_edate, announce_date, winnertxt, winnervalue, notice, entrycopy, badgetag )"
		vQuery = vQuery & "	VALUES("
		vQuery = vQuery & "'" & vDidx & "', '" & vCate42EntrySDate & "', '" & vCate42EntryEDate & "', '" & vCate42AnnounDate & "', '" & vCate42WinnerTxt & "', "
		vQuery = vQuery & "'" & vCate42WinnerValue & "', '" & vCate42Notice & "', '" & vCate42EntryCopy & "', '" & vCate42Badgetag & "')"
		vQuery = vQuery & "END "
		dbget.Execute vQuery
		
		'### THING thingthing 이미지 gubun = 8
		For l=1 To 3
			tmp = tmp + 1
			vCate42Img(tmp)	= Trim(requestCheckVar(Request("cate42Img"&l&""),100))
		Next
		
		tmp = 0
		For l=1 To 3
			tmp = tmp + 1
			If vCate42Img(tmp) <> "" Then
				Call fnImageSave(vDidx, vCate, "8", vCate42Img(tmp), "", "", l, 0)
			End If
		Next
		
		'### THING thingthing 연결배너 PC 이미지 gubun = 21
		Call fnImageSave(vDidx, vCate, "21", vCate42PCLinkBanImg, vCate42PCLinkBanURL, "", "0", "")
		
		'### THING thingthing 연결배너 Mo 이미지 gubun = 22
		Call fnImageSave(vDidx, vCate, "22", vCate42MoLinkBanImg, vCate42MoLinkBanURL, "", "0", "")
		
		'### THING thingthing 그외이름 저장
		tmp = 0
		For l=1 To 10
			tmp = tmp + 1
			vCate42Value(tmp)	= Trim(requestCheckVar(Request("cate42value"&l&""),100))
		Next
		
		tmp = 0
		For l=1 To 10
			tmp = tmp + 1
			If vCate42Value(tmp) <> "" Then
				If l = 1 Then
					vQuery = "DELETE [db_giftplus].[dbo].[tbl_play_thingthing_winlist] where didx = '" & vDidx & "' "
				End If
				vQuery = vQuery & "INSERT INTO [db_giftplus].[dbo].[tbl_play_thingthing_winlist](didx,winnervalue) VALUES('" & vDidx & "','" & vCate42Value(tmp) & "') "
			End If
		Next
		If vQuery <> "" Then
			dbget.Execute vQuery
		End If
		
		vQuery = "DELETE [db_giftplus].[dbo].[tbl_play_item] WHERE didx = '" & vDidx & "' "
		If vCate42Item <> "" Then
			vQuery = vQuery & "INSERT INTO [db_giftplus].[dbo].[tbl_play_item](didx, itemid) VALUES('" & vDidx & "', '" & vCate42Item & "') "
		End If
		dbget.Execute vQuery
	ElseIf vCate = "43" Then '### THING 배경화면
		'### 배경화면 : cate = 43
		Dim vCate43PCImageURL, vCate43MoImageURL, vCate43QRImageURL, vCate43PCDown(3), vCate43PCLink(3), vCate43MoDown(3), vCate43MoLink(3)
		tmp = 0
		vCate43PCImageURL	= requestCheckVar(Request("cate43pcimg"),100)
		vCate43MoImageURL	= requestCheckVar(Request("cate43moimg"),100)
		vCate43QRImageURL	= requestCheckVar(Request("cate43qrimg"),100)
		
		'### THING 배경화면 PC 컨텐츠 이미지 gubun = 20
		Call fnImageSave(vDidx, vCate, "20", vCate43PCImageURL, "", "", "0", "")
		
		'### THING 배경화면 Mo 컨텐츠 이미지 gubun = 9
		Call fnImageSave(vDidx, vCate, "9", vCate43MoImageURL, "", "", "0", "")
		
		'### THING 배경화면 컨텐츠 QR이미지 gubun = 10
		Call fnImageSave(vDidx, vCate, "10", vCate43QRImageURL, "", "", "0", "")
		
		'### PC 다운로드 저장
		tmp = 0
		For l=1 To 3
			tmp = tmp + 1
			vCate43PCDown(tmp)	= Trim(requestCheckVar(Request("cate43pcdown"&l&""),100))
			vCate43PCLink(tmp)	= Trim(requestCheckVar(Request("cate43pclink"&l&""),100))
		Next
		
		tmp = 0
		For l=1 To 3
			tmp = tmp + 1
			If l = 1 Then
				vQuery = "DELETE [db_giftplus].[dbo].[tbl_play_thingdownload] where didx = '" & vDidx & "' and device = 'pc' "
			End If
			vQuery = vQuery & "INSERT INTO [db_giftplus].[dbo].[tbl_play_thingdownload](didx,device,download,link) "
			vQuery = vQuery & "VALUES('" & vDidx & "','pc','" & vCate43PCDown(tmp) & "','" & vCate43PCLink(tmp) & "') "
		Next
		If vQuery <> "" Then
			dbget.Execute vQuery
		End If
		
		'### Mo 다운로드 저장
		tmp = 0
		For l=1 To 3
			tmp = tmp + 1
			vCate43MoDown(tmp)	= Trim(requestCheckVar(Request("cate43modown"&l&""),100))
			vCate43MoLink(tmp)	= Trim(requestCheckVar(Request("cate43molink"&l&""),100))
		Next
		
		tmp = 0
		For l=1 To 3
			tmp = tmp + 1
			If l = 1 Then
				vQuery = "DELETE [db_giftplus].[dbo].[tbl_play_thingdownload] where didx = '" & vDidx & "' and device = 'mo' "
			End If
			vQuery = vQuery & "INSERT INTO [db_giftplus].[dbo].[tbl_play_thingdownload](didx,device,download,link) "
			vQuery = vQuery & "VALUES('" & vDidx & "','mo','" & vCate43MoDown(tmp) & "','" & vCate43MoLink(tmp) & "') "
		Next
		If vQuery <> "" Then
			dbget.Execute vQuery
		End If
	ElseIf vCate = "5" Then '### COMMA
		Dim vCate5PCTopImageURL, vCate5MoTopImageURL, vCate5Directer, vCate5Img(5), vCate5Copy(5), vCate5PCLinkBanImg, vCate5PCLinkBanURL, vCate5MoLinkBanImg, vCate5MoLinkBanURL
		tmp = 0
		vCate5PCTopImageURL		= requestCheckVar(Request("cate5pctopimg"),100)
		vCate5MoTopImageURL		= requestCheckVar(Request("cate5motopimg"),100)
		vCate5Directer			= requestCheckVar(Request("cate5directer"),150)
		vCate5PCLinkBanImg		= requestCheckVar(Request("cate5pclinkbanimg"),100)
		vCate5PCLinkBanURL		= requestCheckVar(Request("cate5pclinkbanurl"),200)
		vCate5MoLinkBanImg		= requestCheckVar(Request("cate5molinkbanimg"),100)
		vCate5MoLinkBanURL		= requestCheckVar(Request("cate5molinkbanurl"),200)
		
		'### comma 컨텐츠 PC 상단 이미지 gubun = 12
		Call fnImageSave(vDidx, vCate, "12", vCate5PCTopImageURL, "", "", "0", "")
		
		'### comma 컨텐츠 Mo 상단 이미지 gubun = 13
		Call fnImageSave(vDidx, vCate, "13", vCate5MoTopImageURL, "", "", "0", "")
		
		'### [db_giftplus].[dbo].[tbl_play_comma] 저장
		vQuery = "IF EXISTS(select didx from [db_giftplus].[dbo].[tbl_play_comma] where didx = '" & vDidx & "') "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	UPDATE [db_giftplus].[dbo].[tbl_play_comma] SET directer = '" & vCate5Directer & "' where didx = '" & vDidx & "' "
		vQuery = vQuery & "END "
		vQuery = vQuery & "ELSE "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	INSERT INTO [db_giftplus].[dbo].[tbl_play_comma](didx, directer) VALUES('" & vDidx & "', '" & vCate5Directer & "')"
		vQuery = vQuery & "END "
		dbget.Execute vQuery
		
		'### comma 연결배너 PC 이미지 gubun = 15
		Call fnImageSave(vDidx, vCate, "15", vCate5PCLinkBanImg, vCate5PCLinkBanURL, "", "0", "")
		
		'### comma 연결배너 Mo 이미지 gubun = 16
		Call fnImageSave(vDidx, vCate, "16", vCate5MoLinkBanImg, vCate5MoLinkBanURL, "", "0", "")
		
		'### comma 컨텐츠(에디터) 이미지 gubun = 14
		For l=1 To 5
			tmp = tmp + 1
			vCate5Img(tmp)	= Trim(requestCheckVar(Request("cate5Img"&l&""),100))
			vCate5Copy(tmp)	= html2db(Trim(Request("cate5copy"&l&"")))
		Next
		
		tmp = 0
		For l=1 To 5
			tmp = tmp + 1
			Call fnImageSave(vDidx, vCate, "14", vCate5Img(tmp), "", vCate5Copy(tmp), l, 0)
		Next
	ElseIf vCate = "6" Then '### HOWHOW
		Dim vCate6VideoURL, vCate6BannSub, vCate6BannTitle, vCate6BannBtnTitle, vCate6BannBtnLink, vCate6Img(4), vCate6Copy(4)
		tmp = 0
		vCate6VideoURL		= requestCheckVar(Request("cate6videourl"),190)
		vCate6BannSub			= requestCheckVar(Request("cate6bannsub"),100)
		vCate6BannTitle		= requestCheckVar(Request("cate6banntitle"),200)
		vCate6BannBtnTitle	= requestCheckVar(Request("cate6bannbtntitle"),100)
		vCate6BannBtnLink	= requestCheckVar(Request("cate6bannbtnlink"),300)
		
		'### [db_giftplus].[dbo].[tbl_play_howhow] 저장
		vQuery = "IF EXISTS(select didx from [db_giftplus].[dbo].[tbl_play_howhow] where didx = '" & vDidx & "') "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	UPDATE [db_giftplus].[dbo].[tbl_play_howhow] SET "
		vQuery = vQuery & "		videourl = '" & vCate6VideoURL & "' "
		vQuery = vQuery & "		,bannsub = '" & vCate6BannSub & "' "
		vQuery = vQuery & "		,banntitle = '" & vCate6BannTitle & "' "
		vQuery = vQuery & "		,bannbtntitle = '" & vCate6BannBtnTitle & "' "
		vQuery = vQuery & "		,bannbtnlink = '" & vCate6BannBtnLink & "' "
		vQuery = vQuery & "	where didx = '" & vDidx & "' "
		vQuery = vQuery & "END "
		vQuery = vQuery & "ELSE "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	INSERT INTO [db_giftplus].[dbo].[tbl_play_howhow](didx, videourl, bannsub, banntitle, bannbtntitle, bannbtnlink) "
		vQuery = vQuery & "VALUES('" & vDidx & "', '" & vCate6VideoURL & "', '" & vCate6BannSub & "', '" & vCate6BannTitle & "', "
		vQuery = vQuery & "'" & vCate6BannBtnTitle & "', '" & vCate6BannBtnLink & "')"
		vQuery = vQuery & "END "
		dbget.Execute vQuery
		
		'### howhow 컨텐츠(에디터) 이미지 gubun = 17
		For l=1 To 4
			tmp = tmp + 1
			vCate6Img(tmp)	= Trim(requestCheckVar(Request("cate6Img"&l&""),100))
			vCate6Copy(tmp)	= html2db(Trim(Request("cate6copy"&l&"")))
		Next
		
		tmp = 0
		For l=1 To 4
			tmp = tmp + 1
			Call fnImageSave(vDidx, vCate, "17", vCate6Img(tmp), "", vCate6Copy(tmp), l, 0)
		Next
		
	End If
		
	
Function fnImageSave(didx, cate, gubun, imgurl, linkurl, imagecopy, sortno, groupnum)
Dim vQuery
	If groupnum = "" Then
		vQuery = "IF EXISTS(select idx from [db_giftplus].[dbo].[tbl_play_image] where didx = '" & didx & "' and cate = '" & cate & "' and gubun = '" & gubun & "') "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	UPDATE [db_giftplus].[dbo].[tbl_play_image] SET "
		vQuery = vQuery & "		imgurl = '" & imgurl & "' "
		vQuery = vQuery & "		,linkurl = '" & linkurl & "' "
		vQuery = vQuery & "		,imagecopy = '" & imagecopy & "' "
		vQuery = vQuery & "		,sortno = '" & sortno & "' "
		vQuery = vQuery & "	where didx = '" & didx & "' and cate = '" & cate & "' and gubun = '" & gubun & "'"
		vQuery = vQuery & "END "
		vQuery = vQuery & "ELSE "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	INSERT INTO [db_giftplus].[dbo].[tbl_play_image](didx, cate, gubun, imgurl, linkurl, imagecopy, sortno) VALUES("
		vQuery = vQuery & "'" & didx & "', '" & cate & "', '" & gubun & "', '" & imgurl & "', '" & linkurl & "', '" & imagecopy & "', '" & sortno & "')"
		vQuery = vQuery & "END "
	ELSE
		vQuery = "IF EXISTS(select idx from [db_giftplus].[dbo].[tbl_play_image] where didx = '" & didx & "' and cate = '" & cate & "' and gubun = '" & gubun & "' and groupnum = '" & groupnum & "' and sortno = '" & sortno & "') "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	UPDATE [db_giftplus].[dbo].[tbl_play_image] SET "
		vQuery = vQuery & "		imgurl = '" & imgurl & "' "
		vQuery = vQuery & "		,linkurl = '" & linkurl & "' "
		vQuery = vQuery & "		,imagecopy = '" & imagecopy & "' "
		vQuery = vQuery & "		,sortno = '" & sortno & "' "
		vQuery = vQuery & "	where didx = '" & didx & "' and cate = '" & cate & "' and gubun = '" & gubun & "' and groupnum = '" & groupnum & "' and sortno = '" & sortno & "'"
		vQuery = vQuery & "END "
		vQuery = vQuery & "ELSE "
		vQuery = vQuery & "BEGIN "
		vQuery = vQuery & "	INSERT INTO [db_giftplus].[dbo].[tbl_play_image](didx, cate, gubun, imgurl, linkurl, imagecopy, sortno, groupnum) VALUES("
		vQuery = vQuery & "'" & didx & "', '" & cate & "', '" & gubun & "', '" & imgurl & "', '" & linkurl & "', '" & imagecopy & "', '" & sortno & "', '" & groupnum & "')"
		vQuery = vQuery & "END "
	End IF
	dbget.Execute vQuery
End Function
%>