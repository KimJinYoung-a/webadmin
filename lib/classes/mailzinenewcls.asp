<%
'###########################################################
' Description :  텐바이텐 메일진 클래스
' History : 2019.06.24 정태훈 생성 (템플릿 개선)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################

class CMailzineListSubItem

	public Fidx
	public Fregdate
	public Ftitle
	public Fisusing
	public fgubun
	public farea
	public fimg1
	public fimg2
	public fimg3
	public fimg4
	public fimgmap1
	public fimgmap2
	public fimgmap3
	public fimgmap4
	public fmngUserid
	public fmemgubun
	public fsecretgubun
	public FreservationDATE
	public Flastupdate

	public Fregtype
	public Fregtype2
	public Freguserid
	public Fmodiuserid
	public Fevt_code
	public fmailergubun

	public function GetRegTypeName()
		if (Fregtype = "1") then
			GetRegTypeName = "수기메일"
		elseif (Fregtype = "2") then
			GetRegTypeName = "주말특가"
		elseif (Fregtype = "3") then
			GetRegTypeName = "기획전"
		elseif (Fregtype = "4") then
			GetRegTypeName = "기획전+엠디픽"
		elseif (Fregtype = "5") then
			GetRegTypeName = "다이어리스토리"
		else
			GetRegTypeName = "ERR"
		end if
	end function

	' 사용금지. 일반 펑션의 GetRegNewTypeName 이거 사용할것.
	public function GetRegNewTypeName()
		if (Fregtype2 = "101") then
			GetRegNewTypeName = "주말특가"
		elseif (Fregtype2 = "102") then
			GetRegNewTypeName = "기획전"
		elseif (Fregtype2 = "103") then
			GetRegNewTypeName = "기획전+엠디픽"
		elseif (Fregtype2 = "104") then
			GetRegNewTypeName = "다이어리스토리"
		else
			GetRegNewTypeName = "ERR"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CMailzineClassItem
	public FclassDate

	public Fitemid1
	public FsalePer1
	public FclassDesc1
	public FclassSubDesc1

	public Fitemid2
	public FsalePer2
	public FclassDesc2
	public FclassSubDesc2

	public Fitemid3
	public FsalePer3
	public FclassDesc3
	public FclassSubDesc3

	public Fregdate
	public FlastUpdate
	public Freguserid
	public Fmodiuserid
	public fmailergubun

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CMailzineList
	public FItemList()
	public FTotalCount
	public FResultCount
	public FRectDesignerID
	public FCurrPage
	public foneitem
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FPCount
	public FrectSDate
	public FrectEDate
	public FrectSearchKey
	public FrectUsing
	public FrectArea

	public FRectRegType
	public FRectDate

	public frectidx
	public frectmailergubun

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public sub MailzineDetail()
		dim sqlStr,code, sqlsearch

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "&frectidx&"" + vbcrlf
		end if
		if FRectRegType <> "" then
			sqlsearch = sqlsearch & " and IsNull(regtype,'1') = '" & FRectRegType & "' " + vbcrlf
		end if
		if frectmailergubun <> "" then
			sqlsearch = sqlsearch & " and mailergubun = '" & frectmailergubun & "' " & vbcrlf
		end if

		sqlStr = "select title,regdate,img1,img2,img3,img4,imgmap1,imgmap2,imgmap3,imgmap4,isusing,gubun,area,mngUserid,memgubun,secretgubun,reservationDATE, IsNull(regtype,'1') as regtype, reguserid, modiuserid, evt_code, IsNull(regtype2,'101') as regtype2" + vbcrlf
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].tbl_mailzine with (readuncommitted)" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount

			set FOneItem = new CMailzineListSubItem

			if  not rsget.EOF  then
				FOneItem.Ftitle		= db2html(rsget("title"))
				FOneItem.Fregdate	= rsget("regdate")
				FOneItem.Fimg1		= rsget("img1")
				FOneItem.Fimg2		= rsget("img2")
				FOneItem.Fimg3		= rsget("img3")
				FOneItem.Fimg4		= rsget("img4")
				FOneItem.Fimgmap1	= db2html(rsget("imgmap1"))
				FOneItem.Fimgmap2	= db2html(rsget("imgmap2"))
				FOneItem.Fimgmap3	= db2html(rsget("imgmap3"))
				FOneItem.Fimgmap4	= db2html(rsget("imgmap4"))
				FOneItem.Fisusing	= rsget("isusing")
				FOneItem.Fgubun		= rsget("gubun")
				FOneItem.Farea		= rsget("area")
				FOneItem.FmngUserid	= rsget("mngUserid")
				FOneItem.FmemGubun = rsget("memGubun")
				FOneItem.Fsecretgubun = rsget("secretgubun")
				FOneItem.FreservationDATE = rsget("reservationDATE")

				FOneItem.Fregtype = rsget("regtype")
				FOneItem.Freguserid = rsget("reguserid")
				FOneItem.Fmodiuserid = rsget("modiuserid")
				FOneItem.Fevt_code = rsget("evt_code")
				FOneItem.Fregtype2 = rsget("regtype2")
			end if
		rsget.Close
	end sub

	'//admin/mailzine/mailzine_list.asp
	public sub MailzineList()
		dim sqlStr, addSql, i

		addSql = ""
		'# 기간검색
		if FrectSDate<>"" then 		addSql = addSql & " and regdate between '" & Replace(FrectSDate,"-",".") & "' and '" & Replace(FrectEDate,"-",".") & "' "
		'# 제목검색
		if FrectSearchKey<>"" then	addSql = addSql & " and title like '%" & FrectSearchKey & "%' "
		'# 노출여부
		if FrectUsing<>"" then		addSql = addSql & " and isusing='" & FrectUsing & "' "
		'# 발송대상
		if FrectArea<>"" then		addSql = addSql & " and area='" & FrectArea & "' "
		if frectmailergubun <> "" then
			addSql = addSql & " and mailergubun = '" & frectmailergubun & "' " & vbcrlf
		end if

		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].tbl_mailzine with (readuncommitted) Where 1=1 and deleteyn='N' " & addSql + vbcrlf

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx, title, regdate, isusing, gubun, area, mngUserid, memgubun, reservationDATE, lastupdate" & vbcrlf
		sqlStr = sqlStr & " , IsNull(regtype,'1') as regtype, IsNull(regtype2,'101') as regtype2, mailergubun" + vbcrlf
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].tbl_mailzine with (readuncommitted)" + vbcrlf
		sqlStr = sqlStr & " where 1=1 and deleteyn='N' " & addSql + vbcrlf
		sqlStr = sqlStr & " order by regdate Desc" + vbcrlf

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CMailzineListSubItem

				FItemList(i).fmemgubun = rsget("memgubun")
				FItemList(i).fmngUserid = rsget("mngUserid")
				FItemList(i).Fidx = rsget("idx")
				FItemList(i).Ftitle = db2html(rsget("title"))
			    FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Fisusing = rsget("isusing")
				FItemList(i).fgubun = rsget("gubun")
				FItemList(i).farea = rsget("area")
				FItemList(i).FreservationDATE = rsget("reservationDATE")
				FItemList(i).Flastupdate = rsget("lastupdate")
				FItemList(i).Fregtype = rsget("regtype")
				FItemList(i).Fregtype2 = rsget("regtype2")
				FItemList(i).fmailergubun = rsget("mailergubun")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// /admin/mailzine/mailzine_class_list.asp
	public sub MailzineClassList()
		dim sqlStr, addSql, i

		addSql = ""
		if (FrectSDate <> "") and (FrectEDate <> "") then
			addSql = addSql + " and classDate between '" & FrectSDate & "' and '" & FrectEDate & "' "
		end if
		if frectmailergubun <> "" then
			addSql = addSql & " and mailergubun = '" & frectmailergubun & "' " & vbcrlf
		end if

		sqlStr = " select count(*) as cnt from [db_sitemaster].[dbo].[tbl_mailzine_class] with (readuncommitted) "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + addSql

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + " * " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_mailzine_class] with (readuncommitted) "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + addSql
		sqlStr = sqlStr + " order by classDate desc "

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CMailzineClassItem

				FItemList(i).FclassDate = rsget("classDate")

				FItemList(i).Fitemid1 = rsget("itemid1")
				FItemList(i).FsalePer1 = rsget("salePer1")
				FItemList(i).FclassDesc1 = rsget("classDesc1")
				FItemList(i).FclassSubDesc1 = rsget("classSubDesc1")

				FItemList(i).Fitemid2 = rsget("itemid2")
				FItemList(i).FsalePer2 = rsget("salePer2")
				FItemList(i).FclassDesc2 = rsget("classDesc2")
				FItemList(i).FclassSubDesc2 = rsget("classSubDesc2")

				FItemList(i).Fitemid3 = rsget("itemid3")
				FItemList(i).FsalePer3 = rsget("salePer3")
				FItemList(i).FclassDesc3 = rsget("classDesc3")
				FItemList(i).FclassSubDesc3 = rsget("classSubDesc3")

				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).FlastUpdate = rsget("lastUpdate")
				FItemList(i).Freguserid = rsget("reguserid")
				FItemList(i).Fmodiuserid = rsget("modiuserid")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub MailzineClassOne()
		dim sqlStr, addSql

		sqlStr = "select top 1 * " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_mailzine_class] with (readuncommitted) "
		sqlStr = sqlStr + " where classDate = '" & FRectDate & "' "

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount

		set FOneItem = new CMailzineClassItem

		if  not rsget.EOF  then

			FOneItem.FclassDate = rsget("classDate")

			FOneItem.Fitemid1 = rsget("itemid1")
			FOneItem.FsalePer1 = rsget("salePer1")
			FOneItem.FclassDesc1 = rsget("classDesc1")
			FOneItem.FclassSubDesc1 = rsget("classSubDesc1")

			FOneItem.Fitemid2 = rsget("itemid2")
			FOneItem.FsalePer2 = rsget("salePer2")
			FOneItem.FclassDesc2 = rsget("classDesc2")
			FOneItem.FclassSubDesc2 = rsget("classSubDesc2")

			FOneItem.Fitemid3 = rsget("itemid3")
			FOneItem.FsalePer3 = rsget("salePer3")
			FOneItem.FclassDesc3 = rsget("classDesc3")
			FOneItem.FclassSubDesc3 = rsget("classSubDesc3")

			FOneItem.Fregdate = rsget("regdate")
			FOneItem.FlastUpdate = rsget("lastUpdate")
			FOneItem.Freguserid = rsget("reguserid")
			FOneItem.Fmodiuserid = rsget("modiuserid")
		end if
		rsget.Close
	end sub

	' 메일진 템플릿 정보 가져오기	' 2020.05.27 정태훈 생성
	public Function fnMailzineTemplateInfo
		Dim strSql
		strSql ="EXEC [db_sitemaster].[dbo].usp_SCM_Mailzine_TemplateInfo_Get " & FRectRegType
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdText
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnMailzineTemplateInfo = rsget.GetRows()
		END IF
		rsget.close
	End Function

	' 메일진 템플릿 디테일 정보 가져오기	' 2020.05.27 정태훈 생성
	public Function fnMailzineTemplateDetail
		Dim strSql
		strSql ="EXEC [db_sitemaster].[dbo].usp_SCM_Mailzine_TemplateDetailInfo_Get " & frectidx
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdText
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnMailzineTemplateDetail = rsget.GetRows()
		END IF
		rsget.close
	End Function

	' 메일진 템플릿 컨텐츠 혼합 정보 가져오기	' 2020.05.27 정태훈 생성
	public Function fnMailzineTemplateContents
		Dim strSql
		strSql ="EXEC [db_sitemaster].[dbo].usp_SCM_Mailzine_TemplateContentsInfo_Get " & FRectRegType & "," & frectidx
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdText
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnMailzineTemplateContents = rsget.GetRows()
		END IF
		rsget.close
	End Function

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

'/텐바이텐 메일링 회원
function mailzine_member_count()
	dim sql

	sql = "select"
	sql = sql & " count(u.userid) as 'memberAllCount'"
	sql = sql & " ,count(case when u.emailok = 'Y' then u.userid end) as 'memberEmailYCount'"
	sql = sql & " , sum(case when l.lastlogin>=dateadd(year,-3,getdate()) then 1 else 0 end) as recent3yearsMemberAllCount"
	sql = sql & " , sum(case when l.lastlogin>=dateadd(year,-3,getdate()) and u.emailok = 'Y' then 1 else 0 end) as recent3yearsMemberEmailYCount"
	sql = sql & " FROM [db_user].dbo.tbl_user_n u with (nolock)"
	sql = sql & " join db_user.dbo.tbl_logindata l with (nolock)"
	sql = sql & " 	on u.userid=l.userid"

	'response.write sql &"<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		if not rsget.EOF  then
			response.write "전체회원: "& rsget("memberAllCount") & " 명"
			response.write "<br>전체회원(텐바이텐수신 Y): "& rsget("memberEmailYCount") & " 명"
			response.write "<br>최근3년로그인: "& rsget("recent3yearsMemberAllCount") & " 명"
			response.write "<br><font color='red'>최근3년로그인(텐바이텐수신 Y): "& rsget("recent3yearsMemberEmailYCount") & " 명</font> <- 회원메일진 대상"
		end if

	rsget.close
end function

'/텐바이텐 비회원 메일링
function mailzine_notmember_count()
	dim sql

	sql = "select"
	sql = sql & " count(idx) as 'notMemberAllCount'"
	sql = sql & " ,count(case when email_10x10 = 'Y' then idx end) as 'notMemberEmailYCount'"
	sql = sql & " from db_user.dbo.tbl_mailzine_notmember with (readuncommitted)"
	sql = sql & " where isusing = 'Y'"

	'response.write sql &"<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF  then
			response.write "전체비회원: "& rsget("notMemberAllCount") & " 명"
			response.write "<br>전체비회원(텐바이텐수신 Y): "& rsget("notMemberEmailYCount") & " 명"
		end if
	rsget.close
end function

'-----------------------------------------------------------------------
' sbGetDesignerid :웹디자인팀 부서번호(12)로 디자이너이름 리스트가져오기
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
Sub sbGetDesignerid(ByVal selName, ByVal sIDValue, ByVal sScript)
Dim strSql, arrList, intLoop, strResult
	strSql = "SELECT p.id, t.username from db_partner.[dbo].tbl_partner as p with (readuncommitted) "
	strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as t with (readuncommitted) on p.id = t.userid "
	strSql = strSql & " WHERE p.part_sn ='12' "
	strSql = strSql & " and p.isUsing ='Y' "
	strSql = strSql & " order by p.level_sn "

	rsget.Open strSql,dbget
		IF not rsget.eof THEN
			arrList = rsget.getRows()
		End IF
	rsget.close

	if isNull(sIDValue) then sIDValue=""

	strResult = "<select name='" & selName & "' " & sScript & " class='select'>" &_
				"<option value=''>선택</option>"

	If isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
			strResult = strResult & "<option value='" & arrList(0,intLoop) & "'"
			if Cstr(arrList(0,intLoop)) = Cstr(sIDValue) then
				strResult = strResult & " selected"
			end if
			strResult = strResult & ">" & arrList(1,intLoop) & "</option>"
		Next
	End IF

	strResult = strResult & "</select>"
	Response.Write strResult
End Sub

function Drawareagubun(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%= selectBoxName %>" <%= changeFlag %>>
		<option value="" <% if selectedId = "" then response.write "selected"%>>선택</option>
		<option value="ten_all" <% if selectedId = "ten_all" then response.write "selected"%>>텐바이텐_전지역</option>
		<option value="ten_metropolitan" <% if selectedId = "ten_metropolitan" then response.write "selected"%>>텐바이텐_수도권</option>
		<option value="ten_metro_jeju" <% if selectedId = "ten_metro_jeju" then response.write "selected"%>>텐바이텐_제주도</option>
		<option value="finger_all" <% if selectedId = "finger_all" then response.write "selected"%>>핑거스_전지역</option>
		<option value="finger_metropolitan" <% if selectedId = "finger_metropolitan" then response.write "selected"%>>핑거스_수도권</option>
		<option value="girl_all" <% if selectedId = "girl_all" then response.write "selected"%>>유아러걸_전지역</option>
		<option value="ten_china" <% if selectedId = "ten_china" then response.write "selected"%>>텐바이텐_중국</option>
	</select>
<%
end function

function getareagubun(tmp)
	if tmp = "ten_all" then
		getareagubun = "텐바이텐_전지역"
	elseif tmp = "ten_metropolitan" then
		getareagubun = "텐바이텐_수도권"
	elseif tmp = "finger_all" then
		getareagubun = "핑거스_전지역"
	elseif tmp = "finger_metropolitan" then
		getareagubun = "핑거스_수도권"
	elseif tmp = "girl_all" then
		getareagubun = "유아러걸_전지역"
	elseif tmp = "ten_metro_jeju" then
		getareagubun = "텐바이텐_제주도"
	elseif tmp = "ten_china" then
		getareagubun = "텐바이텐_중국"
	else
		getareagubun = "지정안됨"
	end if
end function

function Drawisusing(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%= selectBoxName %>" <%= changeFlag %>>
		<option value="" <% if selectedId = "" then response.write "selected"%>>선택</option>
		<option value="Y" <% if selectedId = "Y" then response.write "selected"%>>Y</option>
		<option value="N" <% if selectedId = "N" then response.write "selected"%>>N</option>
	</select>
<%
end function

function DrawMemberGubun(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%= selectBoxName %>" <%= changeFlag %>>
		<option value="" <% if selectedId = "" then response.write "selected"%>>선택</option>
		<option value="member_all" <% if selectedId = "member_all" then response.write "selected"%>>모든회원</option>
		<option value="BLUE" <% if selectedId = "BLUE" then response.write "selected"%>>BLUE</option>
		<option value="FAMILY" <% if selectedId = "FAMILY" then response.write "selected"%>>FAMILY</option>
		<option value="FRIENDS" <% if selectedId = "FRIENDS" then response.write "selected"%>>FRIENDS</option>
		<option value="GREEN" <% if selectedId = "GREEN" then response.write "selected"%>>GREEN</option>
		<option value="ORANGE" <% if selectedId = "ORANGE" then response.write "selected"%>>ORANGE</option>
		<option value="STAFF" <% if selectedId = "STAFF" then response.write "selected"%>>STAFF</option>
		<option value="VIPGOLD" <% if selectedId = "VIPGOLD" then response.write "selected"%>>VIPGOLD</option>
		<option value="VIPSILVER" <% if selectedId = "VIPSILVER" then response.write "selected"%>>VIPSILVER</option>
		<option value="VIPGOLD_SILVER" <% if selectedId = "VIPGOLD_SILVER" then response.write "selected"%>>VIPGOLD_SILVER</option>
		<option value="YELLOW" <% if selectedId = "YELLOW" then response.write "selected"%>>YELLOW</option>
	</select>
<%
End function

function DrawsecretGubun(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%= selectBoxName %>" <%= changeFlag %>>
		<option value="" <% if selectedId = "" then response.write "selected"%>>선택</option>
		<option value="Y" <% if selectedId = "Y" then response.write "selected"%>>Y</option>
		<option value="N" <% if selectedId = "N" then response.write "selected"%>>N</option>
	</select>
<%
end function

'// traking tag 생성
Function MailzineTrakingTag(sendDate,campaignCd)
	Dim trakingTag, campaignName

	if sendDate="" then sendDate = date()
	sendDate = replace(sendDate,"-","")

	Select Case cStr(campaignCd)
		Case "1"
			campaignName = "manual"			'수기메일
		Case "101"
			campaignName = "weekend"		'주말특가
		Case "102"
			campaignName = "event"			'기획전
		Case "103"
			campaignName = "mdpick"			'기획전+MD Pick
		Case "104"
			campaignName = "diarystory"		'다이어리스토리
	end Select

	trakingTag = "rdsite=tmailer"						'lagacy : 유입사이트
	trakingTag = trakingTag & "&utm_source=10x10"		'amplitude : 유입처
	trakingTag = trakingTag & "&utm_medium=mailzine"	'amplitude : 유입방법
	trakingTag = trakingTag & "&utm_campaign=" & sendDate & "_" & campaignName	'amplitude : 유입캠페인

	MailzineTrakingTag = trakingTag
End Function

function DrawMailzineKind(selectBoxName, selectedId, changeFlag)
Dim strSql, arrList, intLoop, strResult
	strSql = "SELECT code_value, code_desc FROM db_sitemaster.dbo.tbl_mailzine_code with (readuncommitted)"
	strSql = strSql & " WHERE code_type='mailzineKind'"
	strSql = strSql & " AND code_using='Y'"
	strSql = strSql & " AND code_dispYN='Y'"
	strSql = strSql & " ORDER BY code_sort ASC"
	rsget.Open strSql,dbget
		IF not rsget.eof THEN
			arrList = rsget.getRows()
		End IF
	rsget.close
%>
	<select name="<%= selectBoxName %>" class="select" <%= changeFlag %>>
		<option value="" <% if selectedId = "" then response.write "selected"%>>선택</option>
<%
	If isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
%>
		
		<option value="<%=arrList(0,intLoop)%>"<% if selectedId = arrList(0,intLoop)  then response.write " selected"%>><%=arrList(1,intLoop)%></option>
<%
		Next
	End IF
%>
	</select>
<%
end function

' 작성구분		' 2020.02.12 한용민 생성 자동화 처리
function GetRegNewTypeName(regtype)
	Dim strSql, tmpregtypename

	regtype = trim(getNumeric(regtype))
	if regtype="" or isnull(regtype) then exit function

	strSql = "SELECT code_value, code_desc" & vbcrlf
	strSql = strSql & " FROM db_sitemaster.dbo.tbl_mailzine_code with (readuncommitted)"
	strSql = strSql & " WHERE code_type='mailzineKind'"
	strSql = strSql & " AND code_using='Y'"
	strSql = strSql & " AND code_dispYN='Y'"
	strSql = strSql & " AND code_value='"& regtype &"'"
	strSql = strSql & " ORDER BY code_sort ASC"

	'response.write strSql & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.eof THEN
			tmpregtypename = rsget("code_desc")
		End IF
	rsget.close

	GetRegNewTypeName=tmpregtypename
end function

' 메일러구분	' 2020.05.28 한용민 생성
function Drawmailergubun(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%= selectBoxName %>" <%= changeFlag %>>
		<option value="" <% if selectedId = "" then response.write "selected"%>>선택</option>
		<option value="EMS" <% if selectedId = "EMS" then response.write "selected"%>>EMS</option>
		<option value="TMS" <% if selectedId = "TMS" then response.write "selected"%>>TMS</option>
	</select>
<%
end function

%>
