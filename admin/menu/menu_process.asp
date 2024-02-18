<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 메뉴관리
' History : 서동석 생성
'			2021.10.19 한용민 수정(수정로그 저장)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/admin/MenuCls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
dim mode, menuid, menupos, viewidx, menuname, menuname_en, linkurl,parentid,divcd, isusing, menucolor
dim param, strMsg, page, SearchKey, SearchString, part_sn, level_sn, childYN, child_sn, logtype, logmsg, adminid
dim SQL, splPsn, splLsn, splCsn, lp1, lp2, useSslYN, criticinfo, saveLog, lv1personal, lv1customerYN, lv2partnerYN, lv3InternalYN
dim orgviewidx, orgmenuname, orgmenuname_en, orglinkurl, orglv1customerYN, orglv2partnerYN, orglv3InternalYN
dim orguseSslYN, orgsaveLog, orgmenucolor, orgparentid, orgisUsing, orgdivcd, opart, orgpartstr, currpartstr, i
menupos		= requestCheckvar(request.Form("menupos"),10)
mode		= requestCheckvar(request.Form("mode"),32)
menuid		= requestCheckvar(request.Form("mid"),32)
viewidx		= requestCheckvar(request.Form("viewidx"),10)
menuname	= request.Form("menuname")
menuname_en	= request.Form("menuname_en")
linkurl		= request.Form("linkurl")
parentid	= requestCheckvar(request.Form("parentid"),10)
part_sn		= requestCheckvar(request.Form("part_sn"),500)      ''array
level_sn	= requestCheckvar(request.Form("level_sn"),500)     ''array
childYN		= requestCheckvar(request.Form("childYN"),10)
isusing		= requestCheckvar(request.Form("isusing"),10)
menucolor	= requestCheckvar(request.Form("menucolor"),10)
page		= requestCheckvar(request.Form("page"),10)
SearchKey	= request.Form("SearchKey")
SearchString = request.Form("SearchString")

divcd		= requestCheckvar(request.Form("divcd"),10)
useSslYN 	= requestCheckvar(request.Form("useSslYN"),10)
criticinfo  = requestCheckvar(request.Form("criticinfo"),10)
saveLog  	= requestCheckvar(request.Form("saveLog"),10)
lv1customerYN 	= requestCheckvar(request.Form("lv1customerYN"),1)
lv2partnerYN 	= requestCheckvar(request.Form("lv2partnerYN"),1)
lv3InternalYN 	= requestCheckvar(request.Form("lv3InternalYN"),1)
adminid=session("ssBctId")
logmsg=""
orgpartstr=""
currpartstr=""
i=0
if (useSslYN = "") then
	useSslYN = "N"
end if

if (criticinfo = "") then
	criticinfo = "0"
end if

if (saveLog = "") then
	saveLog = "0"
end if

if lv1customerYN="" or isnull(lv1customerYN) then lv1customerYN="N"
if lv2partnerYN="" or isnull(lv2partnerYN) then lv2partnerYN="N"
if lv3InternalYN="" or isnull(lv3InternalYN) then lv3InternalYN="N"

'페이지 파라메터
param = "?menupos=" & menupos & "&pid=" & parentid & "& page=" & page & "&SearchKey=" & SearchKey & "&SearchString=" & SearchString

'response.write part_sn & "<br>"
'response.write level_sn & "<br>"
'dbget.close()	:	response.End


'트랜젝션 시작
dbget.beginTrans

'// 메뉴정보 처리 분기 //
Select Case mode
	Case "add"
		strMsg = "메뉴를 등록하였습니다."
		SQL =	"Insert into db_partner.[dbo].tbl_partner_menu " &VbCRLF
		SQL = SQL & "	(viewidx, menuname, linkurl, parentid, menucolor, isusing, divcd, menuname_en, useSslYN, criticinfo, saveLog, lv1customerYN, lv2partnerYN, lv3InternalYN) values " &VbCRLF
		SQL = SQL & "	('" & viewidx & "' " &VbCRLF
		SQL = SQL & "	,'" & menuname & "' " &VbCRLF
		SQL = SQL & "	,'" & linkurl & "' " &VbCRLF
		SQL = SQL & "	,'" & parentid & "' " &VbCRLF
		SQL = SQL & "	,'" & menucolor & "' " &VbCRLF
		SQL = SQL & "	,'" & isusing & "' " &VbCRLF
		SQL = SQL & "	,'" & divcd & "'"&VbCRLF
		SQL = SQL & "	,'" & menuname_en & "'"&VbCRLF
		SQL = SQL & "	,'" & useSslYN & "'"&VbCRLF
		SQL = SQL & "	,'" & criticinfo & "'"&VbCRLF
		SQL = SQL & "	,'" & saveLog & "'"&VbCRLF
		SQL = SQL & "	,'" & lv1customerYN & "'"&VbCRLF
		SQL = SQL & "	,'" & lv2partnerYN & "'"&VbCRLF		
		SQL = SQL & "	,'" & lv3InternalYN & "'"&VbCRLF
		SQL = SQL & "	)"&VbCRLF

		dbget.Execute(SQL)

		SQL = "select @@identity "
		rsget.Open SQL, dbget
			menuid = rsget(0)
		rsget.Close

		logtype="A0"
		logmsg="메뉴권한신규등록"

	Case "modi", "popmodi"
		strMsg = "메뉴정보를 수정하였습니다."

		SQL ="select"
		SQL = SQL & " isnull(viewidx,'') as viewidx, isnull(menuname,'') as menuname, isnull(menuname_en,'') as menuname_en, isnull(linkurl,'') as linkurl"
		SQL = SQL & " , isnull(lv1customerYN,'N') as lv1customerYN, isnull(lv2partnerYN,'N') as lv2partnerYN, isnull(lv3InternalYN,'N') as lv3InternalYN"
		SQL = SQL & " , isnull(useSslYN,'N') as useSslYN, isnull(saveLog,0) as saveLog, isnull(menucolor,'') as menucolor"
		SQL = SQL & " , isnull(parentid,'') as parentid, isnull(isUsing,'') as isUsing, isnull(divcd,'') as divcd"
		SQL = SQL & " from db_partner.[dbo].tbl_partner_menu with (nolock)"
		SQL = SQL & " where id="& menuid &""

		'response.write SQL & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

		IF not rsget.EOF THEN
			orgviewidx = rsget("viewidx")
			orgmenuname = rsget("menuname")
			orgmenuname_en = rsget("menuname_en")
			orglinkurl = rsget("linkurl")
			orglv1customerYN = rsget("lv1customerYN")
			orglv2partnerYN = rsget("lv2partnerYN")
			orglv3InternalYN = rsget("lv3InternalYN")
			orguseSslYN = rsget("useSslYN")
			orgsaveLog = rsget("saveLog")
			orgmenucolor = rsget("menucolor")
			orgparentid = rsget("parentid")
			orgisUsing = rsget("isUsing")
			orgdivcd = rsget("divcd")
		END IF
		rsget.close

		if cstr(orgviewidx)<>cstr(viewidx) then logmsg = logmsg & "/표시순서: "& orgviewidx &" -> "& viewidx &""
		if cstr(orgmenuname)<>cstr(menuname) then logmsg = logmsg & "/메뉴명: "& orgmenuname &" -> "& menuname &""
		if cstr(orgmenuname_en)<>cstr(menuname_en) then logmsg = logmsg & "/메뉴명(영문): "& orgmenuname_en &" -> "& menuname_en &""
		if cstr(orglinkurl)<>cstr(linkurl) then logmsg = logmsg & "/링크URL: "& orglinkurl &" -> "& linkurl &""
		if orglv1customerYN="" then orglv1customerYN="N"
		if lv1customerYN="" then lv1customerYN="N"
		if cstr(orglv1customerYN)<>cstr(lv1customerYN) then
			logmsg = logmsg & "/LV1(고객정보): "& orglv1customerYN &" -> "& lv1customerYN &""
		end if
		if orglv2partnerYN="" then orglv2partnerYN="N"
		if lv2partnerYN="" then lv2partnerYN="N"
		if cstr(orglv2partnerYN)<>cstr(lv2partnerYN) then
			logmsg = logmsg & "/LV2(파트너정보): "& orglv2partnerYN &" -> "& lv2partnerYN &""
		end if
		if orglv3InternalYN="" then orglv3InternalYN="N"
		if lv3InternalYN="" then lv3InternalYN="N"
		if cstr(orglv3InternalYN)<>cstr(lv3InternalYN) then
			logmsg = logmsg & "/LV3(내부정보): "& orglv3InternalYN &" -> "& lv3InternalYN &""
		end if
		if orguseSslYN="" then orguseSslYN="N"
		if useSslYN="" then useSslYN="N"
		if cstr(orguseSslYN)<>cstr(useSslYN) then logmsg = logmsg & "/SSL 사용: "& orguseSslYN &" -> "& useSslYN &""
		if orgsaveLog="" then orgsaveLog="N"
		if saveLog="" then saveLog="N"
		if cstr(orgsaveLog)<>cstr(saveLog) then logmsg = logmsg & "/접속로그 저장: "& orgsaveLog &" -> "& saveLog &""
		if cstr(orgmenucolor)<>cstr(menucolor) then logmsg = logmsg & "/표시색상: "& orgmenucolor &" -> "& menucolor &""
		if cstr(orgparentid)<>cstr(parentid) then logmsg = logmsg & "/상위메뉴: "& orgparentid &" -> "& parentid &""
		if cstr(orgisUsing)<>cstr(isUsing) then logmsg = logmsg & "/사용여부: "& orgisUsing &" -> "& isUsing &""
		if cstr(orgdivcd)<>cstr(divcd) then logmsg = logmsg & "/기존권한: "& orgdivcd &" -> "& divcd &""

		'if left(logmsg,1)="/" then logmsg=mid(logmsg,2,4000)
		if logmsg<>"" then logmsg="메뉴권한수정"&logmsg

		'메뉴 본문내용 수정
		SQL =	"Update db_partner.[dbo].tbl_partner_menu Set " & VbCRLF
		SQL = SQL & "	viewidx		= '" & viewidx & "' " & VbCRLF
		SQL = SQL & "	,menuname	= '" & menuname & "' " & VbCRLF
		SQL = SQL & "	,linkurl	= '" & linkurl & "' " & VbCRLF
		SQL = SQL & "	,parentid	= '" & parentid & "' " & VbCRLF
		SQL = SQL & "	,menucolor	= '" & menucolor & "' " & VbCRLF
		SQL = SQL & "	,isusing	= '" & isusing & "' " & VbCRLF
		SQL = SQL & "	,divcd	    = '" & divcd & "' " & VbCRLF
		SQL = SQL & "	,menuname_en= '" & menuname_en & "' " & VbCRLF
		SQL = SQL & "	,useSslYN	= '" & useSslYN & "' " & VbCRLF
		SQL = SQL & "	,criticinfo = '" & criticinfo & "' " & VbCRLF
		SQL = SQL & "	,saveLog 	= '" & saveLog & "' " & VbCRLF
		SQL = SQL & "	,lv1customerYN 	= '" & lv1customerYN & "' " & VbCRLF
		SQL = SQL & "	,lv2partnerYN 	= '" & lv2partnerYN & "' " & VbCRLF
		SQL = SQL & "	,lv3InternalYN 	= '" & lv3InternalYN & "' Where" & VbCRLF
		SQL = SQL & " id=" & menuid
		dbget.Execute(SQL)

		'종속메뉴번호 접수
		if childYN="Y" then
			SQL = "Select id From db_partner.[dbo].tbl_partner_menu Where parentid=" & menuid
			rsget.Open SQL, dbget
			if Not(rsget.EOF or rsget.BOF) then
				Do Until rsget.EOF
					child_sn = child_sn & rsget(0)
					rsget.MoveNext
					if Not(rsget.EOF) then child_sn = child_sn & ","
				loop
			end if
			rsget.Close
		end if

		logtype="A3"
end Select

Set opart = new CTenByTenMember
	opart.FPagesize = 500
	opart.FCurrPage = 1
	opart.frectmenuid=menuid
	opart.getpartner_menu_part()

orgpartstr=""
if opart.FResultCount>0 then
	for i=0 to opart.FResultCount - 1	
		orgpartstr=orgpartstr & "부서번호" & opart.FitemList(i).fpart_sn & "등급번호" & opart.FitemList(i).flevel_sn
	next
end if
set opart=nothing

'// 메뉴권한 저장 //
'기존 자료 정리
SQL = "Delete From db_partner.dbo.tbl_menu_part Where menu_id=" & menuid
dbget.Execute(SQL)

'수정자료 저장
if part_sn<>"" then splPsn = Split(part_sn, ",")
if level_sn<>"" then splLsn = Split(level_sn, ",")

''IsArray(splPsn) : 추가 서동석 - 권한이 없는경우에도 수정 가능하도록 변경
If IsArray(splPsn) Then
    For lp1=0 to Ubound(splPsn)
    	IF Trim(splPsn(lp1))<>"" and Trim(splLsn(lp1))<>"" THEN
    	SQL =	"Insert Into db_partner.dbo.tbl_menu_part (menu_id, part_sn, level_sn) " &_
    			"Values (" & menuid & ", " & splPsn(lp1) & ", " & splLsn(lp1) & ")"
    	dbget.Execute(SQL)
    	END IF
    Next

    '// 부모메뉴 권한 추가(부모가 있고, 부모에 자식 권한이 없는경우;여집합 추가;허진원2012-03-08) //
    if parentid>0 then

		'1. 상위권한중에 추가되는것 보다 높은 등급이 있으면 삭제
		SQL =	"Delete T1 " &_
				"From db_partner.dbo.tbl_menu_part as T1 " &_
				"	join db_partner.dbo.tbl_menu_part as T2 " &_
				"		on T1.part_sn=T2.part_sn " &_
				"Where T1.menu_id=" & parentid &_
				"	and T2.menu_id=" & menuid &_
				"	and T1.level_sn<T2.level_sn "
		dbget.Execute(SQL)

		'2. 추가되는것 중에 상위권한보다 낮은 등급 또는 신규일때 추가
		SQL =	"insert into db_partner.dbo.tbl_menu_part " &_
				"Select T.* " &_
				"from ( " &_
				"		select " & parentid & " as menu_id, part_sn, level_sn " &_
				"		from ( " &_
				"			select 1 as id, part_sn, level_sn from db_partner.dbo.tbl_menu_part where menu_id=" & menuid &_
				"			union all " &_
				"			select 2 as id, part_sn, level_sn from db_partner.dbo.tbl_menu_part where menu_id=" & parentid &_
				"		) as tt " &_
				"		group by part_sn, level_sn " &_
				"		having max(id)=1 " &_
				"	) as T " &_
				"	left join db_partner.dbo.tbl_menu_part as M " &_
				"		on M.menu_id=" & parentid &_
				"			and M.part_sn=T.part_sn " &_
				"where (M.level_sn is null or M.level_sn<T.level_sn)"
		dbget.Execute(SQL)

    end if

    '// 종속메뉴 권한 저장 //
    if childYN="Y" then
    	if child_sn<>"" then
    		splCsn = Split(child_sn, ",")
    		for lp2=0 to Ubound(splCsn)
    			SQL = "Delete From db_partner.dbo.tbl_menu_part Where menu_id=" & splCsn(lp2)
    			dbget.Execute(SQL)
    			for lp1=0 to Ubound(splPsn)
    				IF Trim(splPsn(lp1))<>"" and Trim(splLsn(lp1))<>"" THEN
    				SQL =	"Insert Into db_partner.dbo.tbl_menu_part (menu_id, part_sn, level_sn) " &_
    						"Values (" & splCsn(lp2) & ", " & splPsn(lp1) & ", " & splLsn(lp1) & ")"
    				dbget.Execute(SQL)
    				END IF
    			next
    		next
    	end if
    end if
end if

Set opart = new CTenByTenMember
	opart.FPagesize = 500
	opart.FCurrPage = 1
	opart.frectmenuid=menuid
	opart.getpartner_menu_part()

currpartstr=""
if opart.FResultCount>0 then
	for i=0 to opart.FResultCount - 1
		currpartstr=currpartstr & "부서번호" & opart.FitemList(i).fpart_sn & "등급번호" & opart.FitemList(i).flevel_sn
	next
end if

set opart=nothing

if cstr(orgpartstr)<>cstr(currpartstr) then
	if mode="modi" or mode="popmodi" then
		logmsg = logmsg & "/지정권한: "& orgpartstr &" -> "& currpartstr &""
		if left(logmsg,8)<>"메뉴권한수정" then logmsg="메뉴권한수정"&logmsg
	else
		logmsg = logmsg & "/지정권한: "& currpartstr &""
	end if
end if

if logmsg<>"" then
	' 메뉴수정로그저장		' 2021.10.19 한용민 생성
	SQL="insert into db_partner.dbo.tbl_partner_menu_log (" & vbcrlf
	SQL=SQL& " menuid,logtype,logmsg,isusing,adminid,regdate)" & vbcrlf
	SQL=SQL& " 		select" & vbcrlf
	SQL=SQL& " 		"& menuid &",N'"& logtype &"',N'"& logmsg &"','Y',N'"& adminid &"',getdate()" & vbcrlf

	dbget.Execute SQL
end if

'오류검사 및 실행
If Err.Number = 0 Then
	dbget.CommitTrans				'커밋(정상)

	if mode="popmodi" or mode="modi" then
		response.write	"<script language='javascript'>" &_
						"	alert('" & strMsg & "');" &_
						"	opener.location.reload();" &_
						"	self.close();" &_
						"</script>"
	else
		response.write	"<script language='javascript'>" &_
						"	alert('" & strMsg & "');" &_
						"	self.location='menu_list.asp" & param & "';" &_
						"</script>"
	end if
Else
    dbget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
