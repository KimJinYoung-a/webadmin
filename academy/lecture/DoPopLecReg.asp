<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스
' History : 2010.05.12 한용민 수정
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim lec_idx, lecOption, lec_title, sellprice , itemsubtotalsum, mileage, itemea
dim buy_name, buy_phone, buy_hp, buy_email, buy_userid, buy_level, buycash , entryname(3), entry_hp(3)
dim paymethod, sitename, adminId, ref_ip, ipkumdiv, MakerId , orderidx, orderserial, rndjumunno
dim SQL, j , mileagegubun
dim matinclude_yn, mat_cost, mat_buying_cost
Dim weclassyn, IsWeClass

	lec_idx = RequestCheckvar(request.Form("lec_idx"),10)					'강좌번호
	lecOption = RequestCheckvar(request.Form("lecOption"),4)				'강좌옵션
	lec_title = html2db(request.Form("lec_title"))							'강좌명
	MakerId = RequestCheckvar(request.Form("makerId"),32)					'강사ID

	matinclude_yn = RequestCheckvar(request.Form("matinclude_yn"),1)		'재료비포함여부(포함 = C)
	mat_cost = RequestCheckvar(request.Form("mat_cost"),10)					'재료비
	mat_buying_cost = RequestCheckvar(request.Form("mat_buying_cost"),10)	'재료비매입가

	'재료비 포함시 수강료+재료비, 이외 수걍료만
	sellprice = RequestCheckvar(request.Form("sellprice"),10)
	'재료비 포함시 수강료매입가+재료비매입가, 이외 수걍료매입가만
	buycash = RequestCheckvar(request.Form("buycash"),10)
	'합계
	itemsubtotalsum = RequestCheckvar(request.Form("itemsubtotalsum"),10)

	mileage = RequestCheckvar(request.Form("mileage"),10)					'마일리지
	itemea = RequestCheckvar(request.Form("itemea"),10)						'수강인원
	buy_userid = RequestCheckvar(request.Form("buy_userid"),32)				'주문자 아이디
	buy_name = Left(html2db(request.Form("buy_name")),16)					'주문자 이름
	buy_phone = RequestCheckvar(request.Form("buy_phone1"),4) & "-" & RequestCheckvar(request.Form("buy_phone2"),4) & "-" & RequestCheckvar(request.Form("buy_phone3"),4)	'주문자 전화번호
	buy_hp = RequestCheckvar(request.Form("buy_hp1"),4) & "-" & RequestCheckvar(request.Form("buy_hp2"),4) & "-" & RequestCheckvar(request.Form("buy_hp3"),4)	'주문자 휴대폰
	buy_email = html2db(request.Form("buy_email"))							'주문자 이메일
	buy_level = RequestCheckvar(request.Form("buy_level"),10)				'주문자 회원등급
	entryname(1) = Left(html2db(request.Form("entryname1")),32)				'수강자#2 이름
	entry_hp(1) = RequestCheckvar(request.Form("entry1_hp1"),4) & "-" & RequestCheckvar(request.Form("entry1_hp2"),4) & "-" & RequestCheckvar(request.Form("entry1_hp3"),4)	'수강자#1 연락처
	entryname(2) = Left(html2db(request.Form("entryname2")),32)				'수강자#3 이름
	entry_hp(2) = RequestCheckvar(request.Form("entry2_hp1"),4) & "-" & RequestCheckvar(request.Form("entry2_hp2"),4) & "-" & RequestCheckvar(request.Form("entry2_hp3"),4)	'수강자#2 연락처
	entryname(3) = Left(html2db(request.Form("entryname3")),32)				'수강자#4 이름
	entry_hp(3) = RequestCheckvar(request.Form("entry3_hp1"),4) & "-" & RequestCheckvar(request.Form("entry3_hp2"),4) & "-" & RequestCheckvar(request.Form("entry3_hp3"),4)	'수강자#3 연락처
	paymethod = RequestCheckvar(request.Form("paymethod"),10)				'결제방법
	ipkumdiv = "4"										'결제 상태 (결제완료)
	sitename = RequestCheckvar(request.Form("sitename"),16)					'사이트명
	adminId = Session("ssBctId ")						'관리자ID
	ref_ip = Left(request.ServerVariables("REMOTE_ADDR"),32)	'접속 IP
	mileagegubun = RequestCheckvar(request.Form("mileagegubun"),10)			'마일리지 적립여부
    
'####### 단체수강에 따른 변수추가.
    weclassyn = RequestCheckvar(request.Form("weclassyn"),1)				'위클래스 여부
    IsWeClass = (weclassyn="Y")
    IF (paymethod="7") THEN ipkumdiv="2"
  	if lec_title <> "" then
		if checkNotValidHTML(lec_title) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
	if buy_email <> "" then
		if checkNotValidHTML(buy_email) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If

Dim vWantStudyName, vWantStudyYear, vWantStudyMonth, vWantStudyDay, vWantStudyAmPm, vWantStudyHour, vWantStudyMin, vWantStudyPlace, vWantStudyWho
vWantStudyName	= Trim(request.Form("wantstudyName"))
vWantStudyName = Replace(vWantStudyName,chr(34),"")
vWantStudyName = Replace(vWantStudyName,"'","")
vWantStudyName = Replace(vWantStudyName,chr(34),"")
vWantStudyYear	= Trim(RequestCheckvar(request.Form("wantstudyYear"),4))
vWantStudyMonth	= Trim(RequestCheckvar(request.Form("wantstudyMonth"),2))
vWantStudyDay	= Trim(RequestCheckvar(request.Form("wantstudyDay"),2))
vWantStudyAmPm	= Trim(RequestCheckvar(request.Form("wantstudyAmPm"),4))
vWantStudyHour	= Trim(RequestCheckvar(request.Form("wantstudyHour"),2))
vWantStudyMin	= Trim(RequestCheckvar(request.Form("wantstudyMin"),2))
vWantStudyPlace	= Trim(request.Form("wantstudyPlace"))
vWantStudyPlace = Replace(vWantStudyPlace,"'","")
vWantStudyPlace = Replace(vWantStudyPlace,chr(34),"")
vWantStudyWho	= Trim(RequestCheckvar(request.Form("wantstudyWho"),6))
if vWantStudyName <> "" then
	if checkNotValidHTML(vWantStudyName) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if vWantStudyPlace <> "" then
	if checkNotValidHTML(vWantStudyPlace) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if mileagegubun = "" then mileagegubun = "ON"
if mileagegubun = "OFF" then mileage = 0

if lecOption="" then lecOption="0000"				'강좌옵션 기본값 지정

'//기본정보 가져옴
dim olecture
set olecture = new CLecture
	olecture.FRectIdx = lec_idx
	olecture.FRectLecOpt = lecOption

	if lec_idx<>"" then
		olecture.GetOneLecture
	end if

IF (Not IsWeClass) THEN  '' 단체 수강은 한정 체크 안함.
    '/// 수강인원수 확인
    SQL = "Select (limit_count-limit_sold) as remainCnt " &_
    	" From [db_academy].[dbo].tbl_lec_item_option " &_
    	" Where lecIdx=" & lec_idx &_
    	"	and lecOption='" & lecOption & "'"
    rsACADEMYget.Open sql, dbACADEMYget, 1
    if rsACADEMYget.EOF or rsACADEMYget.BOF then
    	response.write	"<script language='javascript'>" &_
    					"	alert('강좌정보가 없습니다.');" &_
    					"	self.close();" &_
    					"</script>"
    	rsACADEMYget.Close()
    	dbACADEMYget.Close()
    	response.End
    else
    	if Cint(rsACADEMYget(0))<Cint(itemea) then
    		response.write	"<script language='javascript'>" &_
    						"	alert('신청하신 강좌에 남은인원이 부족합니다.\n\n※현재 남은 인원 : " & rsACADEMYget(0) & "명');" &_
    						"	history.back();" &_
    						"</script>"
    		rsACADEMYget.Close()
    		dbACADEMYget.Close()
    		response.End
    	end if
    end if
    rsACADEMYget.Close
END IF

'트랜젝션 시작
dbACADEMYget.beginTrans

'// 기본정보 저장
Randomize
rndjumunno = CLng(Rnd * 100000) + 1
rndjumunno = CStr(rndjumunno)

SQL =	" Insert into db_academy.dbo.tbl_academy_order_master " &_
		"	(orderserial, jumundiv, userid, accountdiv, regdate, sitename, referip) " &_
		" Values " &_
		"	('" & rndjumunno & "'" &_
		"	,'8', '" & buy_userid & "'" &_
		"	,'" & paymethod & "', getdate() " &_
		"	,'" & sitename & "'" &_
		"	,'" & ref_ip & "')"
dbACADEMYget.Execute(SQL)

SQL = "Select @@identity "
rsACADEMYget.Open sql, dbACADEMYget, 1
	orderidx = rsACADEMYget(0)
rsACADEMYget.Close


'// 주문번호 생성( "B" + 주문일자[4자리] + 주문일련번호[5자리] )
orderserial = "B" + Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),4,256)
orderserial = orderserial & Format00(5,Right(CStr(orderidx),5))


'/// Order_Master 저장 ///
SQL =	" Update db_academy.dbo.tbl_academy_order_master Set " & VbCRLF
SQL =SQL& " 	  orderserial = '" & orderserial & "'" & VbCRLF
SQL =SQL& " 	, accountname = '" & buy_name & "'" & VbCRLF
SQL =SQL& " 	, totalitemno = " & itemea & VbCRLF
SQL =SQL& " 	, totalmileage = " & mileage & VbCRLF
SQL =SQL& " 	, totalsum = " & itemsubtotalsum & VbCRLF
SQL =SQL& " 	, subtotalprice = " & itemsubtotalsum & VbCRLF
SQL =SQL& " 	, ipkumdiv = '" & ipkumdiv & "'" & VbCRLF
SQL =SQL& " 	, cancelyn = 'N'" & VbCRLF
SQL =SQL& " 	, buyname = '" & buy_name & "'" & VbCRLF
SQL =SQL& " 	, buyphone = '" & buy_phone & "'" & VbCRLF
SQL =SQL& " 	, buyhp = '" & buy_hp & "'" & VbCRLF
SQL =SQL& " 	, buyemail = '" & buy_email & "'" & VbCRLF
SQL =SQL& " 	, reqname = '" & buy_name & "'" & VbCRLF
SQL =SQL& " 	, reqphone = '" & buy_phone & "'" & VbCRLF
SQL =SQL& " 	, reqhp = '" & buy_hp & "'" & VbCRLF
SQL =SQL& " 	, reqemail = '" & buy_email & "'" & VbCRLF
SQL =SQL& " 	, userlevel = " & buy_level & VbCRLF
SQL =SQL& " 	, goodsnames = '" & lec_title & "'" & VbCRLF
IF (ipkumdiv="4") THEN
SQL =SQL& " 	, ipkumdate = getdate()" & VbCRLF
END IF
SQL =SQL& " Where idx = " & orderidx
		'response.write sql
dbACADEMYget.Execute(SQL)


Dim iloopNo , iItemNo
IF (IsWeClass) then
    iloopNo=1
    iItemNo=itemea
ELSE
    iloopNo=itemea
    iItemNo=1
ENd IF

'/// Order_detail 저장 (강좌주문건 처리 - 한주문건당 한강좌) ///
for j=0 to iloopNo-1
	SQL =	" Insert into [db_academy].[dbo].tbl_academy_order_detail " &_
			"		( masteridx, orderserial, oitemdiv, itemid, itemoption " &_
			"		, makerid, itemno, itemcost, buycash, itemname, itemoptionname " &_
			"		, entryname, entryhp, vatinclude, mileage,isupchebeasong, issailitem, matcostAdded, matbuycashAdded, matinclude_yn, reducedprice, couponNotAsigncost,weClassYn) " &_
			" Values " &_
			"		( " & orderidx &_
			"		, '" & orderserial & "'" &_
			"		, '90', " & lec_idx &_
			"		, '" & lecOption & "' " &_
			"		, '" & MakerId & "'" &_
			"		, "&iItemNo&", " & sellprice & ", " & buycash &_
			"		, '" & Left(html2db(lec_title),64) & "', '"&olecture.FOneItem.Flecoptionname&"'"

	if j=0 then
		SQL = SQL & " , '" & buy_name & "'" &_
					" , '" & buy_hp & "'"
	else
		SQL = SQL & " , '" & entryname(j) & "'" &_
					" , '" & entry_hp(j) & "'"
	end if

	SQL = SQL &	" , '', " & Mileage &_
				" , 'Y', 'N', " & mat_cost & ", " & mat_buying_cost & ", '" & matinclude_yn & "', " & sellprice & ", " & sellprice 
	IF (IsWeClass) then
	    SQL = SQL &	" ,'Y'"	
	ELSE
	    SQL = SQL &	" ,NULL"
    END IF
	SQL = SQL &	 ")"
	dbACADEMYget.Execute(SQL)
next

IF (IsWeClass) then
	SQL = "INSERT INTO [db_academy].[dbo].[tbl_academy_order_weclass](orderserial, wantstudyName, wantstudyYear, wantstudyMonth, wantstudyDay, " & _
			 "		wantstudyAmPm, wantstudyHour, wantstudyMin, wantstudyPlace, wantstudyWho) " & _
			 "VALUES('" & orderserial & "', '" & vWantStudyName & "', '" & vWantStudyYear & "', '" & vWantStudyMonth & "', '" & vWantStudyDay & "', " & _
			 "		'" & vWantStudyAmPm & "', '" & vWantStudyHour & "', '" & vWantStudyMin & "', '" & vWantStudyPlace & "', '" & vWantStudyWho & "')"
	dbACADEMYget.Execute SQL
ELSE
    '/// 강좌테이블 및 옵션테이블의 인원정보 수정 ///
    SQL = "Update [db_academy].[dbo].tbl_lec_item " &_
    	" Set limit_sold = limit_sold + " & itemea &_
    	" Where idx=" & lec_idx
    dbACADEMYget.Execute(SQL)
    
    SQL = "Update [db_academy].[dbo].tbl_lec_item_option " &_
    	" Set limit_sold = limit_sold + " & itemea &_
    	" Where lecIdx=" & lec_idx &_
    	"	and lecOption='" & lecOption & "'"
    dbACADEMYget.Execute(SQL)
END IF

'/// 오류검사 및 반영 ///
If Err.Number = 0 Then
	dbACADEMYget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('저장 처리되었습니다.');" &_
					"	opener.history.go(0);" &_
					"	self.close();" &_
					"</script>"

Else
    dbACADEMYget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"
End If

%>

<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

