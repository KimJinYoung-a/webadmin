<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ�
' History : 2010.05.12 �ѿ�� ����
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

	lec_idx = RequestCheckvar(request.Form("lec_idx"),10)					'���¹�ȣ
	lecOption = RequestCheckvar(request.Form("lecOption"),4)				'���¿ɼ�
	lec_title = html2db(request.Form("lec_title"))							'���¸�
	MakerId = RequestCheckvar(request.Form("makerId"),32)					'����ID

	matinclude_yn = RequestCheckvar(request.Form("matinclude_yn"),1)		'�������Կ���(���� = C)
	mat_cost = RequestCheckvar(request.Form("mat_cost"),10)					'����
	mat_buying_cost = RequestCheckvar(request.Form("mat_buying_cost"),10)	'������԰�

	'���� ���Խ� ������+����, �̿� �����Ḹ
	sellprice = RequestCheckvar(request.Form("sellprice"),10)
	'���� ���Խ� ��������԰�+������԰�, �̿� ��������԰���
	buycash = RequestCheckvar(request.Form("buycash"),10)
	'�հ�
	itemsubtotalsum = RequestCheckvar(request.Form("itemsubtotalsum"),10)

	mileage = RequestCheckvar(request.Form("mileage"),10)					'���ϸ���
	itemea = RequestCheckvar(request.Form("itemea"),10)						'�����ο�
	buy_userid = RequestCheckvar(request.Form("buy_userid"),32)				'�ֹ��� ���̵�
	buy_name = Left(html2db(request.Form("buy_name")),16)					'�ֹ��� �̸�
	buy_phone = RequestCheckvar(request.Form("buy_phone1"),4) & "-" & RequestCheckvar(request.Form("buy_phone2"),4) & "-" & RequestCheckvar(request.Form("buy_phone3"),4)	'�ֹ��� ��ȭ��ȣ
	buy_hp = RequestCheckvar(request.Form("buy_hp1"),4) & "-" & RequestCheckvar(request.Form("buy_hp2"),4) & "-" & RequestCheckvar(request.Form("buy_hp3"),4)	'�ֹ��� �޴���
	buy_email = html2db(request.Form("buy_email"))							'�ֹ��� �̸���
	buy_level = RequestCheckvar(request.Form("buy_level"),10)				'�ֹ��� ȸ�����
	entryname(1) = Left(html2db(request.Form("entryname1")),32)				'������#2 �̸�
	entry_hp(1) = RequestCheckvar(request.Form("entry1_hp1"),4) & "-" & RequestCheckvar(request.Form("entry1_hp2"),4) & "-" & RequestCheckvar(request.Form("entry1_hp3"),4)	'������#1 ����ó
	entryname(2) = Left(html2db(request.Form("entryname2")),32)				'������#3 �̸�
	entry_hp(2) = RequestCheckvar(request.Form("entry2_hp1"),4) & "-" & RequestCheckvar(request.Form("entry2_hp2"),4) & "-" & RequestCheckvar(request.Form("entry2_hp3"),4)	'������#2 ����ó
	entryname(3) = Left(html2db(request.Form("entryname3")),32)				'������#4 �̸�
	entry_hp(3) = RequestCheckvar(request.Form("entry3_hp1"),4) & "-" & RequestCheckvar(request.Form("entry3_hp2"),4) & "-" & RequestCheckvar(request.Form("entry3_hp3"),4)	'������#3 ����ó
	paymethod = RequestCheckvar(request.Form("paymethod"),10)				'�������
	ipkumdiv = "4"										'���� ���� (�����Ϸ�)
	sitename = RequestCheckvar(request.Form("sitename"),16)					'����Ʈ��
	adminId = Session("ssBctId ")						'������ID
	ref_ip = Left(request.ServerVariables("REMOTE_ADDR"),32)	'���� IP
	mileagegubun = RequestCheckvar(request.Form("mileagegubun"),10)			'���ϸ��� ��������
    
'####### ��ü������ ���� �����߰�.
    weclassyn = RequestCheckvar(request.Form("weclassyn"),1)				'��Ŭ���� ����
    IsWeClass = (weclassyn="Y")
    IF (paymethod="7") THEN ipkumdiv="2"
  	if lec_title <> "" then
		if checkNotValidHTML(lec_title) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
	if buy_email <> "" then
		if checkNotValidHTML(buy_email) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
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
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if vWantStudyPlace <> "" then
	if checkNotValidHTML(vWantStudyPlace) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if mileagegubun = "" then mileagegubun = "ON"
if mileagegubun = "OFF" then mileage = 0

if lecOption="" then lecOption="0000"				'���¿ɼ� �⺻�� ����

'//�⺻���� ������
dim olecture
set olecture = new CLecture
	olecture.FRectIdx = lec_idx
	olecture.FRectLecOpt = lecOption

	if lec_idx<>"" then
		olecture.GetOneLecture
	end if

IF (Not IsWeClass) THEN  '' ��ü ������ ���� üũ ����.
    '/// �����ο��� Ȯ��
    SQL = "Select (limit_count-limit_sold) as remainCnt " &_
    	" From [db_academy].[dbo].tbl_lec_item_option " &_
    	" Where lecIdx=" & lec_idx &_
    	"	and lecOption='" & lecOption & "'"
    rsACADEMYget.Open sql, dbACADEMYget, 1
    if rsACADEMYget.EOF or rsACADEMYget.BOF then
    	response.write	"<script language='javascript'>" &_
    					"	alert('���������� �����ϴ�.');" &_
    					"	self.close();" &_
    					"</script>"
    	rsACADEMYget.Close()
    	dbACADEMYget.Close()
    	response.End
    else
    	if Cint(rsACADEMYget(0))<Cint(itemea) then
    		response.write	"<script language='javascript'>" &_
    						"	alert('��û�Ͻ� ���¿� �����ο��� �����մϴ�.\n\n������ ���� �ο� : " & rsACADEMYget(0) & "��');" &_
    						"	history.back();" &_
    						"</script>"
    		rsACADEMYget.Close()
    		dbACADEMYget.Close()
    		response.End
    	end if
    end if
    rsACADEMYget.Close
END IF

'Ʈ������ ����
dbACADEMYget.beginTrans

'// �⺻���� ����
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


'// �ֹ���ȣ ����( "B" + �ֹ�����[4�ڸ�] + �ֹ��Ϸù�ȣ[5�ڸ�] )
orderserial = "B" + Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),4,256)
orderserial = orderserial & Format00(5,Right(CStr(orderidx),5))


'/// Order_Master ���� ///
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

'/// Order_detail ���� (�����ֹ��� ó�� - ���ֹ��Ǵ� �Ѱ���) ///
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
    '/// �������̺� �� �ɼ����̺��� �ο����� ���� ///
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

'/// �����˻� �� �ݿ� ///
If Err.Number = 0 Then
	dbACADEMYget.CommitTrans				'Ŀ��(����)

	response.write	"<script language='javascript'>" &_
					"	alert('���� ó���Ǿ����ϴ�.');" &_
					"	opener.history.go(0);" &_
					"	self.close();" &_
					"</script>"

Else
    dbACADEMYget.RollBackTrans				'�ѹ�(�����߻���)

	response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"
End If

%>

<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

