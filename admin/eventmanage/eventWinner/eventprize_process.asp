<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��÷��
' History : 2009.04.17 ������ ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<% '��������,���ٳ���,�����Ͽ콺,�ΰŽ�,��Ŭ���ڵ�, �����μ�,��ȭ�̺�Ʈ

'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim eMode, eCode,ekind ,ename, eKindName, esday, eeday, estate, epday, sType
Dim cEvtCont, strSql, tmpCode, j
Dim iranking, srankname, sgiftname, arrwinner, itemid, stitle
dim gcd,rg
Dim cvalue, ctype, mprice, csdate, cedate, tlist, cprice
Dim iErrcnt,iSuccnt

'' ��۱��� ������ �߰�
Dim isupchebeasong, makerid, reqdeliverdate

eMode = Request.Form("mode") '������ ó������
eCode  = Request.Form("eC")	'�̺�Ʈ�ڵ�
sType =  Request.Form("selType")

set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�

	cEvtCont.fnGetEventCont	 '�̺�Ʈ ���� ��������
	ekind =	cEvtCont.FEKind
	ename =	db2html(cEvtCont.FEName)
	eKindName = fnGetEventCodeDesc("eventkind",ekind)
	esday =	cEvtCont.FESDay
	eeday =	cEvtCont.FEEDay
	epday =	cEvtCont.FEPDay
	estate =	cEvtCont.FEState
set cEvtCont = nothing

'--------------------------------------------------------
' �Լ�����
'--------------------------------------------------------
	'#### �̺�Ʈ ��۵�� #################
	Function fnSetSongjang(ByVal rdgubun, ByVal gubuncd, ByVal gubunname, ByVal evtprize_code, ByVal userid,ByVal prizetitle , ByVal itemid, ByVal isupchebeasong, ByVal makerid, ByVal reqdeliverdate)
		if rdgubun="U" then
			strSql = "insert into [db_contents].[dbo].tbl_etc_songjang (gubuncd,gubunname,evtprize_code,userid,prizetitle, isupchebeasong, delivermakerid, reqdeliverdate) "&_
					" values "&_
		 			"  ('" & gubuncd &"','" & gubunname& "',"&evtprize_code &",'"&userid&"','" &prizetitle & "(" & itemid & ")" &"','" & isupchebeasong & "','" & makerid & "','" & reqdeliverdate & "')"
		elseif rdgubun="F" then
			strSql = "insert into [db_contents].[dbo].tbl_etc_songjang (" & vbcrlf
			strSql = strSql & " gubuncd,gubunname,evtprize_code,userid,username,reqname,reqphone,reqhp,reqzipcode,reqaddress1" & vbcrlf
			strSql = strSql & " ,reqaddress2, inputdate, prizetitle, isupchebeasong, delivermakerid, reqdeliverdate" & vbcrlf
			strSql = strSql & " )" & vbcrlf
			strSql = strSql & " 	select distinct '" & gubuncd & "','" & gubunname & "',"&evtprize_code&", u.userid, u.username" & vbcrlf
			strSql = strSql & " 	, u.username, u.userphone, u.usercell, u.zipcode, z.Addr050_si + ' ' + z.Addr050_gu ,u.useraddr" & vbcrlf
			strSql = strSql & " 	, getdate(),'" &prizetitle & "(" & itemid & ")" & "','" & isupchebeasong & "','" & makerid & "'" & vbcrlf
			strSql = strSql & " 	,'" & reqdeliverdate & "'" & vbcrlf
			strSql = strSql & " 	from [db_user].[10x10].tbl_user_n u" & vbcrlf
			strSql = strSql & " 	where u.userid  = '"&userid&"'" & vbcrlf
		end if

		dbget.execute strSql
	End Function

	'###�������� ��� ###################
	Function fnSetUserCoupon(ByVal userid,ByVal coupontype,ByVal couponvalue,ByVal couponname,ByVal minbuyprice,ByVal startdate,ByVal expiredate,ByVal targetitemlist,ByVal couponmeaipprice,ByVal reguserid, ByVal evtprize_code)
		strSql = "insert into [db_user].[10x10].tbl_user_coupon(masteridx,userid,coupontype,couponvalue,couponname "&_
				 " ,minbuyprice,startdate,expiredate,targetitemlist,couponmeaipprice,reguserid,evtprize_code)"&_
				 " values "&_
				 " (0,'"&userid&"','"&coupontype&"','"&couponvalue&"','"&couponname&"','"&minbuyprice&"',"&_
				 "'"&startdate&"','"&expiredate&"','"&targetitemlist&"',"&couponmeaipprice&",'"&reguserid&"',"&evtprize_code&")"
		dbget.execute strSql
	END Function

	'###�̺�Ʈ����, �α� ��� ###################
	Function fnSetEventPrize(ByVal eCode,ByVal evt_ranking,ByVal evt_rankname,ByVal itemid,ByVal evt_giftname,ByVal evt_winner,ByVal AdminID)
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event_prize] ([evt_code], [evt_ranking], [evt_rankname], itemid, [evt_giftname], [evt_winner],  [AdminID]) "&_
				"	 SELECT "&eCode&","&evt_ranking&",'"&evt_rankname&"','"&itemid&"', '"&evt_giftname&"',userid,'"& AdminID&"'"&_
				"		FROM [db_user].[10x10].[tbl_user_n] WHERE userid = '"&evt_winner&"'"
		dbget.execute strSql

		'' SQL 2005������ �۵�����..?
		''strSql = "select SCOPE_IDENTITY() "  '': �۵�����
		''strSql = "select IDENT_CURRENT('[db_event].[dbo].[tbl_event_prize]') " '': �۵�OK
		strSql = "select @@IDENTITY " '': �۵�OK

		rsget.Open strSql, dbget
		tmpCode = rsget(0)
		rsget.Close

		'###�̺�Ʈ �α� ���###################
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event_managelog] ([evt_code], [evtlog_cont], [adminid], [evtlog_regip]) "&_
		 "		VALUES ("&eCode&",'��÷�ڵ��','"& AdminID&"','"&Request.ServerVariables("REMOTE_ADDR")&"')"
		 dbget.execute strSql
	End Function

	Function fnSetEventWinnerLog(ByVal eCode,ByVal evt_ranking,ByVal evt_winner)
		strSQL =" UPDATE [db_event].dbo.tbl_event_winner_log " &_
				" SET rank ='" & evt_ranking & "' " &_
				" WHERE evt_code='" & eCode & "' " &_
				" AND userid='" & evt_winner &"' "
		 dbget.execute strSql
	End Function


'--------------------------------------------------------
' ������ ó��  : �̺�Ʈ��÷ ���̺�, ���, ����, �������̺�(����,�ΰŽ�,����)
'--------------------------------------------------------
   '�⺻
	iranking = Request.Form("sR")
	srankname = html2db(Request.Form("sRN"))
	sgiftname = html2db(Request.Form("sGN"))
	arrwinner = split(Request.Form("sW"),",")
	itemid    = Request.Form("itemid")
	stitle = html2db("["&eKindName&"]"&eName&" ��÷") '//�̺�Ʈ��

	'���
	gcd = "01" '//�̺�Ʈ:01, ��Ÿ:90
	rg = request("rdgubun") '//���������
	isupchebeasong  =  Request("isupchebeasong")
    makerid         =  Request("makerid")
    reqdeliverdate  =  Request("reqdeliverdate")

	'����
	cvalue = request("couponvalue")
	ctype = request("coupontype")
	mprice = request("minbuyprice")
	csdate = request("sDate")&" 00:00:00"
	cedate = request("eDate")&" 23:59:59"
	tlist = request("targetitemlist")
	cprice = request("couponmeaipprice")


	if (Not IsNumeric(cprice)) then cprice=0
	if (cprice="") then cprice=0
	IF CStr(sType) = "2"	THEN sgiftname ="����"
	IF CStr(sType) = "3"	THEN sgiftname =""

	'Ʈ�����
dbget.beginTrans
	iErrcnt = 0
For j = 0 To UBound(arrwinner)
	SELECT CASE eKind
	Case "2" '���ٳ���(�̹��� ������ Ȯ��, 5�־ȿ� ������ ������� ����Ȯ��)

		'1.Check : ��������Ȯ��
		strSql = " SELECT evt_winner FROM  [db_event].[dbo].[tbl_event_prize] WHERE evt_code ="&eCode
		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			rsget.close
			Call sbAlertMsg ("�̹��� ��÷�ڰ� �̹� �����Ǿ����ϴ�", "back", "")
			dbget.close()	:	response.End
		END IF
		rsget.close

		'2.Check : 5�־ȿ� ������ ������� ����Ȯ��
		strSql = " select evt_winner from  [db_event].[dbo].[tbl_event_prize] where evt_code in ( "&_
				"	select top 5 evt_code from [db_event].[dbo].[tbl_event] where evt_kind = 2  order by evt_code desc "&_
				")  and evt_winner = '"&html2db(Trim(arrwinner(j)))&"'"
		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			rsget.close
			Call sbAlertMsg ("5�־ȿ� �ѹ��̻� ��÷�ǽ� ���Դϴ�.", "back", "")
			dbget.close()	:	response.End
		END IF
		rsget.close

		'3.Check : ���ٳ����� ���� �� ������� ����Ȯ��
		strSql = " select userid from [db_contents].[dbo].[tbl_one_comment]  where  userid='"&html2db(Trim(arrwinner(j)))&"' and evt_code="&eCode
		rsget.Open strSql, dbget
		IF (rsget.EOF OR rsget.BOF) THEN
			rsget.close
			Call sbAlertMsg ("�̺�Ʈ �����ڰ� �ƴմϴ�.��÷�ڸ� Ȯ�����ּ���", "back", "")
			dbget.close()	:	response.End
		END IF
		rsget.close

		tmpCode = ""
		'4. �̺�Ʈ���� ���
		   Call fnSetEventPrize (eCode,iranking,srankname,itemid,sgiftname,html2db(Trim(arrwinner(j))),session("ssBctId"))

		'5. ������
		IF  not( tmpCode = ""  or isNull(tmpCode)) tHEN
			fnSetSongjang rg, gcd, stitle, tmpCode, html2db(Trim(arrwinner(j))), sgiftname , itemid, isupchebeasong, makerid, reqdeliverdate
		ELSE
			iErrcnt = iErrcnt + 1
		END IF

		'6.���ٳ��� ���
		strSql = "UPDATE [db_contents].[dbo].[tbl_one_comment] SET winYN='Y' WHERE userid='"&html2db(Trim(arrwinner(j)))&"' and evt_code="&eCode
		dbget.execute strSql

	Case "3" '100% shop
	'Case "5" '�����Ͽ콺
	Case "8" '�������ΰŽ�
	Case Else

		tmpCode = ""
		'1. �̺�Ʈ���� ���
		fnSetEventPrize eCode,iranking,srankname,itemid,sgiftname,html2db(Trim(arrwinner(j))),session("ssBctId")
		fnSetEventWinnerLog eCode,iranking,html2db(Trim(arrwinner(j)))
		'2. ����Ǵ� ���� ���
		IF  not( tmpCode = ""  or isNull(tmpCode)) tHEN
			IF CStr(sType) = "1"	THEN '����ǰ���
			 	fnSetSongjang rg, gcd, stitle, tmpCode, html2db(Trim(arrwinner(j))), sgiftname ,itemid, isupchebeasong, makerid, reqdeliverdate
			ELSEIF  CStr(sType) ="2" THEN '�������
				fnSetUserCoupon html2db(Trim(arrwinner(j))),ctype,cvalue,stitle,mprice,csdate,cedate,tlist,cprice, session("ssBctId"),tmpCode
			END IF
		ELSE
			iErrcnt = iErrcnt + 1
		END IF
	END Select
Next

	IF Err.Number = 0 THEN
		dbget.CommitTrans
%>
		<script language="javascript">
		<!--
			alert("��ϵǾ����ϴ�.");
			opener.location.reload();
			window.close();
		//-->
		</script>
<%
	dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
	END IF

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
