<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_event_winner.asp
' Description :  �̺�Ʈ ��÷���
' History : 2007.02.22 ������ ����
'           2009.04.14 �ѿ�� ����
'           2009.08.06 ������ SMS/�̸��� �߼� �߰�
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/mailLib.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->

<!-- #include virtual="/lib/util/scm_myalarmlib.asp" -->

<% '��������,���ٳ���,�����Ͽ콺,�ΰŽ�,��Ŭ���ڵ�, �����μ�,��ȭ�̺�Ʈ

'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim eMode, eCode, egKindCode, ekind ,ename, eKindName, esday, eeday, estate, epday, sType
Dim cEvtCont, strSql, tmpCode, j, iranking, srankname, sgiftname, arrwinner, itemid, stitle, gcd,rg
Dim cvalue, ctype, mprice, csdate, cedate, tlist, cprice, iErrcnt,iSuccnt, iGiftKindCode, dAStartDate , dAEndDate
Dim iEPCode, sGiveWinner, chkSms, smsCont, chkEmail, emailCont, itemuse_sdate, itemuse_edate, usewrite_sdate
Dim usewrite_edate, return_yn, return_date, itemuse_itemid, itemuse, vUploadType, vTempArr, vDelUser, vRemainCount, vSongJangID

'' ��۱��� ������ �߰�
Dim isupchebeasong, makerid, reqdeliverdate, PrizeCount
Dim jungsan, jungsanValue, vChangeContents, vSCMChangeSQL

'// MY�˸�
dim myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL

eMode 		= Request.Form("mode") 	'������ ó������
eCode  		= Request.Form("eC")	'�̺�Ʈ�ڵ�
egKindCode 	= Request.Form("egKC")	'�������׷��ڵ�(�ΰŽ�/��ȭ�̺�Ʈ ȸ��)
sType 		= "1" '���Ľ����̼� ����
vUploadType	= Request.Form("uploadtype")
PrizeCount	= Request.Form("prizecnt")
If vUploadType = "" Then
	vUploadType = "direct"
End If
if egKindCode = "" then egKindCode = 0

IF eCode = 4 THEN
	strSql= " SELECT evt_name "&_
			" FROM [db_culture_station].[dbo].[tbl_culturestation_event] "&_
			" WHERE evt_code ="&egKindCode
	rsget.Open strSql, dbget
	IF not (rsget.EOF or rsget.BOF) THEN
		ekind = eCode
		ename = db2html(rsget("evt_name"))
	END IF
	rsget.close
ELSE
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
END IF


'--------------------------------------------------------
' ������ ó��  : �̺�Ʈ��÷ ���̺�, ���, ����, �������̺�(����,�ΰŽ�,����)
'--------------------------------------------------------

   '�⺻
	iranking 	= Request.Form("sR")
	srankname 	= html2db(Request.Form("sRN"))

	Dim reqArr, defaltminus, intLoop, ExceptionUser
	If Request.Form("sW") <> "" Then
		reqArr = Request.Form("sW")
		If Right(reqArr,1) = "," Then
			reqArr=left(reqArr,len(reqArr)-1)
		End If
		
		reqArr = Split(reqArr,",")
		PrizeCount = PrizeCount - (ubound(reqArr) + 1)
	Else
		PrizeCount=PrizeCount
	End If

	ExceptionUser = "," & replace(Request.Form("sW"),",","','") & "'"

	'��÷�� �ڵ� ���
	Dim ArrPrizeUser, PrizeUsers
	PrizeUsers=""
	strSQL = "exec [db_culture_station].[dbo].[usp_WWW_CultureStation_AutoPrize_Add] " & PrizeCount & "," & eCode & ",'" & ExceptionUser & "'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF) then
		ArrPrizeUser = rsget.getRows
	end if
	rsget.Close

	If isArray(ArrPrizeUser) Then
		For intLoop = 0 To UBound(ArrPrizeUser,2)
			PrizeUsers = PrizeUsers + ArrPrizeUser(0,intLoop) + ","
		Next
		PrizeUsers = left(PrizeUsers,len(PrizeUsers)-1)
	Else
		Response.Write "<script>alert('��÷�ڰ� �����ϴ�.');history.back();</script>"
		dbget.close()
		Response.End
	End If

	'2013-01-16 ������...��÷��ID�� �������� ","�� ���� ���� �������� �̵��ϰ� ����
	If Right(PrizeUsers,1) = "," Then
		Call sbAlertMsg ("��÷���� �Ǹ����� ,�� ���� �ٽ��Է��ϼ���.\n������� aaa,bbb, -> aaa,bbb", "back", "")
	End If
	'2013-01-16 ������...��÷��ID�� �������� ","�� ���� ���� �������� �̵��ϰ� ���� ��

	arrwinner 	= split(PrizeUsers,",")
	dAStartDate = left(now(),10)
	dAEndDate 	= dateadd("d",14,date())
	If sType = "1" Then
		stitle = html2db(eName & "- " & Trim(Replace(srankname,"��÷","")) & " ��÷") '//OnlyView �� �� �̺�Ʈ��+�����Ī
	Else
		stitle 	= html2db(eName&" ��÷") '//�̺�Ʈ��
	End If
	sGiveWinner =  Request.Form("gUserid")

	'���
	gcd = "01" '//�̺�Ʈ:01, ��Ÿ:90
	rg = request("rdgubun") '//���������
	iGiftKindCode	= Request.Form("iGK")
	sgiftname		= Request.Form("sGKN")	'//����ǰ��
	isupchebeasong  =  Request("isupchebeasong")

	jungsan            = request("jungsan")
	jungsanValue       = request("jungsanValue")

	If jungsan = "" Then
		jungsan = "N"
	Else
		jungsan = "Y"
	End If

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

	'�׽���
	itemuse = Replace(request("itemuse"),"'","")
	If sType = "5" Then
		stitle = itemuse
	End If
	itemuse_sdate = request("itemuse_sdate")
	itemuse_edate = request("itemuse_edate")
	usewrite_sdate = request("usewrite_sdate")
	usewrite_edate = request("usewrite_edate")
	return_yn = request("return_yn")
	return_date = request("return_date")
	itemuse_itemid = request("itemuse_itemid")

	'��÷�� ������
	chkSms = "Y"
	smsCont = "[�ٹ�����] �̺�Ʈ��÷�� �����մϴ�. �������� �� �����ٹ������� Ȯ�����ּ���."
	chkEmail = request("chkEmail")
	emailCont = request("emailCont")

	if (Not IsNumeric(cprice)) then cprice=0
	if (cprice="") then cprice=0


	vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") ��÷�� ���." & vbCrLf
	vChangeContents = vChangeContents & "- ���� = " & sType & ", ��� = " & iranking & ", �����Ī = " & srankname & vbCrLf
	vChangeContents = vChangeContents & "- ��÷Ȯ�αⰣ = " & dAStartDate & " ~ " & dAEndDate & ", ��÷�ڵ�Ϲ�� = " & vUploadType & vbCrLf
	vChangeContents = vChangeContents & "- �������ϱ��� = " & rg & ", ����ǰ�� = " & sgiftname & ", ����û�� = " & reqdeliverdate & vbCrLf
	vChangeContents = vChangeContents & "- ��۱��� = " & isupchebeasong & ", ���꿩�� = " & jungsan & ", ����� = " & jungsanValue & ", ��üID = " & makerid & vbCrLf
	vChangeContents = vChangeContents & "- ����Ÿ�� = " & cvalue & "(" & ctype & "), �ּұ��űݾ� = " & mprice & ", ��ȿ�Ⱓ = " & csdate & " ~ " & cedate & vbCrLf
	vChangeContents = vChangeContents & "- �׽��ͻ�ǰ(" & itemuse_itemid & ") = " & itemuse & ", �׽��ͻ�ǰ���Ⱓ = " & itemuse_sdate & " ~ " & itemuse_edate & vbCrLf
	vChangeContents = vChangeContents & "- ��÷��SMS(" & chkSms & ") = " & smsCont & ", ��÷���̸���(" & chkEmail & ") = " & emailCont & vbCrLf
	vChangeContents = vChangeContents & "- ��÷�� = " & PrizeUsers & vbCrLf

	'####### ��÷�ڸ� ������ ���ε�.
	If vUploadType = "excel" Then
		strSql = "SELECT Top 100 userid FROM [db_temp].[dbo].[tbl_event_winner_excel] WHERE evt_code = '" & eCode & "'"
		rsget.Open strSql,dbget
		If Not rsget.eof Then
			j = 0
			arrwinner = ""
			Do Until rsget.Eof
				vTempArr = vTempArr & rsget("userid")
				vDelUser = vDelUser & "'" & rsget("userid") & "'"
				
				j = j + 1				
				If rsget.RecordCount <> j Then
					vTempArr = vTempArr & ","
					vDelUser = vDelUser & ","
				End If
				rsget.MoveNext
			Loop
			arrwinner = split(vTempArr,",")
		Else
			Response.Write "<script>alert('���� ���ε�� ��÷�ڰ� �����ϴ�.');parent.location.reload();</script>"
			dbget.close()
			Response.End
		End If
		rsget.close
	End If
	
	'Ʈ�����
	dbget.beginTrans
	IF eMode = "C" THEN
		iEPCode		= request("epC")
		sGiveWinner	= request("gUserid")
	
		'���� ��÷�� ���� ����, ���ο� ��÷�� insert
		strSql = "UPDATE  [db_event].[dbo].[tbl_event_prize] SET evtprize_status =6 "&_
				"WHERE evtprize_code= "&iEPCode&" AND evt_winner='"&sGiveWinner&"'"
		dbget.execute strSql
	
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
		END IF
	
	END IF
	
	'2013-01-09 ������ ����	(���� ��÷��ID�� �츮 ���̺� �ִ��� �˻�// ������ ����, ������ ƨ��
	Dim oo
	For oo = 0 to UBound(arrwinner)
		strSql = ""
		strSql = strSql & " SELECT count(*) as cnt FROM db_user.dbo.tbl_user_n where userid='"&html2db(Trim(arrwinner(oo)))&"' " & VBCRLF
		rsget.Open strSql,dbget
			If rsget("cnt") = 0 Then
				rsget.close
				Call sbAlertMsg ("ID : "&html2db(Trim(arrwinner(oo)))&"�� �����ϴ�. �ٽ� �Է��ϼ���", "back", "")
				dbget.close()	:	response.End
			End If
		rsget.Close
	Next
	'2013-01-09 ������ ���� ��
	
		iErrcnt = 0
	For j = 0 To UBound(arrwinner)
		SELECT CASE eKind
		Case "2" '���ٳ���(�̹��� ������ Ȯ��, 5�־ȿ� ������ ������� ����Ȯ��) ### �Ⱦ�. ����Ϸ��� �� �Ʒ� �ּ��κ� ó������ �ٿ��ֱ�.
		Case "3" '100% shop
		'Case "5" '�����Ͽ콺
		Case "8" '�������ΰŽ�
		Case Else
	
			tmpCode = ""
			vSongJangID = ""
			'1. �̺�Ʈ���� ���
			fnSetEventPrize sType,eCode,egKindCode,iranking,srankname,iGiftKindCode,html2db(Trim(arrwinner(j))),dAStartDate,dAEndDate,session("ssBctId"),iEPCode,stitle
	
			'// MY�˸�
			myalarmtitle = "�̺�Ʈ ��÷�� ���ϵ帳�ϴ�!"
			myalarmsubtitle = ename
			if (Len(myalarmsubtitle) > 20) then
				myalarmsubtitle = Left(myalarmsubtitle, 20) & " ..."
			end if
	
			myalarmcontents = "�̺�Ʈ ��÷�ҽ��� �˷��帳�ϴ�."
			myalarmwwwTargetURL = "/my10x10/myeventmaster.asp"
	
			Call MyAlarm_InsertMyAlarm_SCM(html2db(Trim(arrwinner(j))), "007", myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL)

			'2. ����Ǵ� ���� ���
			IF  not( tmpCode = ""  or isNull(tmpCode)) tHEN
				IF CStr(sType) = "3"	THEN '����ǰ���
				 	fnSetSongjang rg, gcd, stitle, tmpCode, html2db(Trim(arrwinner(j))), sgiftname ,iGiftKindCode, isupchebeasong, makerid, reqdeliverdate,jungsanValue, jungsan
				ELSEIF  CStr(sType) ="2" THEN '�������
					fnSetUserCoupon html2db(Trim(arrwinner(j))),ctype,cvalue,stitle,mprice,csdate,cedate,tlist,cprice, session("ssBctId"),tmpCode
				ELSEIF  CStr(sType) ="4" THEN 'Ƽ�ϵ��
					fnSetTicket tmpCode, egKindCode, html2db(Trim(arrwinner(j)))
				ELSEIF  CStr(sType) ="5" THEN '�׽����̺�Ʈ
	
					fnSetSongjang rg, gcd, stitle, tmpCode, html2db(Trim(arrwinner(j))), sgiftname ,0, isupchebeasong, makerid, reqdeliverdate,jungsanValue, jungsan
	
					fnSetTester tmpCode, eCode, html2db(Trim(arrwinner(j))), itemuse_itemid, itemuse, itemuse_sdate, itemuse_edate, usewrite_sdate, usewrite_edate, return_yn, return_date
				END IF
			ELSE
				iErrcnt = iErrcnt + 1
			END IF
	
		END Select
	
		'//��÷�� �޽��� �߼�	### DB Ʈ����Ƕ����� ��ü�� �Ǿ� �߼��� �ȳ����⵵ ��. �׷��� �Ʒ��� �ű�. 20151016 ���ر�.
		'if chkSms="Y" or chkEmail="Y" then
		'	Call fnSendUerMessege(html2db(Trim(arrwinner(j))), chkSms, smsCont, chkEmail, emailCont)
		'end if
		
		
		'### �ּҰ� ���� ������� inputdate �� null, evtprize_status �� 0���� �ٲ���.
		If sType = "3" AND rg = "F" Then
			strSql = "IF EXISTS(select id from [db_sitemaster].[dbo].[tbl_etc_songjang] WHERE id = '" & vSongJangID & "' and (reqzipcode = '' or replace(reqzipcode,' ','') = '-')) " & vbCrLf & _
					 "BEGIN " & vbCrLf & _
					 "	UPDATE [db_sitemaster].[dbo].[tbl_etc_songjang] SET inputdate = Null WHERE id = '" & vSongJangID & "' " & vbCrLf & _
					 "	UPDATE [db_event].[dbo].[tbl_event_prize] SET evtprize_status = 0 WHERE evtprize_code = '" & tmpCode & "' " & vbCrLf & _
					 "END "
			dbget.execute strSql
		End If
	Next
	
	If vUploadType = "excel" Then
		strSql = "DELETE FROM [db_temp].[dbo].[tbl_event_winner_excel] WHERE userid IN(" & vDelUser & ") AND evt_code = '" & eCode & "'"
		dbget.execute strSql
		
		strSql = "SELECT count(userid) FROM [db_temp].[dbo].[tbl_event_winner_excel] WHERE evt_code = '" & eCode & "'"
		rsget.Open strSql,dbget
		vRemainCount = rsget(0)
		rsget.close
	End If
	

	IF Err.Number <> 0 THEN
		dbget.RollBackTrans
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[10]", "back", "")
	Else
		dbget.CommitTrans
		
		For j = 0 To UBound(arrwinner)
			'//��÷�� �޽��� �߼�
			if chkSms="Y" or chkEmail="Y" then
				Call fnSendUerMessege(html2db(Trim(arrwinner(j))), chkSms, smsCont, chkEmail, emailCont)
			end if
		Next

		'### ���� �α� ����(event)
		vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & Request("menupos") & "', "
		vSCMChangeSQL = vSCMChangeSQL & "'" & html2db(vChangeContents) & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		dbget.execute(vSCMChangeSQL)
	END IF
	
	If vUploadType = "excel" Then
		Response.Write "<script type=""text/javascript"">" & vbCrLf
		If vRemainCount > 0 Then
			Response.Write "parent.$('#excelprocing').hide();" & vbCrLf
			Response.Write "parent.$('#excelSubmit').show();" & vbCrLf
			Response.Write "parent.$('#excelprocdetail').html('&nbsp;&nbsp;<font color=red size=3>* ���� ó�� ���� : <strong>"&vRemainCount&"</strong></font>&nbsp;&nbsp;');" & vbCrLf
			Response.Write "parent.jsPageReload();" & vbCrLf
		Else
			Response.Write "alert('��ϵǾ����ϴ�.');" & vbCrLf
			Response.Write "parent.jsPageReload();" & vbCrLf
			Response.Write "parent.window.close();" & vbCrLf
		End If
		Response.Write "</script>" & vbCrLf
	Else
%>
	<script language="javascript">
	<!--
	
		alert("��ϵǾ����ϴ�.");
		opener.location.reload();
		window.close();
	//-->
	</script>
<%
	End If
	
	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' �Լ�����
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	'#### �̺�Ʈ ��۵�� #################
	Function fnSetSongjang(ByVal rdgubun, ByVal gubuncd, ByVal gubunname, ByVal evtprize_code, ByVal userid,ByVal prizetitle , ByVal giftkindcode, ByVal isupchebeasong, ByVal makerid, ByVal reqdeliverdate, ByVal jungsanValue, ByVal jungsan)
		if rdgubun="U" then
			strSql = "insert into [db_sitemaster].[dbo].tbl_etc_songjang (gubuncd,gubunname,evtprize_code,userid,prizetitle,evtprize_giftkindcode, isupchebeasong, delivermakerid, reqdeliverdate, jungsan, jungsanYN) "&_
					" values "&_
		 			"  ('" & gubuncd &"','" & gubunname& "',"&evtprize_code &",'"&userid&"','" &prizetitle&"',"&giftkindcode&",'" & isupchebeasong & "','" & makerid & "','" & reqdeliverdate & "','" & jungsanValue & "','" & jungsan & "')"
		elseif rdgubun="F" then
			strSql = "insert into [db_sitemaster].[dbo].tbl_etc_songjang (" & vbcrlf
			strSql = strSql & " gubuncd,gubunname,evtprize_code,userid,username,reqname,reqphone,reqhp,reqzipcode" & vbcrlf
			strSql = strSql & " ,reqaddress1,reqaddress2, inputdate, prizetitle,evtprize_giftkindcode, isupchebeasong" & vbcrlf
			strSql = strSql & " , delivermakerid, reqdeliverdate, jungsan, jungsanYN"
			strSql = strSql & " )" & vbcrlf
			strSql = strSql & " 	select distinct" & vbcrlf
			strSql = strSql & " 	'" & gubuncd & "','" & gubunname & "',"&evtprize_code&", u.userid, u.username, u.username" & vbcrlf
			strSql = strSql & " 	, u.userphone, u.usercell, u.zipcode, u.zipaddr + ' ' + u.useraddr ,u.useraddr, getdate()," & vbcrlf
			strSql = strSql & " 	'" &prizetitle& "',"&giftkindcode&",'" & isupchebeasong & "','" & makerid & "','" & reqdeliverdate & "'" & vbcrlf
			strSql = strSql & " 	,'" & jungsanValue & "','" & jungsan & "'" & vbcrlf
			strSql = strSql & " 	from [db_user].[dbo].tbl_user_n u" & vbcrlf
			strSql = strSql & " 	where u.userid  = '"&userid&"'" & vbcrlf
		end if
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
		END IF
		
		strSql = "select @@IDENTITY " '': �۵�OK
		rsget.Open strSql, dbget
		vSongJangID = rsget(0)
		rsget.Close
	End Function

	'###�������� ��� ###################
	Function fnSetUserCoupon(ByVal userid,ByVal coupontype,ByVal couponvalue,ByVal couponname,ByVal minbuyprice,ByVal startdate,ByVal expiredate,ByVal targetitemlist,ByVal couponmeaipprice,ByVal reguserid, ByVal evtprize_code)
		strSql = "insert into [db_user].[dbo].tbl_user_coupon(masteridx,userid,coupontype,couponvalue,couponname "&_
				 " ,minbuyprice,startdate,expiredate,targetitemlist,couponmeaipprice,reguserid,evtprize_code)"&_
				 " values "&_
				 " (0,'"&userid&"','"&coupontype&"','"&couponvalue&"','"&couponname&"','"&minbuyprice&"',"&_
				 "'"&startdate&"','"&expiredate&"','"&targetitemlist&"',"&couponmeaipprice&",'"&reguserid&"',"&evtprize_code&")"
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
		END IF
	END Function

	'###�̺�Ʈ���� ��� ###################
	Function fnSetEventPrize(ByVal sType, ByVal eCode,ByVal egKindCode, ByVal evt_ranking,ByVal evt_rankname,ByVal iGiftKindCode,ByVal evt_winner,ByVal dAStartDate,ByVal dAEndDate, ByVal AdminID, ByVal iGiveEPCode,ByVal stitle)
		Dim iprizestatus : iprizestatus = 0
		IF 	(dAEndDate = "" OR (sType="3" and rg="F") )THEN iprizestatus = 3 '��÷Ȯ�αⰣ �� �Է½� Ȯ�λ��·�
		IF iGiveEPCode = "" THEN iGiveEPCode = "NULL"
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event_prize] (evtprize_type, [evt_code],evtgroup_code, [evt_ranking], [evt_rankname], giftkind_code, [evt_winner],  [evtprize_startdate], [evtprize_enddate], [evtprize_status],[AdminID],[give_evtprizecode],evtprize_name) "&_
				"	 SELECT "&sType&","&eCode&","&egKindCode&","&evt_ranking&",'"&evt_rankname&"','"&iGiftKindCode&"', userid, '"&dAStartDate &"','"&dAEndDate&"', "&iprizestatus&", '"& AdminID&"',"&iGiveEPCode&",'"&stitle&"'"&_
				"		FROM [db_user].[dbo].[tbl_user_n] WHERE userid = '"&evt_winner&"'"


		dbget.execute strSql

		strSql = ""
		'//�����̺�Ʈ ��÷�ڹ�ǥ �Ϸ� ó�� '// 2009-04-14 �ѿ�� ��÷��ó��
		if eCode = 4 then

			strSql = "update db_culture_station.dbo.tbl_culturestation_event set"+vbcrlf
			strSql = strSql & " prizeyn = 'Y'"+vbcrlf
			strSql = strSql & " where evt_code = "&egKindCode&""+vbcrlf

			'response.write strSql&"<br>"
			dbget.execute strSql

		'//�Ϲ��̺�Ʈ ��÷�ڹ�ǥ �Ϸ� ó�� '// 2009-04-14 �ѿ�� ��÷��ó��
		else
			strSql = "update db_event.dbo.tbl_event set"+vbcrlf
			strSql = strSql & " prizeyn = 'Y'"+vbcrlf
			strSql = strSql & " where evt_code = "&eCode&""+vbcrlf

			'response.write strSql&"<br>"
			dbget.execute strSql
		end if

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
		END IF

		'' SQL 2005������ �۵�����..?
		''strSql = "select SCOPE_IDENTITY() From [db_event].[dbo].[tbl_event_prize] "  '': �۵�����		'/������.��ü ���� ���� �ѷ���. '/2016.06.02 �ѿ��
		'strSql = "select SCOPE_IDENTITY()"		'/���� sql 2005 ���°� ����. ����Ҳ��� �̷� ���·� ����.	'/2016.06.02 �ѿ��
		''strSql = "select IDENT_CURRENT('[db_event].[dbo].[tbl_event_prize]') " '': �۵�OK
		strSql = "select @@IDENTITY " '': �۵�OK

		rsget.Open strSql, dbget
		tmpCode = rsget(0)
		rsget.Close
	End Function

	'###Ƽ�� ��� ###################
	Function fnSetTicket(ByVal evtprize_code, ByVal egKindCode, ByVal evt_winner)
		strSql = "INSERT INTO [db_culture_station].[dbo].[tbl_ticket_prize] ( [evtprize_code], [cul_evt_code], [evt_winner])"&_
				" VALUES ("&evtprize_code&","&egKindCode&",'"&evt_winner&"') "
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
		END IF
	End Function

	'###�׽��� ��� ###################
	Function fnSetTester(ByVal evtprize_code, ByVal evt_Code, ByVal evt_winner, ByVal itemuse_itemid, ByVal itemuse, ByVal itemuse_sdate, ByVal itemuse_edate, ByVal usewrite_sdate, ByVal usewrite_edate, ByVal return_yn, ByVal return_date)
		strSql = "INSERT INTO [db_event].[dbo].[tbl_tester_event_winner] ( [evtprize_code], [evt_code], [evt_winner], [itemid], [itemname], [itemuse_sdate], [itemuse_edate], [usewrite_sdate], [usewrite_edate], [return_yn], [return_date])"&_
				" VALUES ('"&evtprize_code&"','"&evt_Code&"','"&evt_winner&"','"&itemuse_itemid&"','"&itemuse&"','"&itemuse_sdate&"','"&itemuse_edate&"','"&usewrite_sdate&"','"&usewrite_edate&"','"&return_yn&"','"&return_date&"') "
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
		END IF
	End Function

	'### SMS/�̸��� �߼� ###################
	Sub fnSendUerMessege(userid, chks, scont, chke, econt)
		dim uHp, uMail
		strSql = "Select top 1 usercell, usermail " &_
				" From db_user.dbo.tbl_user_n " &_
				" Where userid='" & userid & "'"
		rsget.Open strSql, dbget
		IF Not(rsget.EOF or rsget.BOF) THEN
			uHp = rsget("usercell")
			uMail = rsget("usermail")
		END IF
		rsget.close

		'SMS �߼�
		if chks="Y" then
			if Not(uHP="" or isNull(uHP)) then Call SendNormalSMS_LINK(uHP,"",scont)
		end if

		'eMail �߼�
		if chke="Y" then
			if Not(uMail="" or isNull(uMail)) then
				Call sendmailCS(uMail, "�̺�Ʈ ��÷ �ȳ��Դϴ�.", replace(econt,vbCrLf,"<br>"))
			end if
		end if
	End Sub
	
'####### ���ٳ��� ó������
'		'1.Check : ��������Ȯ��
'		strSql = " SELECT evt_winner FROM  [db_event].[dbo].[tbl_event_prize] WHERE evt_code ="&eCode
'		rsget.Open strSql, dbget
'		IF not (rsget.EOF or rsget.BOF) THEN
'			rsget.close
'			Call sbAlertMsg ("�̹��� ��÷�ڰ� �̹� �����Ǿ����ϴ�", "back", "")
'			dbget.close()	:	response.End
'		END IF
'		rsget.close
'
'		'2.Check : 5�־ȿ� ������ ������� ����Ȯ��
'		strSql = " select evt_winner from  [db_event].[dbo].[tbl_event_prize] where evt_code in ( "&_
'				"	select top 5 evt_code from [db_event].[dbo].[tbl_event] where evt_kind = 2  order by evt_code desc "&_
'				")  and evt_winner = '"&html2db(Trim(arrwinner(j)))&"'"
'		rsget.Open strSql, dbget
'		IF not (rsget.EOF or rsget.BOF) THEN
'			rsget.close
'			Call sbAlertMsg ("5�־ȿ� �ѹ��̻� ��÷�ǽ� ���Դϴ�.", "back", "")
'			dbget.close()	:	response.End
'		END IF
'		rsget.close
'
'		'3.Check : ���ٳ����� ���� �� ������� ����Ȯ��
'		strSql = " select userid from [db_contents].[dbo].[tbl_one_comment]  where  userid='"&html2db(Trim(arrwinner(j)))&"' and evt_code="&eCode
'		rsget.Open strSql, dbget
'		IF (rsget.EOF OR rsget.BOF) THEN
'			rsget.close
'			Call sbAlertMsg ("�̺�Ʈ �����ڰ� �ƴմϴ�.��÷�ڸ� Ȯ�����ּ���", "back", "")
'			dbget.close()	:	response.End
'		END IF
'		rsget.close
'
'		tmpCode = ""
'		'4. �̺�Ʈ���� ���
'		   Call fnSetEventPrize (sType,eCode,egKindCode,iranking,srankname,iGiftKindCode,html2db(Trim(arrwinner(j))),dAStartDate,dAEndDate, session("ssBctId"),iEPCode,stitle)
'
'		   '// MY�˸�
'		   myalarmtitle = "�̺�Ʈ ��÷�� ���ϵ帳�ϴ�!"
'		   myalarmsubtitle = ename
'		   if (Len(myalarmsubtitle) > 20) then
'			   myalarmsubtitle = Left(myalarmsubtitle, 20) & " ..."
'		   end if
'
'		   myalarmcontents = "�̺�Ʈ ��÷�ҽ��� �˷��帳�ϴ�."
'		   myalarmwwwTargetURL = "/my10x10/myeventmaster.asp"
'
'		   Call MyAlarm_InsertMyAlarm_SCM(html2db(Trim(arrwinner(j))), "007", myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL)
'
'		'5. ������
'		IF  not( tmpCode = ""  or isNull(tmpCode)) tHEN
'			fnSetSongjang rg, gcd, stitle, tmpCode, html2db(Trim(arrwinner(j))), sgiftname , iGiftKindCode, isupchebeasong, makerid, reqdeliverdate,jungsanValue, jungsan
'		ELSE
'			iErrcnt = iErrcnt + 1
'		END IF
'
'		'6.���ٳ��� ���
'		strSql = "UPDATE [db_contents].[dbo].[tbl_one_comment] SET winYN='Y' WHERE userid='"&html2db(Trim(arrwinner(j)))&"' and evt_code="&eCode
'		dbget.execute strSql
'		IF Err.Number <> 0 THEN
'			dbget.RollBackTrans
'			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
'		END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
