<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ���ⱸ�� ��ǰ
' History : 2016.06.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/standing/item_standing_cls.asp"-->
<%
dim strSql, i, lastuserid, menupos, mode, maxsendkey, identikey, itemgubun, uidx, zipcode, reqzipaddr, useraddr, username
dim itemid, itemoption, sendkey, reserveDlvDate, reserveidx, reserveitemgubun, reserveItemID, reserveItemOption, reserveItemName
dim userphone1, userphone2, userphone3, userphone, usercell1, usercell2, usercell3, usercell, isusing, uidxarr, itemno
dim sendstatuscnt, sendstatus, jukyogubun, orderserial, userid, smsyn, tmpSql, michulgoorder, optsellyn, optisusing, standingorderarr
	lastuserid=session("ssBctId")
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	mode = requestcheckvar(request("mode"),32)
	itemgubun = getNumeric(requestcheckvar(request("itemgubun"),10))
	itemid = getNumeric(requestcheckvar(request("itemid"),10))
	itemoption = requestcheckvar(request("itemoption"),32)
	sendkey = getNumeric(requestcheckvar(request("sendkey"),10))
	reserveidx = getNumeric(requestcheckvar(request("reserveidx"),10))
	uidx = getNumeric(requestcheckvar(request("uidx"),10))
	reqzipaddr = requestcheckvar(request("addr1"),128)
	useraddr = requestcheckvar(request("addr2"),128)
	username = requestcheckvar(request("username"),32)
	userphone1 = getNumeric(requestcheckvar(request("userphone1"),4))
	userphone2 = getNumeric(requestcheckvar(request("userphone2"),4))
	userphone3 = getNumeric(requestcheckvar(request("userphone3"),4))
	usercell1 = getNumeric(requestcheckvar(request("usercell1"),4))
	usercell2 = getNumeric(requestcheckvar(request("usercell2"),4))
	usercell3 = getNumeric(requestcheckvar(request("usercell3"),4))
	isusing = requestcheckvar(request("isusing"),10)
	zipcode = requestcheckvar(request("zipcode"),7)
	itemno = getNumeric(requestcheckvar(request("itemno"),10))
	jukyogubun = requestcheckvar(request("jukyogubun"),16)
	orderserial = getNumeric(requestcheckvar(request("orderserial"),11))
	userid = requestcheckvar(request("userid"),16)
	smsyn = requestcheckvar(request("smsyn"),1)

michulgoorder=0
dim referer
	referer = request.ServerVariables("HTTP_REFERER")

if InStr(referer,"10x10.co.kr")<1 and session("ssBctId")<>"tozzinet" then
	response.write "not valid Referer"
    response.end
end if

if itemgubun="" then itemgubun="10"

'//���ⱸ�� �߼� ����� ���� ��������
if mode="standingusersudonginsert" then
	if getNumeric(itemid)="" then
		response.write "��ǰ�ڵ尡 �����ϴ�."
		dbget.close()	:	response.end
	end if
	if itemoption="" then
		response.write "�ɼ��ڵ尡 �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(sendkey)="" then
		response.write "�߼������� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(reserveidx)="" then
		response.write "����ȸ�� Vol.(��ȣ)�� �����ϴ�."
		dbget.close()	:	response.end
	end if

	strSql = "exec db_item.[dbo].[sp_Ten_item_standing_user_insert_sudong] '"& itemgubun &"', "& itemid &", '"& itemoption &"', "& sendkey &", "& reserveidx &""

	'response.write strSql & "<Br>"
	dbget.execute strSql

	sendstatus=0

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.replace('"& getSCMSSLURL &"/admin/itemmaster/standing/pop_standinguser.asp?itemgubun="& itemgubun &"&itemid="& itemid &"&itemoption="& itemoption &"&sendkey="& sendkey &"&sendstatus="& sendstatus &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

'/����
elseif mode="editstandinguser" then
	if getNumeric(uidx)="" then
		response.write "�ϷĹ�ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if jukyogubun="" then
		response.write "���䰡 �����ϴ�."
		dbget.close()	:	response.end
	end if
	if username="" then
		response.write "�̸��� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(zipcode)="" then
		response.write "�����ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if reqzipaddr="" then
		response.write "�ּ�1�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(useraddr)="" then
		response.write "���ּҰ� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(userphone1)="" or getNumeric(userphone2)="" or getNumeric(userphone3)="" then
		response.write "��ȭ��ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(usercell1)="" or getNumeric(usercell2)="" or getNumeric(usercell3)="" then
		response.write "�ڵ��� ��ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if isusing="" then
		response.write "��뿩�ΰ� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(itemno)="" then
		response.write "������ �����ϴ�."
		dbget.close()	:	response.end
	end if

	userphone = userphone1 & "-" & userphone2 & "-" & userphone3
	usercell = usercell1 & "-" & usercell2 & "-" & usercell3

	strSql = "update db_item.[dbo].[tbl_item_standing_user]" & vbcrlf
	strSql = strSql & " set jukyogubun = '"& jukyogubun &"'" & vbcrlf
	strSql = strSql & " , username='"& html2db(username) &"'" & vbcrlf
	strSql = strSql & " , zipcode='"& html2db(zipcode) &"'" & vbcrlf
	strSql = strSql & " , reqzipaddr='"& html2db(reqzipaddr) &"'" & vbcrlf
	strSql = strSql & " , useraddr='"& html2db(useraddr) &"'" & vbcrlf
	strSql = strSql & " , userphone='"& html2db(userphone) &"'" & vbcrlf
	strSql = strSql & " , usercell='"& html2db(usercell) &"'" & vbcrlf
	strSql = strSql & " , isusing='"& isusing &"'" & vbcrlf
	strSql = strSql & " , itemno="& itemno &"" & vbcrlf
	strSql = strSql & " , lastupdate=getdate()" & vbcrlf
	strSql = strSql & " , lastadminid='"& lastuserid &"' where" & vbcrlf
	strSql = strSql & " uidx="& uidx &"" & vbcrlf

	'response.write strSql & "<Br>"
	dbget.execute strSql

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	response.write "	location.replace('"& getSCMSSLURL &"/admin/itemmaster/standing/pop_standinguser_edit.asp?uidx="& uidx &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

'/��߼�
elseif mode="standinguser_re" then
	if getNumeric(uidx)="" then
		response.write "�ϷĹ�ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if jukyogubun="" then
		response.write "���䰡 �����ϴ�."
		dbget.close()	:	response.end
	end if
	if username="" then
		response.write "�̸��� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(zipcode)="" then
		response.write "�����ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if reqzipaddr="" then
		response.write "�ּ�1�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(useraddr)="" then
		response.write "���ּҰ� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(userphone1)="" or getNumeric(userphone2)="" or getNumeric(userphone3)="" then
		response.write "��ȭ��ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(usercell1)="" or getNumeric(usercell2)="" or getNumeric(usercell3)="" then
		response.write "�ڵ��� ��ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if isusing="" then
		response.write "��뿩�ΰ� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(itemno)="" then
		response.write "������ �����ϴ�."
		dbget.close()	:	response.end
	end if

	userphone = userphone1 & "-" & userphone2 & "-" & userphone3
	usercell = usercell1 & "-" & usercell2 & "-" & usercell3

	sendstatuscnt = getsendstatuscnt("05", itemid, itemoption, sendkey, "Y", orderserial, jukyogubun, usercell)
	if sendstatuscnt>0 then
		response.write "�ش� ȸ���� �߼۴�⳪ ��߼۴�� �׸��� �̹� ���� �մϴ�.<Br>�ٽ� Ȯ�� �Ͻð� ��߼� �ϼ���."
		dbget.close()	:	response.end
	end if

	strSql = "insert into db_item.[dbo].[tbl_item_standing_user] (" & vbcrlf
	strSql = strSql & " orgitemid, orgitemoption, sendkey, jukyogubun, orderserial, userid, itemno, sendstatus, senddate, username" & vbcrlf
	strSql = strSql & " , zipcode, reqzipaddr, useraddr, userphone, usercell, isusing, regdate, regadminid, lastupdate ,lastadminid" & vbcrlf
	strSql = strSql & " , rebeasongbeforeuidx" & vbcrlf
	strSql = strSql & " )" & vbcrlf
	strSql = strSql & " 	select" & vbcrlf
	strSql = strSql & " 	su.orgitemid, su.orgitemoption, su.sendkey, '"& jukyogubun &"', su.orderserial, su.userid, "& itemno &", 5, NULL, '"& html2db(username) &"'" & vbcrlf
	strSql = strSql & " 	, '"& html2db(zipcode) &"', '"& html2db(reqzipaddr) &"', '"& html2db(useraddr) &"', '"& html2db(userphone) &"'" & vbcrlf
	strSql = strSql & " 	, '"& html2db(usercell) &"', '"& isusing &"', getdate(), 'SYSTEM', getdate(), 'SYSTEM'" & vbcrlf
	strSql = strSql & " 	, (case" & vbcrlf
	strSql = strSql & " 		when isnull(su.rebeasongbeforeuidx,'')<>'' then su.rebeasongbeforeuidx else su.uidx end) as rebeasongbeforeuidx" & vbcrlf
	strSql = strSql & " 	from db_item.[dbo].[tbl_item_standing_user] su" & vbcrlf
	strSql = strSql & " 	where su.isusing='Y'" & vbcrlf
	strSql = strSql & " 	and su.sendstatus in (3,7)" & vbcrlf
	strSql = strSql & " 	and uidx="& uidx &"" & vbcrlf

	'response.write strSql & "<Br>"
	dbget.execute strSql

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.replace('"& getSCMSSLURL &"/admin/itemmaster/standing/pop_standinguser.asp?itemgubun="& itemgubun &"&itemid="& itemid &"&itemoption="& itemoption &"&sendkey="& sendkey &"&sendstatus=5&menupos="& menupos &"');"
	response.write "	location.replace('"& getSCMSSLURL &"/admin/itemmaster/standing/pop_standinguser_edit.asp?uidx="& uidx &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

'/���� �߼�
elseif mode="standinguser_sudong" then
	if getNumeric(itemid)="" then
		response.write "��ǰ�ڵ尡 �����ϴ�."
		dbget.close()	:	response.end
	end if
	if itemoption="" then
		response.write "�ɼ��ڵ尡 �����ϴ�."
		dbget.close()	:	response.end
	end if
	if jukyogubun="" then
		response.write "���䰡 �����ϴ�."
		dbget.close()	:	response.end
	end if
	if username="" then
		response.write "�̸��� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(zipcode)="" then
		response.write "�����ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if reqzipaddr="" then
		response.write "�ּ�1�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(useraddr)="" then
		response.write "���ּҰ� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(userphone1)="" or getNumeric(userphone2)="" or getNumeric(userphone3)="" then
		response.write "��ȭ��ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(usercell1)="" or getNumeric(usercell2)="" or getNumeric(usercell3)="" then
		response.write "�ڵ��� ��ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if isusing="" then
		response.write "��뿩�ΰ� �����ϴ�."
		dbget.close()	:	response.end
	end if
	if getNumeric(itemno)="" then
		response.write "������ �����ϴ�."
		dbget.close()	:	response.end
	end if

	userphone = userphone1 & "-" & userphone2 & "-" & userphone3
	usercell = usercell1 & "-" & usercell2 & "-" & usercell3

	strSql = "insert into db_item.[dbo].[tbl_item_standing_user] (" & vbcrlf
	strSql = strSql & " orgitemid, orgitemoption, reserveidx, jukyogubun, orderserial, userid, itemno, sendstatus, senddate, username" & vbcrlf
	strSql = strSql & " , zipcode, reqzipaddr, useraddr, userphone, usercell, isusing, regdate, regadminid, lastupdate ,lastadminid" & vbcrlf
	strSql = strSql & " )" & vbcrlf
	strSql = strSql & " 	select" & vbcrlf
	strSql = strSql & " 	orgitemid, orgitemoption, startreserveidx, '"& jukyogubun &"', '"& orderserial &"', '"& userid &"'" & vbcrlf
	strSql = strSql & " 	, "& itemno &", 0, NULL, '"& html2db(username) &"'" & vbcrlf
	strSql = strSql & " 	, '"& html2db(zipcode) &"', '"& html2db(reqzipaddr) &"', '"& html2db(useraddr) &"', '"& html2db(userphone) &"'" & vbcrlf
	strSql = strSql & " 	, '"& html2db(usercell) &"', '"& isusing &"', getdate(), 'SYSTEM', getdate(), 'SYSTEM'" & vbcrlf
	strSql = strSql & " 	from db_item.dbo.tbl_item_standing_item" & vbcrlf
	strSql = strSql & " 	where orgitemid = "& itemid &"" & vbcrlf
	strSql = strSql & " 	and orgitemoption = '"& itemoption &"'" & vbcrlf

	'response.write strSql & "<Br>"
	dbget.execute strSql

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.replace('"& getSCMSSLURL &"/admin/itemmaster/standing/pop_standinguser.asp?itemid="& itemid &"&itemoption="& itemoption &"&sendstatus=05&menupos="& menupos &"');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.end

'/�߼�ó��
elseif mode="savestandingsend" then
	if uidx="" then
		response.write "�ϷĹ�ȣ�� �����ϴ�."
		dbget.close()	:	response.end
	end if

	uidxarr = request.form("uidx")

	'/���� �߼�
	if smsyn="Y" then
		IF application("Svr_Info")<>"Dev" THEN
			'// LMS�߼�
			tmpSql = " insert into [smsdb].db_LgSMS.dbo.mms_msg( "
			tmpSql = tmpSql + " 	subject "
			tmpSql = tmpSql + " 	, phone "
			tmpSql = tmpSql + " 	, callback "
			tmpSql = tmpSql + " 	, status "
			tmpSql = tmpSql + " 	, reqdate "
			tmpSql = tmpSql + " 	, msg "
			tmpSql = tmpSql + " 	, file_cnt "
			tmpSql = tmpSql + " 	, file_path1 "
			tmpSql = tmpSql + " 	, expiretime) "
			tmpSql = tmpSql + " SELECT "
			tmpSql = tmpSql + " 	'" + html2db("[�ٹ�����] ���ⱸ�� �߼۾ȳ�") + "' "
			tmpSql = tmpSql + " 	, m.reqhp "
			tmpSql = tmpSql + " 	, '1644-6030' "
			tmpSql = tmpSql + " 	, '0' "
			tmpSql = tmpSql + " 	, getdate() "
			tmpSql = tmpSql + " 	, convert(varchar(4000),'" + ("�ֹ��Ͻ� ���ⱸ���� ����߼۵Ǿ����ϴ�." & vbCrLf & vbCrLf & "7���̳� �����Գ� Ȯ�� �����ϸ�, ��Ÿ ���ǻ����� ������ : 1644-6030 ���� ���� ��Ź �帳�ϴ�." & vbCrLf & vbCrLf & "�ູ ������ �Ϸ� �����ñ� �ٶ��ϴ� :)") + "') "
			tmpSql = tmpSql + " 	, 0 "
			tmpSql = tmpSql + " 	, null "
			tmpSql = tmpSql + " 	, '43200' "
			tmpSql = tmpSql + " FROM db_item.[dbo].[tbl_item_standing_user] su"
			tmpSql = tmpSql + " join db_order.dbo.tbl_order_master m"
			tmpSql = tmpSql + " 	on su.orderserial = m.orderserial"
			tmpSql = tmpSql + " WHERE su.SendDate is NULL and su.isusing='Y' and su.uidx in ("& uidxarr &") "

			'response.write tmpSql & "<br>"
			dbget.execute tmpSql
		end if
	end if

	strSql = "update su" & vbcrlf
	strSql = strSql & " set su.sendstatus = (case when su.sendstatus=0 then 3" & vbcrlf
	strSql = strSql & " 	when su.sendstatus=5 then 7 else 0 end)" & vbcrlf
	strSql = strSql & " ,su.senddate=getdate()" & vbcrlf
	strSql = strSql & " ,lastupdate=getdate()" & vbcrlf
	strSql = strSql & " ,lastadminid='"& lastuserid &"'" & vbcrlf
	strSql = strSql & " from db_item.[dbo].[tbl_item_standing_user] su where" & vbcrlf
	strSql = strSql & " su.isusing='Y'" & vbcrlf
	strSql = strSql & " and su.uidx in ("& uidxarr &")" & vbcrlf

	'response.write strSql & "<Br>"
	dbget.execute strSql

	' ���������� ���� ���Ϸ�� ����ħ
	strSql = "update d" & vbcrlf
	strSql = strSql & " set d.currstate = '7'" & vbcrlf		' ���Ϸ�
	strSql = strSql & " , d.beasongdate = getdate()" & vbcrlf
	strSql = strSql & " from db_item.dbo.tbl_item_standing_user su" & vbcrlf
	strSql = strSql & " join db_item.dbo.tbl_item_standing_order so" & vbcrlf
	strSql = strSql & " 	on su.orgitemid = so.orgitemid" & vbcrlf
	strSql = strSql & " 	and su.orgitemoption = so.orgitemoption" & vbcrlf
	strSql = strSql & " 	and su.reserveidx = so.reserveidx" & vbcrlf
	strSql = strSql & " join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	strSql = strSql & " 	on su.orderserial = d.orderserial" & vbcrlf
	strSql = strSql & " 	and so.reserveitemid = d.itemid" & vbcrlf
	strSql = strSql & " 	and so.reserveitemoption = d.itemoption" & vbcrlf
	strSql = strSql & " 	and d.cancelyn='A'" & vbcrlf
	strSql = strSql & " 	and d.currstate = '3'" & vbcrlf
	strSql = strSql & " where su.isusing='Y'" & vbcrlf
	strSql = strSql & " and su.uidx in ("& uidxarr &")" & vbcrlf

	'response.write strSql & "<Br>"
	dbget.execute strSql

	' ���������� ���� ���Ϸ�� ����ħ
	strSql = "update m" & vbcrlf
	strSql = strSql & " set m.ipkumdiv = '8', m.beadaldate=isNULL(m.beadaldate,(convert(varchar(19),getdate(),21)))" & vbcrlf		' ���Ϸ�
	strSql = strSql & " from db_item.dbo.tbl_item_standing_user su" & vbcrlf
	strSql = strSql & " join db_item.dbo.tbl_item_standing_order so" & vbcrlf
	strSql = strSql & " 	on su.orgitemid = so.orgitemid" & vbcrlf
	strSql = strSql & " 	and su.orgitemoption = so.orgitemoption" & vbcrlf
	strSql = strSql & " 	and su.reserveidx = so.reserveidx" & vbcrlf
	strSql = strSql & " join db_item.[dbo].[tbl_item_standing_item] si" & vbcrlf
	strSql = strSql & " 	on su.orgitemid = si.orgitemid" & vbcrlf
	strSql = strSql & " 	and su.orgitemoption = si.orgitemoption" & vbcrlf
	strSql = strSql & " 	and su.reserveidx = si.endreserveidx" & vbcrlf	' ������ ȸ�� ���� üũ��
	strSql = strSql & " join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	strSql = strSql & " 	on su.orderserial = d.orderserial" & vbcrlf
	strSql = strSql & " 	and so.reserveitemid = d.itemid" & vbcrlf
	strSql = strSql & " 	and so.reserveitemoption = d.itemoption" & vbcrlf
	strSql = strSql & " 	and d.cancelyn='A'" & vbcrlf
	strSql = strSql & " 	and d.currstate = '7'" & vbcrlf		' �ش� �ǹ߼� ��ǰ�� ���Ϸ� ���� üũ
	strSql = strSql & " join db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
	strSql = strSql & " 	on su.orderserial = m.orderserial" & vbcrlf
	strSql = strSql & " 	and m.cancelyn='N'" & vbcrlf
	strSql = strSql & " 	and m.ipkumdiv = '7'" & vbcrlf		' �Ϻ���� �ΰŸ�
	strSql = strSql & " where su.isusing='Y'" & vbcrlf
	strSql = strSql & " and su.uidx in ("& uidxarr &")" & vbcrlf

	'response.write strSql & "<Br>"
	dbget.execute strSql

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.replace('"& getSCMSSLURL &"/admin/itemmaster/standing/pop_standinguser.asp?itemgubun="& itemgubun &"&itemid="& itemid &"&itemoption="& itemoption &"&sendstatus=05&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end
else
	response.write "<script type='text/javascript'>"
	response.write "	alert('�����ڰ� �����ϴ�.');"
	response.write "</script>"
	dbget.close()	:	response.end
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->