<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : 2010.12.22 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim playSn ,startdate ,enddate ,isusing, playLinkType, evt_code, linkURL, itemid
dim i , mode , plyItemSn, chkCnt
	playSn = request("playSn")
	startdate = request("startdate")
	enddate = request("enddate")
	isusing = request("isusing")
	playLinkType = request("playLinkType")
	evt_code = request("evt_code")
	linkURL = request("linkURL")
	mode = request("mode")
	plyItemSn = request("plyItemSn")
	itemid = request("itemid")

dim referer , sql
referer = request.ServerVariables("HTTP_REFERER")

'//�ű� & ���� 
if mode = "add" then
			
	'//�ű�	
	if playSn = "" then

        '// �ߺ� ��� �˻�
        sql = "select count(playSn) " + vbcrlf	
		sql = sql & " from db_momo.dbo.tbl_momo_playInfo" + vbcrlf
		sql = sql & " where isusing = 'Y' and playStartDate = '"&startdate&"' and playEndDate = '"&enddate&"'"

        rsget.Open sql, dbget, 1
        	chkCnt = rsget(0)
        rsget.Close
		if chkCnt>0  then
			response.write "<script language='javascript'>alert('�ش� ��¥�� ���� ������ �̹� ���� �մϴ�');self.close();</script>"
			dbget.close() : response.end
		end if

		'// �̺�Ʈ ��ȿ�� �˻�
		if playLinkType="E" then
	        sql = "select count(evt_code) " + vbcrlf	
			sql = sql & " from db_event.dbo.tbl_event" + vbcrlf
			sql = sql & " where evt_code='"&evt_code&"' and evt_using = 'Y' and evt_startdate<='"&startdate&"' and evt_enddate>='"&enddate&"'"

	        rsget.Open sql, dbget, 1
	        	chkCnt = rsget(0)
	        rsget.Close
			if chkCnt<=0  then
				response.write "<script language='javascript'>alert('�ش� �Ⱓ�� �����ϴ� �̺�Ʈ�� �����ϴ�.');history.back();</script>"
				dbget.close() : response.end
			end if
		end if

		'// �������� ���� ó��
		sql = "insert into db_momo.dbo.tbl_momo_playInfo (playLinkType, evt_code, linkURL, playStartDate, playEndDate, isusing)" + vbcrlf
		sql = sql & " values (" + vbcrlf
		sql = sql & " '"&playLinkType&"'"
		sql = sql & " ,'"&evt_code&"'"
		sql = sql & " ,'"&html2db(linkURL)&"'"
		sql = sql & " ,'"&html2db(startdate)&"'"
		sql = sql & " ,'"&html2db(enddate)&"'"
		sql = sql & " ,'"&isusing&"'"
		sql = sql & " )"
	
		dbget.execute sql

	'//����	
	else 
	
		sql = "update db_momo.dbo.tbl_momo_playInfo set" + vbcrlf	
		sql = sql & " playLinkType='"&playLinkType&"'" + vbcrlf
		sql = sql & " ,evt_code='"&evt_code&"'" + vbcrlf
		sql = sql & " ,linkURL='"&html2db(linkURL)&"'" + vbcrlf
		sql = sql & " ,playStartDate='"&html2db(startdate)&"'" + vbcrlf
		sql = sql & " ,playEndDate='"&html2db(enddate)&"'" + vbcrlf		
		sql = sql & " ,isusing='"&isusing&"'" + vbcrlf						
		sql = sql & " where playSn = "&playSn&"" + vbcrlf	
		
		dbget.execute sql
		
	end if			

	response.write "<script language='javascript'>"
	response.write "	opener.location.reload();"
	response.write "	alert('OK');"
	response.write "	self.close();"
	response.write "</script>"

'//��ǰ���
elseif mode = "itemAdd" then

	if playSn = "" then
		response.write "<script>alert('�������� ���̵� ���� �����ϴ�.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	'// ���۵� ������ �ڵ尪 Ȯ��
	if Right(itemid,1)="," then
		itemid = Left(itemid,Len(itemid)-1)
	end if

	'// �߰�
	sql = "insert into db_momo.dbo.tbl_momo_playItem" &_
			" (playSn, itemid)" &_
			" select '" + Cstr(playSn) + "', itemid" &_
			" from [db_item].[dbo].tbl_item" &_
			" where itemid in (" + itemid + ")" 
	dbget.execute sql
	
	response.write "<script language='javascript'>"
	response.write "	location.replace('" + referer + "');"		
	response.write "</script>"

'//���� ��ǰ����
elseif mode = "itemDel" then

	if playSn = "" then
		response.write "<script>alert('�������� ���̵� ���� �����ϴ�.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	'// ���۵� ������ �ڵ尪 Ȯ��
	if Right(plyItemSn,1)="," then
		plyItemSn = Left(plyItemSn,Len(plyItemSn)-1)
	end if

	'// ����
	sql = "delete db_momo.dbo.tbl_momo_playItem " + vbcrlf
	sql = sql & " where  plyItemSn in (" & plyItemSn & ") "
	dbget.execute sql
	
	response.write "<script language='javascript'>"
	response.write "	location.replace('" & referer & "');"
	response.write "</script>"

'//�̺�Ʈ��ǰ���
elseif mode = "evtItemAdd" then
			
	if playSn = "" then
		response.write "<script>alert('�������� ���̵� ���� �����ϴ�.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	'// �̺�Ʈ��ǰ ���翩�� Ȯ��
    sql = "select count(itemid) " + vbcrlf	
	sql = sql & " from db_event.dbo.tbl_eventitem" + vbcrlf
	sql = sql & " where evt_code='"&evt_code&"' "

    rsget.Open sql, dbget, 1
    	chkCnt = rsget(0)
    rsget.Close
	if chkCnt<=0  then
		response.write "<script language='javascript'>alert('�ش� �̺�Ʈ�� ��ϵ� ��ǰ�� �����ϴ�.\n�̺�Ʈ���� ��ǰ�� ���� ������ּ���.');history.back();</script>"
		dbget.close() : response.end
	end if

	'// ���� ��ǰ ����
	sql = "delete db_momo.dbo.tbl_momo_playItem " + vbcrlf
	sql = sql & " where playSn=" & playSn
	dbget.execute sql

	'// �̺�Ʈ ��ǰ �߰�
	sql = "insert into db_momo.dbo.tbl_momo_playItem" &_
			" (playSn, itemid)" &_
			" select '" + Cstr(playSn) + "', itemid" &_
			" from [db_event].[dbo].tbl_eventitem" &_
			" where evt_code='"&evt_code&"' "
	dbget.execute sql

	response.write "<script language='javascript'>"
	response.write "	location.replace('" + referer + "');"		
	response.write "</script>"
end if
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
