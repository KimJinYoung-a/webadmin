<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : 2010.11.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim cardidx ,startdate ,enddate ,isusing ,regdate , forecastgubun, link_url, couponidx
dim i , mode , contents , image_url , idx
	cardidx = request("cardidx")		
	startdate = request("startdate")
	enddate = request("enddate")				
	isusing = request("isusing")
	mode = request("mode")
	contents = request("contents")
	image_url = request("image_url")
	idx = request("idx")	
	forecastgubun = request("forecastgubun")
	link_url = request("link_url")
	couponidx = request("couponidx")
	If couponidx = "" Then
		couponidx = "null"
	End If
		
dim referer , sql
referer = request.ServerVariables("HTTP_REFERER")

'//�ű� & ���� 
if mode = "add" then
			
	'//�ű�	
	if cardidx = "" then

        sql = "select top 1" & vbcrlf
		sql = sql & " cardidx ,startdate ,enddate ,isusing ,regdate" + vbcrlf	
		sql = sql & " from db_momo.dbo.tbl_forecast_card" + vbcrlf
		sql = sql & " where isusing = 'Y' and startdate = '"&startdate&"' and enddate = '"&enddate&"'"

        'response.write sqlStr&"<br>"
        rsget.Open sql, dbget, 1
		if not rsget.EOF  then			
			response.write "<script language='javascript'>alert('�ش� ��¥�� ���� ������ �̹� ���� �մϴ�');self.close();</script>"
			dbget.close() : response.end
		end if
		
		sql = ""
		sql = "insert into db_momo.dbo.tbl_forecast_card (startdate,enddate,isusing)" + vbcrlf
		sql = sql & " values (" + vbcrlf		
		sql = sql & " '"&html2db(startdate)&"'"		
		sql = sql & " ,'"&html2db(enddate)&"'"	
		sql = sql & " ,'"&isusing&"'"												
		sql = sql & " )"		
	
		'response.write sql &"<br>"
		dbget.execute sql
		
	'//����	
	else 
	
		sql = "update db_momo.dbo.tbl_forecast_card set" + vbcrlf	
		sql = sql & " startdate='"&html2db(startdate)&"'" + vbcrlf
		sql = sql & " ,enddate='"&html2db(enddate)&"'" + vbcrlf		
		sql = sql & " ,isusing='"&isusing&"'" + vbcrlf						
		sql = sql & " where cardidx = "&cardidx&"" + vbcrlf	
		
		'response.write sql &"<br>"
		dbget.execute sql
		
	end if	
		

	response.write "<script language='javascript'>"
	response.write "	opener.location.reload();"
	response.write "	alert('OK');"
	response.write "	self.close();"
	response.write "</script>"

'//��ǥ���
elseif mode = "detailadd" then
			
	if cardidx = "" then
		response.write "<script>alert('�������� ���̵� ���� �����ϴ�.'); self.close();</script>"
		dbget.close() : response.end
	end if
		
	
	'//�ű�	
	if idx = "" then

        sql = "select top 1" & vbcrlf
		sql = sql & " idx ,cardidx ,forecastgubun ,image_url ,contents ,isusing, link_url" + vbcrlf	
		sql = sql & " from db_momo.dbo.tbl_forecast_card_detail" + vbcrlf
		sql = sql & " where isusing = 'Y' and cardidx = '"&cardidx&"' and forecastgubun = '"&forecastgubun&"'"

        'response.write sqlStr&"<br>"
        rsget.Open sql, dbget, 1
		if not rsget.EOF  then			
			response.write "<script language='javascript'>alert('ī�� �������� �Ѱ����� ��� ���� �մϴ�.');self.close();</script>"
			dbget.close() : response.end
		end if
		
		sql = ""
		sql = "insert into db_momo.dbo.tbl_forecast_card_detail (cardidx ,forecastgubun ,image_url ,contents ,isusing, link_url, couponidx)" + vbcrlf
		sql = sql & " values (" + vbcrlf
		sql = sql & " '"&cardidx&"'"	
		sql = sql & " ,'"&forecastgubun&"'"			
		sql = sql & " ,'"&html2db(image_url)&"'"		
		sql = sql & " ,'"&html2db(contents)&"'"	
		sql = sql & " ,'"&isusing&"'"
		sql = sql & " ,'"&link_url&"'"
		sql = sql & " , "&couponidx&" "
		sql = sql & " )"		
	
		'response.write sql &"<br>"
		dbget.execute sql
		
	'//����	
	else 
	
		sql = "update db_momo.dbo.tbl_forecast_card_detail set" + vbcrlf	
		sql = sql & " image_url='"&html2db(image_url)&"'" + vbcrlf
		sql = sql & " ,contents='"&html2db(contents)&"'" + vbcrlf		
		sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
		sql = sql & " ,link_url='"&link_url&"'" + vbcrlf
		sql = sql & " , couponidx = "&couponidx&" " + vbcrlf
		sql = sql & " where idx = "&idx&"" + vbcrlf	
		
		'response.write sql &"<br>"
		dbget.execute sql
		
	end if	
	
	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	response.write "	location.replace('" + referer + "');"		
	response.write "</script>"
end if
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
