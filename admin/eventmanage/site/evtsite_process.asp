<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/site/evtsite_process.asp
' Description :  �̺�Ʈ html �̹��� ������ó��
' History : 2007.03.27 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
	Dim siteidx, sitelocation, sitetype, sitecont, sitelinktype, sitelink, sitewidth, siteheight, sitedisporder, siteusing
	Dim eMode, strSql

    eMode = Trim(Request.Form("imod"))	
    siteidx = Request.Form("siteidx")
	sitelocation = Request.Form("sitelocation")
	sitetype = Request.Form("selType")
	if sitetype = "text" then
		sitecont = Request.Form("stxt")
	else
		sitecont = Request.Form("evtImg")
	end if	
	sitelinktype = Trim(Request.Form("selLinkType"))
	IF cStr(sitelinktype) = "M" THEN
		sitelink = html2db(Request.Form("sM"))
	ELSE
		sitelink = html2db(Request.Form("sL"))
	END IF	
	
	sitewidth = Request.Form("sW")
	siteheight = Request.Form("sH")
	sitedisporder = Request.Form("sDO")
	siteusing = Request.Form("rdoUse")

'--------------------------------------------------------
' ������ ó��
'--------------------------------------------------------
SELECT Case eMode
Case "I"	 '���
	strSql = "INSERT INTO [db_event].[dbo].[tbl_event_sitemanage]  ( "&_
			" [evtsite_location], [evtsite_type], [evtsite_cont], [evtsite_linktype], [evtsite_link], [evtsite_width], [evtsite_height], [evtsite_disporder], [evtsite_using], [adminid] "&_
			"  ) VALUES (  "&_ 
			 sitelocation&",'"&sitetype&"','"&sitecont&"','"&sitelinktype&"','"&sitelink&"','"&sitewidth&"','"&siteheight&"','"&sitedisporder&"','"&siteusing&"','"&session("ssBctId")&"'"&_
			 " )"
	dbget.execute strSql
		
	IF Err.Number = 0 THEN
		response.redirect("index.asp?menupos="&menupos&"&sitelocation="&sitelocation)
		dbget.close()	:	response.End
	ELSE
		Call sbAlertMsg ("��Ͽ� ������ �߻��Ͽ����ϴ�.", "back", "") 
	END IF	

CASE "U" '����
	strSql = " UPDATE [db_event].[dbo].[tbl_event_sitemanage] SET  "&_
			" 	[evtsite_location] = "&sitelocation&", [evtsite_type]='"&sitetype&"', [evtsite_cont]='"&sitecont&"', [evtsite_linktype]='"&sitelinktype&"'"&_
			"	, [evtsite_link]='"&sitelink&"', [evtsite_width]='"&sitewidth&"', [evtsite_height]='"&siteheight&"', [evtsite_disporder]='"&sitedisporder&"'"&_
			" 	, [evtsite_using]='"&siteusing&"', [adminid] = '"&session("ssBctId")&"'"&_
			" WHERE evtsite_idx = "&siteidx
	dbget.execute strSql
	
	IF Err.Number = 0 THEN
		response.redirect("index.asp?menupos="&menupos&"&sitelocation="&sitelocation)
		dbget.close()	:	response.End
	ELSE
		Call sbAlertMsg ("������ ������ �߻��Ͽ����ϴ�.", "back", "") 
	END IF	
CASE Else
	Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "") 
END SELECT	
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->