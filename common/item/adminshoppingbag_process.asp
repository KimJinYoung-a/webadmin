<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �¶��� & �������� ���� ��ٱ���
' Hieditor : 2011.08.04 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shortagestock_cls.asp" -->
<!-- #include virtual="/lib/classes/items/adminshoppingbag/adminshoppingbag_cls.asp" -->

<%
dim i , mode , bagidxarr , sql ,itemnoarr ,userid ,itemidarr ,itemgubunarr ,itemoptionarr ,shopid ,onoffgubun , menupos
    mode = requestCheckVar(request("mode"),32)
    bagidxarr = request("bagidxarr")
    itemnoarr = request("itemnoarr")    
	userid = session("ssBctId")
	onoffgubun = requestCheckVar(request("onoffgubun"),10)
	shopid = requestCheckVar(request("shopid"),32)
	itemidarr = Request("itemidarr")
	itemgubunarr = Request("itemgubunarr")
	itemoptionarr = Request("itemoptionarr")
	menupos = requestCheckVar(Request("menupos"),10)

'//��ٱ��� �߰�
if mode = "I" then
	if userid = "" or onoffgubun = "" or shopid = "" or itemgubunarr="" or itemidarr="" or itemoptionarr = "" or itemnoarr = "" then
        response.write "<script language='javascript'>"
        response.write "    alert('�ʿ��� ���� �����ϴ�');"
        response.write "</script>"
		dbget.close() : response.end		
	end if

	'//��ٱ��� �߰�
	'response.write userid &"/"& onoffgubun&"/"& shopid&"/"& itemgubunarr&"/"& itemidarr&"/"& itemoptionarr&"/"& itemnoarr &"<Br>"
	putadminshoppingbag_insert userid, onoffgubun, shopid, itemgubunarr, itemidarr, itemoptionarr, itemnoarr , "parent.opener" ,menupos

'//��ٱ��� �ٷ��߰�
elseif mode = "directbagaddarr" then
	if userid = "" or onoffgubun = "" or shopid = "" or itemgubunarr="" or itemidarr="" or itemoptionarr = "" or itemnoarr = "" then
        response.write "<script language='javascript'>"
        response.write "    alert('�ʿ��� ���� �����ϴ�');"
        response.write "</script>"
		dbget.close() : response.end		
	end if

	'//��ٱ��� �߰�
	'response.write userid &"/"& onoffgubun&"/"& shopid&"/"& itemgubunarr&"/"& itemidarr&"/"& itemoptionarr&"/"& itemnoarr &"<Br>"
	putadminshoppingbag_insert userid, onoffgubun, shopid, itemgubunarr, itemidarr, itemoptionarr, itemnoarr , "" ,menupos

    response.write "<script language='javascript'>"
    response.write "    alert('��ٱ��Ͽ� ��ǰ�� ����Ǿ����ϴ�');"
    response.write "    self.close();"
    response.write "</script>"
	dbget.close() : response.end
			    		
'//��ٱ��� �������    
elseif mode = "bagdelarr" then

	if bagidxarr = "" then
        response.write "<script language='javascript'>"
        response.write "    alert('�ε��� ���� �����ϴ�');"
        response.write "</script>"
		dbget.close() : response.end
	end if	
	
	bagidxarr = split(bagidxarr,",")
	
	if isarray(bagidxarr) then
		for i = 0 to ubound(bagidxarr) - 1
			sql = ""		
			sql = "delete from db_temp.dbo.tbl_adminshoppingbag where" + vbcrlf
			sql = sql & " bagidx = "& requestCheckVar(bagidxarr(i),10) &""

			'response.write sql & "<Br>"
			dbget.execute sql		
		next
	end if

    response.write "<script type='text/javascript'>"
    response.write "    alert('OK');"
    response.write "	parent.location.reload();"
    response.write "</script>"
	dbget.close() : response.end

'//�ֹ��� ��ٱ��� ����    
elseif mode = "baginsertdelarr" then

	if bagidxarr = "" then
        response.write "<script type='text/javascript'>"
        response.write "    alert('�ε��� ���� �����ϴ�');"
        response.write "</script>"
		dbget.close() : response.end
	end if	
	
	bagidxarr = split(bagidxarr,",")
	
	if isarray(bagidxarr) then
		for i = 0 to ubound(bagidxarr) - 1
			sql = ""		
			sql = "delete from db_temp.dbo.tbl_adminshoppingbag where" + vbcrlf
			sql = sql & " bagidx = "& requestCheckVar(bagidxarr(i),10) &""
			
			'response.write sql & "<Br>"
			dbget.execute sql		
		next
	end if

	dbget.close() : response.end
		
'//��ٱ��� �������	
elseif mode = "bageditarr" then

	if bagidxarr = "" or itemnoarr = "" then
        response.write "<script type='text/javascript'>"
        response.write "    alert('�ε��� ���� �����ϴ�');"
        response.write "</script>"
		dbget.close() : response.end
	end if
	
	bagidxarr = split(bagidxarr,",")
	itemnoarr = split(itemnoarr,",")


	if isarray(bagidxarr) then
		for i = 0 to ubound(bagidxarr) - 1		
			sql = ""
			sql = "update db_temp.dbo.tbl_adminshoppingbag" + vbcrlf
			sql = sql & " set itemno = "& requestCheckVar(itemnoarr(i),10) &" where" + vbcrlf
			sql = sql & " bagidx = "& requestCheckVar(bagidxarr(i),10) &""
			
			'response.write sql & "<Br>"
			dbget.execute sql		
		next
	end if
	
    response.write "<script type='text/javascript'>"
    response.write "    alert('OK');"
    response.write "	parent.location.reload();"
    response.write "</script>"
	dbget.close() : response.end	
end if    
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->