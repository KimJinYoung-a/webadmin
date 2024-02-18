<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �޴�����
' History : ������ ����
'			2021.10.19 �ѿ�� ����(�����α� ����)
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

'������ �Ķ����
param = "?menupos=" & menupos & "&pid=" & parentid & "& page=" & page & "&SearchKey=" & SearchKey & "&SearchString=" & SearchString

'response.write part_sn & "<br>"
'response.write level_sn & "<br>"
'dbget.close()	:	response.End


'Ʈ������ ����
dbget.beginTrans

'// �޴����� ó�� �б� //
Select Case mode
	Case "add"
		strMsg = "�޴��� ����Ͽ����ϴ�."
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
		logmsg="�޴����ѽűԵ��"

	Case "modi", "popmodi"
		strMsg = "�޴������� �����Ͽ����ϴ�."

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

		if cstr(orgviewidx)<>cstr(viewidx) then logmsg = logmsg & "/ǥ�ü���: "& orgviewidx &" -> "& viewidx &""
		if cstr(orgmenuname)<>cstr(menuname) then logmsg = logmsg & "/�޴���: "& orgmenuname &" -> "& menuname &""
		if cstr(orgmenuname_en)<>cstr(menuname_en) then logmsg = logmsg & "/�޴���(����): "& orgmenuname_en &" -> "& menuname_en &""
		if cstr(orglinkurl)<>cstr(linkurl) then logmsg = logmsg & "/��ũURL: "& orglinkurl &" -> "& linkurl &""
		if orglv1customerYN="" then orglv1customerYN="N"
		if lv1customerYN="" then lv1customerYN="N"
		if cstr(orglv1customerYN)<>cstr(lv1customerYN) then
			logmsg = logmsg & "/LV1(������): "& orglv1customerYN &" -> "& lv1customerYN &""
		end if
		if orglv2partnerYN="" then orglv2partnerYN="N"
		if lv2partnerYN="" then lv2partnerYN="N"
		if cstr(orglv2partnerYN)<>cstr(lv2partnerYN) then
			logmsg = logmsg & "/LV2(��Ʈ������): "& orglv2partnerYN &" -> "& lv2partnerYN &""
		end if
		if orglv3InternalYN="" then orglv3InternalYN="N"
		if lv3InternalYN="" then lv3InternalYN="N"
		if cstr(orglv3InternalYN)<>cstr(lv3InternalYN) then
			logmsg = logmsg & "/LV3(��������): "& orglv3InternalYN &" -> "& lv3InternalYN &""
		end if
		if orguseSslYN="" then orguseSslYN="N"
		if useSslYN="" then useSslYN="N"
		if cstr(orguseSslYN)<>cstr(useSslYN) then logmsg = logmsg & "/SSL ���: "& orguseSslYN &" -> "& useSslYN &""
		if orgsaveLog="" then orgsaveLog="N"
		if saveLog="" then saveLog="N"
		if cstr(orgsaveLog)<>cstr(saveLog) then logmsg = logmsg & "/���ӷα� ����: "& orgsaveLog &" -> "& saveLog &""
		if cstr(orgmenucolor)<>cstr(menucolor) then logmsg = logmsg & "/ǥ�û���: "& orgmenucolor &" -> "& menucolor &""
		if cstr(orgparentid)<>cstr(parentid) then logmsg = logmsg & "/�����޴�: "& orgparentid &" -> "& parentid &""
		if cstr(orgisUsing)<>cstr(isUsing) then logmsg = logmsg & "/��뿩��: "& orgisUsing &" -> "& isUsing &""
		if cstr(orgdivcd)<>cstr(divcd) then logmsg = logmsg & "/��������: "& orgdivcd &" -> "& divcd &""

		'if left(logmsg,1)="/" then logmsg=mid(logmsg,2,4000)
		if logmsg<>"" then logmsg="�޴����Ѽ���"&logmsg

		'�޴� �������� ����
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

		'���Ӹ޴���ȣ ����
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
		orgpartstr=orgpartstr & "�μ���ȣ" & opart.FitemList(i).fpart_sn & "��޹�ȣ" & opart.FitemList(i).flevel_sn
	next
end if
set opart=nothing

'// �޴����� ���� //
'���� �ڷ� ����
SQL = "Delete From db_partner.dbo.tbl_menu_part Where menu_id=" & menuid
dbget.Execute(SQL)

'�����ڷ� ����
if part_sn<>"" then splPsn = Split(part_sn, ",")
if level_sn<>"" then splLsn = Split(level_sn, ",")

''IsArray(splPsn) : �߰� ������ - ������ ���°�쿡�� ���� �����ϵ��� ����
If IsArray(splPsn) Then
    For lp1=0 to Ubound(splPsn)
    	IF Trim(splPsn(lp1))<>"" and Trim(splLsn(lp1))<>"" THEN
    	SQL =	"Insert Into db_partner.dbo.tbl_menu_part (menu_id, part_sn, level_sn) " &_
    			"Values (" & menuid & ", " & splPsn(lp1) & ", " & splLsn(lp1) & ")"
    	dbget.Execute(SQL)
    	END IF
    Next

    '// �θ�޴� ���� �߰�(�θ� �ְ�, �θ� �ڽ� ������ ���°��;������ �߰�;������2012-03-08) //
    if parentid>0 then

		'1. ���������߿� �߰��Ǵ°� ���� ���� ����� ������ ����
		SQL =	"Delete T1 " &_
				"From db_partner.dbo.tbl_menu_part as T1 " &_
				"	join db_partner.dbo.tbl_menu_part as T2 " &_
				"		on T1.part_sn=T2.part_sn " &_
				"Where T1.menu_id=" & parentid &_
				"	and T2.menu_id=" & menuid &_
				"	and T1.level_sn<T2.level_sn "
		dbget.Execute(SQL)

		'2. �߰��Ǵ°� �߿� �������Ѻ��� ���� ��� �Ǵ� �ű��϶� �߰�
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

    '// ���Ӹ޴� ���� ���� //
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
		currpartstr=currpartstr & "�μ���ȣ" & opart.FitemList(i).fpart_sn & "��޹�ȣ" & opart.FitemList(i).flevel_sn
	next
end if

set opart=nothing

if cstr(orgpartstr)<>cstr(currpartstr) then
	if mode="modi" or mode="popmodi" then
		logmsg = logmsg & "/��������: "& orgpartstr &" -> "& currpartstr &""
		if left(logmsg,8)<>"�޴����Ѽ���" then logmsg="�޴����Ѽ���"&logmsg
	else
		logmsg = logmsg & "/��������: "& currpartstr &""
	end if
end if

if logmsg<>"" then
	' �޴������α�����		' 2021.10.19 �ѿ�� ����
	SQL="insert into db_partner.dbo.tbl_partner_menu_log (" & vbcrlf
	SQL=SQL& " menuid,logtype,logmsg,isusing,adminid,regdate)" & vbcrlf
	SQL=SQL& " 		select" & vbcrlf
	SQL=SQL& " 		"& menuid &",N'"& logtype &"',N'"& logmsg &"','Y',N'"& adminid &"',getdate()" & vbcrlf

	dbget.Execute SQL
end if

'�����˻� �� ����
If Err.Number = 0 Then
	dbget.CommitTrans				'Ŀ��(����)

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
    dbget.RollBackTrans				'�ѹ�(�����߻���)

	response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
