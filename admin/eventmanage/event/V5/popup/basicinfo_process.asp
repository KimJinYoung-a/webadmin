<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : basicinfo_process.asp
' Discription : I��(������) �̺�Ʈ �⺻���� ��� ���μ���
' History : 2019.01.22 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%
dim eKind, blnWeb, blnMobile, blnApp, eName, eSdate, eEdate, ePdate
dim eDateView, endlessview, eSale, eGift, eCoupon, eOnlyTen, eComment
dim eBbs, eItemps, eisblogurl, eOneplusone, eFreedelivery, eBookingsell
dim eisNew, ediary, eLevel, eManager, eusing, eMode, eScope, eCode
dim sOpenDate, sCloseDate, sImgregdate, eState, sPartnerid, evt_sortNo
dim isWeb, isMobile, isApp, strSql, eTag, evt_tagkind
dim bannerTypeDiv, bannerCouponTxt, bannerGubun, etcitemid, subcopyK
dim refer, evt_type, evt_kind, estimateSalePrice, itemsort, vChangeContents
dim vSCMChangeSQL, eSalePer, eSaleCPer, eSTime, eETime, marketing_event_kind
refer = request.ServerVariables("HTTP_REFERER")

eMode 	= requestCheckVar(Request.Form("imod"),2) '������ ó������
eCode = requestCheckVar(Request.Form("evt_code"),10)
eKind = requestCheckVar(Request.Form("eventkind"),2)
isWeb = requestCheckVar(Request.Form("blnWeb"),1)
isMobile = requestCheckVar(Request.Form("blnMobile"),1)
isApp = requestCheckVar(Request.Form("blnApp"),1)
eName = html2db(requestCheckVar(Request.Form("sEN"),120))
eSdate = requestCheckVar(Request.Form("sSD"),10)
eSTime = requestCheckVar(Request.Form("sST"),10)
eEdate = requestCheckVar(Request.Form("sED"),10)
eETime = requestCheckVar(Request.Form("sET"),10)
ePdate = requestCheckVar(Request.Form("sPD"),10)
eDateView = requestCheckVar(Request.Form("dateview"),1)
endlessview = requestCheckVar(Request.Form("endlessview"),1)
itemsort = requestCheckVar(Request.Form("itemsort"),2)

eSale = requestCheckVar(Request.Form("chSale"),1)
eGift = requestCheckVar(Request.Form("chGift"),1)
eCoupon	= requestCheckVar(Request.Form("chCoupon"),1)
eOnlyTen = requestCheckVar(Request.Form("chOnlyTen"),1)
eComment = requestCheckVar(Request.Form("chComm"),1)
eBbs = requestCheckVar(Request.Form("chBbs"),1)
eItemps	= requestCheckVar(Request.Form("chItemps"),1)
eisblogurl	= requestCheckVar(Request.Form("isblogurl"),1)
eOneplusone	= requestCheckVar(Request.Form("chOneplusone"),1)
eFreedelivery = requestCheckVar(Request.Form("chFreedelivery"),1)
eBookingsell = requestCheckVar(Request.Form("chBookingsell"),1)
eisNew =requestCheckVar(Request.Form("chNew"),1)
ediary = requestCheckVar(Request.Form("chDiary"),1)
eLevel = requestCheckVar(Request.Form("eventlevel"),1)
eState = requestCheckVar(Request.Form("eventstate"),1)
eManager = requestCheckVar(Request.Form("eventmanager"),1)
eusing = requestCheckVar(Request.Form("using"),1)

bannerTypeDiv = requestCheckVar(Request.Form("bannerTypeDiv"),1)
bannerCouponTxt = Request.Form("bannerCouponTxt")
bannerGubun = requestCheckVar(Request.Form("bannerGubun"),1)
etcitemid = Trim(requestCheckVar(Request.Form("etcitemid"),10)) '��ǰ���� ��ǰ�ڵ�
subcopyK = html2db(requestCheckVar(Request.Form("subcopyK"),500)) '����ī�� �ѱ�
eTag = html2db(requestCheckVar(Replace(Request.Form("eTag")," ",""),300))

evt_type = requestCheckVar(Request.Form("evt_type"),10)
evt_kind = requestCheckVar(Request.Form("evt_kind"),10)
estimateSalePrice = requestCheckVar(Request.Form("estimateSalePrice"),16)

eSalePer = requestCheckVar(Request.Form("sSP"),8)
eSaleCPer = requestCheckVar(Request.Form("sCSP"),8)
marketing_event_kind = requestCheckVar(Request.Form("marketing_event_kind"),1)

if eEdate="" then eEdate=dateadd("yyyy",1,eSdate)

if eSTime <> "00" then ' �ð� ������
    eSdate = eSdate & " " & eSTime & ":00:00"
end if
if eETime <> "00" then ' �ð� ������
    eEdate = eEdate & " " & eETime & ":59:29"
else
    eEdate = eEdate & " 23:59:29"
end if

if endlessview="Y" then
    if eEdate="" then eEdate=eSdate
end if
if eName <> "" then
	if checkNotValidHTML(eName) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if bannerCouponTxt <> "" then
	if checkNotValidHTML(bannerCouponTxt) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if subcopyK <> "" then
	if checkNotValidHTML(subcopyK) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if eTag <> "" then
	if checkNotValidHTML(eTag) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
If Right(eTag,1) = "," Then
    eTag = Left(eTag,(Len(eTag)-1))
End If

if isWeb = "" then isWeb = 0
if isMobile = "" then isMobile = 0
if isApp = "" then isApp = 0
if eDateView = "" then eDateView = 0
if eGift ="" then eGift = 0
if eSale ="" then eSale = 0
if eCoupon ="" then eCoupon = 0
if eOnlyTen ="" then eOnlyTen = 0
if eComment ="" then eComment = 0
if eBbs ="" then eBbs = 0
if eItemps ="" then eItemps = 0
if eisblogurl ="" then eisblogurl = 0
if eOneplusone ="" then	eOneplusone = 0
if eFreedelivery ="" then eFreedelivery = 0
if eBookingsell ="" then eBookingsell = 0
if eisNew ="" then eisNew = 0
if ediary = "" then ediary = 0
if eManager ="" then eManager = 1
if estimateSalePrice="" then estimateSalePrice=0
if bannerTypeDiv="3" then eGift=1
'��÷�� ���� ���� üũ
if eKind<>"28" then
    if eComment=0 and eBbs=0 and eItemps=0 then ePdate=""
end if
Dim strAdd : strAdd = ""
Dim strAdd1 : strAdd1 = ""
Dim istatus
'--gift Ȯ��
Dim strgift	: strgift = ""
Dim igiftcnt : igiftcnt = 0
Dim isAllGiftEvent : isAllGiftEvent = False
Dim strSale	: strSale = ""
Dim arrSale,intSale
dim tempSalePer, tempSaleCPer, title_pc

tempSalePer = "~" & eSalePer & "%"
tempSaleCPer = "~" & eSaleCPer & "%"

If eSalePer <> "" and eSalePer <> "0" Then
    title_pc = eName
    if ((eKind = "1" or  ekind="23" ) and (eSale = "1" or eCoupon="1") and (eSalePer <> "" or eSalePer <> "0" )) then eName = eName &"|"&tempSalePer
Elseif eSaleCPer<>"" And (eSalePer="" or eSalePer="0") Then
    title_pc = eName
    if ((eKind = "1" or  ekind="23" ) and (eSale = "1" or eCoupon="1") and (eSaleCPer <> "" or eSaleCPer <> "0" )) then eName = eName &"|"&tempSaleCPer
End If

'--------------------------------------------------------
' ������ ó��
' I : �̺�Ʈ ������, U: �������, disply���/����
'--------------------------------------------------------
select case eMode
case "BI"
    '�⺻�� ����
    'eState = 0
    eScope = 2
    if eScope=2 then sPartnerid = ""
    if evt_sortNo="" then evt_sortNo="0"
	sOpenDate = "null"
	sCloseDate = "null"
	sImgregdate = "null"
	
    '���°� �����϶� ������ ���
	if eState = 6 or eState = 7 then
		sOpenDate = "getdate()"
	elseif eState = 9 then
		sCloseDate = "getdate()"
	elseif eState = 3 then
	    sImgregdate = "getdate()"	
	end if
	
	'Ʈ����� (1.master���/2.disply���/3.MDTheme���)
	dbget.beginTrans
		'--1.master���
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event] (evt_kind, evt_manager, evt_scope, partner_id,evt_name, evt_startdate, evt_enddate, evt_prizedate, evt_level, evt_state, opendate, closedate, evt_lastupdate, adminid, evt_sortNo , isWeb, isMobile, isApp ,evt_imgregdate, evt_subcopyk, evt_subname) " & vbCrlf
		strSql = strSql + " VALUES("&eKind&","&eManager&","&eScope&",'"&sPartnerid&"','"&eName&"','"&eSdate&"','"&eEdate&"','"&ePdate&"',"&eLevel&","&eState&","&sOpenDate&","&sCloseDate&",getdate(),'"&session("ssBctId")&"',"&evt_sortNo&","&isWeb&","&isMobile&","&isApp&","&sImgregdate&",'" & subcopyK & "','" & subcopyK & "')"
		dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
            response.End 
        end if

        strSql = "select SCOPE_IDENTITY()"
        rsget.Open strSql, dbget, 0
        eCode = rsget(0)
        rsget.Close

        '--2.disply���
        if marketing_event_kind ="1" then
            strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display](evt_code, issale, isgift, iscoupon, isOnlyTen, isOneplusone"&_
                    ", isFreedelivery, isbookingsell, isDiary, isNew, iscomment, isbbs, isitemps, isGetBlogURL, evt_dateview, endlessview, evt_itemsort ,partMDid, estimateSalePrice, SalePer, SaleCPer, marketing_event_kind, evt_execFile_mo, evt_isExec_mo, evt_template, evt_template_mo)" & vbCrlf&_
                    " VALUES (" & eCode & ", " & eSale & "," & eGift & "," & eCoupon & ",'" & eOnlyTen & "'," & eOneplusone & "," & eFreedelivery &_
                    "," & eBookingsell & "," & ediary & "," & eisNew & "," & eComment & "," & eBbs & "," & eItemps & ",'" & eisblogurl & "','" & eDateView & "','" & endlessview & "'," & itemsort & ",'" & session("ssBctId") & "'," & estimateSalePrice & ",'" & eSalePer & "','" & eSaleCPer & "','" & marketing_event_kind & "','/apps/appcom/wish/web2014/event/etc/realtimeevt/pickUp.asp',1,10,11)"
        elseif marketing_event_kind ="2" then
            strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display](evt_code, issale, isgift, iscoupon, isOnlyTen, isOneplusone"&_
                   ", isFreedelivery, isbookingsell, isDiary, isNew, iscomment, isbbs, isitemps, isGetBlogURL, evt_dateview, endlessview, evt_itemsort ,partMDid, estimateSalePrice, SalePer, SaleCPer, marketing_event_kind, evt_execFile_mo, evt_isExec_mo, evt_template, evt_template_mo)" & vbCrlf&_
                   " VALUES (" & eCode & ", " & eSale & "," & eGift & "," & eCoupon & ",'" & eOnlyTen & "'," & eOneplusone & "," & eFreedelivery &_
                   "," & eBookingsell & "," & ediary & "," & eisNew & "," & eComment & "," & eBbs & "," & eItemps & ",'" & eisblogurl & "','" & eDateView & "','" & endlessview & "'," & itemsort & ",'" & session("ssBctId") & "'," & estimateSalePrice & ",'" & eSalePer & "','" & eSaleCPer & "','" & marketing_event_kind & "','/apps/appCom/wish/web2014/event/attendance/index.asp',1,10,11)"
        elseif marketing_event_kind ="5" then
            strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display](evt_code, issale, isgift, iscoupon, isOnlyTen, isOneplusone"&_
                   ", isFreedelivery, isbookingsell, isDiary, isNew, iscomment, isbbs, isitemps, isGetBlogURL, evt_dateview, endlessview, evt_itemsort ,partMDid, estimateSalePrice, SalePer, SaleCPer, marketing_event_kind, evt_execFile_mo, evt_isExec_mo, evt_template, evt_template_mo)" & vbCrlf&_
                   " VALUES (" & eCode & ", " & eSale & "," & eGift & "," & eCoupon & ",'" & eOnlyTen & "'," & eOneplusone & "," & eFreedelivery &_
                   "," & eBookingsell & "," & ediary & "," & eisNew & "," & eComment & "," & eBbs & "," & eItemps & ",'" & eisblogurl & "','" & eDateView & "','" & endlessview & "'," & itemsort & ",'" & session("ssBctId") & "'," & estimateSalePrice & ",'" & eSalePer & "','" & eSaleCPer & "','" & marketing_event_kind & "','/event/only_app/index.asp',1,10,11)"
        elseif marketing_event_kind ="6" then
            strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display](evt_code, issale, isgift, iscoupon, isOnlyTen, isOneplusone"&_
                   ", isFreedelivery, isbookingsell, isDiary, isNew, iscomment, isbbs, isitemps, isGetBlogURL, evt_dateview, endlessview, evt_itemsort ,partMDid, estimateSalePrice, SalePer, SaleCPer, marketing_event_kind, evt_execFile_mo, evt_isExec_mo, evt_template, evt_template_mo)" & vbCrlf&_
                   " VALUES (" & eCode & ", " & eSale & "," & eGift & "," & eCoupon & ",'" & eOnlyTen & "'," & eOneplusone & "," & eFreedelivery &_
                   "," & eBookingsell & "," & ediary & "," & eisNew & "," & eComment & "," & eBbs & "," & eItemps & ",'" & eisblogurl & "','" & eDateView & "','" & endlessview & "'," & itemsort & ",'" & session("ssBctId") & "'," & estimateSalePrice & ",'" & eSalePer & "','" & eSaleCPer & "','" & marketing_event_kind & "','/apps/appcom/wish/web2014/event/etc/secretShop/index.asp',1,10,11)"
        else
            strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display](evt_code, issale, isgift, iscoupon, isOnlyTen, isOneplusone"&_
                    ", isFreedelivery, isbookingsell, isDiary, isNew, iscomment, isbbs, isitemps, isGetBlogURL, evt_dateview, endlessview, evt_itemsort ,partMDid, estimateSalePrice, SalePer, SaleCPer, marketing_event_kind)" & vbCrlf&_
                    " VALUES (" & eCode & ", " & eSale & "," & eGift & "," & eCoupon & ",'" & eOnlyTen & "'," & eOneplusone & "," & eFreedelivery &_
                    "," & eBookingsell & "," & ediary & "," & eisNew & "," & eComment & "," & eBbs & "," & eItemps & ",'" & eisblogurl & "','" & eDateView & "','" & endlessview & "'," & itemsort & ",'" & session("ssBctId") & "'," & estimateSalePrice & ",'" & eSalePer & "','" & eSaleCPer & "','" & marketing_event_kind & "')"
        end if
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
            response.End 
        end if

        '--3.MDTheme��� 
        strSql = "INSERT INTO [db_event].[dbo].[tbl_event_md_theme](evt_code, title_pc, title_mo"
        if eComment then
            strSql = strSql & ", comm_start, comm_end"
        end if
        if eBbs then
            strSql = strSql & ", board_start, board_end"
        end if
        if eItemps then
            strSql = strSql & ", eval_start, eval_end"
        end if
        strSql = strSql & ")" & vbCrlf
        strSql = strSql & " VALUES (" & eCode & ",'" & title_pc & "','" & title_pc & "'"
        if eComment then
            strSql = strSql & ",'" & eSdate & "','" & eEdate & "'"
        end if
        if eBbs then
            strSql = strSql & ",'" & eSdate & "','" & eEdate & "'"
        end if
        if eItemps then
            strSql = strSql & ",'" & eSdate & "','" & eEdate & "'"
        end if
        strSql = strSql & ")"
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[3]", "back", "")
            response.End 
        end if

        if evt_type<>"" and evt_kind<>"" then
            '===========================================================
            '--4.���Ľ����̼��� ��� ����, ���� ����
            strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
            strSql = strSql + " SET eventtype_pc='" & evt_type & "'" & vbCrlf
            strSql = strSql + ", eventtype_mo='" & evt_kind & "'" & vbCrlf
            strSql = strSql + " where evt_code=" & eCode
            dbget.execute strSql
            if Err.Number <> 0 then
                dbget.RollBackTrans 
                Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[4]", "back", "")
                response.End
            end if
        end if
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=1&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case "BU"

 	sOpenDate = requestCheckVar(Request.Form("eOD"),30)
 	sCloseDate =requestCheckVar(Request.Form("eCD"),30)
    sImgregdate=requestCheckVar(Request.Form("eIRD"),30)
    
    eScope = 2
    if eScope=2 then sPartnerid = ""

 	IF ((eState = 6 or eState = 7) and sOpenDate ="" ) THEN 	'����ó���� ����
		strAdd = ", [opendate] = getdate() " & vbCrlf
	ELSEIF (eState = 9 and sCloseDate ="" ) THEN
		strAdd = ", [closedate] = getdate() " & vbCrlf	'����ó���� ����
	END IF
		
	IF (eState = 3 and sImgregdate ="" ) THEN
		strAdd1 = ", [evt_imgregdate] = getdate() " & vbCrlf	'�̹�������� ����
	END IF

	'������ ������ ����� ������ ���� ��¥�� ����
	IF eState = 9 and  datediff("d",eEdate,date()) <0 THEN
			eEdate = date()
	END IF



	'Ʈ����� (1.master����/2.disply����/3.MDTheme����)
	dbget.beginTrans
		'--1.master ����
		strSql = "UPDATE [db_event].[dbo].[tbl_event]" & vbCrlf
        strSql = strSql + " SET evt_kind=" & eKind & vbCrlf
        strSql = strSql + ", evt_manager=" & eManager & vbCrlf
        strSql = strSql + ", evt_name='" & eName & "'" & vbCrlf
        strSql = strSql + ", evt_startdate='" & eSdate & "'" & vbCrlf
        strSql = strSql + ", evt_enddate='" & eEdate & "'" & vbCrlf
        strSql = strSql + ", evt_prizedate='" & ePdate & "'" & vbCrlf
        strSql = strSql + ", evt_level=" & eLevel & vbCrlf
        strSql = strSql + ", evt_lastupdate=getdate()" & vbCrlf
        strSql = strSql + ", adminid='" & session("ssBctId") & "'" & vbCrlf
        strSql = strSql + ", isWeb=" & isWeb & vbCrlf
        strSql = strSql + ", isMobile=" & isMobile & vbCrlf
        strSql = strSql + ", isApp=" & isApp & vbCrlf
        strSql = strSql + ", evt_subcopyk='" & subcopyK & "'" & vbCrlf
        strSql = strSql + ", evt_subname='" & subcopyK & "'" & vbCrlf
        strSql = strSql + ", evt_using='" & eusing & "'" & vbCrlf
        strSql = strSql + ", evt_state=" & eState & vbCrlf
        strSql = strSql + strAdd
        strSql = strSql + strAdd1
        strSql = strSql + " WHERE evt_code=" & eCode
		dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
            response.End 
        end if

        '--2.disply ����
        strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
        strSql = strSql + " SET issale=" & eSale & vbCrlf
        strSql = strSql + ", isgift=" & eGift & vbCrlf
        strSql = strSql + ", iscoupon=" & eCoupon & vbCrlf
        strSql = strSql + ", isOnlyTen=" & eOnlyTen & vbCrlf
        strSql = strSql + ", isOneplusone=" & eOneplusone & vbCrlf
        strSql = strSql + ", isFreedelivery=" & eFreedelivery & vbCrlf
        strSql = strSql + ", isbookingsell=" & eBookingsell & vbCrlf
        strSql = strSql + ", isDiary=" & ediary & vbCrlf
        strSql = strSql + ", isNew=" & eisNew & vbCrlf
        strSql = strSql + ", iscomment=" & eComment & vbCrlf
        strSql = strSql + ", isbbs=" & eBbs & vbCrlf
        strSql = strSql + ", isitemps=" & eItemps & vbCrlf
        strSql = strSql + ", isGetBlogURL=" & eisblogurl & vbCrlf
        strSql = strSql + ", evt_dateview=" & eDateView & vbCrlf
        strSql = strSql + ", endlessview='" & endlessview & "'" & vbCrlf
        strSql = strSql + ", estimateSalePrice='" & estimateSalePrice & "'" & vbCrlf
        strSql = strSql + ", evt_itemsort='" & itemsort & "'" & vbCrlf
        strSql = strSql + ", SalePer='" & eSalePer & "'" & vbCrlf
        strSql = strSql + ", SaleCPer='" & eSaleCPer & "'" & vbCrlf
        strSql = strSql + ", marketing_event_kind='" & marketing_event_kind & "'" & vbCrlf
        if marketing_event_kind ="1" then
        strSql = strSql + ", evt_execFile_mo='/apps/appcom/wish/web2014/event/etc/realtimeevt/pickUp.asp'" & vbCrlf
        elseif marketing_event_kind ="2" then
        strSql = strSql + ", evt_execFile_mo='/apps/appcom/wish/web2014/event/everyday_mileage/index.asp'" & vbCrlf
        elseif marketing_event_kind ="5" then
        strSql = strSql + ", evt_execFile_mo='/event/only_app/index.asp'" & vbCrlf
        elseif marketing_event_kind ="6" then
        strSql = strSql + ", evt_execFile_mo='/apps/appcom/wish/web2014/event/etc/secretShop/index.asp'" & vbCrlf
        end if
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
            response.End 
        end if

        '--3.MDTheme ����
        strSql = "UPDATE [db_event].[dbo].[tbl_event_md_theme]" & vbCrlf
        strSql = strSql + " SET title_pc='" & title_pc & "'" & vbCrlf
        strSql = strSql + ", title_mo='" & title_pc & "'" & vbCrlf
        if eComment then
            strSql = strSql & " ,comm_start='" & eSdate & "'" & vbCrlf
            strSql = strSql & " ,comm_end='" & eEdate & "'" & vbCrlf
        end if
        if eBbs then
            strSql = strSql & " ,board_start='" & eSdate & "'" & vbCrlf
            strSql = strSql & " ,board_end='" & eEdate & "'" & vbCrlf
        end if
        if eItemps then
            strSql = strSql & " ,eval_start='" & eSdate & "'" & vbCrlf
            strSql = strSql & " ,eval_end='" & eEdate & "'" & vbCrlf
        end if
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
            response.End 
        end if

        if evt_type<>"" and evt_kind<>"" then
            '===========================================================
            '--4.���Ľ����̼��� ��� ����, ���� ����
            strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
            strSql = strSql + " SET eventtype_pc='" & evt_type & "'" & vbCrlf
            strSql = strSql + ", eventtype_mo='" & evt_kind & "'" & vbCrlf
            strSql = strSql + " where evt_code=" & eCode
            dbget.execute strSql
            if Err.Number <> 0 then
                dbget.RollBackTrans 
                Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[4]", "back", "")
                response.End
            end if
        end if

        vChangeContents = vChangeContents & "�̺�Ʈ UPDATE " & vbCrLf
        vChangeContents = vChangeContents & "- �̺�Ʈ�� : evt_name = " & eName & ", evt_code = " & eCode & vbCrLf
        vChangeContents = vChangeContents & "- ���� : evt_kind = " & eKind & vbCrLf
        vChangeContents = vChangeContents & "- Ÿ�� : ����issale = " & eSale & ", ����ǰisgift = " & eGift & ", ����iscoupon = " & eCoupon & ", isOnlyTen = " & eOnlyTen & ","
        vChangeContents = vChangeContents & " isOneplusone = " & eOneplusone & ", ������isFreedelivery = " & eFreedelivery & ", �����Ǹ�isbookingsell = " & eBookingsell & ","
        vChangeContents = vChangeContents & " isDiary = " & ediary & ", ��ĪisNew = " & eisNew & vbCrLf
        vChangeContents = vChangeContents & "- ��� : �ڸ�Ʈiscomment = " & eComment & ", �Խ���isbbs = " & eBbs & ", ��ǰ�ı�isitemps = " & eItemps & ", Blog URL isGetBlogURL = " & eisblogurl & vbCrLf
        vChangeContents = vChangeContents & "- �Ⱓ : evt_startdate ~ evt_enddate = " & eSdate & " ~ " & eEdate & vbCrLf
        vChangeContents = vChangeContents & "- ��÷��ǥ�� : evt_prizedate = " & ePdate & vbCrLf
        vChangeContents = vChangeContents & "- ���� : evt_state = " & eState & vbCrLf
        vChangeContents = vChangeContents & "- �߿䵵 : evt_level = " & eLevel & vbCrLf
        vChangeContents = vChangeContents & "- �̺�Ʈ�Ⱓ���⿩�� evt_dateview = " & eDateView & "" & vbCrLf
        vChangeContents = vChangeContents & "- ��ǥ��ǰ�����׹�� : ��ǥ��ǰ�ڵ� = " & etcitemid & vbCrLf

        '-�̺�Ʈ ���¿� ���� ����,����ǰ,���� ���� ����---------------
            IF (eState < 7) THEN  	'������ ���� �߱޴��� ���
                istatus = 0
            ELSEIF (eState <9) THEN
                istatus = 7
            ELSE
                istatus = eState
            END IF
            if eusing="N" then
                istatus = 9
            end if
        '--------------------------------------------------------------

        'IF egift = 0 THEN strgift = ", gift_using = 'N' "'����ǰüũ ���ҽ� �������� �ϴ� ����(2019.09.09-corpse2)

            strSql =" SELECT count(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code = "&eCode&" AND gift_using ='Y' "
            rsget.Open strSql, dbget
            IF not (rsget.EOF or rsget.BOF) THEN
                igiftcnt = rsget(0)
            END IF
            rsget.close
            
            ''��ü ���� �̺�Ʈ ���� CHECK
            strSql =" SELECT count(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code = "&eCode&" AND gift_scope in (1,9) "
            rsget.Open strSql, dbget
            IF not (rsget.EOF or rsget.BOF) THEN
                isAllGiftEvent = rsget(0)>0
            END IF
            rsget.close
            
            '��ü����/���̾������ ���� ����Ǹ� �ȵ�.
            if (isAllGiftEvent) then
                strgift = ""
            end if
            
            if igiftcnt > 0 then
            strSql ="	UPDATE [db_event].[dbo].[tbl_gift] Set gift_name = '"&eName&"' ,gift_startdate ='"&eSdate&"', gift_enddate ='"&eEdate&"', gift_status= "	&istatus&strAdd&_
                    "			, lastupdate = getdate(), adminid = '"&session("ssBctId")&"', site_scope= "&eScope&", partner_id='"&sPartnerid&"' "&strgift&_
                    "		WHERE evt_code = "&eCode
            
            if (istatus=0) then ''��ü����/���̾������ ���� ����Ǹ� �ȵ�.
                strSql = strSql&"  and gift_scope not in (1,9)"  
            end if
            
            dbget.execute strSql
            end if

        '-- sale Ȯ��
            IF eSale = 0 THEN strSale = ", sale_using = 0 "
            strSql = " SELECT sale_code, sale_status FROM [db_event].[dbo].[tbl_sale] WHERE evt_code = "&eCode&" AND sale_using =1 "
            rsget.Open strSql, dbget
            IF not (rsget.EOF or rsget.BOF) THEN
                arrSale = rsget.getRows()
            END IF
            rsget.close

            IF isarray(arrSale)  THEN
                For intSale = 0 To UBound(arrSale,2)
                '������ ��� ���»��°� 6, ������°� 8 �̹Ƿ� ���°� ���� �ʿ�
                if (eState = 7 AND arrSale(1,intSale) >= 6) OR ( eState > 7 AND arrSale(1,intSale) >= 8 )  THEN		istatus = arrSale(1,intSale)
                    strSql ="	UPDATE [db_event].[dbo].[tbl_sale] Set sale_name = '"&eName&"', sale_startdate ='"&eSdate&"', sale_enddate ='"&eEdate&"', sale_status="	&istatus&strAdd&_
                            "			, lastupdate = getdate(), adminid = '"&session("ssBctId")&"'"&strSale&_
                            "		WHERE evt_code = "&eCode&" and sale_code = "&arrSale(0,intSale)
                    dbget.execute strSql
                Next
            END IF

        IF Err.Number = 0 THEN            
            '### ���� �α� ����(event)
            vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
            vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & menupos & "', "
            vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
            dbget.execute(vSCMChangeSQL)
        ELSE
            dbget.RollBackTrans
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
        END IF
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=1&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End

case "PI"
    '�⺻�� ����
    'eState = 0
    eScope = 2
    if eScope=2 then sPartnerid = ""
    if evt_sortNo="" then evt_sortNo="0"
	sOpenDate = "null"
	sCloseDate = "null"
	sImgregdate = "null"
	
    '���°� �����϶� ������ ���
	if eState = 6 or eState = 7 then
		sOpenDate = "getdate()"
	elseif eState = 9 then
		sCloseDate = "getdate()"
	elseif eState = 3 then
	    sImgregdate = "getdate()"	
	end if
	
	'Ʈ����� (1.master���/2.disply���/3.MDTheme���)
	dbget.beginTrans
		'--1.master���
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event] (evt_kind, evt_manager, evt_scope, partner_id,evt_name, evt_startdate, evt_enddate, evt_level, evt_state, opendate, closedate, evt_lastupdate, adminid, evt_sortNo , isWeb, isMobile, isApp ,evt_imgregdate, evt_subcopyK, evt_subname) " & vbCrlf
		strSql = strSql + " VALUES("&eKind&","&eManager&","&eScope&",'"&sPartnerid&"','"&eName&"','"&eSdate&"','"&eEdate&"',"&eLevel&","&eState&","&sOpenDate&","&sCloseDate&",getdate(),'"&session("ssBctId")&"',"&evt_sortNo&","&isWeb&","&isMobile&","&isApp&","&sImgregdate&",'"&subcopyK & "','"&subcopyK&"')"
		dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
            response.End 
        end if

        strSql = "select SCOPE_IDENTITY()"
        rsget.Open strSql, dbget, 0
        eCode = rsget(0)
        rsget.Close

        '--2.disply���
        strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display](evt_code, issale, isgift, iscoupon, isOnlyTen, isOneplusone"&_
                ", isFreedelivery, isbookingsell, isDiary, isNew, iscomment, isbbs, isitemps, isGetBlogURL, evt_dateview, endlessview, bannerTypeDiv, bannerCouponTxt, bannerGubun, etc_itemid, evt_tag, evt_itemsort, partMDid, estimateSalePrice, eventtype_pc, eventtype_mo, mdtheme, mdthememo, SalePer, SaleCPer)" & vbCrlf&_
                " VALUES (" & eCode & ", " & eSale & "," & eGift & "," & eCoupon & ",'" & eOnlyTen & "'," & eOneplusone & "," & eFreedelivery &_
                "," & eBookingsell & "," & ediary & "," & eisNew & "," & eComment & "," & eBbs & "," & eItemps & ",'" & eisblogurl & "','" & eDateView & "','" & endlessview & "','" & bannerTypeDiv & "','" & bannerCouponTxt & "','" & bannerGubun & "','" & etcitemid & "','" & eTag & "'," & itemsort & ",'" & session("ssBctId") & "'," & estimateSalePrice & ", 80, 80, 5, 5,'" & eSalePer & "','" & eSaleCPer & "')"
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
            response.End 
        end if

        '--3.MDTheme��� 
        strSql = "INSERT INTO [db_event].[dbo].[tbl_event_md_theme](evt_code, title_pc, title_mo"
        if eComment then
            strSql = strSql & ", comm_start, comm_end"
        end if
        if eBbs then
            strSql = strSql & ", board_start, board_end"
        end if
        if eItemps then
            strSql = strSql & ", eval_start, eval_end"
        end if
        strSql = strSql & ")" & vbCrlf
        strSql = strSql & " VALUES (" & eCode & ",'" & title_pc & "','" & title_pc & "'"
        if eComment then
            strSql = strSql & ",'" & eSdate & "','" & eEdate & "'"
        end if
        if eBbs then
            strSql = strSql & ",'" & eSdate & "','" & eEdate & "'"
        end if
        if eItemps then
            strSql = strSql & ",'" & eSdate & "','" & eEdate & "'"
        end if
        strSql = strSql & ")"
        dbget.execute strSql

        '================ �̺�Ʈ ����� ��ǰ�̺�Ʈ =================
        '2015-11-04 ����ȭ �߰�
        if bannerTypeDiv="1" then
            evt_tagkind="7"
        elseif bannerTypeDiv="2" then
            evt_tagkind="2"
        elseif bannerTypeDiv="3" then
            evt_tagkind="1"
        elseif bannerTypeDiv="4" then
            evt_tagkind="4"
        elseif bannerTypeDiv="5" then
            evt_tagkind="3"
        elseif bannerTypeDiv="6" then
            evt_tagkind="5"
        elseif bannerTypeDiv="7" then
            evt_tagkind="6"
        end if
        strSql = "IF EXISTS(SELECT evt_code FROM db_event.dbo.tbl_event_mobile_addetc where evt_code = "&eCode&" )"&vbCrlf 
        strSql = strSql& "begin"&vbCrlf 
        strSql = strSql& " UPDATE db_event.dbo.tbl_event_mobile_addetc SET evt_tagkind = '"& evt_tagkind &"', evt_tagopt1 = '"& bannerCouponTxt &"' , etc_opt1 = '"& eName &"' , etc_opt2 = '"& subcopyK &"'  "&vbCrlf 
        strSql = strSql& "	WHERE  evt_code = "&eCode&" "&vbCrlf 
        strSql = strSql& "end"&vbCrlf 
        strSql = strSql& " ELSE "&vbCrlf
        strSql = strSql& "begin"&vbCrlf 
        strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_mobile_addetc (evt_code, evt_tagkind , evt_tagopt1 , etc_opt1 , etc_opt2 )"&vbCrlf 
        strSql = strSql& " VALUES("&eCode&", '"& evt_tagkind &"','"& bannerCouponTxt &"','"& eName &"','"& subcopyK &"')"&vbCrlf 
        strSql = strSql& "end"
        dbget.execute strSql
        '===========================================================

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[3]", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=1&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case "PU"
	'Ʈ����� (1.master����/2.disply����/3.MDTheme����)
    eScope = 2
    if eScope=2 then sPartnerid = ""

 	IF ((eState = 6 or eState = 7) and sOpenDate ="" ) THEN 	'����ó���� ����
		strAdd = ", [opendate] = getdate() " & vbCrlf
	ELSEIF (eState = 9 and sCloseDate ="" ) THEN
		strAdd = ", [closedate] = getdate() " & vbCrlf	'����ó���� ����
	END IF
		
	IF (eState = 3 and sImgregdate ="" ) THEN
		strAdd1 = ", [evt_imgregdate] = getdate() " & vbCrlf	'�̹�������� ����
	END IF

	'������ ������ ����� ������ ���� ��¥�� ����
	IF eState = 9 and  datediff("d",eEdate,date()) <0 THEN
			eEdate = date()
	END IF

	dbget.beginTrans
		'--1.master ����
		strSql = "UPDATE [db_event].[dbo].[tbl_event]" & vbCrlf
        strSql = strSql + " SET evt_kind=" & eKind & vbCrlf
        strSql = strSql + ", evt_manager=" & eManager & vbCrlf
        strSql = strSql + ", evt_name='" & eName & "'" & vbCrlf
        strSql = strSql + ", evt_startdate='" & eSdate & "'" & vbCrlf
        strSql = strSql + ", evt_enddate='" & eEdate & "'" & vbCrlf
        strSql = strSql + ", evt_level=" & eLevel & vbCrlf
        strSql = strSql + ", evt_lastupdate=getdate()" & vbCrlf
        strSql = strSql + ", adminid='" & session("ssBctId") & "'" & vbCrlf
        strSql = strSql + ", isWeb=" & isWeb & vbCrlf
        strSql = strSql + ", isMobile=" & isMobile & vbCrlf
        strSql = strSql + ", isApp=" & isApp & vbCrlf
        strSql = strSql + ", evt_subcopyK='" & subcopyK & "'" & vbCrlf
        strSql = strSql + ", evt_subname='" & subcopyK & "'" & vbCrlf
        strSql = strSql + ", evt_using='" & eusing & "'" & vbCrlf
        strSql = strSql + ", evt_state=" & eState & vbCrlf
        strSql = strSql + strAdd
        strSql = strSql + strAdd1
        strSql = strSql + " WHERE evt_code=" & eCode
		dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
            response.End 
        end if

        '--2.disply ����
        strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
        strSql = strSql + " SET bannerTypeDiv='" & bannerTypeDiv & "'" & vbCrlf
        strSql = strSql + ", bannerCouponTxt='" & bannerCouponTxt & "'" & vbCrlf
        strSql = strSql + ", bannerGubun='" & bannerGubun & "'" & vbCrlf
        strSql = strSql + ", etc_itemid='" & etcitemid & "'" & vbCrlf
        strSql = strSql + ", evt_tag='" & eTag & "'" & vbCrlf
        strSql = strSql + ", estimateSalePrice='" & estimateSalePrice & "'" & vbCrlf
        strSql = strSql + ", evt_itemsort='" & itemsort & "'" & vbCrlf
        strSql = strSql + ", eventtype_pc=80, eventtype_mo=80, mdtheme=5, mdthememo=5" & vbCrlf
        strSql = strSql + ", SalePer='" & eSalePer & "'" & vbCrlf
        strSql = strSql + ", SaleCPer='" & eSaleCPer & "'" & vbCrlf
        strSql = strSql + ", issale=" & eSale & vbCrlf
        strSql = strSql + ", isgift=" & eGift & vbCrlf
        strSql = strSql + ", iscoupon=" & eCoupon & vbCrlf
        strSql = strSql + ", isOnlyTen=" & eOnlyTen & vbCrlf
        strSql = strSql + ", isOneplusone=" & eOneplusone & vbCrlf
        strSql = strSql + ", isFreedelivery=" & eFreedelivery & vbCrlf
        strSql = strSql + ", isbookingsell=" & eBookingsell & vbCrlf
        strSql = strSql + ", isDiary=" & ediary & vbCrlf
        strSql = strSql + ", isNew=" & eisNew & vbCrlf
        strSql = strSql + ", endlessview='" & endlessview & "'" & vbCrlf
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
            response.End 
        end if

        '--3.MDTheme ����
        strSql = "UPDATE [db_event].[dbo].[tbl_event_md_theme]" & vbCrlf
        strSql = strSql + " SET title_pc='" & title_pc & "'" & vbCrlf
        strSql = strSql + ", title_mo='" & title_pc & "'" & vbCrlf
        if eComment then
            strSql = strSql & " ,comm_start='" & eSdate & "'" & vbCrlf
            strSql = strSql & " ,comm_end='" & eEdate & "'" & vbCrlf
        end if
        if eBbs then
            strSql = strSql & " ,board_start='" & eSdate & "'" & vbCrlf
            strSql = strSql & " ,board_end='" & eEdate & "'" & vbCrlf
        end if
        if eItemps then
            strSql = strSql & " ,eval_start='" & eSdate & "'" & vbCrlf
            strSql = strSql & " ,eval_end='" & eEdate & "'" & vbCrlf
        end if
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        '================ �̺�Ʈ ����� ��ǰ�̺�Ʈ =================
        '2015-11-04 ����ȭ �߰�
        if bannerTypeDiv="1" then
            evt_tagkind="7"
        elseif bannerTypeDiv="2" then
            evt_tagkind="2"
        elseif bannerTypeDiv="3" then
            evt_tagkind="1"
        elseif bannerTypeDiv="4" then
            evt_tagkind="4"
        elseif bannerTypeDiv="5" then
            evt_tagkind="3"
        elseif bannerTypeDiv="6" then
            evt_tagkind="5"
        elseif bannerTypeDiv="7" then
            evt_tagkind="6"
        end if
        strSql = "IF EXISTS(SELECT evt_code FROM db_event.dbo.tbl_event_mobile_addetc where evt_code = "&eCode&" )"&vbCrlf 
        strSql = strSql& "begin"&vbCrlf 
        strSql = strSql& " UPDATE db_event.dbo.tbl_event_mobile_addetc SET evt_tagkind = '"& evt_tagkind &"', evt_tagopt1 = '"& bannerCouponTxt &"' , etc_opt1 = '"& eName &"' , etc_opt2 = '"& subcopyK &"'  "&vbCrlf 
        strSql = strSql& "	WHERE  evt_code = "&eCode&" "&vbCrlf 
        strSql = strSql& "end"&vbCrlf 
        strSql = strSql& " ELSE "&vbCrlf
        strSql = strSql& "begin"&vbCrlf 
        strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_mobile_addetc (evt_code, evt_tagkind , evt_tagopt1 , etc_opt1 , etc_opt2 )"&vbCrlf 
        strSql = strSql& " VALUES("&eCode&", '"& evt_tagkind &"','"& bannerCouponTxt &"','"& eName &"','"& subcopyK &"')"&vbCrlf 
        strSql = strSql& "end"
        dbget.execute strSql
        '===========================================================

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
            response.End 
        end if

        vChangeContents = vChangeContents & "�̺�Ʈ UPDATE " & vbCrLf
        vChangeContents = vChangeContents & "- �̺�Ʈ�� : evt_name = " & eName & ", evt_code = " & eCode & vbCrLf
        vChangeContents = vChangeContents & "- ���� : evt_kind = " & eKind & vbCrLf
        vChangeContents = vChangeContents & "- Ÿ�� : ����issale = " & eSale & ", ����ǰisgift = " & eGift & ", ����iscoupon = " & eCoupon & ", isOnlyTen = " & eOnlyTen & ","
        vChangeContents = vChangeContents & " isOneplusone = " & eOneplusone & ", ������isFreedelivery = " & eFreedelivery & ", �����Ǹ�isbookingsell = " & eBookingsell & ","
        vChangeContents = vChangeContents & " isDiary = " & ediary & ", ��ĪisNew = " & eisNew & vbCrLf
        vChangeContents = vChangeContents & "- ��� : �ڸ�Ʈiscomment = " & eComment & ", �Խ���isbbs = " & eBbs & ", ��ǰ�ı�isitemps = " & eItemps & ", Blog URL isGetBlogURL = " & eisblogurl & vbCrLf
        vChangeContents = vChangeContents & "- �Ⱓ : evt_startdate ~ evt_enddate = " & eSdate & " ~ " & eEdate & vbCrLf
        vChangeContents = vChangeContents & "- ��÷��ǥ�� : evt_prizedate = " & ePdate & vbCrLf
        vChangeContents = vChangeContents & "- ���� : evt_state = " & eState & vbCrLf
        vChangeContents = vChangeContents & "- �߿䵵 : evt_level = " & eLevel & vbCrLf
        vChangeContents = vChangeContents & "- �̺�Ʈ�Ⱓ���⿩�� evt_dateview = " & eDateView & "" & vbCrLf
        vChangeContents = vChangeContents & "- ��ǥ��ǰ�����׹�� : ��ǥ��ǰ�ڵ� = " & etcitemid & vbCrLf

        '-�̺�Ʈ ���¿� ���� ����,����ǰ,���� ���� ����---------------
            IF (eState < 7) THEN  	'������ ���� �߱޴��� ���
                istatus = 0
            ELSEIF (eState <9) THEN
                istatus = 7
            ELSE
                istatus = eState
            END IF
            if eusing="N" then
                istatus = 7
            end if
        '--------------------------------------------------------------

        'IF egift = 0 THEN strgift = ", gift_using = 'N' "'����ǰüũ ���ҽ� �������� �ϴ� ����(2019.09.09-corpse2)

            strSql =" SELECT count(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code = "&eCode&" AND gift_using ='Y' "
            rsget.Open strSql, dbget
            IF not (rsget.EOF or rsget.BOF) THEN
                igiftcnt = rsget(0)
            END IF
            rsget.close
            
            ''��ü ���� �̺�Ʈ ���� CHECK
            strSql =" SELECT count(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code = "&eCode&" AND gift_scope in (1,9) "
            rsget.Open strSql, dbget
            IF not (rsget.EOF or rsget.BOF) THEN
                isAllGiftEvent = rsget(0)>0
            END IF
            rsget.close
            
            '��ü����/���̾������ ���� ����Ǹ� �ȵ�.
            if (isAllGiftEvent) then
                strgift = ""
            end if
            
            if igiftcnt > 0 then
            strSql ="	UPDATE [db_event].[dbo].[tbl_gift] Set gift_name = '"&eName&"' ,gift_startdate ='"&eSdate&"', gift_enddate ='"&eEdate&"', gift_status= "	&istatus&strAdd&_
                    "			, lastupdate = getdate(), adminid = '"&session("ssBctId")&"', site_scope= "&eScope&", partner_id='"&sPartnerid&"' "&strgift&_
                    "		WHERE evt_code = "&eCode
            
            if (istatus=0) then ''��ü����/���̾������ ���� ����Ǹ� �ȵ�.
                strSql = strSql&"  and gift_scope not in (1,9)"  
            end if
            
            dbget.execute strSql
            end if

        '-- sale Ȯ��
            IF eSale = 0 THEN strSale = ", sale_using = 0 "
            strSql = " SELECT sale_code, sale_status FROM [db_event].[dbo].[tbl_sale] WHERE evt_code = "&eCode&" AND sale_using =1 "
            rsget.Open strSql, dbget
            IF not (rsget.EOF or rsget.BOF) THEN
                arrSale = rsget.getRows()
            END IF
            rsget.close

            IF isarray(arrSale)  THEN
                For intSale = 0 To UBound(arrSale,2)
                '������ ��� ���»��°� 6, ������°� 8 �̹Ƿ� ���°� ���� �ʿ�
                if (eState = 7 AND arrSale(1,intSale) >= 6) OR ( eState > 7 AND arrSale(1,intSale) >= 8 )  THEN		istatus = arrSale(1,intSale)
                    strSql ="	UPDATE [db_event].[dbo].[tbl_sale] Set sale_name = '"&eName&"', sale_startdate ='"&eSdate&"', sale_enddate ='"&eEdate&"', sale_status="	&istatus&strAdd&_
                            "			, lastupdate = getdate(), adminid = '"&session("ssBctId")&"'"&strSale&_
                            "		WHERE evt_code = "&eCode&" and sale_code = "&arrSale(0,intSale)
                    dbget.execute strSql
                Next
            END IF

        IF Err.Number = 0 THEN
            '### ���� �α� ����(event)
            vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
            vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & menupos & "', "
            vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
            dbget.execute(vSCMChangeSQL)
        ELSE
            dbget.RollBackTrans
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
        END IF
    '===========================================================
	dbget.CommitTrans

	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=1&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case else
	Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->