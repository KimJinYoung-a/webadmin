<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- include virtual="/admin/offshop/shopcscenter/cscenter_mail_Function_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->
<%
dim mode, modeflag2, divcd, reguserid, ipkumdiv ,title, orderno, contents_jupsu
dim finishuser, contents_finish ,requireupche, requiremakerid, ForceReturnByTen ,detailitemlist
dim opentitle, opencontents ,newasid ,isCsMailSend ,IsAllCancel ,CancelValidResultMessage
dim sqlStr, i ,ScanErr ,ResultMsg, ReturnUrl, EtcStr ,ProceedFinish ,returnmethod ,masteridxtmp
dim oordermaster ,buf_requiremakerid , masteridx , csmasteridx , cancelorderno ,GC_IsOLDOrder
dim reqname ,reqphone ,reqhp , reqzipcode ,reqzipaddr ,reqaddress ,comment ,reqemail
	masteridx        = requestCheckVar(request.Form("masteridx"),10)
	csmasteridx   = requestCheckVar(request.Form("csmasteridx"),10)
	mode        = requestCheckVar(request.Form("mode"),32)
	modeflag2   = requestCheckVar(request.Form("modeflag2"),10)
	divcd       = requestCheckVar(request.Form("divcd")	,4)
	ipkumdiv    = requestCheckVar(request.Form("ipkumdiv"),10)
	reguserid   = session("ssbctid")
	finishuser  = reguserid
	title       = requestCheckVar(html2DB(request.Form("title")),128)
	orderno = requestCheckVar(request.Form("orderno"),16)
	contents_jupsu  = requestCheckVar(html2DB(request.Form("contents_jupsu")),800)
	detailitemlist  = html2db(request.Form("detailitemlist"))
	contents_finish = requestCheckVar(html2DB(request.Form("contents_finish")),800)
	requireupche = requestCheckVar(request.Form("requireupche"),1)
	requiremakerid = requestCheckVar(request.Form("requiremakerid"),32)
	ForceReturnByTen = requestCheckVar(request.Form("ForceReturnByTen"),32)
	buf_requiremakerid  = requestCheckVar(request.Form("buf_requiremakerid"),32)
	isCsMailSend = requestCheckVar((request.Form("csmailsend")="on"),10)
	cancelorderno       = requestCheckVar(request.Form("cancelorderno"),16)
	reqname       = requestCheckVar(request.Form("reqname"),32)
	reqphone       = requestCheckVar(request.Form("reqphone"),32)
	reqhp       = requestCheckVar(request.Form("reqhp"),32)
	reqzipcode       = requestCheckVar(request.Form("reqzipcode"),7)
	reqzipaddr       = requestCheckVar(request.Form("reqzipaddr"),128)
	reqaddress       = requestCheckVar(request.Form("reqaddress"),255)
	comment       = request.Form("comment")
	reqemail       = requestCheckVar(request.Form("reqemail"),128)
	
	newasid = -1		
	if (returnmethod="") then returnmethod="R000"
	ScanErr = ""
	ProceedFinish = False

''�ֹ� ����Ÿ
set oordermaster = new COrder
	oordermaster.FRectmasteridx = masteridx	
	oordermaster.fQuickSearchOrderMaster

'response.write "mode : " & mode & "<br>"

'/cs����
if (mode="regcsas") then
	'response.write "divcd : " & divcd & "<br>"

    'CS ���� - ��üa/s
	if (divcd="A030") then

        dbget.beginTrans
 
		'a/s �� ��� ��üa/s ��..   ��üa/s(����ȸ��)�� �ֽ�..  �Ѵ� �귣��id ���� �ؾ���
        if (requiremakerid<>"") then

			'' CS Master ���� ''html2db ������� ����.
	        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)
	   
			'CS Detail ����(���� ��ǰ����)
			Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx ,"Y")
		        	
        	'/��ü a/s ���� ��ü��ۺ��� requireupche : Y
            call RegCSMasterAddUpche_off(csmasteridx, requiremakerid)

			'����� ���(���ּҳ� �����ּ�)
			call Regdelivery_off(csmasteridx, reqname ,reqphone ,reqhp ,reqemail,reqzipcode ,reqzipaddr ,reqaddress ,comment)
			
        	'��üa/s(����ȸ��)
            '' CS Master ���� ''html2db ������� ����.
            newasid = RegCSMaster_off("A031", orderno, reguserid, "��üA/S(����ȸ��)", contents_jupsu, masteridx)
			
			'CS Detail ����(���� ��ǰ����)			
            Call AddCSDetailByArrStr_off(detailitemlist, newasid, orderno ,masteridx ,"N")
			
			'/�±�ȯ ȸ�� ���� ������ ���� requiremaejang : Y
        	call RegCSMasterAddmaejang_off(newasid, requiremakerid)

			'����� ���(��ü�ּ�)
			if isarray(Getpartnerdeliverinfo_off(requiremakerid,"")) then
				call Regdelivery_off(newasid, Getpartnerdeliverinfo_off(requiremakerid,"")(2,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(3,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(4,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(5,0),Getpartnerdeliverinfo_off(requiremakerid,"")(6,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(7,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(8,0) ,comment)			
			end if

			ResultMsg = "��üA/S ���� �� ��üA/S(����ȸ��) ���� �Ϸ�"
        end if

        ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script type='text/javascript'>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"            
            dbget.close()	:	response.End
        End If

    else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
        response.write "<script type='text/javascript'>alert('" + ResultMsg + "');</script>"
        response.write "<script type='text/javascript'>history.back();</script>"
        dbget.close()	:	response.End
    end if

''���� ���� ����
elseif (mode="editcsas") then

    dbget.beginTrans
	
	'' CS master ����
	Call EditCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu ,csmasteridx)

    '' CS Detail ����
    Call EditCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno)

	'����� ���(���ּҳ� �����ּ�)
	call Regdelivery_off(csmasteridx, reqname ,reqphone ,reqhp ,reqemail,reqzipcode ,reqzipaddr ,reqaddress ,comment)
			
    ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd + "&mode=editreginfo"
		
    If (Err.Number = 0) and (ScanErr="") Then
        dbget.CommitTrans
    
    	ResultMsg = ResultMsg + "OK"
    Else
        dbget.RollBackTrans
        response.write "<script type='text/javascript'>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"        
        dbget.close()	:	response.End
    End If

'CS ���� ���� �Ϸ�ó��
elseif (mode="finishcsas") then	    
	'response.write "divcd : " & divcd & "<br>"
			
	'CS ���� ���� �Ϸ�ó�� - �±�ȯ ��� / ���� / ���� �߼� / ��Ÿ /  ���� ���ǻ���	/	��üa/s /	��üa/s(����ȸ��)
    if  (divcd="A000") or (divcd="A001") or (divcd="A002") or (divcd="A009") or (divcd="A006") or (divcd="A005") or (divcd="A700") or (divcd="A030") or (divcd="A031") then
	
		dbget.beginTrans
		
		Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)
		
		ResultMsg = "ó�� �Ϸ�"
		ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
		
		If (Err.Number = 0) and (ScanErr="") Then
		    dbget.CommitTrans
		Else
		    dbget.RollBackTrans
		    response.write "<script type='text/javascript'>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"	        
		    dbget.close()	:	response.End
		End If

	else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
        response.write "<script type='text/javascript'>alert('" + ResultMsg + "');</script>"
        response.write "<script type='text/javascript'>history.back();</script>"
        dbget.close()	:	response.End
    end if

'' ��ü ó���Ϸ� => �������·κ���
elseif (mode="upcheconfirm2jupsu") then
	    
    sqlStr = " select top 1 currstate from db_shop.dbo.tbl_shopjumun_cs_master"
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)
	
	'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if not(rsget("currstate")="B006" or rsget("currstate")="B008") then
	            ResultMsg = "��üó���Ϸᳪ ����ó���Ϸ� ���°� �ƴմϴ�. ���� �Ұ�"                
	        end if
		else
		    ResultMsg = "�ڵ����. ���� �Ұ�"
		end if
	rsget.Close

    if (ResultMsg="") then
        sqlStr = " update db_shop.dbo.tbl_shopjumun_cs_master" + VbCrlf
        sqlStr = sqlStr + "set currstate='B001'" + VbCrlf
        sqlStr = sqlStr + ",contents_jupsu='" + (contents_jupsu) + "'" + VbCrlf
        sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)

		'response.write sqlStr &"<br>"        
        dbget.Execute sqlStr

        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
    else
        response.write "<script type='text/javascript'>alert('" + ResultMsg + "');</script>"
        response.write "<script type='text/javascript'>history.back();</script>"
        dbget.close()	:	response.End
    end if
    
'CS ����
elseif (mode="deletecsas") then
	
    ''Check Valid Delete - ����� B006 ��üó���Ϸ� , B007 �Ϸ� ������ ���(����) �Ұ�
    if (NOT ValidDeleteCS_off(csmasteridx)) then
        response.write "<script type='text/javascript'>alert(" + Chr(34) + "���� ��� ���� ���°� �ƴմϴ�. ������ ���� ���." + Chr(34) + ")</script>"
        response.write "<script type='text/javascript'>history.back()</script>"
        dbget.close()	:	response.End
    end if

    If Not DeleteCSProcess_off(csmasteridx, reguserid) then
        ResultMsg = ResultMsg + "������ ������ ����"
    else
        ResultMsg = ResultMsg + "OK"
    End if
    
    ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd + "&mode=editreginfo"
   
end if

%>

<%
response.write "<script type='text/javascript'>alert('" + ResultMsg + "');</script>"
response.write "<script type='text/javascript'>location.replace('" + ReturnUrl + "');</script>"
response.End
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->