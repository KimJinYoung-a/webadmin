<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �귣�� ��� ����
' Hieditor : 2009.04.07 ������ ����
'			 2010.05.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->

<%
dim mode, makerid, contractType, contractID , contractState , contractEtcContetns , onoffgubun
dim mailfrom, mailto, mailtitle, mailcontent, innerContents ,CurrState,NextState, sendOpenMail
	mode            = request("mode")
	makerid         = request("makerid")
	contractType    = request("contractType")
	contractID      = request("contractID")
	contractEtcContetns = request("contractEtcContetns")
	CurrState       = request("CurrState")
	NextState       = request("NextState")
	sendOpenMail    = request("sendOpenMail")

dim sqlStr , objItem, contractExists , contractContents ,contractNo, contractName, HtmlcontractEtcContetns
dim bufStr, refer
dim ocontract
refer = request.ServerVariables("HTTP_REFERER")

'//�űԵ��
if (mode="regContract") then

    sqlStr = "select contractContents, contractName ,onoffgubun" +vbcrlf
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
    sqlStr = sqlStr & " where contractType=" & contractType
    
    'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        contractContents = db2Html(rsget("contractContents"))
        contractName = db2Html(rsget("contractName"))
        onoffgubun = rsget("onoffgubun")
    end if
    rsget.Close
	
	'//�¶��� ��༭�� ���� ��ϵǾ� �ִ� ��༭ �ִ��� Check
	if onoffgubun = "ON" then    
	    sqlStr = "select count(contractID) as cnt from db_partner.dbo.tbl_partner_contract"
	    sqlStr = sqlStr & " where makerid='" & makerid & "'"
	    sqlStr = sqlStr & " and contractType=" & contractType
	    sqlStr = sqlStr & " and contractState>=0"
	    sqlStr = sqlStr & " and contractState<7"
	    rsget.Open sqlStr,dbget,1
	    if Not rsget.Eof then
	        contractExists = rsget("cnt")>0
	    end if
	    rsget.Close
	    
	    if (contractExists) then
	        response.write "<script>alert('�̹� �������� ���� ������ ������ �ֽ��ϴ�.\n�������� �����Ǵ� �Ϸ��� ��ϰ����մϴ�.');history.back();</script>"
	        dbget.close()	:	response.End
	    end if
	end if
	
    sqlStr = " select * from db_partner.dbo.tbl_partner_contract where 1=0"
    rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	    rsget("makerid")            = makerid       
	    rsget("contractType")       = contractType
        rsget("contractState")      = 0
        rsget("contractName")       = Newhtml2db(contractName)
        rsget("contractEtcContetns")= contractEtcContetns
        rsget("reguserid")          = session("ssBctID")

	rsget.update
	    contractID = rsget("contractID")
	rsget.close
   
    For Each objItem In Request.Form
        ''response.write objItem & "," & Request.Form(objItem) & "<br>"
        if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
            sqlStr = " insert into db_partner.dbo.tbl_partner_contractDetail"
            sqlStr = sqlStr & " (contractID, detailKey, detailValue)"
            sqlStr = sqlStr & " values("
            sqlStr = sqlStr & " " & contractID
            sqlStr = sqlStr & " ,'" & objItem & "'"
            sqlStr = sqlStr & " ,'" & Newhtml2db(Request.Form(objItem)) & "'"
            sqlStr = sqlStr & " )"
            
            dbget.Execute sqlStr
            
            if (objItem="$$CONTRACT_DATE$$") then
                bufStr  = Request.Form(objItem)
                bufStr  = Left(bufStr,4) & "��" & Mid(bufStr,6,2) & "��" & Mid(bufStr,9,2) & "��"
                contractContents = Replace(contractContents,objItem,bufStr)
            else
                contractContents = Replace(contractContents,objItem,Request.Form(objItem))
            end if
            
            if (objItem="$$CONTRACT_DATE$$") then contractNo=Request.Form(objItem)
        end if
    Next
    
    ''��Ÿ������
    if Trim(contractEtcContetns)<>"" then
        HtmlcontractEtcContetns = "<p style='margin:0cm;margin-bottom:.0001pt;text-align:justify;text-justify:"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "inter-ideograph;punctuation-wrap:simple;word-break:break-hangul'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<b><span style='font-size:11.0pt;font-family:����;color:windowtext;layout-grid-mode:line'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "- ��Ÿ������</span></b></p>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<br>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<p class=MsoNormal style='margin-left:5.0pt'><span style='font-size:11.0pt;font-family:����;color:windowtext;layout-grid-mode:line'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & replace(contractEtcContetns,VbCrlf,"<br>")
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "</span></p>"
        
        contractContents = Replace(contractContents,"$$ContractEtcContetns$$",HtmlcontractEtcContetns)
    else
        contractContents = Replace(contractContents,"$$ContractEtcContetns$$","")
    end if
    
    ''��༭ ��ȣ ����. YYYYMMDD(�����)-contractType-contractID
    contractNo = Replace(contractNo,"-","") & "-" & contractType & "-" & contractID
    contractContents = Replace(contractContents,"$$CONTRACT_NO$$",contractNo)

    sqlStr = " update db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
    sqlStr = sqlStr & " ,contractNo='" & contractNo & "'"
    sqlStr = sqlStr & " where contractID=" & contractID
    
    dbget.Execute sqlStr
        
    response.write "<script>alert('��� �Ǿ����ϴ�.\n\nȮ�� �Ͻ��� ��ü����(�߼�) ���� �����Ͻñ� �ٶ��ϴ�.');</script>"
    response.write "<script>location.replace('/admin/member/contractReg.asp?makerid="& makerid & "&ContractID=" & ContractID & "');</script>"
    dbget.close()	:	response.End

'//����
elseif (mode="editContract") then
    ''���� ���ɻ��� Check
    sqlStr = "select contractType, contractState from db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " where  contractID=" & contractID
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        contractState   = rsget("contractState")
        contractType    = rsget("contractType")
        contractExists = (contractState>=3)
    end if
    rsget.Close
    
    if (contractExists) then
        response.write "<script>alert('���� ���� ���°� �ƴմϴ�.\n������ ���� ���.');history.back();</script>"
        dbget.close()	:	response.End
    end if
    
    
    sqlStr = "select t.contractContents, t.contractName from "
    sqlStr = sqlStr & " db_partner.dbo.tbl_partner_contractType t,"
    sqlStr = sqlStr & " db_partner.dbo.tbl_partner_contract c"
    sqlStr = sqlStr & " where c.contractID=" & contractID
    sqlStr = sqlStr & " and c.contractType=t.contractType"
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        contractContents = db2Html(rsget("contractContents"))
        contractName = db2Html(rsget("contractName"))
    end if
    rsget.Close
    
    
    
    sqlStr = " update db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " set contractEtcContetns='" & Newhtml2db(contractEtcContetns) & "'"
    sqlStr = sqlStr & " where contractID=" & contractID
    
    dbget.Execute sqlStr
    
    For Each objItem In Request.Form
        ''response.write objItem & "," & Request.Form(objItem) & "<br>"
        if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
            sqlStr = " update db_partner.dbo.tbl_partner_contractDetail"
            sqlStr = sqlStr & " set detailValue='" & Newhtml2db(Request.Form(objItem)) & "'"
            sqlStr = sqlStr & " where contractID=" & contractID
            sqlStr = sqlStr & " and detailKey='" & objItem & "'"
            
            dbget.Execute sqlStr
            
            if (objItem="$$CONTRACT_DATE$$") then
                bufStr  = Request.Form(objItem)
                bufStr  = Left(bufStr,4) & "��" & Mid(bufStr,6,2) & "��" & Mid(bufStr,9,2) & "��"
                contractContents = Replace(contractContents,objItem,bufStr)
            else
                contractContents = Replace(contractContents,objItem,Request.Form(objItem))
            end if
            
            if (objItem="$$CONTRACT_DATE$$") then contractNo=Request.Form(objItem)
        end if
    Next
    
    ''��Ÿ������
    if Trim(contractEtcContetns)<>"" then
        HtmlcontractEtcContetns = "<p style='margin:0cm;margin-bottom:.0001pt;text-align:justify;text-justify:"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "inter-ideograph;punctuation-wrap:simple;word-break:break-hangul'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<b><span style='font-size:11.0pt;font-family:����;color:windowtext;layout-grid-mode:line'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "- ��Ÿ������</span></b></p>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<br>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<p class=MsoNormal style='margin-left:5.0pt'><span style='font-size:11.0pt;font-family:����;color:windowtext;layout-grid-mode:line'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & replace(contractEtcContetns,VbCrlf,"<br>")
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "</span></p>"
        
        contractContents = Replace(contractContents,"$$ContractEtcContetns$$",HtmlcontractEtcContetns)
    else
        contractContents = Replace(contractContents,"$$ContractEtcContetns$$","")
    end if
    
    ''��༭ ��ȣ ����. YYYYMMDD(�����)-contractType-contractID
    contractNo = Replace(contractNo,"-","") & "-" & contractType & "-" & contractID
    contractContents = Replace(contractContents,"$$CONTRACT_NO$$",contractNo)

    sqlStr = " update db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
    sqlStr = sqlStr & " ,contractNo='" & contractNo & "'"
    sqlStr = sqlStr & " ,contractName='" & Newhtml2db(contractName) & "'"
    if (contractState=-2) then
        sqlStr = sqlStr & " ,contractState=0"
    end if
    sqlStr = sqlStr & " where contractID=" & contractID
    
    dbget.Execute sqlStr
    
    response.write "<script>alert('���� �Ǿ����ϴ�.\n\nȮ�� �Ͻ��� ��߼� �Ͻñ� �ٶ��ϴ�.');</script>"
    response.write "<script>opener.location.reload(); location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
    
elseif (mode="stateChange") then
    ''CurrState, NextState

    set ocontract = new CPartnerContract
	    ocontract.FRectContractID = ContractID
	    ocontract.FRectMakerid = makerid
	    ocontract.getOneContract
	        
	    contractName = ocontract.FOneItem.FcontractName
	    contractNo   = ocontract.FOneItem.FcontractNo
	    contractType = ocontract.FOneItem.FcontractType
           
    ''�¿��� �����ε�
    sqlStr = "select contractContents, contractName ,onoffgubun" +vbcrlf
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
    sqlStr = sqlStr & " where contractType=" & contractType
    
    'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        onoffgubun = rsget("onoffgubun")
    end if
    rsget.Close
         
    ''�����ΰ�� , ���Ϲ߼� Check �Ǿ� ������ 
    if (NextState="1") and (sendOpenMail="on") then
        sqlStr = "select IsNULL(p.email,'') as email from [db_partner].[dbo].tbl_partner p,"
        sqlStr = sqlStr & " db_partner.dbo.tbl_partner_contract c"
        sqlStr = sqlStr & " where c.contractID=" & contractID
        sqlStr = sqlStr & " and c.makerid=p.id"
        
        'response.write sqlStr &"<br>"
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            mailto = db2Html(rsget("email"))
        end if
        rsget.Close
        
        mailto = Trim(mailto)

        if (mailto="") or (InStr(mailto,"@")<0) or (Len(mailto)<8) then
            response.write "<script>alert('��ü ����� E���� �ּҰ� ��ȿ���� �ʽ��ϴ�.\n�귣���������� E���� ���� �� ����Ͻñ� �ٶ��ϴ�.');</script>"
            response.write "<script>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if
        
        sqlStr = "select IsNULL(p.usermail,'') as email from db_partner.dbo.tbl_user_tenbyten p,"
        sqlStr = sqlStr & " db_partner.dbo.tbl_partner_contract c"
        sqlStr = sqlStr & " where c.contractID=" & contractID
        sqlStr = sqlStr & " and c.reguserid=p.userid"
        
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            mailfrom = db2Html(rsget("email"))
        end if
        rsget.Close
        
        mailfrom = Trim(mailfrom)
        
        if (mailfrom="") or (InStr(mailfrom,"@")<0) or (Len(mailfrom)<8) then
            response.write "<script>alert('10x10 ����� E���� �ּҰ� ��ȿ���� �ʽ��ϴ�.���� �������� E���� ���� �� ����Ͻñ� �ٶ��ϴ�.');</script>"
            response.write "<script>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if
             
        mailtitle = "[�ٹ�����]��ü ��༭�� ���� �Ǿ����ϴ�."
        innerContents = ""
        
        '' ������ ���� ���� ������
        innerContents = innerContents & " �ȳ��ϼ���" & "<br>"
        innerContents = innerContents & "(��)�ٹ����ٰ� ���� �ο����� ������ �Ǿ� �ݰ����ϴ�." & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "�Ʒ��� ���� ����� ����ǿ��� " & "<br>"
        innerContents = innerContents & "��༭ ��������� �Ĳ��� �о��ֽ� �� " & "<br>"
        innerContents = innerContents & "������ ���߾� ��༭�� �������� �߼��� �ֽø� �����ϰڽ��ϴ�." & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "��༭ �� : " & contractName  & "<br>"
        innerContents = innerContents & "��༭ ��ȣ : " & contractNo  & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        
        if onoffgubun = "ON" then
	        innerContents = innerContents & "�� ��༭ ������ : �¶��� ���� �� �귣�� " & "<br>"
	        innerContents = innerContents & "- �������ο��� �����ϴ� �귣��� ��󿡼� ���ܵ˴ϴ�. " & "<br>"
	        innerContents = innerContents & "(�������� �����ÿ��� �������� ����ڰ� ���������� �����帳�ϴ�.)" & "<br>"
    	else
       		innerContents = innerContents & "�� ��༭ ������ : �������� ���� �� �귣�� " & "<br>"    	
    	end if
        
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "�� ��༭ �ٿ��� " & "<br>"
        innerContents = innerContents & "- �Ʒ� [��༭ �ٿ�ε�] Ŭ���Ͽ� �ٿ���� �� ����Ȯ�� �� ��������� �������ּ���!!" & "<br>"
        innerContents = innerContents & " (�ٽ� �ٿ�ε� �����÷��� ���� �α��� �� ���� ��� [��ü��༭ �ٿ�ε�]�� �̿��Ͽ� �ּ��� )" & "<br>"
        innerContents = innerContents & "- ��༭ ���� �ϴ� ����� [���flow �ٿ�ε�] �ٿ������ �� �� ������ ���ֽñ� �ٶ��ϴ�." 
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<a href='" & manageUrl & "/designer/company/popContract.asp?ContractID=" & ContractID & "' target='_blank'><b><font color=blue>[��ü��༭ �ٿ�ε�]</font></b></a>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<a href='" & manageUrl & "/designer/company/contractflow.ppt' target='_blank'><b><font color=blue>[���flow �ٿ�ε�]</font></b></a>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"

        innerContents = innerContents & "�� �ʼ� Ȯ�λ��� (�ݵ�� �ι� ���� Ȯ�����ּ���) " & "<br>"
        innerContents = innerContents & "������, ������ �� �ΰ����� �´� �� �� Ȯ�����ּž� �մϴ�." & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "�� ��ü�������� �ʼ� ������� (��!! ���� �����ϼž� �� �κ�)" & "<br>"
        innerContents = innerContents & "- ǥ��(ù��)�� ������� ���� : ���¾�ü�� ��ǥ�̻� �Ǵ� ����� ������ �����Ͻô� ����� ����" & "<br>"

        if (contractType=5) then
			innerContents = innerContents & "- ���å���� ����" & "<br>"
        end if

        innerContents = innerContents & "- ������ ���� '��'�� ��ǥ�̻� �ֹε�Ϲ�ȣ �� �ּ� ���� : ����ڵ������ ��ǥ�� �ֹι�ȣ �� �ּҿ��� �մϴ�." & "<br>"
        innerContents = innerContents & "- ���λ������ ��� ���������� ��ǥ�̻� �ֹι�ȣ �� �ּҸ� �����ϼŵ� �Ǹ�, '��'����� ���� '��' ����� ������ ������ �˴ϴ�." & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "�� �������� : " & "<br>"
        
        innerContents = innerContents & "�� ��༭�ٿ�ε�" & "<br>"
        innerContents = innerContents & "�� ���¾�ü���� ��༭ Ȯ���� ���� / 2�� ����߼� " & "<br>"
        innerContents = innerContents & "�� �ٹ����ٿ��� ��༭ ���� ����Ȯ��" & "<br>"
        innerContents = innerContents & "�� �ٹ����ٿ��� ���¾�ü�� ��༭ 1�� �߼� / ���Ϸ�" & "<br>"
        
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "�� ��༭ �����ô� ��" & "<br>"
        
        if onoffgubun = "ON" then     
            innerContents = innerContents & "�ּ� : (03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� " & "<br>"
            innerContents = innerContents & "����� : " & ocontract.FoneItem.Fusername  & "<br>"
	        innerContents = innerContents & "tel : " & ocontract.FoneItem.Finterphoneno  & " (���� " & ocontract.FoneItem.Fextension & ") / ���� : "& ocontract.FoneItem.Fdirect070 &"<br>"
	        innerContents = innerContents & "fax : 02-2179-9244 <br>"	            
        
        '/�¶��� ��࿡ ���������� ���� ���� ����.. ������ ����� ���� �Ұ���.. �ھƳ���
        else
            innerContents = innerContents & "�ּ� : (03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� �������� �繫�� " & "<br>"
            innerContents = innerContents & "����� : �̿��� �븮<br>"
	        innerContents = innerContents & "tel : 02-554-2033 (���� 222) / ���� : 070-7515-5422<br>"
            innerContents = innerContents & "fax : 02-2179-9058 <br>"
            innerContents = innerContents & "mail: john6136@10x10.co.kr<br>"
        end if

        innerContents = innerContents & "<br>"    
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "�� ����߼۽� �Բ� �����ž� �� ����" & "<br>"
        innerContents = innerContents & "- ���ε� ��༭ 2��" & "<br>"
        innerContents = innerContents & "- �������� �纻" & "<br>"
        innerContents = innerContents & "- ����� ����� �纻" & "<br>"
        innerContents = innerContents & "- �ΰ����� ���� (��༭�� ������ ����)" & "<br>"
        
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "�� �� Ÿ " & "<br>"
        innerContents = innerContents & "�ٹ����� ���� �����ϴ� �귣�� ���̵� 2�� �̻��� ��� " & "<br>"
        innerContents = innerContents & "��༭�� �� �귣�� ���̵𸶴� �ۼ��� ���ּž� �ϸ�, " & "<br>"
        innerContents = innerContents & "���ü���(����ڵ����,�ΰ�����,��������)�� 1�θ� �ּŵ� �˴ϴ�. " & "<br>"
        innerContents = innerContents & "���� ���̵�� �н����带 �ο����� ������ ��� ��翥�𿡰� ������ �ֽñ� �ٶ��ϴ�. " & "<br>"
               
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "�� ��༭ ���� ���� �ñ��� ���� �� ���MD���� ���� �Ͻñ� �ٶ��ϴ�."
        
        innerContents = innerContents & "<br>"

        mailcontent = "<html>"
    	mailcontent = mailcontent + "<head>"
    	mailcontent = mailcontent + "<title>�ٹ����� ��ü ��༭ ����</title>"
    	mailcontent = mailcontent + "<meta http-equiv='Content-Type' content='text/html; charset=euc-kr'>"
    	mailcontent = mailcontent + "<style>"
    	mailcontent = mailcontent + ".text {"
    	mailcontent = mailcontent + ""
    	mailcontent = mailcontent + "font-family: 'Verdana', 'Arial', 'Helvetica', 'sans-serif';"
    	mailcontent = mailcontent + "font-size: 12px;"
    	mailcontent = mailcontent + "line-height: 130%;"
    	mailcontent = mailcontent + "color: #333333;"
    	mailcontent = mailcontent + "}"
    	mailcontent = mailcontent + "</style>"
    	mailcontent = mailcontent + "</head>"
    	mailcontent = mailcontent + "<body bgcolor=#FFFFFF leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
    	mailcontent = mailcontent + "<table width=573 border=0 cellpadding=0 cellspacing=0>"
    	mailcontent = mailcontent + "<tr>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_01.gif' width='45' height='114'></td>"
    	mailcontent = mailcontent + "<td> <a href='http://www.10x10.co.kr' target='_blank'><img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_02.gif' width='479' height='114' border='0'></a></td>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_03.gif' width='49' height='114'></td>"
    	mailcontent = mailcontent + "</tr>"
    	mailcontent = mailcontent + "<tr>"
    	mailcontent = mailcontent + "<td background='http://www.10x10.co.kr/apps/mail_form/images/mail_form_04.gif'></td>"
    	mailcontent = mailcontent + "<td bgcolor='#F7F7F7' align='center'>"
    	mailcontent = mailcontent + "<table border='0' cellpadding='0' cellspacing='0' height='200' width='90%'>"
    	mailcontent = mailcontent + "<tr>"
    	mailcontent = mailcontent + "<td class='text'>"
    	mailcontent = mailcontent + innerContents
    	mailcontent = mailcontent + "</td>"
    	mailcontent = mailcontent + "</tr>"
    	mailcontent = mailcontent + "</table>"
    	mailcontent = mailcontent + "</td>"
    	mailcontent = mailcontent + "<td background='http://www.10x10.co.kr/apps/mail_form/images/mail_form_06.gif'></td>"
    	mailcontent = mailcontent + "</tr>"
    	mailcontent = mailcontent + "<tr>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_07.gif' width='45' height='107'></td>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_10.gif' width='479' height='107'></td>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_09.gif' width=49 height='107'></td>"
    	mailcontent = mailcontent + "</tr>"
    	mailcontent = mailcontent + "</table>"
    	mailcontent = mailcontent + "</body>"
    	mailcontent = mailcontent + "</html>"
    	
        Call SendMail(mailfrom, mailto, mailtitle, mailcontent)
        
        response.write "<script>alert('(" & mailfrom & ")���Ϸ� ��ü ����ڿ��� �̸���(" & mailto & ")�� �߼� �Ͽ����ϴ�.');</script>"
    
    end if
    set ocontract = Nothing
    
    sqlStr = " update db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " set contractState=" & NextState & ""
    if (NextState="7") then
        sqlStr = sqlStr & " ,finishdate=getdate()"
    elseif (NextState="0") then
        sqlStr = sqlStr & " ,confirmdate=NULL"
    end if
    sqlStr = sqlStr & " where contractID=" & contractID
    
    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
    
    
    response.write "<script>alert('���� ������ ���� �Ͽ����ϴ�.');</script>"
    response.write "<script>opener.location.reload(); location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
    
else
    response.write "<script>alert('���ǵ��� �ʾҽ��ϴ�. - " & mode & "');</script>"
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->