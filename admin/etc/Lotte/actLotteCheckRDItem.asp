<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
	'// ��������
	dim LotteGoodNo, LotteStatCd
	dim strSql, actCnt, lp, waitCnt, rjtCnt
    dim AssignedCNT
    dim param2 : param2=request("param2")

	actCnt = 0			'�ǰ��ŰǼ�
    waitCnt = 0
	on Error Resume Next

	'// �Ե����� �ӽõ�� ��ǰ ��ȸ(������,���ο�û,�����,������û / ���οϷ�,�ݷ�,���κҰ� ����)
	strSql = "Select top 100 itemid, LotteTmpGoodNo From db_item.dbo.tbl_lotte_regItem "
	strSql = strSql & " Where LotteStatCd in ('10','20','51','52') "
	'strSql = strSql & " Where LotteStatCd in ('10','51','52') "		'���ο�û�� ����(ù���:10, ���ο�û����:20 �̹Ƿ�, ��ü ��ǰ�� ������ ���� �ʾ��� ��� ���� ��ǰ�� ��� Ȯ���ϰ� �Ǵ� ��Ȳ�� �߻��Ǿ� ������ 10���� ���� �� ������ ���� ���� ���)
	strSql = strSql & " and dateDiff(hh,IsNULL(lastConfirmdate,'2001-01-01'),getdate())>5"  ''���ο�û���¿��� ��� ��û�ϹǷ�.
	strSql = strSql & " and LotteTmpGoodNo is Not NULL"
	''strSql = strSql & " and LotteTmpGoodNo='34561137'"
	strSql = strSql & " order by IsNULL(lastConfirmdate,'2001-01-01') asc, regdate asc" ''asc"  ù��� ����. LotteStatCd,

	if (param2="0") then
	    strSql = "Select top 100 r.itemid, r.LotteTmpGoodNo From db_item.dbo.tbl_lotte_regItem r"
	    strSql = strSql & "     Join db_item.dbo.tbl_item i"
	    strSql = strSql & "     on r.itemid=i.itemid"
	    strSql = strSql & "     and (i.sellyn<>'Y' or (i.sellcash>isNULL(lottePrice,0)))"
	    strSql = strSql & " Where r.LotteStatCd in ('20') " ''���ο�û
	    strSql = strSql & " and dateDiff(hh,IsNULL(r.lastConfirmdate,'2001-01-01'),getdate())>5"  ''���ο�û���¿��� ��� ��û�ϹǷ�.
	    strSql = strSql & " and r.LotteTmpGoodNo is Not NULL"
	    strSql = strSql & " order by IsNULL(r.lastConfirmdate,'2001-01-01') asc, r.regdate asc"
	end if

	rsget.Open strSql,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		'// �Ե����� ���õ�� ��ǰ ��ȸ
		Do Until rsget.EOF
			Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "GET", lotteAPIURL & "/openapi/getRdToPrGoodsNoApi.lotte?subscriptionId=" & lotteAuthNo & "&goods_req_no=" & rsget("LotteTmpGoodNo"), false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send()
			If objXML.Status = "200" Then

				'//���޹��� ���� Ȯ��
				'Response.contentType = "text/xml; charset=euc-kr"
				'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
				'response.End

				'XML�� ���� DOM ��ü ����
				Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				if Err<>0 then
					Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
					dbget.Close: Response.End
				end if

				LotteGoodNo		= Trim(xmlDOM.getElementsByTagName("goods_no").item(0).text)			'���û�ǰ��ȣ
				LotteStatCd		= Trim(xmlDOM.getElementsByTagName("conf_stat_cd").item(0).text)		'���������ڵ�
''rw "LotteStatCd="&LotteStatCd
				'// ����
				strSql = "Update db_item.dbo.tbl_lotte_regItem "
				strSql = strSql & "	Set lastConfirmdate=getdate()"
				if (LotteStatCd<>"") then
    				strSql = strSql & "	,LotteStatCd='" & LotteStatCd & "'"
    			end if

				if (LotteGoodNo>"0") and (LotteGoodNo<>"") then
					strSql = strSql & " , LotteGoodNo='" & LotteGoodNo & "'"
				end if

				strSql = strSql & " Where itemid='" & rsget("itemid") & "'"
				dbget.Execute strSql,AssignedCNT
				if (LotteStatCd="30") then
				    actCnt = actCnt+AssignedCNT
				elseif (LotteStatCd="20") then
				    waitCnt = waitCnt+AssignedCNT
				elseif (LotteStatCd="40") then
				    rjtCnt = rjtCnt+AssignedCNT
				end if
''rw AssignedCNT&":"&LotteStatCd&":"&LotteGoodNo
				Set xmlDOM = Nothing
			else
				Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
				dbget.Close: Response.End
			end if
			Set objXML = Nothing

			rsget.MoveNext
		Loop

	end if

	rsget.Close

	'##### DB ���� ó�� #####
    If Err.Number = 0 Then
    	if actCnt>0 or waitCnt>0 or rjtCnt>0 then
    	    if (IsAutoScript) then
    	        rw  "OK|"&actCnt & "���� ����." & waitCnt& "���� ���δ��." & rjtCnt & "�� �ݷ�"
    	    else
    	        IF (session("ssBctID")="icommang") then
    	            rw actCnt & "���� ����." & waitCnt& "���� ���δ��." & rjtCnt & "�� �ݷ�"
    	        else
        		    Response.Write "<script language=javascript>alert('" & actCnt & "���� ���������� ���ŵǾ����ϴ�.');parent.history.go(0);</script>"
        	    end if
        	end if
    	else
    	    if (IsAutoScript) then
    	        rw  "OK|"&actCnt & "���� ����." & waitCnt& "���� ���δ��."  & rjtCnt & "�� �ݷ�"
    	    else
        		Response.Write "<script language=javascript>alert('������ �ӽõ�� ��ǰ�� �����ϴ�.');parent.history.go(0);</script>"
        	end if
    	end if
    Else
        if (IsAutoScript) then
            rw "S_ERR|ó�� �߿� ������ �߻��߽��ϴ�"
        else
            Response.Write "<script language=javascript>alert('ó�� �߿� ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���.');</script>"
        end if
    End If

	on Error Goto 0


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->