<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
	'// ��������
	dim DispNo, DispNm, DispLrgNm, DispMidNm, DispSmlNm, DispThnNm
	dim CateCnt, CateInfo
	dim strSql, actCnt, disp_tp_cd, arrMDGrNo, lp

	actCnt = 0			'�ǰ��ŰǼ�
	disp_tp_cd = requestCheckVar(request("disptpcd"),10)  ''"10"	'����Ÿ���ڵ�(10:�Ϲݸ���, 11:�귣�����, 12:��������)

	'// MD��ǰ�� �ڵ� ����
	strSql = "Select Distinct groupCode From db_temp.dbo.tbl_lotte_MDCateGrp " ''Where isUsing='Y'"
	rsget.Open strSql,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			reDim arrMDGrNo(rsget.recordCount)
			For lp=0 to (rsget.recordCount-1)
				arrMDGrNo(lp)=rsget(0)
				rsget.MoveNext
			Next
		else
			Response.Write "<script language=javascript>alert('��ϵ� MD��ǰ���� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
			rsget.Close: dbget.Close: Response.End
		end if
	rsget.Close

	on Error Resume Next

	'// Ʈ������ ����
	dbget.beginTrans

	'��� MD��뿩�� ����
	strSql = "update db_temp.dbo.tbl_lotte_Category Set isUsing='N', lastupdate=getdate() Where isUsing='Y' and disptpcd='"&disp_tp_cd&"'"
	dbget.Execute(strSql)

	'// �Ե����� ����ī�װ� ��ȸ
	for lp=0 to ubound(arrMDGrNo)-1
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", lotteAPIURL & "/openapi/searchDispCatListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&disp_tp_cd=" & disp_tp_cd & "&md_gsgr_no=" & arrMDGrNo(lp), false
'		rw lotteAPIURL & "/openapi/searchDispCatListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&disp_tp_cd=" & disp_tp_cd & "&md_gsgr_no=" & arrMDGrNo(lp)
'		response.end
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
	
			'//���޹��� ���� Ȯ��
			'Response.contentType = "text/xml; charset=euc-kr"
			'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
	
			'XML�� ���� DOM ��ü ����
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
			
			CateCnt = xmlDOM.getElementsByTagName("CategoryCount").item(0).text		'�����
			if Err<>0 then
				Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
				dbget.RollBackTrans: dbget.Close: Response.End
			end if

			if CateCnt>0 then

				'// CateInfo Loop
				Set CateInfo = xmlDOM.getElementsByTagName("CategoryInfo")
				for each SubNodes in CateInfo
					DispNo		= Trim(SubNodes.getElementsByTagName("DispNo").item(0).text)		'ī�װ� �ڵ�
					DispNm		= Trim(SubNodes.getElementsByTagName("DispNm").item(0).text)		'ī�װ���(�����з�)
					DispLrgNm	= Trim(SubNodes.getElementsByTagName("DispLrgNm").item(0).text)		'��з���
					DispMidNm	= Trim(SubNodes.getElementsByTagName("DispMidNm").item(0).text)		'�ߺз���
					DispSmlNm	= Trim(SubNodes.getElementsByTagName("DispSmlNm").item(0).text)		'�Һз���
					DispThnNm	= Trim(SubNodes.getElementsByTagName("DispThnNm").item(0).text)		'���з���

					'MD���翩�� Ȯ��
					strSql = "Select count(DispNo) From db_temp.dbo.tbl_lotte_Category Where DispNo='" & DispNo & "'"
					rsget.Open strSql,dbget,1

					if rsget(0)>0 then
						'// ���� -> �����
						strSql = "update db_temp.dbo.tbl_lotte_Category "
						strSql = strSql & " Set isUsing='Y'"
						strSql = strSql & " , groupCode='" & arrMDGrNo(lp) & "'"
						strSql = strSql & " , disptpcd='"&disp_tp_cd&"'"
						strSql = strSql & " , DispNm='"&DispNm&"'"
						strSql = strSql & " , DispLrgNm='"&html2db(DispLrgNm)&"'"
						strSql = strSql & " , DispMidNm='"&html2db(DispMidNm)&"'"
						strSql = strSql & " , DispSmlNm='"&html2db(DispSmlNm)&"'"
						strSql = strSql & " , DispThnNm='"&html2db(DispThnNm)&"'"
						strSql = strSql & "  Where DispNo='" & DispNo & "'"
						dbget.Execute(strSql)
						actCnt = actCnt+1
					else
						'// ���� -> �űԵ��
						strSql = "Insert into db_temp.dbo.tbl_lotte_Category (DispNo, DispNm, DispLrgNm, DispMidNm, DispSmlNm, DispThnNm, disptpcd, groupCode) values " &_
								" ('" & DispNo & "'" &_
								", '" & html2db(DispNm) & "'" &_
								", '" & html2db(DispLrgNm) & "'" &_
								", '" & html2db(DispMidNm) & "'" &_
								", '" & html2db(DispSmlNm) & "'" &_
								", '" & html2db(DispThnNm) & "'" &_
								", '" & html2db(disp_tp_cd) & "'" &_
								", '" & arrMDGrNo(lp) & "')"
						dbget.Execute(strSql)
						actCnt = actCnt+1
					end if

					rsget.Close
				Next
				Set CateInfo = Nothing

			end if
	
			Set xmlDOM = Nothing
		else
			Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
			dbget.RollBackTrans: dbget.Close: Response.End
		end if
		Set objXML = Nothing

	Next

	'##### DB ���� ó�� #####
    If Err.Number = 0 Then
    	dbget.CommitTrans				'Ŀ��(����)
    	Response.Write "<script language=javascript>alert('" & actCnt & "���� ���������� ���ŵǾ����ϴ�.');parent.history.go(0);</script>"
    Else
        dbget.RollBackTrans				'�ѹ�(�����߻���)
        Response.Write "<script language=javascript>alert('�ڷ� ���� �߿� ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���.');</script>"
    End If

	on Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->