<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<%
Dim DispNo, DispNm, DispLrgNm, DispMidNm, DispSmlNm, DispThnNm
Dim CateCnt, CateInfo
Dim strSql, actCnt, disp_tp_cd, arrMDGrNo, lp
actCnt = 0			'�ǰ��ŰǼ�
disp_tp_cd = requestCheckVar(request("disptpcd"),10)  ''"10"	'����Ÿ���ڵ�(10:�Ϲݸ���, 11:�귣�����, 12:��������)
'// MD��ǰ�� �ڵ� ����
strSql = "SELECT Distinct groupCode FROM db_temp.dbo.tbl_lotteiMall_MDCateGrp WHERE isUsing = 'Y'"
rsget.Open strSql,dbget,1
If Not(rsget.EOF or rsget.BOF) then
	ReDim arrMDGrNo(rsget.recordCount)
	For lp = 0 to (rsget.recordCount - 1)
		arrMDGrNo(lp)=rsget(0)
		rsget.MoveNext
	Next
Else
	Response.Write "<script language=javascript>alert('��ϵ� MD��ǰ���� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
	rsget.Close: dbget.Close: Response.End
End If
rsget.Close

'on Error Resume Next
dbget.beginTrans

'��� MD��뿩�� ����
strSql = "update db_temp.dbo.tbl_lotteiMall_Category Set isUsing='N', lastupdate=getdate() Where isUsing='Y' and disptpcd='"&disp_tp_cd&"'"
dbget.Execute(strSql)
'// �Ե����̸� ����ī�װ� ��ȸ
for lp = 0 to ubound(arrMDGrNo)-1
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", ltiMallAPIURL & "/openapi/searchDispCatListOpenApi.lotte?subscriptionId=" & ltiMallAuthNo & "&disp_tp_cd=" & disp_tp_cd & "&md_gsgr_no=" & arrMDGrNo(lp), false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				CateCnt = xmlDOM.getElementsByTagName("CategoryCount").item(0).text		'�����
				If Err <> 0 Then
					Response.Write "<script language=javascript>alert('�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
					dbget.RollBackTrans: dbget.Close: Response.End
				End If

				If CInt(CateCnt) > 0 Then
					Set CateInfo = xmlDOM.getElementsByTagName("CategoryInfoList")
						For each SubNodes in CateInfo
							DispNo		= Trim(SubNodes.getElementsByTagName("DispNo").item(0).text)		'ī�װ� �ڵ�
							DispNm		= Trim(SubNodes.getElementsByTagName("DispNm").item(0).text)		'ī�װ���(�����з�)
							DispLrgNm	= Trim(SubNodes.getElementsByTagName("DispLrgNm").item(0).text)		'��з���
							DispMidNm	= Trim(SubNodes.getElementsByTagName("DispMidNm").item(0).text)		'�ߺз���
							DispSmlNm	= Trim(SubNodes.getElementsByTagName("DispSmlNm").item(0).text)		'�Һз���
							DispThnNm	= Trim(SubNodes.getElementsByTagName("DispThnNm").item(0).text)		'���з���

							'MD���翩�� Ȯ��
							strSql = "Select count(DispNo) From db_temp.dbo.tbl_lotteiMall_Category Where DispNo='" & DispNo & "' and groupCode = '" & arrMDGrNo(lp) & "' "
							rsget.Open strSql,dbget,1
							If rsget(0) > 0 Then
								'// ���� -> �����
								strSql = "update db_temp.dbo.tbl_lotteiMall_Category Set isUsing='Y', groupCode='" & arrMDGrNo(lp) & "', disptpcd='"&disp_tp_cd&"' Where DispNo='" & DispNo & "'  and groupCode = '" & arrMDGrNo(lp) & "' "
								dbget.Execute(strSql)
							Else
								'// ���� -> �űԵ��
								strSql = "Insert into db_temp.dbo.tbl_lotteiMall_Category (DispNo, DispNm, DispLrgNm, DispMidNm, DispSmlNm, DispThnNm, disptpcd, groupCode) values " &_
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
							End If
							rsget.Close
						Next
					Set CateInfo = Nothing
				End If
			Set xmlDOM = Nothing
		Else
			Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
			dbget.RollBackTrans: dbget.Close: Response.End
		End If
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