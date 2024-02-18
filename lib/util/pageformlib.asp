<%
'=========================================================
' 2011 New ����¡ �Լ�
' 2011.03.21 ���ر� ����
' 2012.03.26 ������ DIV���̾ƿ����� ����
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' sbDisplayPaging_New(���� ��������ȣ, �� ���ڵ� ����, ���������� ���̴� ��ǰ ����(select top ��), ������ ��ϴ���(ex.10������������ or 5�������� ����), js �������̵� �Լ���)
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' ������ �̵� js �Լ����� strJsFuncName ���� ���Ƿ� ���ϰ� ������ ��ȣ�� ��Ƽ� �ѱ�. �� �������� ����¡ ���� form�� ����ų� ��Ī���� ���� ���ų� �Ͽ� post �Ǵ� get���� �ѱ�.
' �� �������� �پ��� ���ݵ�� ���� ���� �ڵ� ����� ���� ��� ȯ�濡 �������� �κи� ����.
' ���������: �߼�ī�װ�����Ʈ(/shopping/category_list.asp), �������ΰŽ�(/designfingers/designfingers_main.asp, /designfingers/designfingers.asp)(ajax����), �귣�����(/street/index.asp), �̺�Ʈ, ��Ŭ���ڵ�
'=========================================================

Function fnDisplayPaging_New(strCurrentPage, intTotalRecord, intRecordPerPage, intBlockPerPage, strJsFuncName)
	'���� ����
	Dim intCurrentPage, strCurrentPath, vPageBody
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'���� ������ ����
	intCurrentPage = strCurrentPage		'���� ������ ��

	'�ش��������� ǥ�õǴ� ������������ ������������ ����
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage

	'�� ������ �� ����
	intTotalPage =   int((intTotalRecord-1)/intRecordPerPage) +1
	''eastone �߰�
	if (intTotalPage<1) then intTotalPage=1

	vPageBody = ""
	strJsFuncName = trim(strJsFuncName)

	vPageBody = vPageBody & "<div class=""paging"">" & vbCrLf

	'## ù ������
	vPageBody = vPageBody & "	<a href=""#"" title=""ù ������"" class=""first arrow"" onclick=""" & strJsFuncName & "(1);return false;""><span style=""cursor:pointer;"">�� ó�� �������� �̵�</span></a>" & vbCrLf

	'## ���� ������
	If intStartBlock > 1 Then
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" class=""prev arrow"" onclick=""" & strJsFuncName & "(" & intStartBlock-1 & ");return false;""><span style=""cursor:pointer;"">������������ �̵�</span></a>" & vbCrLf
	Else
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" class=""prev arrow"" onclick=""return false;""><span style=""cursor:pointer;"">������������ �̵�</span></a>" & vbCrLf
	End If

	'## ���� ������
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For
			If Int(intLoop) = Int(intCurrentPage) Then
				vPageBody = vPageBody & "	<a href=""#"" title=""" & intLoop & " ������"" class=""current"" onclick=""" & strJsFuncName & "(" & intLoop & ");return false;""><span style=""cursor:pointer;"">" & intLoop & "</span></a>" & vbCrLf
			Else
				vPageBody = vPageBody & "	<a href=""#"" title=""" & intLoop & " ������"" onclick=""" & strJsFuncName & "(" & intLoop & ");return false;""><span style=""cursor:pointer;"">" & intLoop & "</span></a>" & vbCrLf
			End If
		Next
	Else
		vPageBody = vPageBody & "	<a href=""#"" title=""1 ������"" class=""current"" onclick=""" & strJsFuncName & "(1);return false;""><span style=""cursor:pointer;"">1</span></a>" & vbCrLf
	End If

	'## ���� ������
	If Int(intEndBlock) < Int(intTotalPage) Then	'####### ����������
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" class=""next arrow"" onclick=""" & strJsFuncName & "(" & intEndBlock+1 & ");return false;""><span style=""cursor:pointer;"">���� �������� �̵�</span></a>" & vbCrLf
	Else
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" class=""next arrow"" onclick=""return false;""><span style=""cursor:pointer;"">���� �������� �̵�</span></a>" & vbCrLf
	End If

	'## ������ ������
	vPageBody = vPageBody & "	<a href=""#"" title=""������ ������"" class=""end arrow"" onclick=""" & strJsFuncName & "(" & intTotalPage & ");return false;""><span style=""cursor:pointer;"">�� ������ �������� �̵�</span></a>" & vbCrLf

	vPageBody = vPageBody & "</div>" & vbCrLf

	vPageBody = vPageBody & "<div class=""pageMove"">" & vbCrLf
	vPageBody = vPageBody & "	<input type=""number"" value=""" & intCurrentPage & """ min=""1"" max=""" & intTotalPage & """ style=""width:24px;"" />/" & intTotalPage & "������ <a href=""#"" onclick=""fnDirPg" & strJsFuncName & "($(this).prev().val()); return false;"" class=""btn btnS2 btnGry2""><em class=""whiteArr01 fn"">�̵�</em></a>" & vbCrLf
	vPageBody = vPageBody & "</div>" & vbCrLf
	vPageBody = vPageBody & "<script>" & vbCrLf
	vPageBody = vPageBody & "function fnDirPg" & strJsFuncName & "(pg) {" & vbCrLf
	vPageBody = vPageBody & "	if(pg>0 && pg<=" & intTotalPage & ") " & strJsFuncName & "(pg);" & vbCrLf
	vPageBody = vPageBody & "}" & vbCrLf
	vPageBody = vPageBody & "</script>" & vbCrLf

	fnDisplayPaging_New = vPageBody
End Function

'//����ó�� ������ �ؽ�Ʈ�ڽ��� �ٷ� ���� ���� ����		'//2015.03.31 �ѿ�� �߰�
Function fnDisplayPaging_New_nottextboxdirect(strCurrentPage, intTotalRecord, intRecordPerPage, intBlockPerPage, strJsFuncName)
	'���� ����
	Dim intCurrentPage, strCurrentPath, vPageBody
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'���� ������ ����
	intCurrentPage = strCurrentPage		'���� ������ ��

	'�ش��������� ǥ�õǴ� ������������ ������������ ����
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage

	'�� ������ �� ����
	intTotalPage =   int((intTotalRecord-1)/intRecordPerPage) +1
	''eastone �߰�
	if (intTotalPage<1) then intTotalPage=1

	vPageBody = ""

	vPageBody = vPageBody & "<div class=""paging"">" & vbCrLf

	'## ù ������
	vPageBody = vPageBody & "	<a href=""#"" title=""ù ������"" class=""first arrow"" onclick=""" & strJsFuncName & "(1);return false;""><span style=""cursor:pointer;"">�� ó�� �������� �̵�</span></a>" & vbCrLf

	'## ���� ������
	If intStartBlock > 1 Then
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" class=""prev arrow"" onclick=""" & strJsFuncName & "(" & intStartBlock-1 & ");return false;""><span style=""cursor:pointer;"">������������ �̵�</span></a>" & vbCrLf
	Else
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" class=""prev arrow"" onclick=""return false;""><span style=""cursor:pointer;"">������������ �̵�</span></a>" & vbCrLf
	End If

	'## ���� ������
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For
			If Int(intLoop) = Int(intCurrentPage) Then
				vPageBody = vPageBody & "	<a href=""#"" title=""" & intLoop & " ������"" class=""current"" onclick=""" & strJsFuncName & "(" & intLoop & ");return false;""><span style=""cursor:pointer;"">" & intLoop & "</span></a>" & vbCrLf
			Else
				vPageBody = vPageBody & "	<a href=""#"" title=""" & intLoop & " ������"" onclick=""" & strJsFuncName & "(" & intLoop & ");return false;""><span style=""cursor:pointer;"">" & intLoop & "</span></a>" & vbCrLf
			End If
		Next
	Else
		vPageBody = vPageBody & "	<a href=""#"" title=""1 ������"" class=""current"" onclick=""" & strJsFuncName & "(1);return false;""><span style=""cursor:pointer;"">1</span></a>" & vbCrLf
	End If

	'## ���� ������
	If Int(intEndBlock) < Int(intTotalPage) Then	'####### ����������
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" class=""next arrow"" onclick=""" & strJsFuncName & "(" & intEndBlock+1 & ");return false;""><span style=""cursor:pointer;"">���� �������� �̵�</span></a>" & vbCrLf
	Else
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" class=""next arrow"" onclick=""return false;""><span style=""cursor:pointer;"">���� �������� �̵�</span></a>" & vbCrLf
	End If

	'## ������ ������
	vPageBody = vPageBody & "	<a href=""#"" title=""������ ������"" class=""end arrow"" onclick=""" & strJsFuncName & "(" & intTotalPage & ");return false;""><span style=""cursor:pointer;"">�� ������ �������� �̵�</span></a>" & vbCrLf

	vPageBody = vPageBody & "</div>" & vbCrLf

	fnDisplayPaging_New_nottextboxdirect = vPageBody
End Function

'// �ű� ����¡��
Function fnDisplayPaging_New2017(strCurrentPage, intTotalRecord, intRecordPerPage, intBlockPerPage, strJsFuncName)
	'���� ����
	Dim intCurrentPage, strCurrentPath, vPageBody
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'���� ������ ����
	intCurrentPage = strCurrentPage		'���� ������ ��

	'�ش��������� ǥ�õǴ� ������������ ������������ ����
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage

	'�� ������ �� ����
	intTotalPage =   int((intTotalRecord-1)/intRecordPerPage) +1
	''eastone �߰�
	if (intTotalPage<1) then intTotalPage=1

	vPageBody = ""


	'## ���� ������
	If intStartBlock > 1 Then
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" onclick=""" & strJsFuncName & "(" & intStartBlock-1 & ");return false;"">[prev]</a>" & vbCrLf
	Else
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" onclick=""return false;""><span style=""cursor:pointer;"">[prev]</span></a>" & vbCrLf
	End If

	'## ���� ������
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For
			If Int(intLoop) = Int(intCurrentPage) Then
				vPageBody = vPageBody & "	<a href=""#"" title=""" & intLoop & " ������"" onclick=""" & strJsFuncName & "(" & intLoop & ");return false;""><span class='cRd1'>" & intLoop & "</span></a>" & vbCrLf
			Else
				vPageBody = vPageBody & "	<a href=""#"" title=""" & intLoop & " ������"" onclick=""" & strJsFuncName & "(" & intLoop & ");return false;"">" & intLoop & "</a>" & vbCrLf
			End If
		Next
	Else
		vPageBody = vPageBody & "	<a href=""#"" title=""1 ������"" onclick=""" & strJsFuncName & "(1);return false;""><span class='cRd1'>1</span></a>" & vbCrLf
	End If

	'## ���� ������
	If Int(intEndBlock) < Int(intTotalPage) Then	'####### ����������
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" onclick=""" & strJsFuncName & "(" & intEndBlock+1 & ");return false;"">[next]</a>" & vbCrLf
	Else
		vPageBody = vPageBody & "	<a href=""#"" title=""���� ������"" class=""next arrow"" onclick=""return false;"">[next]</a>" & vbCrLf
	End If

	fnDisplayPaging_New2017 = vPageBody
End Function

'// 2013�⿡ �Ϻκ� ���� ����¡�� > ���� ������ ��ȯ (���� ����)
Function fnDisplayPaging_2013(strCurrentPage, intTotalRecord, intRecordPerPage, intBlockPerPage, strJsFuncName)
	'// ������
	fnDisplayPaging_2013 = fnDisplayPaging_New(strCurrentPage, intTotalRecord, intRecordPerPage, intBlockPerPage, strJsFuncName)
End Function
%>
