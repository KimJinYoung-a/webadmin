<%
'�Է�: strCurrentPage-���� ������ ������, intCurrentPage-����������, intTotalRecord-�� �˻��Ǽ�
'			, intRecordPerPage-���������� �������� ���ڵ� ��, intBlockPerPage-�� ��������
Sub sbDisplayPaging(ByVal strCurrentPage, ByVal intCurrentPage, ByVal intTotalRecord, ByVal intRecordPerPage, ByVal intBlockPerPage,ByVal menupos)

	'���� ����
	Dim strCurrentPath
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'���� ������ ��
	strCurrentPath = Request.ServerVariables("Script_Name")

	'�ش��������� ǥ�õǴ� ������������ ������������ ����
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage

	'�� ������ �� ����
	intTotalPage =  -(int(-(intTotalRecord/intRecordPerPage)))

	'�� ���� & hidden �Ķ���� ����
	Response.Write	"<form name='frmPaging' method='get' action ='" & strCurrentPath & "'>" &_
							"<input type='hidden' name='" & strCurrentPage & "' />"

	'�Ķ���� ����(��: �˻���)�� hidden �Ķ���ͷ� �����Ѵ�
	strParamName = ""
	For Each strParamName In Request.Form
		If strParamName <> strCurrentPage Then

			'hidden �Ķ���� ���� �Ķ���� �˿�
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.Form(strParamName),50) & "' />"
		End If
	Next
	strParamName = ""

	For Each strParamName In Request.Querystring
		If strParamName <> strCurrentPage Then
			'hidden �Ķ���� ���� �Ķ���� �˿�
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.QueryString(strParamName),50) & "' />"
		END IF
	Next


	'���� ������ �̹��� ����
	If intStartBlock > 1 Then
		Response.Write "<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intStartBlock - intBlockPerPage & ";document.frmPaging.submit();' onfocus='this.blur();'>[prev]</a>"

	Else
		Response.Write "[prev]"
	End If


	'����¡ ���
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For

			If Int(intLoop) <> Int(intStartBlock) Then Response.Write "|"

			If Int(intLoop) = Int(intCurrentPage) Then		'���� ������
				Response.Write "&nbsp;<span class='cRd1'>" & intLoop & "</span>&nbsp;"
			Else															'�� �� ������
				Response.Write "&nbsp;<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intLoop & ";document.frmPaging.submit();'><font class='text01'>" & intLoop & "</font></a>&nbsp;"
			End If

		Next
	Else		'�� �������� ���� �Ҷ�
		Response.Write "&nbsp;<span class='cRd1'>1</span>&nbsp;"
	End If

	'���� ������ �̹��� ����
	If Int(intEndBlock) < Int(intTotalPage) Then
		Response.Write "<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intEndBlock+1 & ";document.frmPaging.submit();'onfocus='this.blur();'>[next]</a>"
	Else
		Response.Write "[next]"
	End If

	Response.Write "</form>"

End Sub
%>
