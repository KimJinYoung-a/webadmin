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
							"<input type='hidden' name='" & strCurrentPage & "'>" 
		
	'�Ķ���� ����(��: �˻���)�� hidden �Ķ���ͷ� �����Ѵ�
	strParamName = ""
	For Each strParamName In Request.Form	
		If strParamName <> strCurrentPage Then
			
			'hidden �Ķ���� ���� �Ķ���� �˿�
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.Form(strParamName),50) & "'>"
		End If
	Next
	strParamName = ""
	
	For Each strParamName In Request.Querystring
		If strParamName <> strCurrentPage Then			
			'hidden �Ķ���� ���� �Ķ���� �˿�
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.QueryString(strParamName),50) & "'>"		
		END IF	
	Next
		
	Response.Write "<table border='0' cellpadding='0' cellspacing='0' class='a'><tr align='center'><td>"

	'���� ������ �̹��� ����
	If intStartBlock > 1 Then
		Response.Write "<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intStartBlock - intBlockPerPage & ";document.frmPaging.submit();' onfocus='this.blur();'>[pre]</a>" 
							   
	Else
		Response.Write "[pre]"
	End If

	Response.Write "</td><td>&nbsp;"
	
	'����¡ ���
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For
			
			If Int(intLoop) <> Int(intStartBlock) Then Response.Write "|"
			
			If Int(intLoop) = Int(intCurrentPage) Then		'���� ������
				Response.Write "&nbsp;<span class='text01'><strong>" & intLoop & "</strong></span>&nbsp;"
			Else															'�� �� ������
				Response.Write "&nbsp;<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intLoop & ";document.frmPaging.submit();'><font class='text01'>" & intLoop & "</font></a>&nbsp;"
			End If
		
		Next
	Else		'�� �������� ���� �Ҷ�
		Response.Write "&nbsp;<span class='text01'><strong>1</strong></span>&nbsp;"
	End If

	Response.Write "&nbsp;</td><td>"

	'���� ������ �̹��� ����
	If Int(intEndBlock) < Int(intTotalPage) Then
		Response.Write "<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intEndBlock+1 & ";document.frmPaging.submit();'onfocus='this.blur();'>[next]</a>"  
	Else
		Response.Write "[next]"
	End If
	
	Response.Write "</td></tr></table></form>"

End Sub 
%> 