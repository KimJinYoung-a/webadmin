<%
'####################################################
' Description :  �������� �����ڵ� Ŭ����
' History : 2010.03.09 �ѿ�� ����
'####################################################

Class CEventCommonCode_off
public FCodeType
public FCodeValue
public FCodeDesc
public FCodeUsing
public FCodeSort

	'//�����ڵ� ����Ʈ : �̺�Ʈ Ÿ�Կ� �ش��ϴ� ���� ��������
	public Function fnGetEventCodeList_off
		IF FCodeType = "" THEN Exit Function
		Dim strSql
		
		strSql = "SELECT code_type, code_value, code_desc, code_using, code_sort "&_
				" From [db_shop].[dbo].[tbl_event_off_commoncode] "&_
				" WHERE code_type = '"&FCodeType&"' Order by code_sort "
		
		'response.write strSql &"<br>"		
		rsget.Open strSql,dbget
		
		IF not rsget.EOF THEN
			fnGetEventCodeList_off = rsget.getRows()
		End IF
		rsget.Close		
	End Function
	
	'//������ �ڵ� ���� ��������
	public Function fnGetEventCodeCont_off
		IF FCodeValue = "" or FCodeType = ""  THEN Exit Function				
		Dim strSql
		strSql =" SELECT code_type, code_value, code_desc, code_using, code_sort "&_
				" From  [db_shop].[dbo].[tbl_event_off_commoncode] "&_
				" WHERE code_value = "&FCodeValue&" and code_type ='"&FCodeType&"'"		
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FCodeType 	= rsget("code_type")
			FCodeValue 	= rsget("code_value")
			FCodeDesc 	= rsget("code_desc")
			FCodeUsing 	= rsget("code_using")
			FCodeSort 	= rsget("code_sort")
		End IF			
		rsget.Close		
	End Function
End Class
%>