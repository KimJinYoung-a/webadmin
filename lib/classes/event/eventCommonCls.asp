<%
	Class CEventCommonCode
	public FCodeType
	public FCodeValue
	public FCodeDesc
	public FCodeUsing
	public FCodeSort
	public FCodeDispYN
	
		'//�����ڵ� ����Ʈ : �̺�Ʈ Ÿ�Կ� �ش��ϴ� ���� ��������
		public Function fnGetEventCodeList
			IF FCodeType = "" THEN Exit Function
			Dim strSql
			strSql = "SELECT code_type, code_value, code_desc, code_using, code_sort, code_dispYN "&_
					" From [db_event].[dbo].[tbl_event_commoncode] "&_
					" WHERE code_type = '"&FCodeType&"' Order by code_sort "
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetEventCodeList = rsget.getRows()
			End IF
			rsget.Close		
		End Function
		
		'//������ �ڵ� ���� ��������
		public Function fnGetEventCodeCont
			IF FCodeValue = "" or FCodeType = ""  THEN Exit Function				
			Dim strSql
			strSql =" SELECT code_type, code_value, code_desc, code_using, code_sort, code_dispYN "&_
					" From  [db_event].[dbo].[tbl_event_commoncode] "&_
					" WHERE code_value = "&FCodeValue&" and code_type ='"&FCodeType&"'"		
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				FCodeType 	= rsget("code_type")
				FCodeValue 	= rsget("code_value")
				FCodeDesc 	= rsget("code_desc")
				FCodeUsing 	= rsget("code_using")
				FCodeSort 	= rsget("code_sort")
				FCodeDispYN = rsget("code_dispYN")
			End IF			
			rsget.Close		
		End Function
	End Class
	
	Sub sbOptCodeType(ByVal selCodeType)
		Dim arrSelCode(19), i
		arrSelCode(0)	= Split("eventkind|�̺�Ʈ����","|")
		arrSelCode(1)	= Split("eventtype|�̺�Ʈ����","|")
		arrSelCode(2)	= Split("eventlevel|�߿䵵","|")
		arrSelCode(3)	= Split("eventmanager|��ü","|")
		arrSelCode(4)	= Split("eventscope|����","|")
		arrSelCode(5)	= Split("eventstate|����","|")
		arrSelCode(6)	= Split("eventview|[PC-WEB]ȭ�����ø�","|")
		arrSelCode(7)	= Split("eventview_mo|[Mobile/App]ȭ�����ø�","|")
		arrSelCode(8)	= Split("salestatus|���λ���","|")
		arrSelCode(9)	= Split("salemargin|���θ���","|")
		arrSelCode(10)	= Split("giftscope|����ǰ����","|")
		arrSelCode(11)	= Split("gifttype|����ǰŸ��","|")
		arrSelCode(12)	= Split("giftstatus|����ǰ����","|")
		arrSelCode(13)	= Split("itemsort|��ǰ���ļ���","|")
		arrSelCode(14)	= Split("itemaddtype|��ǰ�������","|")
		arrSelCode(15)	= Split("evtprizestatus|��ǰ�̹���ũ��","|")
		arrSelCode(16)	= Split("evtprizetype|��÷����","|")
		arrSelCode(17)	= Split("evtprizestatus|��÷�ڻ���","|")
		arrSelCode(18)	= Split("designerstatus|�����̳ʻ���","|")

		for i=0 to ubound(arrSelCode)
			if isArray(arrSelCode(i)) then
			Response.Write "<option value=""" & arrSelCode(i)(0) &""" " & chkIIF(Cstr(selCodeType)=arrSelCode(i)(0),"selected","") & ">" & arrSelCode(i)(1) & "</option>" & vbCrlF
			end if
		next
	End Sub
%>