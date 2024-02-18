<%
	Class CEventCommonCode
	public FCodeType
	public FCodeValue
	public FCodeDesc
	public FCodeUsing
	public FCodeSort
	public FCodeDispYN
	public FkindCode
	public FcontentsCode
	public FcontentsEa
	public FRectIDX
	public FRectkindCode
	
		'//�����ڵ� ����Ʈ : �̺�Ʈ Ÿ�Կ� �ش��ϴ� ���� ��������
		public Function fnGetEventCodeList
			IF FCodeType = "" THEN Exit Function
			Dim strSql
			strSql = "SELECT code_type, code_value, code_desc, code_using, code_sort, code_dispYN "&_
					" From [db_sitemaster].[dbo].[tbl_mailzine_code] "&_
					" WHERE code_using='Y' and code_dispYN='Y' and code_type = '"&FCodeType&"' Order by code_sort "
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetEventCodeList = rsget.getRows()
			End IF
			rsget.Close		
		End Function

		'//�����ڵ� ����Ʈ : ������ ������ �ش��ϴ� ���ø� ����Ʈ ��������
		public Function fnGetTemplateList
			IF FRectkindCode = "" THEN Exit Function
			Dim strSql
			strSql = "SELECT M.idx, M.contentsCode, C.code_desc, M.contentsEa, M.sortidx " & vbCrlF
			strSql = strSql & " From [db_sitemaster].[dbo].[tbl_mailzine_contents_manage] as M" & vbCrlF
			strSql = strSql & " Left Join [db_sitemaster].[dbo].[tbl_mailzine_code] as C ON M.contentsCode=C.code_value" & vbCrlF
			strSql = strSql & " WHERE M.kindCode='"&FRectkindCode&"' Order by M.sortidx asc, M.idx asc"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetTemplateList = rsget.getRows()
			End IF
			rsget.Close		
		End Function

		'//������ �ڵ� ���� ��������
		public Function fnGetEventCodeCont
			IF FCodeValue = "" or FCodeType = ""  THEN Exit Function				
			Dim strSql
			strSql =" SELECT code_type, code_value, code_desc, code_using, code_sort, code_dispYN "&_
					" From  [db_sitemaster].[dbo].[tbl_mailzine_code] "&_
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
	
		'//������ ���ø� ���� ��������
		public Function fnGetTemplateCont
			IF FRectIDX="" THEN Exit Function				
			Dim strSql
			strSql =" SELECT kindCode, contentsCode, contentsEa, sortidx"&_
					" From [db_sitemaster].[dbo].[tbl_mailzine_contents_manage] "&_
					" WHERE idx="&FRectIDX
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				FkindCode 	= rsget("kindCode")
				FcontentsCode 	= rsget("contentsCode")
				FcontentsEa 	= rsget("contentsEa")
			End IF
			rsget.Close		
		End Function
	End Class
	
	Sub sbOptCodeType(ByVal selCodeType)
		Dim arrSelCode(19), i
		arrSelCode(0)	= Split("mailzineKind|������ ����","|")
		arrSelCode(1)	= Split("contentsKind|������ ����","|")
		arrSelCode(2)	= Split("mailzineState|������ �ۼ� ����","|")
		for i=0 to ubound(arrSelCode)
			if isArray(arrSelCode(i)) then
			Response.Write "<option value=""" & arrSelCode(i)(0) &""" " & chkIIF(Cstr(selCodeType)=arrSelCode(i)(0),"selected","") & ">" & arrSelCode(i)(1) & "</option>" & vbCrlF
			end if
		next
	End Sub

	Sub sbMailzineKindType(ByVal selCodeType)
		Dim strSql
		strSql = " SELECT code_value, code_desc" & vbCrlF
		strSql = strSql & " From  [db_sitemaster].[dbo].[tbl_mailzine_code] " & vbCrlF
		strSql = strSql & " Where code_type='mailzineKind' And code_using='Y' And code_dispYN='Y'" & vbCrlF
		strSql = strSql & " Order By code_sort ASC"		
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			do until rsget.EOF
				Response.Write "<option value=""" & rsget("code_value") &"""" & chkIIF(Cstr(selCodeType)=Cstr(rsget("code_value"))," selected","") & ">" & rsget("code_desc") & "</option>" & vbCrlF
				rsget.movenext
			loop
		End IF			
		rsget.Close	
	End Sub

	Sub sbContentsKindType(ByVal selCodeType)
		Dim strSql
		strSql = " SELECT code_value, code_desc" & vbCrlF
		strSql = strSql & " From  [db_sitemaster].[dbo].[tbl_mailzine_code] " & vbCrlF
		strSql = strSql & " Where code_type='contentsKind' And code_using='Y' And code_dispYN='Y'" & vbCrlF
		strSql = strSql & " Order By code_sort ASC"		
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			do until rsget.EOF
				Response.Write "<option value=""" & rsget("code_value") &""" " & chkIIF(Cstr(selCodeType)=Cstr(rsget("code_value")),"selected","") & ">" & rsget("code_desc") & "</option>" & vbCrlF
				rsget.movenext
			loop
		End IF			
		rsget.Close	
	End Sub
%>