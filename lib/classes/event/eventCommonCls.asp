<%
	Class CEventCommonCode
	public FCodeType
	public FCodeValue
	public FCodeDesc
	public FCodeUsing
	public FCodeSort
	public FCodeDispYN
	
		'//공통코드 리스트 : 이벤트 타입에 해당하는 내용 가져오기
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
		
		'//선택한 코드 내용 가져오기
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
		arrSelCode(0)	= Split("eventkind|이벤트종류","|")
		arrSelCode(1)	= Split("eventtype|이벤트유형","|")
		arrSelCode(2)	= Split("eventlevel|중요도","|")
		arrSelCode(3)	= Split("eventmanager|주체","|")
		arrSelCode(4)	= Split("eventscope|범위","|")
		arrSelCode(5)	= Split("eventstate|상태","|")
		arrSelCode(6)	= Split("eventview|[PC-WEB]화면템플릿","|")
		arrSelCode(7)	= Split("eventview_mo|[Mobile/App]화면템플릿","|")
		arrSelCode(8)	= Split("salestatus|할인상태","|")
		arrSelCode(9)	= Split("salemargin|할인마진","|")
		arrSelCode(10)	= Split("giftscope|사은품범위","|")
		arrSelCode(11)	= Split("gifttype|사은품타입","|")
		arrSelCode(12)	= Split("giftstatus|사은품상태","|")
		arrSelCode(13)	= Split("itemsort|상품정렬순서","|")
		arrSelCode(14)	= Split("itemaddtype|상품관리방법","|")
		arrSelCode(15)	= Split("evtprizestatus|상품이미지크기","|")
		arrSelCode(16)	= Split("evtprizetype|당첨구분","|")
		arrSelCode(17)	= Split("evtprizestatus|당첨자상태","|")
		arrSelCode(18)	= Split("designerstatus|디자이너상태","|")

		for i=0 to ubound(arrSelCode)
			if isArray(arrSelCode(i)) then
			Response.Write "<option value=""" & arrSelCode(i)(0) &""" " & chkIIF(Cstr(selCodeType)=arrSelCode(i)(0),"selected","") & ">" & arrSelCode(i)(1) & "</option>" & vbCrlF
			end if
		next
	End Sub
%>