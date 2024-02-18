<%
'####################################################
' Description :  오프라인 공통코드 클래스
' History : 2010.03.09 한용민 생성
'####################################################

Class CEventCommonCode_off
public FCodeType
public FCodeValue
public FCodeDesc
public FCodeUsing
public FCodeSort

	'//공통코드 리스트 : 이벤트 타입에 해당하는 내용 가져오기
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
	
	'//선택한 코드 내용 가져오기
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