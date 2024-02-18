<%
'###########################################################
' Description : 정산공용함수
' History : 서동석 생성
'###########################################################

function DrawBillSiteCombo(icompname,icompvalue)
    Dim strSql, intLoop, arrBill

    strSql = " select billsitecode, billsitename"
    strSql = strSql & " from db_jungsan.dbo.tbl_tax_asp_Info with (nolock)"
    strSql = strSql & " where isValid='Y'"
    strSql = strSql & " order by billsitecode"

    rsget.CursorLocation = adUseClient
    rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
	IF Not rsget.EOF THEN
		arrBill = rsget.getRows()	
	END IF	
	rsget.Close
%>
    <select name='<%= icompname %>'>
    <option value="">선택
<%	
		For intLoop =0 To UBOund(arrBill,2)
%>
	    <option value="<%=arrBill(0,intLoop)%>" <%IF icompvalue = arrBill(0,intLoop) THEN%>selected<%END IF%>><%=arrBill(1,intLoop)%></option>
<%		
		Next
%>
    </select>
<%    

end function

%>