<%
'####################################################
' Description :  �������� �̺�Ʈ �Լ�����
' History : 2010.03.09 �ѿ�� ����
'####################################################

'//�������θ��� ���� ������� ����
Sub drawSelectBoxOffShop_off(selectBoxName,selectedId)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%>>�����ϼ���</option>
	<option value='all' <%if selectedId="all" then response.write " selected"%>>��ü�������</option>
	<%
		query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
		query1 = query1 & " where isusing='Y' "
		query1 = query1 & " and userid<>'streetshop000'"
		query1 = query1 & " and userid<>'streetshop800'"
		query1 = query1 & " and userid<>'streetshop870'"
	
		rsget.Open query1,dbget,1
		
		if  not rsget.EOF  then
		rsget.Movefirst
		
		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("userid")) then
		tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
		rsget.close
	response.write("</select>")
end sub

'//�������θ��� ���� ������� ������
Sub drawSelectBoxoneOffShop_off(selectBoxName,selectedId)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%>>�����ϼ���</option>	
	<%
		query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
		query1 = query1 & " where isusing='Y' "
		query1 = query1 & " and userid<>'streetshop000'"
		query1 = query1 & " and userid<>'streetshop800'"
		query1 = query1 & " and userid<>'streetshop870'"
	
		rsget.Open query1,dbget,1
		
		if  not rsget.EOF  then
		rsget.Movefirst
		
		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("userid")) then
		tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
		rsget.close
	response.write("</select>")
end sub

'//�������θ��� ���� ������� ������ , �ؿܸ��� ������
Sub drawSelectBoxinternalOffShop_off(selectBoxName,selectedId)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%>>�����ϼ���</option>	
	<%
		query1 = " select userid,shopname"
		query1 = query1 & " from [db_shop].[dbo].tbl_shop_user"		
		query1 = query1 & " where isusing='Y' "
		query1 = query1 & " and userid<>'streetshop000'"
		query1 = query1 & " and userid<>'streetshop800'"
		query1 = query1 & " and userid<>'streetshop870'"
		query1 = query1 & " and shopdiv <> '7'"
		
		rsget.Open query1,dbget,1
		
		if  not rsget.EOF  then
		rsget.Movefirst
		
		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("userid")) then
		tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
		rsget.close
	response.write("</select>")
end sub

'//�̺�Ʈ �����ڵ� ���ú����� ����
Function fnSetEventCommonCode_off
 On Error Resume Next
	Dim strSql, arrList, intLoop
	Dim intI, intJ, arrCode(), strtype
	strSql = " SELECT code_type, code_value, code_desc " + vbcrlf
	strSql = strSql & " FROM [db_shop].[dbo].[tbl_event_off_commoncode] WHERE code_using ='Y' Order by code_type, code_sort "			
	rsget.Open strSql,dbget
	IF not rsget.EOF THEN
		arrList = rsget.getRows()
	END IF		
	rsget.close
				
	intJ = 0		
	For intI = 0 To UBound(arrList, 2)	
		If intI > 0 AND intI <> Ubound(arrList, 2) Then
			Application.Lock			
			Application(Trim(strtype)) = arrCode
			Application.UnLock
		End If
				
		If strtype <> Trim(arrList(0, intI)) And Not IsEmpty(strtype) Then intJ = 0	' �ε��� �ʱ�ȭ								
		ReDim Preserve arrCode(Ubound(arrList)-1, intJ)	' �迭 Ȯ��	
		strtype = Trim(arrList(0, intI))  	' ���� ���� 				 			
		arrCode(0, intJ) = Trim(arrList(1, intI)) ' �ڵ� ����
		arrCode(1, intJ) = Trim(arrList(2, intI)) ' �ڵ�� ����						 
		
		intJ = intJ + 1 ' �ε��� ����
			
		If intI = Ubound(arrList, 2) Then
			Application.Lock			
			Application(Trim(strtype)) = arrCode
			Application.UnLock
		End If
				
	Next
End Function

'//Ư�������� �����ڵ尪�� �迭���� Ư������ �ڵ�� ��������
Function fnGetCommCodeArrDesc_off(ByVal arrCode, ByVal iCodeValue)
	Dim intLoop		
	IF iCodeValue = "" or isNull(iCodeValue) THEN iCodeValue = -1
	For intLoop =0 To UBound(arrCode,2)		
		IF Cint(iCodeValue) = arrCode(0,intLoop) THEN				
			fnGetCommCodeArrDesc_off = arrCode(1,intLoop)
			Exit For
		END IF	
	Next	
End Function

'//sbGetOptStatusCodeValue_off(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
'//�̺�Ʈ ���°� �����ڵ� ���ú��� select ����ȭ - ���簪�� ���º��� ���������� �̵� ���ϵ���
 Sub sbGetOptStatusCodeValue_off(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList= fnSetCommonCodeArr_off(sType, True) 
 	iValue  = selValue
 	IF  isNull(selValue) THEN 		selValue = ""
 	IF selValue = ""  THEN	iValue = 0
 	
%>
	<select name="<%=sType%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">����</option>
    <%END IF%>
<% 	
 	For intLoop =0 To UBound(arrList,2)
 	 IF Cint(arrList(0,intLoop)) >= Cint(iValue) OR sViewOpt THEN
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=replace(arrList(1,intLoop),"���¿���","����")%></option>
<%   END IF
	Next %>
	<select>
<%
 End Sub

'//sbGetOptEventCodeValue_off(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
'// : �̺�Ʈ �����ڵ� ���ú��� select ����ȭ 
 Sub sbGetOptEventCodeValue_off(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList = Application(sType)
 	iValue  = selValue
 	IF  isNull(selValue) THEN 		selValue = ""
 	IF selValue = ""  THEN	iValue = 0
 	
%>
	<select name="<%=sType%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">����</option>
    <%END IF%>
<% 	
 	For intLoop =0 To UBound(arrList,2) 	 
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
<%   
	Next %>
	</select>
<%
 End Sub

''// �̺�Ʈ �����ڵ� ��������
 Function fnSetCommonCodeArr_off(ByVal code_type, ByVal blnUse) 
	Dim strSql, arrList, intLoop, strAdd
	Dim intI, intJ, arrCode(), strtype
	strAdd = ""
	IF blnUse THEN
		strAdd= " and code_using ='Y' "
	END IF	
	strSql = " SELECT code_value, code_desc " + vbcrlf
	strSql = strSql & " FROM [db_shop].[dbo].[tbl_event_off_commoncode] " + vbcrlf
	strSql = strSql & " WHERE code_type='"&code_type&"' "&strAdd&"" + vbcrlf
	strSql = strSql & " Order by code_type, code_sort "						
	rsget.Open strSql,dbget
	IF not rsget.EOF THEN
		fnSetCommonCodeArr_off = rsget.getRows()
	END IF		
	rsget.close	
End Function	

''// �̺�Ʈ ��ǰ�� ����� �Ǿ� �ֳ� üũ
 Function geteventcheckitem(ByVal eCode) 
	Dim strSql
	
	geteventcheckitem = false

	strSql = "select top 1 evt_code" + vbcrlf
	strSql = strSql &"  from db_shop.dbo.tbl_eventitem_off" + vbcrlf
	strSql = strSql &"  where evt_code = "&eCode&"" + vbcrlf
	
	'response.write strSql				
	rsget.Open strSql,dbget
	
	IF not rsget.EOF THEN
		geteventcheckitem = true
	END IF		
	rsget.close		
End Function

'//sbGetOptCommonCodeArr_off(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
'//Ư�������� �����ڵ尪�� �迭���� select ó��
 Sub sbGetOptCommonCodeArr_off(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal blnUse, ByVal sScript)
   Dim arrCode, intLoop, iValue 	
   	arrCode= fnSetCommonCodeArr_off(sType, blnUse) 
 	iValue  = selValue
 	IF  isNull(selValue) THEN 	selValue = ""
 	IF selValue = ""  THEN	iValue = 0 	
%>
	<select name="<%=sType%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">����</option>
    <%END IF%>
<% 	IF isArray(arrCode) THEN
 	For intLoop =0 To UBound(arrCode,2) 	 
%>
	<option value="<%=arrCode(0,intLoop)%>" <%If CStr(selValue) = CStr(arrCode(0,intLoop)) THEN%>selected<%END IF%>><%=arrCode(1,intLoop)%></option>
<%   
	Next 
	End IF
	%>
	<select>
<%
 End Sub

'//���MD ����Ʈ�������� (���� �̸�,���� �̻�)
 Sub sbGetMDid_off(ByVal selName, ByVal sIDValue, ByVal sScript)
  Dim strSql, arrList, intLoop
   strSql = " SELECT p.id, isNull(p.company_name,u.username) as company_name from db_partner.[dbo].tbl_partner as p " + vbcrlf
   strSql = strSql & " inner join db_partner.[dbo].tbl_user_tenbyten as u on p.id = u.userid "
   strSql = strSql & " WHERE  p.part_sn ='13' and p.posit_sn>='4' and p.posit_sn<='8' and p.isUsing ='Y' order by p.level_sn"
   rsget.Open strSql,dbget
   IF not rsget.eof THEN
   	arrList = rsget.getRows()	
   End IF
   rsget.close
%>
	<select name="<%=selName%>" <%=sScript%>>
	<option value="">����</option>
<%   
   If isArray(arrList) THEN
   		For intLoop = 0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%if Cstr(arrList(0,intLoop)) = Cstr(sIDValue) then %>selected<%end if%>><%=arrList(1,intLoop)%></option>
<%   		
   		Next
   End IF
%>
	</select>
<%	
 End Sub
 
Sub sbOptCodeType(ByVal selCodeType)
%>	
	<option value="evt_kind" <%IF Cstr(selCodeType)="evt_kind" THEN%>selected<%END IF%>>�̺�Ʈ����</option>	
	<option value="evt_state" <%IF Cstr(selCodeType)="evt_state" THEN%>selected<%END IF%>>����</option>	
<%			
End Sub
  
'//sbGetOptGiftCodeValue(������, ���ð�, ��ũ��Ʈ,�̺�Ʈ�ڵ�)
'//����ǰ�����ڵ� ���ú��� select ����ȭ 
 Sub sbGetOptGiftCodeValue_off(ByVal sType, ByVal selValue, ByVal sScript,ByVal eCode)
   Dim arrList, intLoop
 	arrList = fnSetCommonCodeArr_off(sType, True) 
 
%>
<select name="<%=sType%>" <%=sScript%>>	
	<% For intLoop =0 To UBound(arrList,2) %>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
	<% Next %>
<select>
<% 
End Sub

'//�Ŵ뱸��
function getracknum(boxname,stats)
dim i
%>
<select name="<%=boxname%>">	
	<option value='' <% if stats="" then response.write " selected" %>>����</option>
	<% for i = 1 to 20 %>
	<option value='<%=i%>' <% if stats=i then response.write " selected" %>><%=i%></option>
	<% next %>
<%
end function

'�������¿� ���� ���԰� ����
Function fnSetSaleSupplyPrice(ByVal MarginType, ByVal MarginValue, ByVal orgPrice, ByVal orgSupplyPrice, ByVal salePrice,comm_cd)
	Dim orgMRate
	if orgPrice <>0 then '�� ������
		orgMRate = 100-fix(orgSupplyPrice/orgPrice*10000)/100
	end if
		
	SELECT CASE MarginType
		Case 1	'���ϸ���					
			fnSetSaleSupplyPrice = salePrice-fix(salePrice * (orgMRate/100))
		Case 2	'��ü�δ�
			fnSetSaleSupplyPrice = salePrice-(orgPrice-orgSupplyPrice)
		Case 3	'�ݹݺδ�
			fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
		Case 4	'10x10�δ�
			fnSetSaleSupplyPrice = orgSupplyPrice
		Case 5	'��������
			fnSetSaleSupplyPrice = salePrice - fix(salePrice*(MarginValue/100))
		Case 6	'��ü��Ź�ݹݺδ�/�������ٹ����ٺδ�
			if comm_cd = "B012" then
				fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
			else
				fnSetSaleSupplyPrice = orgSupplyPrice
			end if		
		Case 7	'��ü��Ź,�����Ź,�ٹ�������Ź �ݹݺδ�/�������ٹ����ٺδ�
			if comm_cd = "B011" or comm_cd = "B012" or comm_cd = "B013" then
				fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
			else
				fnSetSaleSupplyPrice = orgSupplyPrice
			end if		
	END SELECT	
End Function

'//��¥ ���� '/�̺�Ʈ�Ⱓ,����Ⱓ
function draweventmaechul_datefg(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>����</option>	
	<option value='event' <%if selectedId="event" then response.write " selected"%>>�̺�Ʈ�Ⱓ</option>
	<option value='jumun' <%if selectedId="jumun" then response.write " selected"%>>����Ⱓ</option>
</select>
<%
end function

'/���� ó��(��ü)
function offitemsaleSet_all()
dim sql
		
	sql = "exec db_shop.dbo.sp_Ten_item_SetPrice_off "
	
	'response.write sql &"<Br>"
	dbget.execute sql
end function
%>