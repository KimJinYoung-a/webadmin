<%
'####################################################
' Page : /lib/event_function.asp
' Description :  �̺�Ʈ �Լ� 
' History : 2007.02.07 ������ ����
'####################################################
 
'-----------------------------------------------------------------------  
' 1. sbGetDesignerid :���������� �μ���ȣ(12)�� �����̳��̸� ����Ʈ��������
' 2007.02.07 ������ ����
'-----------------------------------------------------------------------
Sub sbGetDesignerid(ByVal selName, ByVal sIDValue, ByVal sScript)
	Dim strSql, arrList, intLoop
	strSql = " SELECT userid, username from db_partner.[dbo].tbl_user_tenbyten WHERE  part_sn ='12' and isUsing=1" & vbcrlf

	' ��翹���� ó��	' 2018.10.16 �ѿ��
	strSql = strSql & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	strSql = strSql & "	and userid <> ''" & vbcrlf  
	strSql = strSql & "	order by posit_sn, empno" & vbcrlf

	'response.write strSql & "<br>"
   rsget.Open strSql,dbget,1
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
 

'-----------------------------------------------------------------------  
' 2.fnSetEventCommonCode : �̺�Ʈ �����ڵ� ���ú����� ����
' 2007.02.07 ������ ����
'----------------------------------------------------------------------- 
 Function fnSetEventCommonCode
	Dim strSql, arrList, intLoop
	Dim intI, intJ, arrCode(), strtype
	strSql = " SELECT code_type, code_value, code_desc FROM [db_event].[dbo].[tbl_event_commoncode] WHERE code_using ='Y' "
	rsget.Open strSql,dbget,1
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
	
'--------------------------------------------------------------------------------   
' 3.sbGetOptEventCodeValue(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
' : �̺�Ʈ �����ڵ� ���ú��� select ����ȭ 
' 2007.02.07 ������ ����
'--------------------------------------------------------------------------------  
 Sub sbGetOptEventCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop
 	arrList = Application(sType)
 	IF  isNull(selValue) THEN selValue = ""
%>
	<select class="select" name="<%=sType%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">����</option>
    <%END IF%>
<% 	
 	For intLoop =0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
<%  Next %>
	<select>
<%
 End Sub

'-----------------------------------------------------------------------  
' 4.fnGetEventCodeDesc : �̺�Ʈ �����ڵ� ���� ���� �̸� ��������
' 2007.02.07 ������ ����
'----------------------------------------------------------------------- 
 Function fnGetEventCodeDesc(ByVal sType, ByVal selValue)

	Dim arrList, intLoop
	arrList = Application(sType)	 
		
	' ���ø����̼� ���� ���� üũ 
	If IsEmpty(arrList) OR not isArray(arrList) Then Exit Function

	' �ڵ� ����
	selValue = Trim(selValue)

	' �ڵ�� ã�� ��ƾ
	For intLoop = 0 To Ubound(arrList, 2)
		' �ڵ�� �� �� ������ ��ȯ	
			If CStr(selValue) = CStr(arrList(0, intLoop)) Then fnGetEventCodeDesc = arrList(1, intLoop) : Exit For
	Next
End Function

'-----------------------------------------------------------------------  
' 5.sbGetOptCategoryLarge : ī�װ� ��з���������
' 2007.02.07 ������ ����
' 2008.04.15 ������ ����; ��ī�װ��� ����
'----------------------------------------------------------------------- 
 Sub sbGetOptCategoryLarge(byval selectBoxName,byval selectedId, ByVal strEtc)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" <%=strEtc%>>   	
     <option value="" <% if selectedId="" then response.write " selected"%>>����</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"& db2html(rsget("code_nm")) &"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

'-----------------------------------------------------------------------  
' 6.sbGetStaticEvent : �����̺�Ʈ ���� ��������
' 2007.02.07 ������ ����
'----------------------------------------------------------------------- 
	Sub sbGetStaticEvent(ByVal selValue)
%>
	<select class="select" name="selStatic">
	<option value="">����</option>
	<option value="�����ǰ�ı�" <%IF Trim(selValue) = "�����ǰ�ı�" THEN%>selected<%END IF%>>�����ǰ�ı�</option>
	<option value="���ٳ���" <%IF Trim(selValue) = "���ٳ���" THEN%>selected<%END IF%>>���ٳ���</option>
	<option value="�������" <%IF Trim(selValue) = "�������" THEN%>selected<%END IF%>>�������</option>
	<option value="100%SHOP" <%IF Trim(selValue) = "100%SHOP" THEN%>selected<%END IF%>>100% SHOP</option>
	<option value="�����Ͽ콺" <%IF Trim(selValue) = "�����Ͽ콺" THEN%>selected<%END IF%>>�����Ͽ콺</option>
	</select>
<%		
	End Sub
	
'-----------------------------------------------------------------------  
' 7.GetImageFolerName : ������ ��������
' 2007.02.07 ������ ����
'----------------------------------------------------------------------- 	
	function GetImageFolerName(byval itemid)
		GetImageFolerName = "0" + CStr(Clng(itemid\10000))
	end function
	
'-----------------------------------------------------------------------  
' 8.fnSetDispUrl : �̺�Ʈ ������ ���� ��ũ ����
' 2007.02.07 ������ ����
'----------------------------------------------------------------------- 	
	Function fnSetDispUrl(ByVal sKind, ByVal eCode)
		SELECT CASE sKind
			CASE 7
				fnSetDispUrl = "/admin/sitemaster/weekly_codi_detail.asp?mode=add&eC="&eCode
			CASE ELSE
				fnSetDispUrl = "event_modify.asp?eC="&eCode&"&menupos="&menupos
		END SELECT		
	End Function
	
'-----------------------------------------------------------------------  
' 9.sbAlertMsg : �˸����� �� ������ �̵� ó��
' 2007.02.07 ������ ����
'----------------------------------------------------------------------- 	
	Sub sbAlertMsg(byVal strMsg, ByVal strUrl, ByVal strTarget)
		Dim strLink
		IF strUrl = "close" THEN	'�˾� â �������
			strLink = strTarget & ".close();"
		ELSEIF strUrl ="back" THEN	'���� �������� �̵�
			strLink = "history.back(-1);"
		ELSE
			strLink = strTarget & ".location.href='" & strUrl & "';"
		END IF		
%>
	<script language="javascript">
	<!--
		alert("<%=strMsg%>");
		<%=strLink%>;
	//-->
	</script>
<%		dbget.close()	:	response.End
	End Sub

'--------------------------------------------------------------------------------   
' 10.sbGetOptGiftCodeValue(������, ���ð�,  �׷켱�ñ��� �������, ��ũ��Ʈ)
' : ����ǰ�����ڵ� ���ú��� select ����ȭ 
' ���ܻ���: �׷켱���� ��� ������ view 
' 2007.05.09 ������ ����
'--------------------------------------------------------------------------------  
 Sub sbGetOptGiftCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop
 	arrList = Application(sType)
 	IF  isNull(selValue) THEN selValue = ""
%>
	<select class="select" name="<%=sType%>" <%=sScript%>>	
	<% if selValue="1" then %>
	        <option value="1" selected >��ü����</option>
	<% end if %>
	
<% 	
 	For intLoop =0 To UBound(arrList,2)  	
 	 IF ((NOT sViewOpt and  Cstr(arrList(0,intLoop)) <> "4") OR sViewOpt )THEN 	 	
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
<%   END IF	
	Next %>
	<select>
<%
 End Sub
%>