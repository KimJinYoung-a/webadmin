<%
'####################################################
' Page : /lib/event_function.asp
' Description :  �̺�Ʈ �Լ�
' History : 2007.02.07 ������ ����
'####################################################

'-----------------------------------------------------------------------
' 1. sbGetDesignerid :���������� �μ���ȣ(12)�� �����̳��̸� ����Ʈ��������
' 2007.02.07 ������ ����
'2011.01.18 ������ ���� : �μ����� ���̺���(tbl_partner -> tbl_user_tenbyten)
'-----------------------------------------------------------------------
 Sub sbGetDesignerid(ByVal selName, ByVal sIDValue, ByVal sScript)
  Dim strSql, arrList, intLoop
   strSql = " SELECT userid,  username"
   strSql = strSql & "  from db_partner.dbo.tbl_user_tenbyten with (noLock) "
   strSql = strSql & " WHERE part_sn ='12' and isUsing=1 and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> ''  order by posit_sn, empno"

   rsget.Open strSql,dbget
   IF not rsget.eof THEN
   	arrList = rsget.getRows()
   End IF
   rsget.close
%>
	<select name="<%=selName%>" <%=sScript%> class="Select">
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
 On Error Resume Next
	Dim strSql, arrList, intLoop
	Dim intI, intJ, arrCode(), strtype
	strSql = " SELECT code_type, code_value, code_desc FROM [db_event].[dbo].[tbl_event_commoncode] WHERE code_using ='Y' Order by code_type, code_sort "
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

'--------------------------------------------------------------------------------
' 3.sbGetOptEventCodeValue(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
' : �̺�Ʈ �����ڵ� ���ú��� select ����ȭ
' 2007.02.07 ������ ����
'--------------------------------------------------------------------------------
 Sub sbGetOptEventCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList = Application(sType)
 	iValue  = selValue
 	IF  isNull(selValue) THEN 		selValue = ""
 	IF selValue = ""  THEN	iValue = 0

%>
	<select name="<%=sType%>" id="<%=sType%>" <%=sScript%> class="select">
	<%IF sViewOpt THEN%>
	<option value="">����</option>
    <%END IF%>
<%
IF isArray(arrList) then
 	For intLoop =0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>>
	<%
		If arrList(1,intLoop) = "100px" Then
			Response.Write "130px"
		ElseIf arrList(1,intLoop) = "150px" Then
			Response.Write "180px"
		ElseIf arrList(1,intLoop) = "200px" Then
			Response.Write "400px"
		ElseIf arrList(1,intLoop) = "155px" Then
			Response.Write "270px"
		ElseIf arrList(1,intLoop) = "160px" Then
			Response.Write "320px"
		Else
			Response.Write arrList(1,intLoop)
		End If
	%>
	</option>
<%
	Next
end if	
	 %>
	</select>
<%
 End Sub
'--------------------------------------------------------------------------------
' 3-1.sbGetOptEventCodeValue(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
' : �̺�Ʈ �����ڵ� ���ú��� select ����ȭ, ������ ���� �߰�
' 2007.02.07 ������ ����
'--------------------------------------------------------------------------------
 Sub sbGetVarOptEventCodeValue(ByVal sVarName, ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList = Application(sType)
 	iValue  = selValue
 	IF  isNull(selValue) THEN 		selValue = ""
 	IF selValue = ""  THEN	iValue = 0

%>
	<select name="<%=sVarName%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">����</option>
    <%END IF%>
<%
 	For intLoop =0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>>
	<%
		If arrList(1,intLoop) = "100px" Then
			Response.Write "130px"
		ElseIf arrList(1,intLoop) = "150px" Then
			Response.Write "180px"
		ElseIf arrList(1,intLoop) = "200px" Then
			Response.Write "400px"
		ElseIf arrList(1,intLoop) = "155px" Then
			Response.Write "270px"
		ElseIf arrList(1,intLoop) = "160px" Then
			Response.Write "320px"
		Else
			Response.Write arrList(1,intLoop)
		End If
	%>
	</option>
<%
	Next %>
	</select>
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
'-----------------------------------------------------------------------
 Sub sbGetOptCategoryLarge(byval selectBoxName,byval selectedId, ByVal strEtc)
   dim tmp_str,query1
   %><select name="<%=selectBoxName%>" <%=strEtc%>>
     <option value="" <% if selectedId="" then response.write " selected"%>>����</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large"
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsget.Open query1,dbget
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
' 5.sbGetOnlyOptCategoryLarge : ī�װ� ��з���������
' 2007.02.07 ������ ����
'-----------------------------------------------------------------------
 Sub sbGetOnlyOptCategoryLarge(byval selectedId)
   dim tmp_str,query1
   %>
     <option value="" <% if selectedId="" then response.write " selected"%>>����</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large"
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsget.Open query1,dbget
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

end Sub
'-----------------------------------------------------------------------
' 6.sbGetStaticEvent : �����̺�Ʈ ���� ��������
' 2007.02.07 ������ ����
'-----------------------------------------------------------------------
	Sub sbGetStaticEvent(ByVal selValue)
%>
	<select name="selStatic">
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
	<script type="text/javascript">
		alert("<%=strMsg%>");
		<%=strLink%>;
	</script>
<%		response.End
	End Sub

'--------------------------------------------------------------------------------
' 10.sbGetOptGiftCodeValue(������, ���ð�,  �׷켱�ñ��� �������, ��ũ��Ʈ,�̺�Ʈ�ڵ�)
' : ����ǰ�����ڵ� ���ú��� select ����ȭ
' ���ܻ���: �׷켱���� ��� ������ view
' 2007.05.09 ������ ����
'--------------------------------------------------------------------------------
 Sub sbGetOptGiftCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript,ByVal eCode)
   Dim arrList, intLoop
 	arrList = fnSetCommonCodeArr(sType, True)
 	IF  isNull(selValue) THEN selValue = ""
%>
    <select name="<%=sType%>" class="select" <%=sScript%>>
	<% if selValue="1" then %>
	        <option value="1" selected >��ü����</option>
	<% end if %>

<%
 	For intLoop =0 To UBound(arrList,2)

 	IF ((NOT sViewOpt and  Cstr(arrList(0,intLoop)) <> "4") OR sViewOpt) AND (not (Cstr(eCode) <> ""  AND  Cstr(arrList(0,intLoop)) = "5") ) AND ( not (Cstr(eCode) ="" AND ( Cstr(arrList(0,intLoop)) = "2" OR Cstr(arrList(0,intLoop)) = "6") )) THEN
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
<%   END IF
	Next %>
	<% if selValue="9" then %>
	        <option value="9" selected >���̾���Ž�(+�ٹ��)</option>
	 <% end if %>
	</select>
<%
 End Sub

'--------------------------------------------------------------------------------
' 10.sbOptPartner
'	:Ư�� ���޸�
' 2008.03.24 ������ ����
'--------------------------------------------------------------------------------
Sub sbOptPartner(ByVal selPartner)
	Dim strSql, arrList,intLoop, selvalue
	strSql = "SELECT id, company_name from [db_partner].[dbo].tbl_partner where userdiv=999 and isusing='Y'"
	 rsget.Open strSql,dbget
	 IF not rsget.Eof THEN
	 	arrList = rsget.getRows()
	 END IF
	 rsget.close

	 IF isArray(arrList) THEN
	 	For intLoop =0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(selPartner) = arrList(0,intLoop) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%> (<%=arrList(0,intLoop)%>)</option>
<%	 	Next
	 END IF
End Sub
'--------------------------------------------------------------------------------
' 11.sbGetOptStatusCodeValue(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
' : �̺�Ʈ ���°� �����ڵ� ���ú��� select ����ȭ - ���簪�� ���º��� ���������� �̵� ���ϵ���
' 2007.02.07 ������ ����
'--------------------------------------------------------------------------------
 Sub sbGetOptStatusCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList= fnSetCommonCodeArr(sType, True)
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
	</select>
<%
 End Sub

'--------------------------------------------------------------------------------
' 11-1.sbGetOptStatusCodeAuth(������, ���ð�, ��뱸��, ��ũ��Ʈ)
' 2013.04.25 ������ ����
'--------------------------------------------------------------------------------
 Sub sbGetOptStatusCodeAuth(sType, selValue, sViewOpt, sScript)
   Dim arrList, intLoop, iValue
 	arrList= fnSetCommonCodeArr(sType, True)
 	iValue  = selValue
 	IF isNull(selValue) THEN	selValue = ""
 	IF selValue="" THEN			iValue = 0

	'// �̺�Ʈ ������ ����(��������; MD��Ʈ/������ ���� �� ������)
	Dim uMng : uMng=false
	if (session("ssAdminLsn")<=3 and (session("ssAdminPsn")=11 or session("ssAdminPsn")=14)) or (session("ssAdminLsn")=1) then
		uMng = true
	end if

	Response.Write "<select name='" & sType & "' " & sScript & ">" & vbCrLf

	For intLoop =0 To UBound(arrList,2)
		Select Case sViewOpt
			Case "N"
				'# �űԵ�� �� (��ϴ��, ���ο�û�� ���)
				If Cint(arrList(0,intLoop))=0 or Cint(arrList(0,intLoop))=2 then
					Response.Write "<option value='" & arrList(0,intLoop) & "' " & chkIIF(CStr(selValue)=CStr(arrList(0,intLoop)),"selected","") & ">" & replace(arrList(1,intLoop),"���¿���","����")  & "</option>" & vbCrLf
				end if
			Case Else
				if uMng then
					'//�̺�Ʈ �����ڶ�� ��ü ���� ���
					IF Cint(arrList(0,intLoop)) >= Cint(iValue) or (Cint(iValue)<=2 and Cint(arrList(0,intLoop))=1) THEN
			 	 		Response.Write "<option value='" & arrList(0,intLoop) & "' " & chkIIF(CStr(selValue)=CStr(arrList(0,intLoop)),"selected","") & ">" & replace(arrList(1,intLoop),"���¿���","����")  & "</option>" & vbCrLf
					END IF
				else
					IF Cint(iValue)>2 THEN
						IF Cint(arrList(0,intLoop)) >= Cint(iValue) THEN
			 	 			Response.Write "<option value='" & arrList(0,intLoop) & "' " & chkIIF(CStr(selValue)=CStr(arrList(0,intLoop)),"selected","") & ">" & replace(arrList(1,intLoop),"���¿���","����")  & "</option>" & vbCrLf
			 	 		End if
					ElseIf (Cint(arrList(0,intLoop))<>1 and Cint(arrList(0,intLoop))<=2 and Cint(arrList(0,intLoop)) >= Cint(iValue)) or (Cint(iValue)=1 and Cint(arrList(0,intLoop))=1) then
						Response.Write "<option value='" & arrList(0,intLoop) & "' " & chkIIF(CStr(selValue)=CStr(arrList(0,intLoop)),"selected","") & ">" & replace(arrList(1,intLoop),"���¿���","����")  & "</option>" & vbCrLf
					END IF
				end if
		End Select
	Next

 	 Response.Write "</select>" & vbCrLf

 End Sub

'--------------------------------------------------------------------------------
' 11-2.sbGetOptStatusCodeSort(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
' : �̺�Ʈ ���°� �����ڵ� ���ú��� select ����ȭ - ���簪�� ���º��� ���������� �̵� ���ϵ���
' 2015.07.09 ������ ����
'--------------------------------------------------------------------------------
 Sub sbGetOptStatusCodeSort(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
   Dim strSql
  
  	iValue  = selValue
 	IF  isNull(selValue) THEN 		selValue = ""
 	IF selValue = ""  THEN	iValue = 0
 	strSql = ""    
	strSql = " SELECT code_value, code_desc,code_dispYN "&vbcrlf
	strSql = strSql & " FROM [db_event].[dbo].[tbl_event_commoncode] "&vbcrlf
	strSql = strSql & " WHERE code_type='"&sType&"' and code_dispYN ='Y' and code_using ='Y' "&vbcrlf
	strSql = strSql & "    and code_sort >= ( select code_Sort from db_event.dbo.tbl_Event_commoncode where code_type='"&sType&"' and  code_dispYN ='Y' and code_value = '"&iValue&"' and code_using ='Y')"&vbcrlf
	strSql = strSql &" Order by code_dispYN desc, code_type, code_sort "
     rsget.Open strSql,dbget
	IF not rsget.EOF THEN
	    arrList = rsget.getRows()
	 END IF
	rsget.close   
%>
	<select name="<%=sType%>" <%=sScript%> class="select">
	<%IF sViewOpt THEN%>
	<option value="">����</option>
    <%END IF%>
<%
 	For intLoop =0 To UBound(arrList,2) 
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=replace(arrList(1,intLoop),"���¿���","����")%></option>
<%    
	Next %>
	</select>
<%
 End Sub
'-----------------------------------------------------------------------
' 12.fnGetCommCodeArrDesc : Ư�������� �����ڵ尪�� �迭���� Ư������ �ڵ�� ��������
'-----------------------------------------------------------------------
	Function fnGetCommCodeArrDesc(ByVal arrCode, ByVal iCodeValue)
		Dim intLoop
		IF iCodeValue = "" or isNull(iCodeValue) THEN iCodeValue = -1
		For intLoop =0 To UBound(arrCode,2)
	 
			IF Cint(iCodeValue) = arrCode(0,intLoop) THEN
				fnGetCommCodeArrDesc = arrCode(1,intLoop)
				Exit For
			END IF
		Next
	End Function

'-----------------------------------------------------------------------
' 13.fnSetCommonCodeArr : �̺�Ʈ �����ڵ� ��������
' 2007.02.07 ������ ����
'-----------------------------------------------------------------------
 Function fnSetCommonCodeArr(ByVal code_type, ByVal blnUse)
	Dim strSql, arrList, intLoop, strAdd
	Dim intI, intJ, arrCode(), strtype
	strAdd = ""
	IF blnUse THEN
		strAdd= " and code_using ='Y' "
	END IF
	strSql = " SELECT code_value, code_desc,code_dispYN FROM [db_event].[dbo].[tbl_event_commoncode] WHERE code_type='"&code_type&"' and code_dispYN ='Y' "&strAdd&_
			" Order by code_dispYN desc, code_type, code_sort "
	rsget.Open strSql,dbget
	IF not rsget.EOF THEN
		fnSetCommonCodeArr = rsget.getRows()
	END IF
	rsget.close
End Function

'--------------------------------------------------------------------------------
' 14.sbGetOptCommonCodeArr(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
' : Ư�������� �����ڵ尪�� �迭���� select ó��
' 2008.04.15 ������ ����
'--------------------------------------------------------------------------------
 Sub sbGetOptCommonCodeArr(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal blnUse, ByVal sScript)
   Dim arrCode, intLoop, iValue
   	arrCode= fnSetCommonCodeArr(sType, blnUse)
 	iValue  = selValue
 	IF  isNull(selValue) THEN 	selValue = ""
 	IF selValue = ""  THEN	iValue = 0
%>
	<select name="<%=sType%>" class="select" <%=sScript%>>
	<%
		IF sViewOpt THEN
			Response.Write "<option value="""">����</option>" &vbCrLf
		END IF
		IF sType="eventkind" THEN
			'if (session("ssAdminPsn")="11" or session("ssAdminPsn")="21")   then 'MD�μ���� (��������,��ü,��ǰ,�귣��,���̾,�׽���,�űԵ����̳�) 
				Response.Write "<option value=""1,12,13,23,27,28,29""  " & CHKIIF(CStr(selValue) = "1,12,13,23,27,28,29", "selected", "") & ">#�����׸�</option>" &vbCrLf 
			'end if
			'Response.Write "<option value=""1,12,13,16,17,23,24,28"" " & CHKIIF(CStr(selValue) = "1,12,13,16,17,23,24,28", "selected", "") & " >#�Ϲ� �̺�Ʈ</option>" &vbCrLf
			'Response.Write "<option value=""19,25,26"" " & CHKIIF(CStr(selValue) = "19,25,26", "selected", "") & " >#����� or �� ����</option>" &vbCrLf
		END IF
	%>
<% 	IF isArray(arrCode) THEN
 	For intLoop =0 To UBound(arrCode,2)
%>
	<option value="<%=arrCode(0,intLoop)%>" <%If CStr(selValue) = CStr(arrCode(0,intLoop)) THEN%>selected<%END IF%> <%if arrCode(2,intLoop) ="N" then%>style="color:gray;"<%end if%>> <%=arrCode(1,intLoop)%></option>
<%
	Next
	End IF
	%>
	</select>
<%
 End Sub

 '--------------------------------------------------------------------------------
' 15.fnSetStatusDesc
' : ���°��� ���� ���¸�
' 2008.04.15 ������ ����
'--------------------------------------------------------------------------------
	Function fnSetStatusDesc(ByVal FState, ByVal FSDate, ByVal FEDate, ByVal FStateDesc)
		IF FState = "7" AND datediff("d",FSDate,date()) >= 0 and datediff("d",FEDate,date()) <=0 THEN
			fnSetStatusDesc = "����"
		ELSEIF FState ="7" AND datediff("d",FEDate,date()) > 0 THEN
			fnSetStatusDesc = "����"
		ELSE
			fnSetStatusDesc = FStateDesc
		END IF
	End Function

 '-----------------------------------------------------------------------
' 16. sbGetMDid :���MD ����Ʈ�������� (���� �̸�,���ް�� �̻�)
' 2010.01.25 ������ ���� / '2011.01.18 ������ ���� : �μ����� ���̺���(tbl_partner -> tbl_user_tenbyten)
'-----------------------------------------------------------------------
 Sub sbGetMDid(ByVal selName, ByVal sIDValue, ByVal sScript)
  Dim strSql, arrList, intLoop
   strSql = " SELECT userid, username"
   strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "
   strSql = strSql & " WHERE part_sn ='11' and  posit_sn>='4' and  posit_sn<='12' and   isUsing=1  and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '' order by posit_sn, empno"

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


  '-----------------------------------------------------------------------
' 17. sbGetMKTid :���MKT,MD ����Ʈ�������� (���� �̸�,���� �̻�)
' 2011.06.20 ���ر� ����
'-----------------------------------------------------------------------
 Sub sbGetMKTid(ByVal selName, ByVal sIDValue, ByVal sScript)
  Dim strSql, arrList, intLoop
   strSql = " SELECT userid, username"
   strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "
   strSql = strSql & " WHERE part_sn in('11','14') and  posit_sn>='4' and  posit_sn<='8' and   isUsing=1  and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '' order by part_sn, posit_sn, empno"

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

	Sub sbGetwork(ByVal selName, ByVal sIDValue, ByVal sScript)
		Dim strSql, arrList, intLoop
		strSql = " SELECT userid, username "
		strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "
		strSql = strSql & " WHERE userid = '"&sIDValue&"' and userid <> '' "
		rsget.Open strSql,dbget
		IF not rsget.eof THEN
			arrList = rsget.getRows()
		End IF
		rsget.close

		IF isArray(arrList) THEN
%>
			<input type="text" class="text" name="doc_workername" value="<%=arrList(1,0)%>" size="10" readonly>
			<input type="button" class="button" value="����" onClick="workerlist()">
			<input type="button" class="button" value="X" onClick="workerDel()">
<% 			If selName = "selMKTId" Then %>
				<input type="hidden" name="selMKTId" value="<%=arrList(0,0)%>">
<%			Else %>
				<input type="hidden" name="selMId" value="<%=arrList(0,0)%>">
<%
			End If
		Else
%>
			<input type="text" class="text" name="doc_workername" value="" size="10" readonly>
			<input type="button" class="button" value="����" onClick="workerlist()">
			<input type="button" class="button" value="X" onClick="workerDel()">

<%			If selName = "selMKTId" Then %>
				<input type="hidden" name="selMKTId" value="">
<%			Else %>
				<input type="hidden" name="selMId" value="">
<%			End If
		End IF
	End Sub

'--------------------------------------------------------------------------------
' 14.GetEvnetKindName(������, ���ð�)
' 2019.01.22 ������ ����
'--------------------------------------------------------------------------------
Sub GetEvnetKindName(ByVal sType, ByVal selValue)
	Dim arrCode, intLoop
	arrCode= fnSetCommonCodeArr(sType, True)
	IF isArray(arrCode) THEN
		For intLoop=0 To UBound(arrCode,2)
			if selValue = arrCode(0,intLoop) then
				if (sType="jobkind" or sType="placekind") and arrCode(0,intLoop)>1 then
					response.write arrCode(1,intLoop)
				end if
			end if
		Next
	End IF
End Sub

%>
