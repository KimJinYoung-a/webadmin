<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : OkCashbag����
' History : ������ ����
'			2023.03.22 �ѿ�� ����(���� ���� ���̵� ���� �ִºκ� ���� ���� ������ �ڵ�ȭ. �ҽ� ǥ���ڵ�� ����.)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/othermall/okcashbagCls.asp"-->
<%
'if (NOT C_ADMIN_AUTH) then
'    response.write "�����ڸ� ���� �����մϴ�. ������ ���� ���"
'    dbget.Close() :response.end
'end if

' ������ �̰ų� ���߿�� �̰ų� ������Ʈ �ϰ��
If not(C_ADMIN_AUTH or C_SYSTEM_Part or C_partnership_part) Then
    response.write "������ �� �ش� ����ڸ�  ���� �����մϴ�. ������ ���� ���"
    dbget.Close() :response.end
end if

dim ArrIDX
ArrIDX = request("arod")

dim sSdate,sEdate, userid, orderserial, SearchDateType, vRdSite
sSdate 		= requestCheckVar(Request("iSD"),10)
sEdate 		= requestCheckVar(Request("iED"),10)
userid 		= requestCheckVar(Request("uId"),32)
orderserial	= requestCheckVar(Request("oSn"),12)
SearchDateType = requestCheckVar(request("dType"),2)
IF SearchDateType="" THEN SearchDateType="od"

vRdSite		= requestCheckVar(Request("rdsite"),10)
If vRdSite = "" Then
	vRdSite = "okcashbag"
End If

dim OrderType
OrderType = requestCheckVar(Request("otp"),2)
IF OrderType="" Then OrderType="no"

dim CurrPage
CurrPage = requestCheckVar(request("pg"),3)
IF CurrPage="" THEN CurrPage =1

dim sPageSize
	sPageSize = 10000	' ���� �ڵ��� �̷��� �Ǿ� �־ ��¿�� ��� �켱 ���� 1������ �ھƳ���.. �� �̻� �þ��� ����¡ ������ getrows �� �޾ƿ;���.

dim oCash,intLp
Set oCash = New CashbagCls
oCash.FCurrPage=CurrPage
oCash.FPageSize=sPageSize
oCash.FArrIDX = ArrIDX
oCash.FStartDate 	= sSdate
oCash.FEndDate 		= sEdate
oCash.Forderserial 	= orderserial
oCash.FOrderType 	= OrderType
oCash.FSearchType	= SearchDateType
oCash.FRdSite		= vRdSite

IF OrderType="N" Then 		'//����� ������Ʈ
	oCash.updateNormalOrder()
ELSEIF OrderType ="C" Then	'//��Ұ�  ������Ʈ
	oCash.updateCancelOrder()
ELSEIF OrderType="UN" or OrderType ="UC" Then '// ��� �� ���� (����,���)
	oCash.getUpdatedOrder()
END IF

downPersonalInformation_rowcnt=oCash.FTotalCount
%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
dim SaveFilename
SaveFilename = "okcashbag.xls"

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_" & SaveFilename & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '���ۻ�뿩��
%>

<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<style type="text/css">
.mso {mso-number-format:"\@";}
		</style>

	</head>
	<body>
		<table width="100%" align="center" border="1" cellpadding="3" cellspacing="1" class="mso">
			<tr bgcolor="<%= adminColor("sky") %>">
	<!--<td align="center" width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll(this.checked);"></td>-->
	<td align="center" width="100" >�ֹ���ȣ</td>
	<td align="center" width="80">��ٱ��Ϲ�ȣ</td>
	<td align="center">�Ѱ����ݾ�</td>
	<td align="center">�ֹ�����</td>
	<td align="center">�������</td>
	<td align="center">�ֹ���</td>
	<td align="center">ĳ�����ȣ</td>
	<td align="center">��������Ʈ</td>
			</tr>
<% IF oCash.FResultcount<=0 Then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center"> ��ġ�ϴ� ����Ÿ�� �����ϴ�.</td>
	</tr>
<% ELSE %>

	<% FOR intLp=0 To oCash.FResultcount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center" style="mso-number-format:'\@'"><%= oCash.FItemList(IntLp).FOrderSerial %></td>
		<td align="center" style="mso-number-format:'\@'"><%= oCash.FItemList(IntLp).FShoppingBagNo %></td>
		<td align="center" style="mso-number-format:'\@'"><%= FormatNumber(oCash.FItemList(IntLp).FPointCash,0) %></td>
		<td align="center" style="mso-number-format:'\@'"><%= replace(DateValue (oCash.FItemList(IntLp).FRegdate),"-","") %></td>
		<td align="center" style="mso-number-format:'\@'"><%= replace(DateValue (oCash.FItemList(IntLp).FBeadaldate),"-","") %></td>
		<td align="center" style="mso-number-format:'\@'"><%= oCash.FItemList(IntLp).FBuyName %></td>
		<td align="center" style="mso-number-format:'\@'"><%= oCash.FItemList(IntLp).FCashBagCardNo %></td>
		<td align="center" style="mso-number-format:'\@'"><%= FormatNumber(oCash.FItemList(IntLp).FPoint,0) %></td>
	</tr>

	<%
        if intLp mod 500 = 0 then
            Response.Flush		' ���۸��÷���
        end if
	NEXT
	%>
<% End IF %>
		</table>
	</body>
</html>
<script>opener.document.location.reload();</script>

<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
