<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->

<%
	function getMatchXsiteOrderitemInfo(ioutMallorderSeq)
		Dim sqlStr 
		sqlStr = "select orderItemName, orderItemOptionName, orderItemOption "
		sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_tmpOrder"
		sqlStr = sqlStr & " where outMallorderSeq='"&ioutMallorderSeq&"'"

		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			getMatchXsiteOrderitemInfo = rsget.getRows()
		end if
		rsget.close
	end function
'Matchitemoption �� ������Ʈ

	Dim vOutMallOrderSeq, vMatchItemID, vIsOption
	Dim vmatchitemoption
	vOutMallOrderSeq	= requestCheckvar(request("outMallorderSeq"),32)
	vMatchItemID		= requestCheckvar(request("Matchitemid"),32)
	vMatchitemoption     = requestCheckvar(request("matchitemoption"),10)

	Dim oitemoption, i
	set oitemoption = new CItemOption
	oitemoption.FRectItemID = vMatchItemID
	If vMatchItemID <> "" Then
		oitemoption.GetItemOptionInfo
	End If

	Dim arrRows, orderItemOptionName, orderItemName
	if (vOutMallOrderSeq<>"") then
		arrRows = getMatchXsiteOrderitemInfo(vOutMallOrderSeq)
		if isArray(arrRows) then
			orderItemOptionName = ArrRows(1,0)
			orderItemName = ArrRows(0,0)
		end if
	end if
%>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<script language="javascript">
function ModiOrder()
{
	if(frmOption.matchitemoption.value == "")
	{
		alert("���� �ɼ��� �������ּ���");
		frmOption.matchitemoption.focus();
		return;
	}
	frmOption.mode.value="optmatch";
	frmOption.submit();
}

function ModiOrderOption0000(){
    if (confirm('�ɼ� �������� �����Ͻðڽ��ϱ�?')){
        frmOption.mode.value="optnone";
	    frmOption.submit();
    }
}

</script>

<table width="100%" align="left" cellpadding="3" cellspacing="0" class="table_tl">

	<tr height="25">
		<td class="td_br" colspan="2">
			<b>���� �ɼ� ����</b>
		</td>
	</tr>
	<form name="frmOption" method="post" action="popMatchItemOptionProc.asp">
	<input type="hidden" name="outMallorderSeq" value="<%=vOutMallOrderSeq%>">
	<input type="hidden" name="Matchitemid" value="<%=vMatchItemID%>">
	<input type="hidden" name="mode" value="optmatch">
	<input type="hidden" name="tmpisusing">

	<tr>
		<td width="80" align="right" class="td_br_tablebar">�ɼ�:</td>
		<td class="td_br">
		<select name="matchitemoption">
		<%
			If oitemoption.FResultCount > 0 Then
				Response.Write "<option value="""">-����-</option>"
				For i=0 To oitemoption.FResultCount - 1
					If oitemoption.FITemList(i).FOptIsUsing = "Y" Then
						Response.Write "<option value=""" & oitemoption.FITemList(i).FItemOption & """>" & oitemoption.FITemList(i).FOptionName & "</option>"
						vIsOption = "o"
'2014-07-03 18:25 ������ �ϴ� else�� �ּ�ó��..�ɼ� �����ÿ� ��������� �ɼǵ� �����̵�..�Է��ϴ� �� �����߻� ���..
'2016-02-24 14:50 ������ �ϴ� else�� �ּ�Ǯ��..���� FF�� �����ϴ� �ɼ� ������ �̻��� �ɼ����� ��Ī���Ѽ� ������..
'				  �׿����� ������� �ɼ��� ���ý� alert�� ��� �� CS�� �����϶�� ���;� �� �� ����..
				    else
				        Response.Write "<option value=""" & oitemoption.FITemList(i).FItemOption & """ style='color:#CCCCCC'>" & oitemoption.FITemList(i).FOptionName & "(�������)</option>"
						vIsOption = "o"
					End If
				Next
			Else
				Response.Write "<option value="""">�ɼǾ���</option>"
				vIsOption = "x"
			End If

			set oitemoption = Nothing
		%>
		</select>

		&nbsp;<%=orderItemOptionName%> | <%=orderItemName%>
		</td>
	</tr>
	<tr>
		<td class="td_br" colspan="2">
			�� �ɼ� ��������� �ɼ��� ���ý� CS�� �������ּ���
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2" class="td_br">
		<% If vIsOption = "o" Then %>
		    <input type="button" class="button" value="����" onClick="ModiOrder();">
		    <input type="button" class="button" value="���" onClick="self.close();">
		<% Else %>
			<input type="button" class="button" value="�ݱ�" onClick="window.close()">
			<% if (vMatchitemoption<>"0000") then %>
			&nbsp;<input type="button" class="button" value="�ɼǾ������� ����" onClick="ModiOrderOption0000()">
			<% end if %>
		<% End If %>
		</td>
	</tr>
	</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->