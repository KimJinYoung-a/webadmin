<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �������� ������
' History : ���ʻ����ڸ�
'			2017.04.13 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<%
dim masteridx
masteridx = requestCheckVar(request("masteridx"),10)

%>
<script language='javascript'>
function popItemWindow(iid,frm){
	if (frmarr.masteridx.value == "")	{
		alert("������ ������ �������ּ���...");
		frmarr.masteridx.focus();
		return;
	}
	else{
	window.open("/admin/pop/viewitemlist.asp?designerid=" + iid + "&target=" + frm, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	}
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function AddIttems(){
	if (frmarr.masteridx.value == "")	{
		alert("������ ������ �������ּ���...");
		frmarr.masteridx.focus();
		return;
	}
	if (confirm(frmarr.itemid.value + '�������� �߰��Ͻðڽ��ϱ�?')){
		frmarr.itemid.value = frmarr.itemid.value;
		frmarr.mode.value="add";
		frmarr.submit();
	}
}

function AddIttems2(){
	if (frmarr.masteridx.value == "")	{
		alert("������ ������ �������ּ���...");
		frmarr.masteridx.focus();
		return;
	}
	if (frmarr.itemidarr.value == ""){
		alert("�����۹�ȣ�� �Է����ּ���...");
		frmarr.itemidarr.focus();
		return;
	}
	if (confirm(frmarr.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
		frmarr.itemid.value = frmarr.itemidarr.value;
		frmarr.mode.value="add";
		frmarr.submit();
	}
}

</script>
<table width="650" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frmarr" method="post" action="/admin/offshop/lib/domailzinemdchoice.asp">
	<input type="hidden" name="mode">
	<input type="hidden" name="itemid">
	<tr>
		<td class="a">
			���������� : <% DrawSelectBoxMailzine masteridx %>&nbsp;&nbsp;&nbsp;<input type="button" value="������ �߰�" onclick="popItemWindow('','frmarr.itemid')" class="button">
		</td>
	</tr>
	<tr>
		<td class="a">
			<table width=100% border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td><input type="text" name="itemidarr" value="" size="76" class="input"></td>
				<td width="100" align="right"><input type="button" value="������ ���� �߰�" onclick="AddIttems2()" class="button"></td>
			</tr>
			</table><br>(�������� �޸�(,)�� �־��ּ��� ex:41080,40780,40759,)
		</td>
	</tr>
	</form>
</table>

<%
'������ ����
Sub DrawSelectBoxMailzine(byval selectedId)
   dim tmp_str,query1
   %><select name="masteridx" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%>>����</option><%
   query1 = " select idx,regdate from [db_shop].[dbo].tbl_shopmaster_mail"
   query1 = query1 + " where isusing = 'Y'"
   query1 = query1 + " order by regdate desc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("idx")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("idx")&"' "&tmp_str&">"&FormatDate(rsget("regdate"),"0000.00.00")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->