<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ����ڻ���� �����ڵ�
' History : 2008�� 06�� 27�� �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<%
dim gubuntype, gubuncd, typename, gubunname, isusing, orderno ,mode ,ocodelist ,ocodeone ,i ,idx
dim page
	gubuntype  = requestCheckVar(Request("gubuntype"),10)
	gubuncd = requestCheckVar(Request("gubuncd"),2)
	idx = requestCheckVar(Request("idx"),10)
	page = request("page")
	if page="" then page=1

mode = "I"

set ocodelist = new cequipmentcode
	ocodelist.FPageSize = 20
	ocodelist.FCurrPage = page
	ocodelist.frectgubuntype = gubuntype
	ocodelist.getequipmentcodelist

set ocodeone = new cequipmentcode
	ocodeone.frectidx = idx

	if idx <> "" then
		ocodeone.getequipmentcodedetail

		if ocodeone.FTotalCount > 0 then
			idx = ocodeone.FOneItem.fidx
			gubuntype = ocodeone.FOneItem.fgubuntype
			gubuncd = ocodeone.FOneItem.fgubuncd
			typename = ocodeone.FOneItem.ftypename
			gubunname = ocodeone.FOneItem.fgubunname
			isusing = ocodeone.FOneItem.fisusing
			orderno = ocodeone.FOneItem.forderno

			mode = "U"
		else
			idx = ""
		end if
	end if

if orderno = "" then orderno = 0
if isusing = "" then isusing = "Y"
%>

<script language="javascript">

	// �ڵ�Ÿ�� �����̵�
	function jsSetCode(idx,fgubuntype){
		self.location.href = "/common/equipment/popmanagecode.asp?idx="+idx+"&gubuntype="+fgubuntype;
	}

	//�ڵ� �˻�
	function jsSearch(){
		document.frmSearch.submit();
	}

	//�ڵ� ���
	function jsRegCode(){
		var frm = document.frmReg;

		if(!frm.gubuntype.value) {
			alert("����Ÿ�� ������ �ּ���");
			frm.gubuntype.focus();
			return false;
		}

		if(!frm.gubuncd.value) {
			alert("���ڵ带 �Է��� �ּ���");
			frm.gubuncd.focus();
			return false;
		}

		if(!frm.gubunname.value) {
			alert("���ڵ���� �Է��� �ּ���");
			frm.gubunname.focus();
			return false;
		}

		if(!frm.orderno.value) {
			alert("���ļ����� �Է��� �ּ���");
			frm.orderno.focus();
			return false;
		}

		return true;
	}

</script>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//�ڵ� ��� �� ����-->
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<form name="frmReg" method="post" action="/common/equipment/popManageCodeprocess.asp" onSubmit="return jsRegCode();">
		<input type="hidden" name="mode" value="<%=mode%>">
		<tr>
			<td>	+ �ڵ� ��� �� ����</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<tr height="25">
					<td bgcolor="#EFEFEF" width=100 align="center">��ȣ</td>
					<td bgcolor="#FFFFFF">
						<%=idx%><input type="hidden" name="idx" value="<%=idx%>">
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width=100 align="center">����</td>
					<td bgcolor="#FFFFFF">
						<% drawequipmentCodeType "gubuntype" ,gubuntype, "" %>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" align="center">���ڵ�</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="2" maxlength="2" name="gubuncd" value="<%=gubuncd%>"> (ex : MO)
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" align="center">���ڵ��</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="32" maxlength="32" name="gubunname" value="<%=gubunname%>"> (ex : ��񱸺�)
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" align="center">���ļ���</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="4" maxlength="10" name="orderno" value="<%=orderno%>"> ���ڰ� �������� �켱����˴ϴ�.
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" align="center">��뿩��</td>
					<td bgcolor="#FFFFFF">
						<input type="radio" value="Y" name="isusing" <%IF isusing ="Y" THEN%> checked<%END IF%>>���
						<input type="radio" value="N" name="isusing" <%IF  isusing ="N" THEN%> checked<%END IF%>>������
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="right">
				<input type="image" src="/images/icon_save.gif">
			</td>
		</tr>
		<tr>
			<td colspan="2"><hr width="100%"></td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2">+ �ڵ� ����Ʈ</td>
</tr>
<form name="frmSearch" method="get">
<tr>
	<td>
		����Ÿ�� :
		<% drawequipmentCodeType "gubuntype" ,gubuntype, " onChange='jsSearch();'" %>
	</td>
	<td align="right"><a href="javascript:jsSetCode('','');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<tr bgcolor="#EFEFEF" align="center">
			<td>��ȣ</td>
			<td>����Ÿ�Ը�</td>
			<td>���ڵ�</td>
			<td>���ڵ��</td>
			<td>���ļ���</td>
			<td>��뿩��</td>
			<td>���</td>
		</tr>
		<% if ocodelist.fresultcount > 0 then %>
		<% for i = 0 to ocodelist.fresultcount - 1 %>
		<% if ocodelist.FItemList(i).fisusing = "Y" then %>
			<tr bgcolor="#ffffff" align="center">
		<% else %>
			<tr bgcolor="silver" align="center">
		<% end if %>
			<td><%=ocodelist.FItemList(i).fidx%></td>
			<td><%=ocodelist.FItemList(i).ftypename%> (<%=ocodelist.FItemList(i).fgubuntype%>)</td>
			<td><%=ocodelist.FItemList(i).fgubuncd%></td>
			<td><%=ocodelist.FItemList(i).fgubunname%></td>
			<td><%=ocodelist.FItemList(i).forderno%></td>
			<td><%=ocodelist.FItemList(i).fisusing%></td>
			<td>
				<input type="button" value="����" onClick="javascript:jsSetCode('<%=ocodelist.FItemList(i).fidx%>','<%=ocodelist.FItemList(i).fgubuntype%>');" class="input_b">
			</td>
		</tr>
		<% next %>

		<%ELSE%>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="10">��ϵ� ������ �����ϴ�.</td>
		</tr>
		<%End if%>

		</table>
	</td>
</tr>
</form>
</table>

<%
set ocodelist = nothing
set ocodeone = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
