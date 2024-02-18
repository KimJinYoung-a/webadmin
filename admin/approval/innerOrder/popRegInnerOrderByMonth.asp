<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : ������û�� ���
' History : 2011.03.14 ������  ����
' 0 ��û/1 ������/ 5 �ݷ�/7 ����/ 9 �Ϸ�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim yyyy1, mm1, tmpdate

yyyy1 = requestCheckvar(Request("yyyy1"),32)
mm1 = requestCheckvar(Request("mm1"),32)

if yyyy1="" then
	tmpdate = CStr(Now)

	tmpdate = DateAdd("m", -1, tmpdate)

	yyyy1 = Left(tmpdate, 4)
	mm1 = Mid(tmpdate, 6, 2)
end if

%>
<script language="javascript">

function jsRegInsertShopChulgo(frm) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?\n\n������ �ð��� �ҿ�˴ϴ�.(5~10��)") == true) {
		frm.mode.value = "reginsertshopchulgo";
		frm.submit();
	}
}

function jsRegInsertUpcheShopMaeip(frm) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?\n\n������ �ð��� �ҿ�˴ϴ�.(5~10��)") == true) {
		frm.mode.value = "reginsertupcheshopmaeip";
		frm.submit();
	}
}

function jsRegInsertUpcheShopWitak(frm) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?\n\n������ �ð��� �ҿ�˴ϴ�.(5~10��)") == true) {
		frm.mode.value = "reginsertupcheshopwitak";
		frm.submit();
	}
}

function jsRegInsertShopWitakSell(frm) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?\n\n������ �ð��� �ҿ�˴ϴ�.(5~10��)") == true) {
		frm.mode.value = "reginsertshopwitaksell";
		frm.submit();
	}
}

function jsRegInsertPartToOnline(frm) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reginsertparttoonline";
		frm.submit();
	}
}

function jsRegInsertPartToOffline(frm) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reginsertparttooffline";
		frm.submit();
	}
}

function jsRegInsertAll(frm, target) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reginsertall";
		frm.target.value = target;
		frm.submit();
	}
}

</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<form name="frm" method="post" action="popRegInnerOrderByMonth_process.asp">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="target" value="">
		<tr>
			<td width="100%">
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30><b>��/����/���κμ��� ���ΰŷ� �ϰ�����</b></td>
				</tr>
				<tr>
					<td>
						�ŷ��� : <% Call DrawYMBox(yyyy1, mm1) %>
						&nbsp;
						(���κμ� = ������ or ���̶�� ��)
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="100%">
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 01. �¶����Ǹ�(���̶��)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '01');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 02. �¶��θ���(���̶��)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '02');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 03. ������(ON��ǰ)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '03');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 04. ��Ÿ����(ON��ǰ)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '04');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 05. ������(OFF��ǰ)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '05');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 06. ��Ÿ����(OFF��ǰ)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '06');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 07. ������(��Ź��ǰ)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '07');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 08. �������
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '08');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 09. ��ü��Ź
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '09');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 10. ��Ÿ����
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '10');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 11. ������(��һ�ǰ)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '11');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 12. ��Ÿ����(��һ�ǰ)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '12');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 13. �����Ǹ�(��һ�ǰ)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '13');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 14. ��Ÿ�Ǹ�(��һ�ǰ)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="����" onClick="jsRegInsertAll(frm, '14');"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="100%">
				<br>* �� ���ΰŷ� �̿��� ���ΰŷ��� "<font color=red>�������м�>>�������μ��ͼ��Ӹ�</font>" ���� �ϰ������˴ϴ�.
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
