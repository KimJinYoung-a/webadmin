

<table width="100%" class="a">
<tr>
	<td valign="top">
		������ǰ ��� :
		<input class="button" type="button" id="btnRegSel" value="��ǰ���" onClick="ebaySelectRegProcess('<%= gubun %>');">&nbsp;&nbsp;
		<br><br>
		������ǰ ���� :

	</td>
	<td align="right" valign="top">
		<br><br>
		���û�ǰ��
		<Select name="chgSellYn" class="select">
			<option value="N">�Ǹ�����</option>
			<option value="Y">�Ǹ�</option>
		</Select>(��)��
		<input class="button" type="button" id="btnSellYn" value="����" onClick="AuctionSellYnProcess(frmReg.chgSellYn.value);">
	</td>
</tr>
</table>