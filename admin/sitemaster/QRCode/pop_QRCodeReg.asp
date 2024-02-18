<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2011.01.11 - ������ ����
'			2022.07.04 �ѿ�� ����(isms�������������, �ҽ�ǥ��ȭ)
' Discription : QR�ڵ� ���
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/qrCodeCls.asp"-->
<%
	Dim qrSn, QRDiv, QRTitle, QRImage, countYN, QRContent, isUsing

	qrSn		= requestCheckVar(getNumeric(request("qrSn")),10)

	if qrSn<>"" then
		dim oQR
		set oQR = New CQRCode
		oQR.FCurrPage = 1
		oQR.FPageSize=1
		oQR.FRectQRSn = qrSn
		oQR.GetQRCode
	
		if oQR.FResultCount>0 then
			QRDiv		= oQR.FItemList(0).FQRDiv
			QRTitle		= ReplaceBracket(oQR.FItemList(0).FqrTitle)
			QRImage		= oQR.FItemList(0).FqrImage
			countYN		= oQR.FItemList(0).FcountYn
			QRContent	= ReplaceBracket(oQR.FItemList(0).FqrContent)
			isUsing		= oQR.FItemList(0).FisUsing
		end if
	
		set oQR = Nothing
	end if
%>
<script type="text/javascript">
<!--
	//�ڵ� ���
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.QRTitle.value) {
			alert("�ڵ���� �Է��� �ּ���");
			frm.QRTitle.focus();
			return false;
		}
		if(!frm.QRDiv.value) {
			alert("�ڵ屸���� ������ �ּ���");
			frm.QRDiv.focus();
			return false;
		}
		if(!frm.QRContent.value) {
			alert("�ڵ峻���� �Է��� �ּ���");
			frm.QRContent.focus();
			return false;
		}

		if(confirm("�Է��� ������ ��Ȯ�մϱ�?")) {
			return true;
		}
		return false;
	}

	// �α׻�� ���� ����
	function chgCountYn(sw) {
		if(sw=="Y") {
			frmReg.QRDiv.disabled=false;
			frmReg.QRDiv.value="";
		} else {
			frmReg.QRDiv.disabled=true;
			frmReg.QRDiv.value="1";
		}
	}

	<% if qrSn="" then %>
	window.onload = (event) => {
		chgCountYn('N');
	};
	<% end if %>
//-->
</script>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//�ڵ� ��� �� ����-->	
		<%'�ϴ� staticImgUrl -> uploadUrl�� ����..2016-10-17 ������  %>
		<form name="frmReg" method="post" action="<%=uploadUrl%>/linkweb/mobile/captureQRcode_proc.asp" onSubmit="return jsRegCode();" enctype="MULTIPART/FORM-DATA">
		<input type="hidden" name="qrSn" value="<%=qrSn%>">
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<tr>			
			<td><b>QR�ڵ� ��� �� ����</b></td>
		</tr>	
		<tr>
			<td>	
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<% IF qrSn <> "" THEN%>	
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">�ڵ��ȣ</td>
					<td bgcolor="#FFFFFF"><%=qrSn%></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">QR�ڵ�</td>
					<td bgcolor="#FFFFFF"><img src="<%=QRImage%>"></td>
				</tr>
				<%END IF%>
			<% if countYN<>"N" then %>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">�ڵ��</td>
					<td bgcolor="#FFFFFF"><input type="text" size="32" maxlength="64" name="QRTitle" value="<%=QRTitle%>" class="text"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">�α׻��</td>
					<td bgcolor="#FFFFFF">
						<label><input type="radio" value="Y" name="countYN" onclick="chgCountYn(this.value)" onfocus="this.blur()" <%IF countYN="Y" THEN%>checked<%END IF%>>���</label>
						<label><input type="radio" value="N" name="countYN" onclick="chgCountYn(this.value)" onfocus="this.blur()" <%IF countYN="N" or countYN="" THEN%>checked<%END IF%>>������</label>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">QR����</td>
					<td bgcolor="#FFFFFF"><% DrawSelectBoxQRDiv "QRDiv", QRDiv %></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">�ڵ峻��</td>
					<td bgcolor="#FFFFFF"><textarea name="QRContent" class="textarea" style="width:100%; height:40px;"><%=QRContent%></textarea></td>
				</tr>
				<% IF qrSn="" THEN%>	
				<tr>
					<td bgcolor="#EFEFEF" align="center">QR�ڵ� ������</td>
					<td bgcolor="#FFFFFF">
						<label><input type="radio" value="L" name="qrQuality" onfocus="this.blur()">7%</label>
						<label><input type="radio" value="M" name="qrQuality" onfocus="this.blur()" checked>15%(�⺻)</label>
						<label><input type="radio" value="Q" name="qrQuality" onfocus="this.blur()">25%</label>
						<label><input type="radio" value="H" name="qrQuality" onfocus="this.blur()">30%</label>
					</td>
				</tr>
				<% End IF %>
			<% else %>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">�ڵ��</td>
					<td bgcolor="#FFFFFF"><%=QRTitle%></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">URL</td>
					<td bgcolor="#FFFFFF"><%=QRContent%></td>
				</tr>
			<% end if %>
				<tr>
					<td bgcolor="#EFEFEF" align="center">��뿩��</td>
					<td bgcolor="#FFFFFF">
						<label><input type="radio" value="Y" name="isUsing" onfocus="this.blur()" <%IF isUsing="Y" or isUsing="" THEN%>checked<%END IF%>>���</label>
						<label><input type="radio" value="N" name="isUsing" onfocus="this.blur()" <%IF isUsing="N" THEN%>checked<%END IF%>>������</label>
					</td>
				</tr>
				</table>		
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td align="left"><a href="javascript:self.close()"><img src="/images/icon_cancel.gif" border="0"></a></td>
					<td align="right"><input type="image" src="/images/icon_save.gif"></td>
				</tr>
				</table>
			</td>
		</tr>	
		</table>
		</form>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->