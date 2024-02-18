<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2011.01.11 - 허진원 생성
'			2022.07.04 한용민 수정(isms보안취약점수정, 소스표준화)
' Discription : QR코드 등록
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
	//코드 등록
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.QRTitle.value) {
			alert("코드명을 입력해 주세요");
			frm.QRTitle.focus();
			return false;
		}
		if(!frm.QRDiv.value) {
			alert("코드구분을 선택해 주세요");
			frm.QRDiv.focus();
			return false;
		}
		if(!frm.QRContent.value) {
			alert("코드내용을 입력해 주세요");
			frm.QRContent.focus();
			return false;
		}

		if(confirm("입력한 내용이 정확합니까?")) {
			return true;
		}
		return false;
	}

	// 로그사용 여부 변경
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
	<td colspan="2"><!--//코드 등록 및 수정-->	
		<%'하단 staticImgUrl -> uploadUrl로 수정..2016-10-17 김진영  %>
		<form name="frmReg" method="post" action="<%=uploadUrl%>/linkweb/mobile/captureQRcode_proc.asp" onSubmit="return jsRegCode();" enctype="MULTIPART/FORM-DATA">
		<input type="hidden" name="qrSn" value="<%=qrSn%>">
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<tr>			
			<td><b>QR코드 등록 및 수정</b></td>
		</tr>	
		<tr>
			<td>	
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<% IF qrSn <> "" THEN%>	
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">코드번호</td>
					<td bgcolor="#FFFFFF"><%=qrSn%></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">QR코드</td>
					<td bgcolor="#FFFFFF"><img src="<%=QRImage%>"></td>
				</tr>
				<%END IF%>
			<% if countYN<>"N" then %>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">코드명</td>
					<td bgcolor="#FFFFFF"><input type="text" size="32" maxlength="64" name="QRTitle" value="<%=QRTitle%>" class="text"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">로그사용</td>
					<td bgcolor="#FFFFFF">
						<label><input type="radio" value="Y" name="countYN" onclick="chgCountYn(this.value)" onfocus="this.blur()" <%IF countYN="Y" THEN%>checked<%END IF%>>사용</label>
						<label><input type="radio" value="N" name="countYN" onclick="chgCountYn(this.value)" onfocus="this.blur()" <%IF countYN="N" or countYN="" THEN%>checked<%END IF%>>사용안함</label>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">QR구분</td>
					<td bgcolor="#FFFFFF"><% DrawSelectBoxQRDiv "QRDiv", QRDiv %></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">코드내용</td>
					<td bgcolor="#FFFFFF"><textarea name="QRContent" class="textarea" style="width:100%; height:40px;"><%=QRContent%></textarea></td>
				</tr>
				<% IF qrSn="" THEN%>	
				<tr>
					<td bgcolor="#EFEFEF" align="center">QR코드 오차율</td>
					<td bgcolor="#FFFFFF">
						<label><input type="radio" value="L" name="qrQuality" onfocus="this.blur()">7%</label>
						<label><input type="radio" value="M" name="qrQuality" onfocus="this.blur()" checked>15%(기본)</label>
						<label><input type="radio" value="Q" name="qrQuality" onfocus="this.blur()">25%</label>
						<label><input type="radio" value="H" name="qrQuality" onfocus="this.blur()">30%</label>
					</td>
				</tr>
				<% End IF %>
			<% else %>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">코드명</td>
					<td bgcolor="#FFFFFF"><%=QRTitle%></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">URL</td>
					<td bgcolor="#FFFFFF"><%=QRContent%></td>
				</tr>
			<% end if %>
				<tr>
					<td bgcolor="#EFEFEF" align="center">사용여부</td>
					<td bgcolor="#FFFFFF">
						<label><input type="radio" value="Y" name="isUsing" onfocus="this.blur()" <%IF isUsing="Y" or isUsing="" THEN%>checked<%END IF%>>사용</label>
						<label><input type="radio" value="N" name="isUsing" onfocus="this.blur()" <%IF isUsing="N" THEN%>checked<%END IF%>>사용안함</label>
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