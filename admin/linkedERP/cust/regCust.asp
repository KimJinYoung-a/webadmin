<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 거래처 정보
' History : 2011.04.21 정윤정 생성
'			2016.07.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/custCls.asp"-->
<%
Dim clsC, arrBank, intB,sMode, sEmail, sDispSeq, sPostCd, sAddr,sCORPYN,sBRNType
Dim sCustCd, sCustNm, sARYN, sAPYN, sCoYN, sBizNo, sCeoNm, sTaxType, sBSCD, sINTP, sTelNo, sFaxNo
Dim sEmpno,sEmpNm, sPos,sDeptNM, sEmpTel, sEmpHP, sEmpEmail,sPSGB
Dim arrBankNm, arrBankCd, arrSavMN,arrAcctNo, intLoop, intDef
Dim sBank_cd,sacc_no, ssav_mn, arrEmp,arrAcct
Dim sCustNm7,sBizNo7,sPos7,sDeptNm7,sEmpTel7,sEmpHP7,sEmpEmail7,sDispSeq7,sEmpNm7
dim srectEmpno,srectBankno,srectAcctno
Dim isReadOnly : isReadOnly = (requestCheckvar(Request("rO"),10)<>"")
	sCustCd = requestCheckvar(Request("hidCcd"),13)
	srectEmpno= requestCheckvar(Request("hidEno"),10)
	srectBankno= requestCheckvar(Request("hidBno"),8)
	srectAcctno= requestCheckvar(Request("hidAno"),30)
sPSGB = "1"
sMode ="I"

set clsC = new CCust

	IF sCustCd <> "" THEN
	 	sMode = "U"
		clsC.FCustCd = sCustCd
		clsC.FRectempno = srectEmpno
		clsC.FRectBankNo = srectBankno
		clsC.FRectAcctNo = srectAcctno
		clsC.fnGetCustData
		sCORPYN =clsC.FCORPYN
		sBRNType =clsC.FCUSTBRNTYPE
		sPSGB = clsC.FPSGB
		IF isNull(sCORPYN) or sCORPYN ="" then sCORPYN ="Y"
		IF (sCORPYN="Y") then sPSGB="1" ''법인이면 무조건 사업자. // ERP에서 수정시 법인 이면 빈값으로 들어감..

		IF sPSGB = "2" THEN	'개인일때
	  		sCustNm7	= clsC.FCustNM
			sBizNo7		= clsC.FBizNo
			sPos7		= clsC.FPos
			sDeptNm7	= clsC.FDEPT_NM
			sEmpTel7	= clsC.FTELNO
			sEmpHP7		= clsC.FHP_NO
			sEmpEmail7	= clsC.FEMAIL
			sDispSeq7	= clsC.FDispSeq
			sEmpNm7 	= clsC.FEMP_NM
		ELSE
			sCustNM 	= clsC.FCustNM
			sBizNo  	= clsC.FBizNo
			sDispSeq	= clsC.FDispSeq
			sPos 		= clsC.FPos
			sDEPTNM 	= clsC.FDEPT_NM
			sEmpTel 	= clsC.FSTelNo
			sEmpHP 		= clsC.FHP_NO
			sEmpEmail 	= clsC.FSEmail
			sEmpNm 		= clsC.FEMP_NM
		END IF

		sARYN		= clsC.FARYN
		sAPYN	 	= clsC.FAPYN
		sCeoNM  	= clsC.FCeoNM
		sEMAIL  	= clsC.FEMAIL
		sTELNO  	= clsC.FTELNO
		sFAXNO  	= clsC.FFAXNO
		sTAXTYPE	= clsC.FTAXTYPE
		sBSCD   	= clsC.FBSCD
		sINTP   	= clsC.FINTP
		sPostCD 	= clsC.FPostCD
		sADDR   	= clsC.FADDR
		sEmpNo 		= clsC.FEMP_NO

		sBank_cd	= clsC.FBank_cd
		sacc_no 	= clsC.Facct_no
		ssav_mn		= clsC.Fsav_mn

		IF isNull(sBank_cd) THEN sBank_cd = ""
		'arrEmp	=clsC.fnGetCustSaleorList	-리스트 처리 여부 추후선택..현재 1명씩만 보여주도록
		'arrAcct =clsC.fnGetCustAcctList
	END IF

	arrBank = clsC.fnGetBankList
set clsC = nothing

IF sDispSeq = "" THEN sDispSeq = "10000"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script type="text/javascript">

//우편번호 불러오기
function jsSetPC(){
	var winPC = window.open("/lib/searchzip3.asp?target=frmC","PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	winPC.focus();
}

//우편번호 set
	function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".sPCd").value = post1 + post2;

    eval(frmname + ".sAddr").value = addr+" "+dong;
}


function jsSubmitCust(){
	if(document.frmC.rdoRT[1].checked){
		if(jsChkBlank(document.frmC.scnm7.value)){
 		alert("거래처명을  입력해주세요");
 		document.frmC.scnm7.focus();
 		return;
 		}

 		if(jsChkBlank(document.frmC.sem7.value)){
 		alert("담당자명을  입력해주세요");
 		document.frmC.sem7.focus();
 		return;
 		}

 		if(jsChkBlank(document.frmC.sBno71.value)){
 		alert("주민번호을  입력해주세요");
 		document.frmC.sBno71.focus();
 		return;
 		}

 			if(jsChkBlank(document.frmC.sBno72.value)){
 		alert("주민번호을  입력해주세요");
 		document.frmC.sBno72.focus();
 		return;
 		}
	}else{
	if(jsChkBlank(document.frmC.scnm.value)){
 		alert("거래처명을  입력해주세요");
 		document.frmC.scnm.focus();
 		return;
 		}

 		if(jsChkBlank(document.frmC.sBno.value)){
 		alert("사업자번호을  입력해주세요");
 		document.frmC.sBno.focus();
 		return;
 		}
 	}

    if (document.frmC.selBC.value=="10000005"){
        alert("외환은행 선택 불가 : KEB 하나은행으로 변경되었습니다.");
 		document.frmC.sBno.focus();
 		return;
    }

    if (confirm('등록하시겠습니까?')){
    	document.frmC.submit();
    }
}

//등록폼 화면 선택
 function jsSetRegFrm(iType){
 	if(iType=="2"){
 		document.all.dType1.style.display = "none";
 		document.all.dType2.style.display = "";
 	}else{
 		document.all.dType1.style.display = "";
 		document.all.dType2.style.display = "none";
 	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
	<tr>
	<td><strong>거래처 신규등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<form name="frmC" method="post" action="procCust.asp">
			<input type="hidden" name="ver" value="">
			<input type="hidden" name="hidM" value="<%=sMode%>">
			<input type="hidden" name="hidCcd" value="<%=sCustCd%>">
			<input type="hidden" name="hidENo" value="<%=sEmpno%>">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">등록타입 </td>
				<td bgcolor="#FFFFFF"> <input type="radio" name="rdoRT" value="1" <%IF sPSGB="1" THEN%>checked<%END IF%> onClick="jsSetRegFrm(1);">사업자 | <input type="radio" name="rdoRT" value="2" <%IF sPSGB="2" THEN%>checked<%END IF%> onClick="jsSetRegFrm(2);">개인</td></td>
			</tr>
		</table>
</tr>
<tr>
	<td>
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0" >
			<tr>
				<td>
					<div id="dType1" style="display:<%IF sPSGB="2" THEN%>none<%END IF%>;"><!-- 거래처등록 폼-->
					<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" border="0" >
					<tr>
				 		<td>기본정보</td>
					</tr>
					<tr>
						<td>
							<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"  bgcolor="<%= adminColor("tablebg") %>">
						 <tr>
								<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">거래처명</td>
								<td colspan="3"  bgcolor="#FFFFFF"  ><input type="text" name="scnm" value="<%=sCustNm%>" size="20" class="text">(공백없이)</td>
							</tr>
							<tr>
								<td   bgcolor="<%= adminColor("tabletop") %>" align="center">거래처분류</td>
								<td    bgcolor="#FFFFFF">
									<select name="selBRNT">
									<%sbOptCustType sBRNType%>
									</select>
								</td>
								<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">거래처구분</td>
								<td   bgcolor="#FFFFFF">
									<input type="checkbox" name="chkAR" value="Y" <%IF sARYN="Y" THEN%>checked<%END IF%>>매출
									<input type="checkbox" name="chkAP" value="Y" <%IF sAPYN="Y" THEN%>checked<%END IF%>>매입
								</td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">사업자구분</td>
								<td  bgcolor="#FFFFFF"><input type="radio" name="rdoCo" value="Y" <%IF sCORPYN ="Y" or sCORPYN="" THEN%>checked<%END IF%>>법인
									<input type="radio" name="rdoCo" value="N" <%IF sCORPYN ="N" THEN%>checked<%END IF%>>개인</td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">사업자번호</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sBno" value="<%=sBizNo%>" size="15" maxlength="13"  <%= CHKIIF(sCORPYN="Y" and Len(replace(sBizNo,"-",""))=10,"readonly class='text_ro'","class='text'") %> > (-없이 숫자만)</td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">대표자</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sceonm" value="<%=sCeoNm%>" size="10" class="text"></td>
								<td  bgcolor="<%= adminColor("tabletop") %>" align="center">과세유형/휴폐업</td>
								<td  bgcolor="#FFFFFF">
									<select name="selTType">
										<option value="1" <%IF sTaxType="1" or  sTaxType="" THEN%>selected<%END IF%>>일반과세자</option>
										<option value="2" <%IF sTaxType="2" THEN%>selected<%END IF%>>간이과세자</option>
										<option value="3" <%IF sTaxType="3" THEN%>selected<%END IF%>>면세과세자</option>
										<option value="4" <%IF sTaxType="4" THEN%>selected<%END IF%>>비영리과세자</option>
										<option value="5" <%IF sTaxType="5" THEN%>selected<%END IF%>>기타사업자</option>
										<option value="9" <%IF sTaxType="9" THEN%>selected<%END IF%>>휴업</option>
										<option value="10" <%IF sTaxType="10" THEN%>selected<%END IF%>>폐업</option>
									</select>
								</td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">업태</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sBS" value="<%=sBSCD%>" size="20" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">종목</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sIN" value="<%=sINTP%>" size="20" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">전화번호</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sTNo" value="<%=sTelNo%>" size="15" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">팩스번호</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sFNo" value="<%=sFaxNo%>" size="15" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">이메일</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sE" value="<%=sEmail%>" size="30" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">출력순서</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sDS" value="<%=sDispSeq%>" size="5" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">주소</td>
								<td bgcolor="#FFFFFF" colspan="3">
									<input type="text" name="sPCd" value="<%=sPostCd%>" size="6" class="text_ro">
									<input type="button" class="button" value="검색" onClick="FnFindZipNew('frmC','G')">
									<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frmC','G')">
									<% '<input type="button" class="button" value="검색(구)"  onClick="jsSetPC();"> %>
									<input type="text" name="sAddr" value="<%=sAddr%>" size="60" class="text">
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td style="padding-top:10px;">담당자정보</td>
				</tR>
				<tr>
					<td>
						<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"  bgcolor="<%= adminColor("tablebg") %>">
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">성명</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sENm" value="<%=sEmpNm%>" size="10" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">직위</td>
								<td  bgcolor="#FFFFFF" width="300"><input type="text" name="sEP" value="<%=sPos%>" size="10" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">부서</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sDNm" value="<%=sDeptNm%>" size="20" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">전화</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sETN" value="<%=sEmpTel%>" size="15" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">휴대폰</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sEHp" value="<%=sEmpHP%>" size="15" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">이메일</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sEE" value="<%=sEmpEmail%>" size="30" class="text"></td>
							</tr>
						</table>
					</td>
			</tr>
		</table>
		</div>
		<div id="dType2" style="display:<%IF sPSGB="1" THEN%>none<%END IF%>;"><!-- 직원등록 폼-->
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" border="0" >
			<tr>
			<td>기본정보</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"  bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">거래처명</td>
					<td  bgcolor="#FFFFFF"  ><input type="text" name="scnm7" value="<%=sCustNm7%>" size="20" class="text"></td>
					<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">담당자명</td>
					<td  bgcolor="#FFFFFF"  ><input type="text" name="sem7" value="<%=sEmpNm7%>" size="20" class="text"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">거래처분류</td>
					<td bgcolor="#FFFFFF">
						<select name="selBRNT7">
						<%sbOptCustType sBRNType%>
						</select>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="100">주민번호</td>
					<td  bgcolor="#FFFFFF">
						<input type="text" name="sBno71" value="<%IF sBizNo7 <> "" THEN%><%=left(sBizNo7,6)%><%END IF%>" size="10" maxlength="6" class="text">-
						<input type="password" name="sBno72" value="<%IF sBizNo7 <> "" THEN%><%=mid(sBizNo7,7,7)%><%END IF%>" size="10" maxlength="7" class="text">
						</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">직위</td>
					<td  bgcolor="#FFFFFF" width="300"><input type="text" name="sEP7" value="<%=sPos7%>" size="10" class="text"></td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">부서</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sDNm7" value="<%=sDeptNm7%>" size="20" class="text"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">전화</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sTNo7" value="<%=sEmpTel7%>" size="15" class="text"></td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">휴대폰</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sEHp7" value="<%=sEmpHP7%>" size="15" class="text"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">이메일</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sE7" value="<%=sEmpEmail7%>" size="30" class="text"></td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">출력순서</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sDS7" value="<%=sDispSeq7%>" size="5" class="text"></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
			</div>
		</td>
	</tr>
	<tr>
		<td style="padding-top:10px;">계좌정보 </td>
	</tr>
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"  bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td  bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">은행명</td>
					<td  bgcolor="#FFFFFF">
						<select name="selBC">
							<option value="">--선택--</option>
							<%IF isArray(arrBank) THEN
								For intB = 0 To UBound(arrBank,2)
							%>
							<option value="<%=arrBank(0,intB)%>" <%IF Cstr(sBank_cd) = Cstr(arrBank(0,intB)) THEN%>selected<%END IF%>><%=arrBank(1,intB)%></option>
							<%
								Next
								END IF%>
						</select>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="80"> 계좌번호:</td>
					<td   bgcolor="#FFFFFF"> <input type="text" name="sAN" value="<%=sAcc_no%>" size="20" class="text"></td>
					<td  bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">예금주</td>
					<Td  bgcolor="#FFFFFF"><input type="text" name="sSN" value="<%=sSav_mn%>" size="10" class="text"></td>
				</tr>
			</table>
		</td>
		</tr>
		<tr>
			<td align="center" colspan="3">
			<% if Not isReadOnly then %>
				<input type="button" class="button" value="등록" onClick="jsSubmitCust();">
			<% end if%>
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
</table>
</body>
</html>
