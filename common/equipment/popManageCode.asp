<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 장비자산관리 공통코드
' History : 2008년 06월 27일 한용민 생성
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

	// 코드타입 변경이동
	function jsSetCode(idx,fgubuntype){
		self.location.href = "/common/equipment/popmanagecode.asp?idx="+idx+"&gubuntype="+fgubuntype;
	}

	//코드 검색
	function jsSearch(){
		document.frmSearch.submit();
	}

	//코드 등록
	function jsRegCode(){
		var frm = document.frmReg;

		if(!frm.gubuntype.value) {
			alert("구분타입 선택해 주세요");
			frm.gubuntype.focus();
			return false;
		}

		if(!frm.gubuncd.value) {
			alert("상세코드를 입력해 주세요");
			frm.gubuncd.focus();
			return false;
		}

		if(!frm.gubunname.value) {
			alert("상세코드명을 입력해 주세요");
			frm.gubunname.focus();
			return false;
		}

		if(!frm.orderno.value) {
			alert("정렬순서를 입력해 주세요");
			frm.orderno.focus();
			return false;
		}

		return true;
	}

</script>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//코드 등록 및 수정-->
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<form name="frmReg" method="post" action="/common/equipment/popManageCodeprocess.asp" onSubmit="return jsRegCode();">
		<input type="hidden" name="mode" value="<%=mode%>">
		<tr>
			<td>	+ 코드 등록 및 수정</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<tr height="25">
					<td bgcolor="#EFEFEF" width=100 align="center">번호</td>
					<td bgcolor="#FFFFFF">
						<%=idx%><input type="hidden" name="idx" value="<%=idx%>">
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width=100 align="center">구분</td>
					<td bgcolor="#FFFFFF">
						<% drawequipmentCodeType "gubuntype" ,gubuntype, "" %>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" align="center">상세코드</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="2" maxlength="2" name="gubuncd" value="<%=gubuncd%>"> (ex : MO)
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" align="center">상세코드명</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="32" maxlength="32" name="gubunname" value="<%=gubunname%>"> (ex : 장비구분)
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" align="center">정렬순서</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="4" maxlength="10" name="orderno" value="<%=orderno%>"> 숫자가 작을수록 우선노출됩니다.
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" align="center">사용여부</td>
					<td bgcolor="#FFFFFF">
						<input type="radio" value="Y" name="isusing" <%IF isusing ="Y" THEN%> checked<%END IF%>>사용
						<input type="radio" value="N" name="isusing" <%IF  isusing ="N" THEN%> checked<%END IF%>>사용안함
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
	<td colspan="2">+ 코드 리스트</td>
</tr>
<form name="frmSearch" method="get">
<tr>
	<td>
		구분타입 :
		<% drawequipmentCodeType "gubuntype" ,gubuntype, " onChange='jsSearch();'" %>
	</td>
	<td align="right"><a href="javascript:jsSetCode('','');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<tr bgcolor="#EFEFEF" align="center">
			<td>번호</td>
			<td>구분타입명</td>
			<td>상세코드</td>
			<td>상세코드명</td>
			<td>정렬순서</td>
			<td>사용여부</td>
			<td>비고</td>
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
				<input type="button" value="수정" onClick="javascript:jsSetCode('<%=ocodelist.FItemList(i).fidx%>','<%=ocodelist.FItemList(i).fgubuntype%>');" class="input_b">
			</td>
		</tr>
		<% next %>

		<%ELSE%>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="10">등록된 내용이 없습니다.</td>
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
