<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출
' History : 2011.12.27 한용민 생성
'						2014.01.03 정윤정 수정 검색조건 및 필드 추가
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/delaytaxcls.asp"-->
<%
dim i, j
dim yyyy1, mm1, yyyy2, mm2,yyyy3,mm3
dim yyyymm, endyyyymm, issueyyymm, makerid ,offgubun, issuegubun
dim designer,groupid ,vPurchaseType,erpCustcd, jgubun, companynoYN
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	yyyy3 = requestCheckVar(request("yyyy3"),4)
	mm3 = requestCheckVar(request("mm3"),2)
	offgubun = requestCheckVar(request("offgubun"),3)
	issuegubun = requestCheckVar(request("issuegubun"),1)
	designer = requestCheckVar(request("designer"),32)
	groupid  = requestCheckVar(request("groupid"),32)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	erpCustcd = requestCheckVar(request("erpCustcd"),16)
	jgubun   = requestCheckVar(request("jgubun"),10)
	companynoYN = requestCheckVar(request("companynoYN"),1)
Dim jacctcdexists : jacctcdexists =requestCheckVar(request("jacctcdexists"),10)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 =  format00(2,cstr(Month(now())))
if (yyyy2="") then yyyy2 = yyyy1
if (mm2="") then mm2 = mm1

yyyymm = yyyy1 + "-" + mm1
endyyyymm = yyyy2 + "-" + mm2
if yyyy3 <> "" then
issueyyymm = yyyy3 + "-" + mm3
end if

If offgubun <> "ON" Then
'	companynoYN = ""
End If

dim ocdelaytax
set ocdelaytax = new CDelayTax
	ocdelaytax.FRectStartYYYYMM = yyyymm
	ocdelaytax.FRectEndYYYYMM = endyyyymm
	ocdelaytax.FRectIssueYYYYMM = issueyyymm
	ocdelaytax.FRectGubun = offgubun
	ocdelaytax.FRectIssueGubun = issuegubun
	ocdelaytax.FRectdesigner = designer
	ocdelaytax.FRectGroupid = groupid
	ocdelaytax.FRectPurchaseType = vPurchaseType
	ocdelaytax.FRecterpCustcd = erpCustcd
    ocdelaytax.FRectJGubun = jgubun
    ocdelaytax.FRectCompanynoYN = companynoYN
	ocdelaytax.FRectJacctcdExists = jacctcdexists
	ocdelaytax.GetDelayTaxDetailList

%>

<script type="text/javascript">

function formSubmit(page) {
	frm.page.value=page;
	frm.submit();
}
function popSearchGroupID(frmname,compname){
    var popwin = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&compname=" + compname,"popSearchGroupID","width=800 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		구분 :
		<select class="select" name="offgubun">
		<option value="ON" <% if (offgubun = "ON") then %>selected<% end if %> >온라인</option>
		<option value="OFF" <% if (offgubun = "OFF") then %>selected<% end if %> >오프라인</option>
		<option value="ETC" <% if (offgubun = "ETC") then %>selected<% end if %> >기타매출</option>
		</select>
		&nbsp;
		정산월 :  <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
		&nbsp;
		발행월 : <% Call DrawYMBoxdynamic("yyyy3", yyyy3, "mm3", mm3, "") %>
		&nbsp;
		발행구분 :
		<select class="select" name="issuegubun">
		<option value="1" <% if (issuegubun = "1") then %>selected<% end if %> >정상발행</option>
		<option value="2" <% if (issuegubun = "2") then %>selected<% end if %> >발행이전</option>
		<option value="9" <% if (issuegubun = "9") then %>selected<% end if %> >기타발행(선발행)</option>
		</select>
		&nbsp;
		정산방식구분 :
        <% drawSelectBoxJGubun "jgubun",jgubun %>
		&nbsp;&nbsp;
		* 텐바이텐 사업자 여부 : 
        <select name="companynoYN" class="select">
			<option value="">전체
			<option value="Y" <%= CHKIIF(companynoYN="Y","selected","") %> >사업자만
			<option value="N" <%= CHKIIF(companynoYN="N","selected","") %> >사업자제외
		</select>
		&nbsp;&nbsp;
		<input type="checkbox" name="jacctcdexists" <%= CHKIIF(jacctcdexists="on","checked","") %> >계정과목 존재 정산만 보기
        

	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="formSubmit('1');">
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" >
		구매유형 : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
		&nbsp;&nbsp;
		브랜드ID : <% drawSelectBoxDesignerwithName "designer",designer  %>&nbsp;&nbsp;
	  업체(그룹코드) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >
		<input type="button" class="button" value="Code검색" onclick="popSearchGroupID(this.form.name,'groupid');" >&nbsp;&nbsp;
		ERP연계코드 : <input type="text" class="input" name="erpCustcd" value="<%=erpCustcd%>" size="16">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* 최대 3,000건까지 표시됩니다.
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		검색결과 : <b><%= formatnumber(ocdelaytax.FTotalCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">정산월</td>
	<td width="150">브랜드</td>
	<td width="150">사업자명</td>
	<td>정산담당자<br/>휴대폰</td>
	<td width="60">ERP연계코드</td>
	<td width="60">그룹코드</td>
	<td width="60">정산의<br/>과세구분</td>
	<td width="60">과세</td>
	<td width="170">이세로</td>
	<td width="80">등록일</td>
	<td width="80">발행일</td>
	<td width="80">입금일</td>
	<td width="80">정산액<br>(발행액)</td>
	<td width="80">구매유형</td>
	<td>사업부문</td>
	<td>매출계정</td>
	<td>상태</td>
	<td>비고</td>
</tr>
<%
if ocdelaytax.FresultCount > 0 then
%>
	<%
	for i=0 to ocdelaytax.FresultCount-1
	%>
		<tr bgcolor="#FFFFFF" align="center">
			<td nowrap><%= ocdelaytax.FItemList(i).Fyyyymm %></td>
			<td><%= ocdelaytax.FItemList(i).Fmakerid %></td>
			<td><%= ocdelaytax.FItemList(i).Fcompany_name %></td>
			<td><%= ocdelaytax.FItemList(i).fjungsan_hp %></td>
			<td><%= ocdelaytax.FItemList(i).Ferpcust_cd %></td>
			<td><%= ocdelaytax.FItemList(i).Fgroupid %></td>
			<td><%= ocdelaytax.FItemList(i).getTaxTypeName %></td>
			<td><%= ocdelaytax.FItemList(i).Fjungsan_gubun %></td>
			<td><%= ocdelaytax.FItemList(i).Feserotaxkey %></td>
			<td  nowrap><%= ocdelaytax.FItemList(i).Ftaxinputdate %></td>
			<td  nowrap><%= ocdelaytax.FItemList(i).Ftaxregdate %></td>
			<td  nowrap><%= ocdelaytax.FItemList(i).Fipkumdate %></td>
		<td  nowrap align="right"><%= FormatNumber(ocdelaytax.FItemList(i).FjungsanPrice,0)  %></td>
			<td><%= ocdelaytax.FItemList(i).FpurchasetypeName %></td>
			<td><%= ocdelaytax.FItemList(i).FbizsectionName %></td>
			<td><%= ocdelaytax.FItemList(i).FselltypeName %></td>
			<td><%= ocdelaytax.FItemList(i).GetFinishFlagName %></td>
			<td></td>
		</tr>
	<% next %>
		<tr height="25" bgcolor="#ffffff">
			<td colspan="12" align="center">합계</td>
			<td align="right"><%=FormatNumber(ocdelaytax.FTot_jungsanPrice,0)%></td>
			<td colspan="5"></td>
		</tr>
<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="18">검색결과가 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set ocdelaytax = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp"-->