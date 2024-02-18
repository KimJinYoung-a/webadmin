<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 거래처 정보
' History : 2011.04.21 정윤정 생성
'			2019.05.16 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/custCls.asp"-->
<%
Dim clsC
Dim arrList, intLoop
Dim iCUSTgbn,sCUSTtype,iSearchType,sSearchText,iPageSize,iCurrPage
Dim  iTotCnt,iTotalPage
Dim opnType
Dim allacct

iCUSTgbn 		= requestCheckvar(Request("rdoCgbn"),4)
sCUSTtype 	= requestCheckvar(Request("selCT"),3)
iSearchType = requestCheckvar(Request("selSTp"),4)
sSearchText = requestCheckvar(Request("sSTx"),100)
opnType     = requestCheckvar(Request("opnType"),10)
allacct     = requestCheckvar(Request("allacct"),10)

if (iSearchType="5") then
    sSearchText = replace(sSearchText,"-","")
end if

iPageSize 	= 20
iCurrPage 	= requestCheckvar(Request("iCP"),10)
IF iCurrPage = "" THEN iCurrPage =1
Set clsC = new CCust
	clsC.FCUSTgbn        =iCUSTgbn
	clsC.FCUSTtype       =sCUSTtype
	clsC.FARAP_TYPE 		 = "2" '계좌 지급처 정보만
	clsC.FSearchType     =iSearchType
	clsC.FSearchText     =sSearchText
	clsC.FPageSize       =iPageSize
	clsC.FCurrPage       =iCurrPage
	clsC.FRectAllacct    =allacct
	arrList = clsC.fnGetCustList
	iTotCnt	= clsC.FTotCnt
Set clsC  = nothing
 iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<script language="javascript">
    //alert('2016/04/30 sERP 업그레이드 작업중입니다. 등록하지 마세요. 서동석 문의 요망.');
<!--
   //검색
   function jsSearch(){
    document.frm.submit();
   }
	 function jsRegCust(sCCd,sENo, sBNo, sANo){
	 	var winR  = window.open("regCust.asp?hidCcd="+sCCd+"&hidENo="+sENo+"&hidBNo="+sBNo+"&hidANo="+sANo,"popR","width=900, height=700, resizable=yes, scrollbars=yes");
	 	winR.focus();
	}

	 // 페이지 이동
function jsGoPage(iCP)
	{
		document.frmReg.iCP.value=iCP;
		document.frmReg.submit();
	}

	//erp 목록 수신
	function jsGetErp(){
		document.frmErp.submit();
	}
	
	//계좐번호 사용안함처리
	function jsDelAcc(sCCd,  sBNo, sANo){
		if(confirm("거래처의 계좌번호를 삭제하시겠습니까?\n\n( 거래처가 아닌 계좌번호만 삭제됩니다. ) ")){ 
		document.frmAcc.hidCcd.value = sCCd;
		document.frmAcc.hidBNo.value = sBNo;
		document.frmAcc.hidANo.value = sANo;
		document.frmAcc.target="ifrPrc";
		document.frmAcc.submit();
	}
	}
//-->
</script>
<form name="frmErp" method="post" action="procCust.asp">
	<input type="hidden" name="hidM" value="R">
</form>
<form name="frmAcc" method="post" action="procCust.asp">
	<input type="hidden" name="hidM" value="DA">
	<input type="hidden" name="hidCcd" value="">
	<input type="hidden" name="hidBNo" value="">
	<input type="hidden" name="hidANo" value="">
</form>
<iframe name="ifrPrc" id="ifrPrc" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
	<tr>
	<td><strong>거래처  선택</strong> <br>
		 <hr width="100%">
		</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
		<form name="frm" method="get" action="popGetCust.asp">
		<input type="hidden" name="opnType" value="<%= opnType %>">

		<tr align="center" bgcolor="#FFFFFF" >
			<td rowspan="2" width="80" bgcolor="#EEEEEE"> 검색조건 </td>
			<td align="left">&nbsp;
				거래처구분:
				<input type="radio" name="rdoCgbn" value="0" <%IF iCUSTgbn ="" or iCUSTgbn = "0" then%>checked<%end if%>>전체
				<input type="radio" name="rdoCgbn" value="1" <%IF iCUSTgbn = "1" then%>checked<%end if%>>매출
				<input type="radio" name="rdoCgbn" value="2" <%IF iCUSTgbn = "2" then%>checked<%end if%>>매입
				&nbsp; &nbsp;
				거래처분류:
				<select name="selCT">
					<option value="" >전체</option>
					<% sbOptCustType sCUSTtype %>
				</select>
			 &nbsp; &nbsp;
			 <select name="selSTp">
			 	<option value="">-선택-</option>
			 	<option value="1" <%IF iSearchType="1" THEN%>selected<%END IF%>>거래처코드</option>
			 	<option value="2" <%IF iSearchType="2" THEN%>selected<%END IF%>>거래처명</option>
			 	<option value="3" <%IF iSearchType="3" THEN%>selected<%END IF%>>대표자</option>
			 	<option value="4" <%IF iSearchType="4" THEN%>selected<%END IF%>>담당자</option>
			 	<option value="5" <%IF iSearchType="5" THEN%>selected<%END IF%>>사업자번호</option>
				</select>
			 	: <input type="text" name="sSTx" size="20" value="<%=sSearchText%>">
			 &nbsp; &nbsp;
			 <input type="checkbox" name="allacct" <%=CHKIIF(allacct="on","checked","") %> >전체계좌보기
			</td>
			<td rowspan="2" width="50" bgcolor="#EEEEEE">
				<input type="button" class="button_s" value="검색" onClick="jsSearch();">
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td align="right"><input type="button" class="button" value="신규등록" onclick="jsRegCust('','','','');"><%IF C_MngPart OR C_ADMIN_AUTH or C_PSMngPart THEN%>&nbsp;<span><input type="button" class="button" value="ERP목록수신" onClick="jsGetErp();"></span><%END IF%></td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<form name="frmReg" method="post">
			<input type="hidden" name="iCP" value="">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>거래처코드</td>
				<td>구분</td>
				<td>분류</td>
				<td>거래처(직원)명</td>
				<td>담당자</td>
				<td>대표자</td>
				<td>사업자(주민)번호</td>
				<td>전화번호</td>
				<td>은행명</td>
				<td>계좌번호</td>
				<td>예금주</td>
				<td>처리</td>
			</tr>
			<%
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">
				<td><%=arrList(0,intLoop)%></td>
				<td><%IF arrList(2,intLoop)="Y" THEN%>매출<%END IF%> <%IF arrList(3,intLoop)="Y" THEN%>매입<%END IF%></td>
				<td><%=fnGetCustTypeName(arrList(1,intLoop))%></td>
				<td><%=arrList(4,intLoop)%></td>
				<td><%=arrList(19,intLoop)%></td>
				<td><%=arrList(6,intLoop)%></td>
				<td><%IF arrList(20,intLoop) = "2" THEN%><%=left(arrList(5,intLoop),6)%>******<%ELSE%><%=arrList(5,intLoop)%><%END IF%></td>
				<td><%=arrList(8,intLoop)%></td>
				<td><%=arrList(17,intLoop)%></td>
				<td><%=arrList(12,intLoop)%></td>
				<td><%=arrList(14,intLoop)%></td>
				<% if (opnType="eTax") then %>
					<td><input type="button" value="선택" class="button" onClick="opener.jsSetCust('<%=arrList(0,intLoop)%>','<%=Replace(arrList(4,intLoop),"'","")%>','<%=arrList(6,intLoop)%>','<%=arrList(5,intLoop)%>');self.close();">
				<% elseif (opnType="eTaxdetail") then %>
					<td><input type="button" value="선택" class="button" onClick="opener.jsSetCust('<%=arrList(0,intLoop)%>','<%=Replace(arrList(4,intLoop),"'","")%>','<%=arrList(6,intLoop)%>','<%=arrList(5,intLoop)%>','<%=arrList(22,intLoop)%>','<%=arrList(23,intLoop)%>','<%=arrList(24,intLoop)%>','<%=arrList(7,intLoop)%>','<%=arrList(8,intLoop)%>');self.close();">
				<% else %>
					<td><input type="button" value="선택" class="button" onClick="opener.jsSetCust('<%=arrList(0,intLoop)%>','<%=Replace(arrList(4,intLoop),"'","")%>','<%=arrList(17,intLoop)%>','<%=arrList(12,intLoop)%>','<%=arrList(14,intLoop)%>');self.close();">
				<% end if %>
				<input type="button" value="수정" class="button" onClick="jsRegCust('<%=arrList(0,intLoop)%>','<%=arrList(18,intLoop)%>','<%=arrList(11,intLoop)%>','<%=arrList(12,intLoop)%>')">
				 
				<input type="button" value="계좌번호삭제" class="button" onClick="jsDelAcc('<%=arrList(0,intLoop)%>','<%=arrList(11,intLoop)%>','<%=arrList(12,intLoop)%>')">

				</td>
			</tr>
		<%		Next
			ELSE%>
			<tr  align="center" bgcolor="#FFFFFF">
				<td colspan="13">등록된 내용이 없습니다.</td>
			</tr>
		<%END IF%>
		</table>
	</form>
	</td>
</tr>
<%
	Dim iStartPage,iEndPage,iX,iPerCnt
iPerCnt = 10

iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrPage mod iPerCnt) = 0 Then
	iEndPage = iCurrPage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<tr height="25" >
	<td colspan="15" align="center">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
		    <tr valign="bottom" height="25">
		        <td valign="bottom" align="center">
		         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
				<% else %>[pre]<% end if %>
		        <%
					for ix = iStartPage  to iEndPage
						if (ix > iTotalPage) then Exit for
						if Cint(ix) = Cint(iCurrPage) then
				%>
					<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
				<%		else %>
					<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
				<%
						end if
					next
				%>
		    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
				<% else %>[next]<% end if %>
		        </td>
		    </tr>
		</table>
	</td>
</tr>
</table>
<!-- 페이지 끝 -->
</body>
</html>
 <!-- #include virtual="/lib/db/dbclose.asp" -->
