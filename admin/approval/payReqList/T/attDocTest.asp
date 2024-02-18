<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : 결제요청서 등록
' History : 2011.03.14 정윤정  생성
' 0 요청/1 진행중/ 5 반려/7 승인/ 9 완료
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"--> 
<!-- #include virtual="/lib/classes/approval/payManagerCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"--> 
<script language='javascript'>
function jsAttachDoc(v1,v2){
    var iURI = "popAddDoc.asp?idx="+v1;
    var popwin = window.open(iURI,'popAddDoc','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popTmsBaCUST(v1){
    var iURI = "/admin/approval/comm/popTmsBaCust.asp?cust_cd="+v1;
    var popwin = window.open(iURI,'popTmsBaCUST','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >  
<tr>
	<td>
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0" >
        <tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="#E6E6E6" align="center">
					<td rowspan="2" width="80">첨부서류</td>
					<td width="120">서류구분</td>
					<td>관련내용</td>
					<td><input type="button" value="추가" onClick="jsAttachDoc('idx','ret')"></td>
				</tr>
				<tr  bgcolor="#FFFFFF">
					<td align="center" valign="top" width="120">
						 <select name="paydoctype">
						 <option value="">선택
						 <option value="1">세금계산서
						 <option value="2">현금영수증
						 <option value="3">(기타)영수증
						 <option value="11">사업자등록증 사본
                         <option value="12">통장사본
                         <option value="21">거래명세표
                         <option value="99">기타파일
						 </select>
					</td>
					<td> 
					    <input type="button" value="거래처선택" onclick="popTmsBaCUST('');">
					    
					    <!--
						<div id="dFile"> 
						</div>
						<input type="text" name="sL" size="60" maxlength="120"><input type="button" value="파일첨부" class="button" onClick="jsAttachFile();"> 
                        -->						
					</td>
					<td> 
					    삭제
				    </td>
				</tr>
				</table>
			</td>
		</tr>
</table> 
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" --> 
