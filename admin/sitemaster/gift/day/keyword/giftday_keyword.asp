<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트
' History : 2014.03.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftday_cls.asp"-->
<%
dim keywordtype, keywordidx, keywordname, sortno, isusing, regdate
dim cgiftday, arrList, page, intLoop
	keywordidx = requestCheckVar(getNumeric(request("keywordidx")),10)
	page = requestCheckVar(getNumeric(request("page")),10)

keywordtype="1"
if page="" then page=1

SET cgiftday = New Cgiftday_list
	If keywordidx <> "" Then
		cgiftday.FRectkeywordidx = keywordidx
		cgiftday.FRectkeywordtype = keywordtype
		cgiftday.getgiftdaykeywordDetail
	
		keywordtype = cgiftday.FOneItem.Fkeywordtype
		keywordidx = cgiftday.FOneItem.Fkeywordidx
		keywordname = ReplaceBracket(cgiftday.FOneItem.Fkeywordname)
		sortno = cgiftday.FOneItem.Fsortno
		isusing = cgiftday.FOneItem.Fisusing
		regdate = cgiftday.FOneItem.Fregdate
	End If
	
	cgiftday.FPageSize = 500
	cgiftday.FCurrpage = page
	arrList = cgiftday.getgiftdaykeywordList
	
if sortno="" then sortno=99
if isusing="" then isusing="Y"
%>

<script type='text/javascript'>

	function jsUpdateCode(keywordidx){	
		self.location.href = "giftday_keyword.asp?keywordidx="+keywordidx+"&menupos=<%=menupos%>";
	}
	
	//코드 등록
	function jsRegCode(){
		var frm = document.frmReg;			 
		if(!frm.keywordname.value) {
			alert("코드명을 입력해 주세요");
			frm.keywordname.focus();
			return false;
		}
		if(!frm.sortno.value) {
			alert("정렬순서를 입력해 주세요");
			frm.sortno.focus();
			return false;
		}

		return true;
	}

</script>

<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//코드 등록 및 수정-->	
		<form name="frmReg" method="post" action="/admin/sitemaster/gift/day/keyword/giftday_keyword_process.asp" onSubmit="return jsRegCode();" style="margin:0px;">	
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="keywordtype" value="<%= keywordtype %>">
		<input type="hidden" name="mode" value="keywordedit">
		<input type="hidden" name="keywordidx" value="<%= keywordidx %>">
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<tr>			
			<td>	+ 코드 등록 및 수정</td>
		</tr>	
		<tr>
			<td>	
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<tr>
					<td bgcolor="#EFEFEF"   align="center">코드명</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="15" maxlength="16" name="keywordname" value="<%=keywordname%>">
						&nbsp;&nbsp;* ' 또는 " 는 입력이 안됩니다.
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">코드 정렬순서</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="4" maxlength="10" name="sortno" value="<%=sortno%>">
						&nbsp;&nbsp;* 숫자가 작을수록 상단에 있습니다.
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">사용여부</td>
					<td bgcolor="#FFFFFF">
					<input type="radio" value="Y" name="isusing" <%IF isusing ="Y" THEN%>checked<%END IF%>>사용 
					<input type="radio" value="N" name="isusing" <%IF isusing ="N" THEN%>checked<%END IF%>>사용안함 </td>
				</tr>
				</table>		
			</td>
		</tr>
		<tr>
			<td align="right"><input type="image" src="/images/icon_save.gif"> 
				<a href="javascript:jsSetCode('')"><img src="/images/icon_cancel.gif" border="0"></a></td>
		</tr>	
		<tr>
			<td colspan="2"><hr width="100%"></td>
		</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<form name="frmSearch" method="post" action="giftday_keyword.asp" style="margin:0px;">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="keywordtype" value="<%= keywordtype %>">
	<td colspan="2">+ 코드 리스트</td>
</tr>
<tr>
	<td>
	</td>
	<td align="right"><a href="javascript:jsSetCode('');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr>
	<td colspan="2">	
		<div id="divList" style="height:410px;overflow-y:scroll;">	
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">				
		<tr bgcolor="#EFEFEF" align="center">			
			<td width="50">코드값</td>
			<td>코드명</td>
			<td>정렬순서</td>
			<td>사용여부</td>
			<td>처리</td>
		</tr>
		
		<%If isArray(arrList) THEN%>
			<%For intLoop = 0 To UBound(arrList,2)%>
			<tr bgcolor="#FFFFFF" align="center">			
				<td><%=arrList(1,intLoop)%></td>
				<td><%= ReplaceBracket(arrList(2,intLoop)) %></td>
				<td><%=arrList(3,intLoop)%></td>
				<td><%=arrList(4,intLoop)%></td>
				<td>
					<input type="button" value="수정" onClick="javascript:jsUpdateCode('<%=arrList(1,intLoop)%>');" class="input_b">				
				</td>
			</tr>
			<%Next%>
		<%ELSE%>	
			<tr bgcolor="#FFFFFF">			
				<td colspan="5" align="center">등록된 내용이 없습니다.</td>
			</tr>
		<%End if%>
		</table>
		</div>
	</td>
	</form>
</tr>
</table>

<% 
SET cgiftday = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->