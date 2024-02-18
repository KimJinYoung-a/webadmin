<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/optManager/optManagerCls.asp"-->
<%
Dim oOptManager, page, mallid, makerid, itemid, isReged, notmallid
Dim i, iPerCnt, arrList
Dim iStartPage, iEndPage, iTotalPage, ix, iTotCnt, iPageSize
Dim newCode
iPerCnt = 10
iPageSize = 20

page = request("page")
mallid = request("mallid")
notmallid = request("notmallid")
makerid = request("makerid")
itemid = request("itemid")
isReged = request("isReged")

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If


If page = "" Then page = 1

SET oOptManager = new COptManager
	oOptManager.FCurrPage			= page
	oOptManager.FPageSize			= iPageSize
	oOptManager.FRectMallid			= mallid
	oOptManager.FRectNotMallid		= notmallid
	oOptManager.FRectMakerid		= makerid
	oOptManager.FRectItemid			= itemid
	oOptManager.FRectIsReged		= isReged
	oOptManager.FRectCDL			= request("cdl")
	oOptManager.FRectCDM			= request("cdm")
	oOptManager.FRectCDS			= request("cds")
	arrList = oOptManager.getoOptManagerItemList
	iTotCnt	= oOptManager.FTotalCount
SET oOptManager = nothing

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<script language="javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function regOptMall() {
	var chkSel=0;
	var chkmall = 0;

	try {
		if(frmSvArr.chkmall.length>1) {
			for(var i=0;i<frmSvArr.chkmall.length;i++) {
				if(frmSvArr.chkmall[i].checked) chkmall++;
			}
		} else {
			if(frmSvArr.chkmall.checked) chkmall++;
		}
		if(chkmall<=0) {
			alert("선택한 몰이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert(e);
		alert("몰이 없습니다.");
		return;
	}

	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert(e);
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개의 상품을 등록 하시겠습니까?\n\n등록된 상품은 추가금액상품관리에서 확인가능합니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "I";
        document.frmSvArr.action = "/admin/etc/optManager/optManagerProc.asp"
        document.frmSvArr.submit();
    }
}
function checkismall(comp){
	if(comp.name =="notmallid"){
		if(comp.value=="reset"){
			comp.value = "";
			comp.form.mallid.disabled=false;
		}else{
			comp.form.mallid.value = "";
			comp.form.mallid.disabled=true;
		}
	}else if(comp.name =="mallid"){
		if(comp.value=="reset"){
			comp.value = "";
			comp.form.notmallid.disabled=false;
		}else{
			comp.form.notmallid.value = "";
			comp.form.notmallid.disabled=true;
		}
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>
		&nbsp;
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		몰 구분 :
		<select name="mallid" class="select" onchange="checkismall(this);">
			<option value="">전체</option>
			<option value="gsshop" 		<%= chkiif(mallid = "gsshop", "selected", "")%>>GSShop</option>
			<option value="lotteimall"	<%= chkiif(mallid = "lotteimall", "selected", "")%>>Lotteimall</option>
			<option value="reset"		<%= chkiif(mallid = "reset", "selected", "")%>>검색Reset</option>
		</select>
		&nbsp;
		미등록 :
		<select name="notmallid" class="select" onchange="checkismall(this);">
			<option value="">-Choice-</option>
			<option value="gsshop" 		<%= chkiif(notmallid = "gsshop", "selected", "")%>>GSShop</option>
			<option value="lotteimall"	<%= chkiif(notmallid = "lotteimall", "selected", "")%>>Lotteimall</option>
			<option value="reset"		<%= chkiif(notmallid = "reset", "selected", "")%>>검색Reset</option>
		</select>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="mallid" value="<%=mallid%>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		아래 체크박스에 체크 후 등록하세요<br>
		<input type="checkbox" name="chkmall" value="gsshop">GSShop
		<input type="checkbox" name="chkmall" value="lotteimall">lotteimall
		<input type="button" value="등록" class="button" onclick="regOptMall();">
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(iTotCnt,0) %></b>&nbsp;&nbsp;페이지 : <b><%=page%> / <%=iTotalPage%></b></td>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="80">상품코드</td>
	<td width="60">옵션코드</td>
	<td width="100">브랜드ID</td>
	<td>상품명<font color="BLUE">[옵션명]</font></td>
	<td width="100">판매가</td>
	<td width="100">옵션추가금액</td>
	<td width="100">예상판매가</td>
	<td width="300">적용몰</td>
</tr>
<% If iTotCnt > 0 Then %>
<% For i = 0 To UBound(arrList,2) %>
<%
	newCode = CStr(arrList(0,i))&"_"&CStr(arrList(1,i))
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= newCode %>"></td>
	<td><%= arrList(0,i) %></td>
	<td><%= arrList(1,i) %></td>
	<td align="center"><%= arrList(2,i) %></td>
	<td align="left">
		<%= arrList(3,i) %>&nbsp;<font color="BLUE">[<%= arrList(4,i) %>]</font><br>
		<font color="purple">상품명변경 :</font> <input type="text" value="<%= arrList(9,i) %>" size="50" style="color:red" name="newitemname|<%=newCode%>">
	</td>
	<td align="center"><%= FormatNumber(arrList(6,i),0) %></td>
	<td align="center"><%= FormatNumber(arrList(7,i),0) %></td>
	<td align="center"><%= FormatNumber(arrList(8,i),0) %></td>
	<td align="center"><%= arrList(10,i) %></td>
</tr>
<% Next %>
<% End If %>
<%
iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

If (page mod iPerCnt) = 0 Then
	iEndPage = page
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
         <% if (iStartPage-1 )> 0 then %><a href="javascript:goPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% else %>[pre]<% end if %>

        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(page) then
		%>
    		<font color="red">[<%= ix %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= ix %>');">[<%= ix %>]</a>
    		<% end if %>
    	<% next %>

    	<% if Clng(iTotalPage) > Clng(iEndPage)  then %>
    		<a href="javascript:goPage('<%= ix %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->