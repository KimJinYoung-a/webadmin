<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/checknoticls.asp"-->
<%
Dim itemid, page, notignb, isconfirmed
Dim oChkNoti, i
page    				= request("page")
itemid  				= request("itemid")
notignb					= request("notignb")
isconfirmed				= request("isconfirmed")

If page = "" Then page = 1

'텐바이텐 상품코드 엔터키로 검색되게
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

SET oChkNoti = new CNoti
	oChkNoti.FCurrPage					= page
	oChkNoti.FPageSize					= 20
	oChkNoti.FRectItemID				= itemid
	oChkNoti.FRectNotignb				= notignb
	oChkNoti.FRectIsconfirmed			= isconfirmed
	oChkNoti.getCheckNotiList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function fnCheckValidAll(bool, comp){
    var frm = comp.form;

    if (!comp.length){
        if (comp.disabled==false){
            comp.checked = bool;
            AnCheckClick(comp);
        }
    }else{
        for (var i=0;i<comp.length;i++){
            if (comp[i].disabled==false){
                comp[i].checked = bool;
                AnCheckClick(comp[i]);
            }
        }
    }
}
// 선택된 상품 일괄 확인
function NotiSelectConfirmProcess() {
	var chkSel=0;
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
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 확인 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "I";
        document.frmSvArr.action = "/admin/ordermaster/checkNotiprocess.asp"
        document.frmSvArr.submit();
    }
}

function popjChkList(iitemid){
    var iurl = '/admin/etc/extsitejungsan_check.asp?menupos=1&itemid='+iitemid;
    var pop = window.open(iurl,'popjChkList','resizable=yes');
    popjChkList.focus();
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
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		구분 : 
		<select name="notignb" class="select">
			<option value="">전체</option>
			<option value="11" <%= CHkIIF(notignb="11","selected","") %>>판매가</option>
		</select>
		&nbsp;
		확인여부 :
		<select name="isconfirmed" class="select">
			<option value="">전체</option>
			<option value="1" <%= CHkIIF(isconfirmed="1","selected","") %>>확인완료</option>
			<option value="0" <%= CHkIIF(isconfirmed="0","selected","") %>>확인전</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<input type="button" class="button_s" value="확인" onClick="NotiSelectConfirmProcess();">
<p>
<!-- 액션 끝 -->
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		검색결과 : <b><%= FormatNumber(oChkNoti.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oChkNoti.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onclick="fnCheckValidAll(this.checked,frmSvArr.cksel);"></td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td width="60">구분</td>
	<td width="80">알림횟수</td>
	<td width="100">체크값</td>
	<td width="120">등록일</td>
	<td width="120">최종Check일</td>
	<td>최종Check내용</td>
	<td width="70">확인여부</td>
	<td width="70">금액변동</td>
	<td width="120">최종확인일</td>
	<td width="100">최종확인자</td>
</tr>

<% For i=0 to oChkNoti.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oChkNoti.FItemList(i).FItemID %>" <%= Chkiif(oChkNoti.FItemList(i).Fisconfirmed ="1", "disabled","") %>></td>
	<td align="center"><a href="<%=wwwURL%>/<%=oChkNoti.FItemList(i).FItemID%>" target="_blank"><%= oChkNoti.FItemList(i).FItemID %></a></td>
	<td align="center"><%= oChkNoti.FItemList(i).getNotignbStr %></td>
	<td align="center"><%= oChkNoti.FItemList(i).FNoticnt %></td>
	<td align="center"><%= oChkNoti.FItemList(i).FChkData %></td>
	<td align="center"><%= oChkNoti.FItemList(i).FRegdate %></td>
	<td align="center"><%= oChkNoti.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oChkNoti.FItemList(i).FNotistr %></td>
	<td align="center"><%= oChkNoti.FItemList(i).getConfirmedStr %></td>
	<td align="center"><input type="button" value="보기" onclick="popjChkList('<%= oChkNoti.FItemList(i).FItemID %>')"></td>
	<td align="center"><%= oChkNoti.FItemList(i).FLastconfirmDT %></td>
	<td align="center"><%= oChkNoti.FItemList(i).FLastconfirmUser %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oChkNoti.HasPreScroll then %>
		<a href="javascript:goPage('<%= oChkNoti.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oChkNoti.StartScrollPage to oChkNoti.FScrollCount + oChkNoti.StartScrollPage - 1 %>
    		<% if i>oChkNoti.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oChkNoti.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oChkNoti = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
