<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제요청서 리스트
' History : 2011.10.13 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<!-- #include virtual="/lib/classes/approval/innerPartcls.asp"-->
<%

dim i, page, research

page = requestCheckvar(Request("page"),32)
research = requestCheckvar(Request("research"),32)

if (page = "") then
	page = 1
end if



'==============================================================================
dim oinnerpart
set oinnerpart = New CInnerPart

oinnerpart.FCurrPage = page
oinnerpart.FPageSize = 20

oinnerpart.GetInnerPartList

%>



 <script language="javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">

function jsNewRegPart() {
	var winR = window.open("popRegInnerPart.asp","jsNewRegPart","width=500, height=300, resizable=yes, scrollbars=yes");
	winR.focus();
}

function jsModifyPart(idx) {
	var winR = window.open("popRegInnerPart.asp?idx=" + idx,"jsModifyPart","width=500, height=300, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popViewInnerOrderDetail(masteridx) {
	var winR = window.open("popViewInnerOrderDetail.asp?idx="+masteridx,"popViewInnerOrderDetail","width=500, height=700, resizable=yes, scrollbars=yes");
	winR.focus();
}

function jsSearch(){
 document.frm.submit();
}

function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function checkThis(comp) {
	var frm = comp.form;

    AnCheckClick(comp);

    if (comp.checked != true) {
    	frm.chkAll.checked = false;
    }
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function jsDelSelected(frm) {

	var checkeditemfound = false;
	for (var i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];

		if (e.type == "checkbox") {
			if (e.name == "chk") {
				if (e.checked == true) {
					checkeditemfound = true;
					break;
				}
			}
		}
	}

	if (checkeditemfound == false) {
		alert("선택된 내역이 없습니다.");
		return;
	}

    if (confirm('선택 내역을 삭제하시겠습니까?') == true) {
	    frm.mode.value="delselectedarr";
	    frm.submit();
	}
}

</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="innerPartList.asp">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="research" value="on">
			<input type="hidden" name="page" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="1" width="100" height="30" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					<select name="gubun">
					<option value="">--선택--</option>
					<option value="S">매장</option>
					<option value="M">매입부서</option>
					</select>
				</td>
				<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<tr>
	<td>
	    <table width="100%" cellspacing="0" cellpadding="0">
	    <tr>
	        <td align="left">
	        	<input type="button" class="button" value=" 내부부서 등록 " onClick="jsNewRegPart();">
	        </td>
	        <td align="right">
	        </td>
	    </tr>
	    </table>
	</td>
</tr>

<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a"   border="0">
		<Form name="frmAct" method="post" action="innerpart_process.asp">
		<input type="hidden" name="mode" value="">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td width="40">IDX</td>
					<td width="60">구분</td>
					<td>ERP부서명</td>
					<td width="80">ERP부서코드</td>
					<td width="100">어드민부서코드</td>
					<td width="80">작성일</td>
					<td>비고</td>
				</tr>
				<%IF oinnerpart.FResultCount > 0 THEN %>
				<% for i = 0 to (oinnerpart.FResultCount - 1) %>
				<tr bgcolor="#FFFFFF" align="center">
					<td><a href="javascript:jsModifyPart(<%= oinnerpart.FItemList(i).Fidx %>);"><%= oinnerpart.FItemList(i).Fidx %></a></td>
					<td><font color="<%= oinnerpart.FItemList(i).GetDivcdColor %>"><%= oinnerpart.FItemList(i).GetDivcdName %></font></td>
					<td align=left><a href="javascript:jsModifyPart(<%= oinnerpart.FItemList(i).Fidx %>);"><%= oinnerpart.FItemList(i).FBIZSECTION_NM %></a></td>
					<td align=left><%= oinnerpart.FItemList(i).FBIZSECTION_CD %></td>
					<td align=left><%= oinnerpart.FItemList(i).Fscmid %></td>
					<td><%= Left(oinnerpart.FItemList(i).Fregdate, 10) %></td>
					<td></td>
				</tr>
				<%
					Next
				%>
				<%
					ELSE
				%>
				<tr bgcolor="#FFFFFF">
					<td colspan="7" align="center">등록된 내역이 없습니다.</td>
				</tr>
				<%END IF%>
				</table>
			</td>
		</tr>
        </form>
	    <tr align="center" bgcolor="#FFFFFF">
	        <td colspan="7">
	            <% if oinnerpart.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oinnerpart.StartScrollPage-1 %>')">[pre]</a>
	    		<% else %>
	    			[pre]
	    		<% end if %>

	    		<% for i=0 + oinnerpart.StartScrollPage to oinnerpart.FScrollCount + oinnerpart.StartScrollPage - 1 %>
	    			<% if i>oinnerpart.FTotalpage then Exit for %>
	    			<% if CStr(page)=CStr(i) then %>
	    			<font color="red">[<%= i %>]</font>
	    			<% else %>
	    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>

	    		<% if oinnerpart.HasNextScroll then %>
	    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
	        </td>
	    </tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->