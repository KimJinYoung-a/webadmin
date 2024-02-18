<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/member/tenbyten/companyCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%

dim i, j, k
dim idx
dim adminuserid, mode

	idx = requestcheckvar(request("idx"),10)

adminuserid = session("ssBctId")

if (idx = "") then
	mode = "ins"
else
	mode = "mod"
end if

dim oCompanyCalendar
set oCompanyCalendar = new CCompanyCalendar
	oCompanyCalendar.FRectIdx = idx

	oCompanyCalendar.getCompanyCalendarItem()


dim oCompanyCalendarDetail
set oCompanyCalendarDetail = new CCompanyCalendar
	oCompanyCalendarDetail.FCurrPage = 1
	oCompanyCalendarDetail.FPageSize = 200
	oCompanyCalendarDetail.FRectIdx = idx

	if (idx <> "") then
		oCompanyCalendarDetail.getPartOrMemberList()
	end if


dim cMember
Set cMember = new CTenByTenMember
	cMember.Fempno = session("ssBctSn")
	cMember.fnGetMemberData


dim oneDateType : oneDateType = True

if (Left(oCompanyCalendar.FOneItem.FstartDate,10) <> Left(oCompanyCalendar.FOneItem.FendDate,10)) then
	oneDateType = False
end if

%>

<script type="text/javascript">

function jsRegCalendarItem() {
	var from = document.frm;

	if (frm.title.value == "") {
		alert("제목을 입력하세요.");
		frm.title.focus();
		return;
	}

	if (frm.contents.value == "") {
		alert("내용을 입력하세요.");
		frm.contents.focus();
		return;
	}

	if (frm.startDate.value == "") {
		alert("기간을 입력하세요.");
		frm.startDate.focus();
		return;
	}

	if (frm.dateType[0].checked == true) {
		frm.endDate.value = frm.startDate.value;
	}

	if (frm.endDate.value == "") {
		alert("기간을 입력하세요.");
		frm.endDate.focus();
		return;
	}

	if (confirm("저장하시겠습니까?") != true) {
		return;
	}

	frm.action="calendar_process.asp";

	frm.submit();
}

function jsPopCal(fName,sName) {
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false) {
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function jsPopSelectPart() {
	var pop = window.open("popSelectPart.asp", "jsPopSelectPart","width=500,height=150,scrollbars=no");
	pop.focus();
}

function jsPopSelectMember() {
	var pop = window.open("popSelectMember.asp", "jsPopSelectMember","width=700,height=600,scrollbars=yes");
	pop.focus();
}

function addPartItem(valName, valId) {
	return jsCheckAndAddItem("tbl_part", "department_id", valName, valId);
}

function addMemberItem(valName, valId) {
	return jsCheckAndAddItem("tbl_member", "empno", valName, valId);
}

function jsCheckAndAddItem(tblname, elename, valName, valId) {
	var objTbl = eval(tblname);
	var objEle = eval("document.frm." + elename);
	var lenRow = objTbl.rows.length;

	if (objEle != undefined) {
		if (lenRow == 1) {
			if (objEle.value == valId) {
				return "이미 추가되어 있습니다.";
			}
		} else if (lenRow > 1) {
			for(var i = 0; i < objEle.length; i++) {
				if (objEle[i].value == valId) {
					return "이미 추가되어 있습니다.";
				}
			}
		}
	}

	var oRow = objTbl.insertRow(lenRow);
	var oCell1 = oRow.insertCell(0);
	var oCell2 = oRow.insertCell(1);

	oCell1.innerHTML = valName + "<input type='hidden' name='" + elename + "' value='" + valId + "'>";
	oCell2.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick=\"jsDelItem('" + tblname + "', '" + elename + "', " + valId + ")\" align=absmiddle>";

	return "";
}

function jsDelItem(tblname, elename, valId) {
	var objTbl = eval(tblname);
	var objEle = eval("document.frm." + elename);
	var lenRow = objTbl.rows.length;

	if (objEle != undefined) {
		if (lenRow == 1) {
			//if (objEle.value == valId) {
				objTbl.deleteRow(0);
				return;
			//}
		} else if (lenRow > 1) {
			for(var i = 0; i < objEle.length; i++) {
				if (objEle[i].value == valId) {
					objTbl.deleteRow(i);
					return;
				}
			}
		}
	}
}

function jsShowHideEndDate(showhide) {
	if (showhide == true) {
		document.getElementById("endDate").style.display = 'block';
	} else {
		document.getElementById("endDate").style.display = 'none';
	}
}

</script>

<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" style="margin:0px;">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<input type="hidden" name="mode" value="<%= mode %>">
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">IDX</td>
	<td>
		<%= oCompanyCalendar.FOneItem.Fidx %>
		<input type="hidden" name="idx" value="<%= oCompanyCalendar.FOneItem.Fidx %>">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">제목</td>
	<td>
		<input type="text" class="text" name="title" value="<%= oCompanyCalendar.FOneItem.Ftitle %>" size="64" maxlength=128>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">내용</td>
	<td>
		<textarea class="textarea" name="contents" cols=80 rows=5><%= oCompanyCalendar.FOneItem.Fcontents %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">날짜</td>
	<td>
		<table border="0" align="left" class="a" cellpadding="0" cellspacing="1">
			<tr>
				<td>
					<input type="text" class="text" name="startDate" value="<%= oCompanyCalendar.FOneItem.FstartDate %>" size="10" maxlength="10" onClick="jsPopCal('frm','startDate');" style="cursor:hand;">
				</td>
				<td>
					<div id="endDate" style="display:none">
						&nbsp;
						~
						&nbsp;
						<input type="text" class="text" name="endDate" value="<%= oCompanyCalendar.FOneItem.FendDate %>" size="10" maxlength="10" onClick="jsPopCal('frm','endDate');" style="cursor:hand;">
					</div>
				</td>
				<td>
					&nbsp;
					<input type="radio" name="dateType" value="1" onClick="jsShowHideEndDate(false)" <% if (oneDateType = True) then %>checked<% end if %> > 1일
					<input type="radio" name="dateType" value="A" onClick="jsShowHideEndDate(true)" <% if (oneDateType <> True) then %>checked<% end if %> > 기간
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">우선순위</td>
	<td>
		<select class="select" name="importantLevel">
			<option value="0" <% if (oCompanyCalendar.FOneItem.FimportantLevel = "0") then %>selected<% end if %> >없음</option>
			<option value="10" <% if (oCompanyCalendar.FOneItem.FimportantLevel = "10") then %>selected<% end if %> >낮음</option>
			<option value="20" <% if (oCompanyCalendar.FOneItem.FimportantLevel = "20") then %>selected<% end if %> >보통</option>
			<option value="30" <% if (oCompanyCalendar.FOneItem.FimportantLevel = "30") then %>selected<% end if %> >높음</option>
		</select>
	</td>
</tr>
<input type="hidden" name="openLevel" value="0">
<!--
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">공개범위</td>
	<td>
		<select class="select" name="openLevel">
			<option value="0" <% if (oCompanyCalendar.FOneItem.FopenLevel = "0") then %>selected<% end if %> >없음</option>
			<option value="10" <% if (oCompanyCalendar.FOneItem.FopenLevel = "10") then %>selected<% end if %> >부서공개</option>
			<option value="20" <% if (oCompanyCalendar.FOneItem.FopenLevel = "20") then %>selected<% end if %> >전체공개</option>
		</select>
	</td>
</tr>
-->
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">공유부서</td>
	<td>
		<table class=a>
			<tr>
				<td>
					<table name='tbl_part' id='tbl_part' class=a>
						<%
						if (oCompanyCalendarDetail.FResultCount > 0) then
							for i = 0 to oCompanyCalendarDetail.FResultCount - 1
								if Not IsNull(oCompanyCalendarDetail.FItemList(i).Fdepartment_id) then
									Response.Write "<tr>"
									Response.Write "<td>" & oCompanyCalendarDetail.FItemList(i).FdepartmentNameFull & "<input type='hidden' name='department_id' value=" & oCompanyCalendarDetail.FItemList(i).Fdepartment_id & "></td>"
									Response.Write "<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick=""jsDelItem('tbl_part', 'department_id', '" & oCompanyCalendarDetail.FItemList(i).Fdepartment_id & "')"" align=absmiddle></td>"
									Response.Write "</tr>"
								end if
							next
						end if
						%>
					</table>
				</td>
				<td>
					<input type="button" value=" 추 가 " onclick="jsPopSelectPart();" class="button">
					&nbsp;
					<input type="button" value="내부서" onclick="addPartItem('<%= cMember.FdepartmentNameFull %>', '<%= cMember.Fdepartment_id %>');" class="button">
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">공유인원</td>
	<td>
		<table class=a>
			<tr>
				<td>
					<table name='tbl_member' id='tbl_member' class=a>
						<%
						if (oCompanyCalendarDetail.FResultCount > 0) then
							for i = 0 to oCompanyCalendarDetail.FResultCount - 1
								if Not IsNull(oCompanyCalendarDetail.FItemList(i).Fempno) then
									Response.Write "<tr>"
									Response.Write "<td>" & oCompanyCalendarDetail.FItemList(i).Fdepartmentname & " - " & oCompanyCalendarDetail.FItemList(i).Fusername & "&nbsp;" & oCompanyCalendarDetail.FItemList(i).Fposit_name & "<input type='hidden' name='empno' value='" & oCompanyCalendarDetail.FItemList(i).Fempno & "'></td>"
									Response.Write "<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick=""jsDelItem('tbl_member', 'empno', '" & oCompanyCalendarDetail.FItemList(i).Fempno & "')"" align=absmiddle></td>"
									Response.Write "</tr>"
								end if
							next
						end if
						%>
					</table>
				</td>
				<td>
					<input type="button" value=" 추 가 " onclick="jsPopSelectMember();" class="button">
					&nbsp;
					<input type="button" value="나자신" onclick="addMemberItem('<%= cMember.Fusername %>', '<%= cMember.Fempno %>');" class="button">
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">사용여부</td>
	<td>
		<select class="select" name="useYN">
			<option value="Y" <% if (oCompanyCalendar.FOneItem.FuseYN = "Y") then %>selected<% end if %> >사용</option>
			<option value="N" <% if (oCompanyCalendar.FOneItem.FuseYN = "N") then %>selected<% end if %> >사용안함</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">등록자</td>
	<td>
		<%= oCompanyCalendar.FOneItem.Freguserid %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">최종수정</td>
	<td>
		<%= oCompanyCalendar.FOneItem.Flastupdate %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="35">
	<td align="center" colspan="2">
		<input type="button" value="저장" onclick="jsRegCalendarItem();" class="button">
	</td>
</tr>
</table>
</form>

<%
set oCompanyCalendar = nothing
set oCompanyCalendarDetail = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
