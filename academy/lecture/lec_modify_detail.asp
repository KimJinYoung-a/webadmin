<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%

dim orderserial

orderserial = RequestCheckvar(request("orderserial"),16)


'==============================================================================
dim mode, sqlStr
dim detailidxlist, entrynamelist, entryhplist

mode = RequestCheckvar((request("mode"),10)
detailidxlist = html2db(request("detailidxlist"))
entrynamelist = html2db(request("entrynamelist"))
entryhplist = html2db(request("entryhplist"))
if detailidxlist <> "" then
	if checkNotValidHTML(detailidxlist) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if entrynamelist <> "" then
	if checkNotValidHTML(entrynamelist) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if	
if entryhplist <> "" then
	if checkNotValidHTML(entryhplist) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if (mode = "modify") then
    detailidxlist = split(detailidxlist, "|")
    entrynamelist = split(entrynamelist, "|")
    entryhplist = split(entryhplist, "|")

	for i = 0 to UBound(detailidxlist)
		if (trim(detailidxlist(i)) <> "") then
            sqlStr = " update [db_academy].[dbo].tbl_academy_order_detail "
            sqlStr = sqlStr + " set entryname = '" + CStr(trim(entrynamelist(i))) + "', entryhp = '" + CStr(trim(entryhplist(i))) + "' "
            sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and detailidx = " + CStr(trim(detailidxlist(i))) + " "
            rsACADEMYget.Open sqlStr,dbACADEMYget,1
		end if
	next

    response.write "<script>alert('저장되었습니다.');</script>"
    'dbget.close()	:	response.End
end if

'==============================================================================
dim ojumun
set ojumun = new CRequestLecture

ojumun.FRectOrderSerial = orderserial
ojumun.GetRequestLectureMasterOne


'==============================================================================
dim ojumundetail
set ojumundetail = new CRequestLecture

ojumundetail.FRectOrderSerial = orderserial
ojumundetail.CRequestLectureDetailList


'==============================================================================
dim olecture
set olecture = new CLecture
olecture.FRectIdx = ojumun.FOneItem.Fitemid

if (olecture.FRectIdx = "") then
    olecture.FRectIdx = "0"
end if
olecture.GetOneLecture


'==============================================================================
dim olecschedule
set olecschedule = new CLectureSchedule
olecschedule.FRectidx = ojumun.FOneItem.Fitemid
if (olecschedule.FRectIdx = "") then
    olecschedule.FRectIdx = "0"
end if

olecschedule.GetOneLecSchedule


'==============================================================================
if (Left(now, 10)  >= Left(olecture.FOneItem.Flec_startday1, 10)) then
    response.write "<script>alert('강좌시작 이전에만 수정이 가능합니다.'); opener.focus(); window.close();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
dim ix, i

%>
<script>
function SubmitSave() {
    var result_detailidx = "";
    var result_entryname = "";
    var result_entryhp = "";
    var e;

    for (var i = 0; i < frm.length; i++) {
        e = frm.elements[i];

        if (e.name == "detailidx") {
            result_detailidx = result_detailidx + "|" + e.value;
        }

        if (e.name == "entryname") {
            if (e.value == "") {
                alert("수강생 이름을 입력하세요.");
                return;
            }
            result_entryname = result_entryname + "|" + e.value;
        }

        if (e.name == "entryhp") {
            if (e.value == "") {
                alert("수강생 이름을 입력하세요.");
                return;
            }
            result_entryhp = result_entryhp + "|" + e.value;
        }
    }

    if (confirm("저장하시겠습니까?") == true) {
        frm.detailidxlist.value = result_detailidx;
        frm.entrynamelist.value = result_entryname;
        frm.entryhplist.value = result_entryhp;

        frm.submit();
    }
}

function CloseWindow() {
    opener.location.reload();
    opener.focus();
    window.close();
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
    <tr>
        <td>

                <!-- 신청인원 정보 -->
			<table width="100%" height="35" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
				<tr height="10" valign="bottom">
				    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
				    <td background="/images/tbl_blue_round_02.gif"></td>
				    <td background="/images/tbl_blue_round_02.gif"></td>
				    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
				</tr>
				<tr height="25">
				    <td background="/images/tbl_blue_round_04.gif"></td>
				    <td background="/images/tbl_blue_round_06.gif">
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>신청인원 정보</b>
				    </td>
				    <td align="right" background="/images/tbl_blue_round_06.gif">
				      <input type="button" value=" 저장 " onClick="SubmitSave()">
				      <input type="button" value=" 닫기 " onClick="CloseWindow()">
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
			</table>
			<table width="100%" height="185" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
				<tr>
					<td height="5" background="/images/tbl_blue_round_04.gif"></td>
				    <td></td>
				    <td></td>
				    <td></td>
				    <td></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<form name="frm" method="post" onSubmit="return false;">
				<input type="hidden" name="mode" value="modify">
				<input type="hidden" name="orderserial" value="<%= ojumun.FOneItem.FOrderSerial %>">
				<input type="hidden" name="detailidxlist" value="">
				<input type="hidden" name="entrynamelist" value="">
				<input type="hidden" name="entryhplist" value="">
<% for i = 0 to ojumundetail.FResultCount - 1 %>
				<input type="hidden" name="detailidx" value="<%= ojumundetail.FItemList(i).Fdetailidx %>">
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td width="60">수강생</td>
				    <td width="100">
				        <input type="text" name="entryname" value="<%= ojumundetail.FItemList(i).Fentryname %>" size="8" >
				    </td>
				    <td width="60">핸드폰</td>
				    <td>
				        <input type="text" name="entryhp" value="<%= ojumundetail.FItemList(i).Fentryhp %>" size="20" >
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
<% next %>
				</form>
				<tr>
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td></td>
				    <td></td>
				    <td></td>
				    <td></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="10" valign="top">
					<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
				</tr>
			</table>
			<!-- 신청인원 정보 -->

		</td>
	</tr>
</table>


<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->