<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->

<%
dim lec_idx
lec_idx = RequestCheckvar(request("lec_idx"),10)


dim olecture
set olecture = new CLecture
olecture.FRectIdx = lec_idx

if lec_idx<>"" then
	olecture.GetOneLecture
end if


dim olecschedule
set olecschedule = new CLectureSchedule
olecschedule.FRectidx = lec_idx
olecschedule.FRectOptCd = "0000"
if lec_idx<>"" then
	olecschedule.GetOneLecSchedule
end if


dim i
%>
<script language='javascript'>
function SaveItem(frm){
	if (confirm('수정하시겠습니까?')){
		frm.submit();
	}
}


function popLecDateEdit(frm){
	var popwin = window.open('popLecOptionEdit.asp?lec_idx=<%=lec_idx%>','popLecDateEdit','width=700,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popLecMapEdit(frm){
	var popwin = window.open('','popLecMapEdit','width=600,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	강좌코드 : <input type="text" name="lec_idx" value="<%= lec_idx %>" Maxlength="12" size="12">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<% if olecture.FResultCount <1 then %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
	<td align="center">[ 검색 결과가 없습니다. ]</td>
</tr>
</table>
<% else %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name=frmlec method=post action="/academy/lecture/lib/do_leciteminfoedit.asp">
<input type="hidden" name="lec_idx" value="<%= lec_idx %>">
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">강좌코드</td>
		<td colspan="2"><%= lec_idx %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">강좌월구분</td>
		<td colspan="2"><b><%= olecture.FOneItem.Flec_date %></b></td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">강좌명</td>
		<td colspan="2"><%= olecture.FOneItem.Flec_title %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">검색어</td>
		<td colspan="2"><input type="text" name="keyword" value="<%= olecture.FOneItem.Fkeyword %>" size="50" maxlength="40"></td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">브랜드</td>
		<td colspan="2"><%= olecture.FOneItem.Flecturer_id %> (<%= olecture.FOneItem.Flecturer_name %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF"><td colspan="3"></td></tr>
	<tr  bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">수강료/매입가</td>
		<td width="300" colspan="2">
		<%= FormatNumber(olecture.FOneItem.Flec_cost,0) %> / <%= FormatNumber(olecture.FOneItem.Fbuying_cost,0) %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">재료비</td>
		<td bgcolor="#FFFFFF" colspan="2">
		<% if olecture.FOneItem.Fmatinclude_yn="Y" then %>
		포함(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
		<% else %>
		별도(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
		<% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">마일리지</td>
		<td colspan="2">
		<%= olecture.FOneItem.Fmileage %> (point)
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">마감여부</td>
		<td colspan="2">
		<% if olecture.FOneItem.IsSoldOut then %>
		<font color="#CC3333"><b>마감</b></font>
		<% else %>
		접수중
		<% end if %>
		<br> (마감기준 : 접수마감, 접수기간이외, 신청인원 정원초과, 전시안함, 사용안함 )
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">접수여부</td>
		<td colspan="2">
		<% if olecture.FOneItem.Freg_yn="Y" then %>
		<input type="radio" name="reg_yn" value="Y" checked >접수중
		<input type="radio" name="reg_yn" value="N">접수마감
		<% else %>
		<input type="radio" name="reg_yn" value="Y">접수중
		<input type="radio" name="reg_yn" value="N" checked ><font color="#CC3333">접수마감</font>
		<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="#DDDDFF">접수기간</td>
		<td colspan="2">
		<input type="text" name="reg_startday" value="<%= olecture.FOneItem.Freg_startday %>" size="10" maxlength="10">
		~
		<input type="text" name="reg_endday" value="<%= olecture.FOneItem.Freg_endday %>" size="10" maxlength="10">

		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td width=100 bgcolor="#DDDDFF">총정원-총신청 <br>= 총 남은인원</td>
		<td bgcolor="#FFFFFF" colspan="2">
		  <input type="text" size="3" maxlength="3" name="limit_count" value="<%= olecture.FOneItem.Flimit_count %>" readonly style="background-color='#EEEEEE'"> 명
		-
		  <input type="text" size="3" maxlength="3" name="limit_sold" value="<%= olecture.FOneItem.Flimit_sold %>" readonly style="background-color='#EEEEEE'"> 명
		=
		<input type="text" size="3" value="<%= olecture.FOneItem.GetRemainNo %>" readonly style="background-color='#EEEEEE'"> 명
		</td>
	</tr>

	<tr bgcolor="#FFFFFF" >
		<td width=100 bgcolor="#DDDDFF">강좌당 최소인원</td>
		<td bgcolor="#FFFFFF" colspan="2">
		<input type="text" size="3" maxlength="3" name="min_count" value="<%= olecture.FOneItem.Fmin_count %>" > 명
		</td>
	</tr>
	<tr bgcolor="#FFFFFF"><td colspan="3"></td></tr>
	<tr bgcolor="#FFFFFF" >
		<td width=100 bgcolor="#DDDDFF" rowspan="<%= olecschedule.FResultCount+1  %>">강의횟수 및 시간</td>
		<td bgcolor="#FFFFFF" colspan="2">
		<%= olecture.FOneItem.Flec_count %>회 &nbsp;&nbsp;&nbsp;<%= olecture.FOneItem.Flec_time %>시간
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="#FFFFFF" colspan="2">
		1. <%= olecture.FOneItem.Flec_startday1 & " ~ " & olecture.FOneItem.Flec_endday1 %>
		<% if (olecture.FOneItem.Flec_startday1<>olecschedule.FItemList(0).Fstartdate) or (olecture.FOneItem.Flec_endday1<>olecschedule.FItemList(0).Fenddate) then %>
		<br><b><%= olecschedule.FItemList(0).Fstartdate %></b> ~ <b><%= olecschedule.FItemList(0).Fenddate %></b>
		<% end if %>
		</td>
	</tr>
	<% for i=1 to olecschedule.FResultCount-1 %>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="#FFFFFF" colspan="2">
		<%= (i+1) & ". " & olecschedule.FItemList(i).Fstartdate & " ~ " & olecschedule.FItemList(i).Fenddate %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" >
		<td width=100 bgcolor="#DDDDFF">강의일정</td>
		<td bgcolor="#FFFFFF">
			등록 일정수 (<%= olecture.FOneItem.FoptionCnt %>)
		</td>
		<td align="right"><input type="button" value="강의 일정 수정" onclick="popLecDateEdit('<%= lec_idx %>');"></td>
	</tr>
	<tr bgcolor="#FFFFFF"><td colspan="3"></td></tr>
	<tr bgcolor="#FFFFFF" >
		<td width=100 bgcolor="#DDDDFF" >전시여부</td>
		<td colspan="2">
		<% if olecture.FOneItem.Fdisp_yn="Y" then %>
		<input type="radio" name="disp_yn" value="Y" checked >전시
		<input type="radio" name="disp_yn" value="N">전시안함
		<% else %>
		<input type="radio" name="disp_yn" value="Y">전시
		<input type="radio" name="disp_yn" value="N" checked ><font color="#CC3333">전시안함</font>
		<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td width=100 bgcolor="#DDDDFF" >사용여부</td>
		<td colspan="2">
		<% if olecture.FOneItem.Fisusing="Y" then %>
		<input type="radio" name="isusing" value="Y" checked >사용
		<input type="radio" name="isusing" value="N">사용안함
		<% else %>
		<input type="radio" name="isusing" value="Y">사용
		<input type="radio" name="isusing" value="N" checked ><font color="#CC3333">사용안함</font>
		<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td width=100 bgcolor="#DDDDFF" >약도</td>
		<td >
		<% if IsNULL(olecture.FOneItem.Flec_mapimg) or (olecture.FOneItem.Flec_mapimg="") then %>

		<% else %>
		<img src="<%= olecture.FOneItem.Flec_mapimg %>" width="200">
		<% end if %>

		</td>
		<td align="right"><input type="button" value="약도지정" onclick="popLecMapEdit('<%= lec_idx %>');"></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td width=100 bgcolor="#DDDDFF" >등록일</td>
		<td colspan="2">
		<%= olecture.FOneItem.Fregdate %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td width=100 bgcolor="#DDDDFF" >이미지</td>
		<td colspan="2">
		<img src="<%= olecture.FOneItem.Foblong_img4 %>" width="150" align="absmiddle">
		</td>
	</tr>



</form>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4"><input type="button" value="저장" onclick="SaveItem(frmlec)"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<% end if %>
<!-- 표 하단바 끝-->
<%
set olecschedule = Nothing
set olecture = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->