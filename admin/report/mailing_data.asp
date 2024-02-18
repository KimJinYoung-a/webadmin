<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메일 통계
' History : 2007.08.27 한용민 생성
' History : 2016.12.07 유태욱 디자인 및 데이터 리스트 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
response.write "사용중지 매뉴 입니다. 매출통계v2>>메일진통계관리를 사용해 주세요."
response.end

dim page, i, omd
	page = requestcheckvar(getNumeric(request("page")),10)

if page="" then page=1

set omd = New CMailzine
	omd.FCurrPage = page
	omd.FPageSize=100
	omd.GetMailingList
%>
<script type="text/javascript">

function TnMailDataReg(frm){
	if(frm.title.value == ""){
		alert("발송이름을 적어주세요");
		frm.title.focus();
	}
	else if(frm.gubun.value == ""){
		alert("발송구분을 적어주세요");
		frm.gubun.focus();
	}
	else if(frm.startdate.value == ""){
		alert("발송시작시간을 적어주세요");
		frm.startdate.focus();
	}
	else if(frm.enddate.value == ""){
		alert("발송종료시간을 적어주세요");
		frm.enddate.focus();
	}
	else if(frm.reenddate.value == ""){
		alert("재발송종료시간을 적어주세요");
		frm.reenddate.focus();
	}
	else if(frm.totalcnt.value == ""){
		alert("총대상자수를 적어주세요");
		frm.totalcnt.focus();
	}
	else if(frm.realcnt.value == ""){
		alert("실발송통수를 적어주세요");
		frm.realcnt.focus();
	}
	else if(frm.realpct.value == ""){
		alert("실발송비율을 적어주세요");
		frm.realpct.focus();
	}
	else if(frm.filteringcnt.value == ""){
		alert("필터링통수를 적어주세요");
		frm.filteringcnt.focus();
	}
	else if(frm.filteringpct.value == ""){
		alert("필터링비율을 적어주세요");
		frm.filteringpct.focus();
	}
	else if(frm.successcnt.value == ""){
		alert("성공발송통수를 적어주세요");
		frm.successcnt.focus();
	}
	else if(frm.successpct.value == ""){
		alert("성공율을 적어주세요");
		frm.successpct.focus();
	}
	else if(frm.failcnt.value == ""){
		alert("실패발송통수를 적어주세요");
		frm.failcnt.focus();
	}
	else if(frm.failpct.value == ""){
		alert("실패율을 적어주세요");
		frm.failpct.focus();
	}
	else if(frm.opencnt.value == ""){
		alert("오픈통수를 적어주세요");
		frm.opencnt.focus();
	}
	else if(frm.openpct.value == ""){
		alert("오픈율을 적어주세요");
		frm.openpct.focus();
	}
	else if(frm.noopencnt.value == ""){
		alert("미오픈통수를 적어주세요");
		frm.noopencnt.focus();
	}
	else if(frm.noopenpct.value == ""){
		alert("미오픈율을 적어주세요");
		frm.noopenpct.focus();
	}
	else{
		frm.submit();
	}
}

</script>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= omd.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
	<td>발송 이름</td>
	<td>메일 제목</td>
	<td>총 대상자수</td>
	<td>성공 발송수</td>
	<td>오픈 통수</td>
	<td>클릭 통수</td>
	<td>발송 시간</td>
	<td>완료 시간</td>
	<td>메일러</td>
	<td>ETC</td>
</tr>
<% if omd.FResultCount>0 then %>
	<% for i=0 to omd.FResultCount-1 %>
	<tr bgcolor="FFFFFF">
		<td width="200"><% = omd.FItemList(i).Ftitle %></td>
		<td width="200"><% = omd.FItemList(i).fsubject %></td>
		<td align="right"><% = FormatNumber(omd.FItemList(i).Ftotalcnt,0) %></td>
		<td align="right"><% = FormatNumber(omd.FItemList(i).fsuccesscnt,0) %></td>
		<td align="right"><% = FormatNumber(omd.FItemList(i).fopencnt,0) %></td>
		<td align="right"><% = FormatNumber(omd.FItemList(i).fclickcnt,0) %></td>
		<td align="center"><% = omd.FItemList(i).Fstartdate %></td>
		<td align="center"><% = omd.FItemList(i).Fenddate %></td>
		<td align="center"><% = omd.FItemList(i).fmailergubun %></td>
		<td align="center"><a href="mailing_data_reg.asp?idx=<% = omd.FItemList(i).Fidx %>&mode=edit">상세내용보기</a></td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% End If %>
</table>

<% set omd = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->