<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->

<%
dim yyyy1, mm1, dd1

dim taxregdate, isusual, jungsandate, isipkumfinish, dategubun, ipkumdate
taxregdate      = request("taxregdate")
isusual         = request("isusual")
jungsandate     = request("jungsandate")
isipkumfinish   = request("isipkumfinish")

yyyy1           = request("yyyy1")
mm1             = request("mm1")
dd1             = request("dd1")

dategubun		= request("dategubun")
ipkumdate		= request("ipkumdate")

if dategubun="" then dategubun="Ipkum" end if
if isipkumfinish="" then isipkumfinish="Y" end if




if taxregdate<>"" then
    yyyy1   = Left(taxregdate,4)
    mm1     = Mid(taxregdate,6,2)
    dd1     = Mid(taxregdate,9,2)
elseif yyyy1<>"" then
    taxregdate = yyyy1 + "-" + mm1 + "-" + dd1
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))

if (taxregdate="") then
    taxregdate = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
end if

if (ipkumdate="") then
    ipkumdate = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
end if

dim ojungsan
set ojungsan = new CUpcheJungsan


if dategubun="Ipkum" then
	ojungsan.FRectIpkumDate = ipkumdate
elseif dategubun="Tax" then
    ojungsan.FRectTaxRegDate = taxregdate
end if

if isusual="Y" then
    ojungsan.FRectGubun ="EE"
elseif isusual="N" then
    ojungsan.FRectGubun ="FF"
end if

if (jungsandate<>"") then
    ojungsan.FRectjungsandate = jungsandate
end if

if isipkumfinish="Y" then
    ojungsan.FRectfinishflag = "7"
elseif isipkumfinish="N" then
    ojungsan.FRectfinishflag = "3"
elseif isipkumfinish="A" then
    ojungsan.FRectfinishflag = "ALL"
end if

'if jungsanmonth<>"" then
'    ojungsan.FRectYYYYMM = jungsanmonth
'end if

ojungsan.JungsanFixedList



dim i, ipsum
%>

<script language='javascript'>
function PopTaxPrintReDirect(itax_no, makerid){
	var popwinsub = window.open("/admin/upchejungsan/red_taxprint.asp?tax_no=" + itax_no + "&makerid=" + makerid,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
        	<!--
        	정산대상월 :
        	&nbsp;
        	-->
        	<select name="dategubun" class="select">
        	<option value="Tax" <% if dategubun="Tax" then response.write "selected" %> >발행일
        	<option value="Ipkum" <% if dategubun="Ipkum" then response.write "selected" %> >입금일
        	</select>
        	<% DrawOneDateBox yyyy1, mm1, dd1 %>
        	&nbsp;
        	정상발행구분 :
        	<select name="isusual" class="select">
	        	<option value="" <% if isusual="" then response.write "selected" %> >전체
	        	<option value="Y" <% if isusual="Y" then response.write "selected" %> >정상발행
	        	<option value="N" <% if isusual="N" then response.write "selected" %> >이월발행
        	</select>
        	&nbsp;
        	정산일 :
        	<select name="jungsandate" class="select">
	        	<option value="" <% if jungsandate="" then response.write "selected" %> >전체
	        	<option value="15일" <% if jungsandate="15일" then response.write "selected" %> >15일
	        	<option value="말일" <% if jungsandate="말일" then response.write "selected" %> >말일
	        	<option value="수시" <% if jungsandate="수시" then response.write "selected" %> >수시
        	</select>
        	&nbsp;
        	진행상태
        	<select name="isipkumfinish" class="select">
	        	<option value="A" <% if isipkumfinish="A" then response.write "selected" %> >확정이상
	        	<option value="N" <% if isipkumfinish="N" then response.write "selected" %> >정산확정
	        	<option value="Y" <% if isipkumfinish="Y" then response.write "selected" %> >입금완료
        	</select>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ojungsan.FresultCount %></b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="60">정산월</td>
        <td width="70">발행일</td>
        <td width="20"></td>
        <td width="70">입금일</td>
        <td>브랜드ID</td>
        <td>사업자명</td>
        <td width="80">사업자번호</td>
        <td width="80">정산금액</td>
        <td width="50">진행상태</td>
        <td width="40">결제일</td>
        <td width="50">은행</td>
        <td>계좌</td>
        <td>(현재계좌)</td>
    </tr>
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalSuplycash
%>

	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
	    <td><%= ojungsan.FItemList(i).FYYYYMM %></td>
	    <td><%= ojungsan.FItemList(i).FTaxRegDate %></td>
	    <td>
	    	<% if Not IsNULL(ojungsan.FItemList(i).Fneotaxno) then %>
			<a href="javascript:PopTaxPrintReDirect('<%= ojungsan.FItemList(i).Fneotaxno %>','<%= ojungsan.FItemList(i).Fdesignerid %>')"><img src="/images/icon_print02.gif" width="14" height="14" border=0></a>
			<% else %>
			<% end if %>
		</td>
	    <td>
	        <% if IsNULL(ojungsan.FItemList(i).FIpkumdate) or (ojungsan.FItemList(i).FIpkumdate="1900-01-01") then %>
	        <% else %>
	            <%= ojungsan.FItemList(i).FIpkumdate %>
	        <% end if %>
	    </td>
	    <td><%= ojungsan.FItemList(i).FDesignerid %></td>
	    <td><%= ojungsan.FItemList(i).Fcompany_name %></td>
	    <td><%= ojungsan.FItemList(i).Fcompany_no %></td>
	    <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
	    <td><font color="<%= ojungsan.FItemList(i).GetStatecolor %>"><%= ojungsan.FItemList(i).GetStateName %></font></td>
	    <td><%= ojungsan.FItemList(i).Fjungsan_date %></td>

        <% if ojungsan.FItemList(i).Fipkum_bank = "홍콩샹하이" then %>
		<td>HSBC</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "단위농협" then %>
		<td>농협</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "제일" then %>
		<td>sc제일</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "시티" then %>
		<td>한국씨티</td>
		<% else %>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
        <td>
            (
            <% if ojungsan.FItemList(i).Fjungsan_bank = "홍콩샹하이" then %>
            HSBC
    		<% elseif ojungsan.FItemList(i).Fjungsan_bank = "단위농협" then %>
    		농협
    		<% elseif ojungsan.FItemList(i).Fjungsan_bank = "제일" then %>
    		sc제일
    		<% elseif ojungsan.FItemList(i).Fjungsan_bank = "시티" then %>
    		한국씨티
    		<% else %>
    		<%= ojungsan.FItemList(i).Fjungsan_bank %>
    		<% end if %>
            <%= ojungsan.FItemList(i).Fjungsan_acctno %>)
        </td>

	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="5"></td>
	</tr>
</table>

<%
set ojungsan = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->