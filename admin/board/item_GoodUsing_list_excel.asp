<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : 상품후기 관리
' History	:  2021.11.29 한용민 생성
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/board/lib/classes/itemGoodUsingCls.asp" -->
<%
Dim page, SearchKey1, SearchKey2, selStatus, lp, lp2, sDt, eDt, dispcate, selPoint, chkTerm, srtMethod,blnPhotomode, chkFirst, searchKeyword
Dim strDel, makerid, orderserial, menupos, i, arrlist
	page 			= requestCheckvar(Request("page"),10)
	SearchKey1 = requestCheckvar(Request("SearchKey1"),32)
	SearchKey2 = requestCheckvar(Request("SearchKey2"),10)
	selStatus = requestCheckvar(Request("selStatus"),1)
	chkTerm 	= requestCheckvar(Request("chkTerm"),10)
	chkFirst 	= requestCheckvar(Request("chkFirst"),2)
	srtMethod = requestCheckvar(Request("srtMethod"),10)
	sDt = requestCheckvar(Request("sDt"),10)
	eDt = requestCheckvar(Request("eDt"),10)
	dispcate = requestCheckvar(Request("disp"),18)
	selPoint = requestCheckvar(Request("selPoint"),10)
	blnPhotomode = requestCheckvar(Request("photomode"),5)
	makerid     = requestCheckvar(request("makerid"),32)
	orderserial = requestCheckvar(request("orderserial"),12)
	searchKeyword = requestCheckvar(request("keyword"),30)
    menupos = requestCheckvar(Request("menupos"),10)

'기본값 지정
if page="" then page=1
if selStatus="" then selStatus="Y"
if srtMethod="" then srtMethod="idxDcd"
if sDt="" and chkTerm="" then sDt = date()
if eDt="" and chkTerm="" then eDt = date()

'// 상품 후기 목록
dim oGoodUsing
Set oGoodUsing = new CGoodUsing
	oGoodUsing.FPagesize = 20000
	oGoodUsing.FCurrPage = page
	oGoodUsing.FRectSearchKey1 = SearchKey1
	oGoodUsing.FRectSearchKey2 = SearchKey2
	oGoodUsing.FRectselStatus = selStatus
	oGoodUsing.FRectStartDt = sDt
	oGoodUsing.FRectEndDt = eDt
	oGoodUsing.FRectDispcate = dispcate
	oGoodUsing.FRectPoint = selPoint
	oGoodUsing.FRectPhotoMode = blnPhotomode
	oGoodUsing.FRectSort = srtMethod
	oGoodUsing.FRectMakerid = makerid
	oGoodUsing.FRectOrderserial = orderserial
	oGoodUsing.FRectFirst = chkFirst
	oGoodUsing.FRectKeyword = searchKeyword
	oGoodUsing.GetGoodUsingList_excel
    arrlist = oGoodUsing.farrlist

Response.Buffer=true
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_상품후기_" & Left(CStr(now()),10) & "_" & page & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
downFilemenupos=menupos
downPersonalInformation_rowcnt=oGoodUsing.ftotalcount
%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<style type="text/css">
 td {font-size:8.0pt;}
 .txt {mso-number-format:"\@";}
 .num {mso-number-format:"0";}
 .prc {mso-number-format:"\#\,\#\#0";}
</style>
</head>
<body>
<!--[if !excel]>　　<![endif]-->
<div align=center x:publishsource="Excel">

<table width="100%" border="1" align="center" class="a csH15" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="table-layout:fixed">
<tr bgcolor="#FFFFFF" align="left" >
	<td colspan="12">
		검색결과 : <b><%=FormatNumber(oGoodUsing.FResultCount,0)%></b> / Page : <b><%=page%></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>카테고리</td>
    <td>상품코드</td>
    <td>상품명</td>
	<td>옵션코드</td>
	<td>옵션명</td>
    <td>판매가</td>
    <td>작성자</td>
    <td>점수</td>
    <td>후기내용</td>
    <td>포토</td>
    <td>작성일</td>
    <td>상태</td>
</tr>
<% if isarray(arrlist) then %>
<%
for lp=0 to ubound(arrlist,2)
if arrlist(4,lp)="Y" then
    strDel = "일반"
else
    strDel = "삭제"
end if
%>
    <tr bgcolor="#FFFFFF" align="center" >
 	<td>
		<%= replace(arrlist(17,lp),"^^","》") %>
	</td>
	<td>
		<%= arrlist(7,lp) %>
	</td>
	<td class="txt">
		<%=db2html(arrlist(6,lp))%>
	</td>
	<td class="txt">
		<%=arrlist(18,lp)%>
	</td>
	<td class="txt">
		<%=db2html(arrlist(19,lp))%>
	</td>
	<td>
		<%= formatNumber(arrlist(16,lp), 0) %>
	</td>
	<td class="txt">
		<%= arrlist(1,lp) %>
	</td>
	<td align="left">
		<%= arrlist(8,lp) %>
	</td>
	<td class="txt">
		<%=db2html(arrlist(3,lp))%> 
	</td>
	<td>
		<% IF Not(arrlist(14,lp)="" or isNull(arrlist(14,lp))) Then %>
			http://imgstatic.10x10.co.kr/goodsimage/<%= GetImageSubFolderByItemid(arrlist(7,lp)) + "/" + arrlist(14,lp) %>
		<% End IF %>
		<% IF Not(arrlist(15,lp)="" or isNull(arrlist(15,lp))) Then %>
			http://imgstatic.10x10.co.kr/goodsimage/<%= GetImageSubFolderByItemid(arrlist(7,lp)) + "/" + arrlist(15,lp) %>
		<% End IF %>
	</td>
	<td><%=left(arrlist(13,lp),10)%></td>
	<td><%=strDel%></td>
    </tr>
<%
if i mod 100 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
end if
%>

</table>
</div>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
