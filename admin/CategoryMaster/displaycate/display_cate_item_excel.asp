<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : 전시카테고리 상품 엑셀다운로드
' History	:  2021.07.12 한용민 생성
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	Dim cDisp, i, vDepth, vCateCode, vCurrpage, vPageSize, vIsThisCate, vParam, vSearch, vNotCateReg, dispCate, vOnlyBasic,arrlist
	vCurrPage	= NullFillWith(Request("cpg"), "1")
	vDepth 		= NullFillWith(Request("depth_s"), "1")
	vCateCode 	= Request("catecode_s")
	vIsThisCate	= Request("isthiscate")
	vPageSize	= NullFillWith(Request("pagesize"), 20)
	vSearch		= Request("search")
	vNotCateReg	= Request("notcatereg")
	vOnlyBasic	= request("onlybasic")
	dispCate	= Request("disp")

	Dim makerid, cdl, cdm, cds, itemid_s, itemname, keyword, sellyn, usingyn, danjongyn, limityn, sailyn, deliverytype, sortDiv, mustCate
	makerid		= request("makerid")
	cdl 		= request("cdl")
	cdm 		= request("cdm")
	cds 		= request("cds")
	itemid_s	= requestCheckvar(request("itemid_s"),1500)
	itemname	= request("itemname")
	keyword		= request("keyword")
	sellyn      = request("sellyn")
	usingyn     = request("usingyn")
	danjongyn   = request("danjongyn")
	limityn     = request("limityn")
	sailyn      = request("sailyn")
	deliverytype = request("deliverytype")
	sortDiv		= request("sortDiv")
	mustCate	= request("mustCate")

	if sortDiv = "" then sortDiv = "new"

	if itemid_s<>"" then
	dim iA ,arrTemp,arrItemid
	itemid_s = replace(itemid_s,",",chr(10))
	itemid_s = replace(itemid_s,chr(13),"")
	arrTemp = Split(itemid_s,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if trim(arrTemp(iA))<>"" then
			'상품코드 유효성 검사(2008.08.05;허진원)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	itemid_s = left(arrItemid,len(arrItemid)-1)
end if

	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 100000
	cDisp.FRectDepth = vDepth
	If vIsThisCate <> "" Then
		cDisp.FRectCateCode = vCateCode
	End IF
	cDisp.FRectMakerId 		= makerid
	cDisp.FRectItemID 		= itemid_s
	cDisp.FRectCDL 			= cdl
	cDisp.FRectCDM 			= cdm
	cDisp.FRectCDS 			= cds
	cDisp.FRectItemName 	= itemname
	cDisp.FRectKeyword 		= keyword
	cDisp.FRectSellYN		= sellyn
	cDisp.FRectIsUsing		= usingyn
	cDisp.FRectDanjongyn	= danjongyn
	cDisp.FRectLimityn		= limityn
	cDisp.FRectSailYn		= sailyn
	cDisp.FRectDeliveryType	= deliverytype
	cDisp.FRectSortDiv = SortDiv
	cDisp.FRectNotCateReg	= vNotCateReg
	cDisp.FRectOnlyBasic	= vOnlyBasic
	cDisp.FSearchDispCate	= dispCate
	cDisp.FRectMustCate		= mustCate
	cDisp.GetDispCateItemList_notpaging()
arrlist = cDisp.farrlist

Response.Buffer=true
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_전시카테고리상품_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
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

		<table width="100%" border="1" align="center" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
		<tr align="center" bgcolor="#F3F3FF" height="30">
			<td>Maker ID</td>
			<td>상품코드</td>
			<td>상품명</td>
			<td>지정된카테고리</td>
		</tr>
		<%
		If cDisp.FResultCount = 0 Then
		%>
			<tr>
				<td colspan="6" height="30" bgcolor="#FFFFFF" align="center">검색된 상품이 없습니다.</td>
			</tr>
		<%
		Else
			For i=0 To cDisp.FResultCount-1
		%>
			<tr bgcolor="#FFFFFF">
				<td align="center"><%= arrlist(3,i) %></td>
				<td  align="center"><%= arrlist(0,i) %></td>
				<td><%= replace(replace(replace( arrlist(1,i) ,",",""),vbcrlf,""),"<br>","") %></td>
				<td>
					<%= replace(replace(replace( fnCateCodeNameSplit_excel(arrlist(4,i),arrlist(0,i)) ,",",""),vbcrlf,""),"<br>","") %>
                </td>
			</tr>
		<%
            if i mod 1000 = 0 then
                Response.Flush		' 버퍼리플래쉬
            end if
			Next
		%>
		<%
		End If
		%>
		</table>
		
</div>
</body>
</html>
<% SET cDisp = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->