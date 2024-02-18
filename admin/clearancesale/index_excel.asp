<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : Ŭ����� ���� Excel���� ����
'	History		: 2022.02.11 ����; ������
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/ClearanceSale/ClearanceSaleCls.asp"-->
<%
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_CLEARANCE_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '���ۻ�뿩��
%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<%
Dim i, idx
Dim FResultCount, iTotCnt, iCurrentpage
Dim itemid, rectitemid, itemname, makerid, usingyn, sellyn, limityn, catecode, sailyn, itemcouponyn
dim iSalePercent

	idx = request("idx")
	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
	itemid      = requestCheckvar(request("itemid"),255)
	rectitemid  = requestCheckvar(request("rectitemid"),255)
	itemname    = requestCheckvar(request("itemname"),64)
	makerid     = requestCheckvar(request("makerid"),32)
	sellyn      = requestCheckvar(request("sellyn"),10)
	usingyn     = requestCheckvar(request("usingyn"),10)
	limityn     = requestCheckvar(request("limityn"),10)
	catecode    = requestCheckvar(request("catecode"),10)
	sailyn      = requestCheckvar(request("sailyn"),10)
	itemcouponyn = requestCheckvar(request("itemcouponyn"),10)

if iCurrentpage="" then iCurrentpage=1

if rectitemid<>"" then
	dim iA ,arrTemp,arrrectitemid
  rectitemid = replace(rectitemid,chr(13),"")
	arrTemp = Split(rectitemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrrectitemid = arrrectitemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrrectitemid)>0 then
		rectitemid = left(arrrectitemid,len(arrrectitemid)-1)
	else
		if Not(isNumeric(rectitemid)) then
			rectitemid = ""
		end if
	end if
end if

dim oclear
set oclear = new CClaearanceitem
	oclear.FPageSize = 5000
	oclear.FRectItemid		= rectitemid
	oclear.FRectSellYN		= sellyn
	oclear.FRectIsusing		= usingyn
	oclear.FRectMakerid		= makerid
	oclear.FRectLimityn		= limityn
	oclear.FRectCatecode		= catecode
	oclear.FRectSaleYN		= sailyn
	oclear.FRectItemcouponYN	= itemcouponyn
	oclear.FRectitemname	= itemname
	oclear.FCurrPage = iCurrentpage
	oclear.fnGetclaearanceitemList
iTotCnt = oclear.FTotalCount
%>
<% '����Ʈ--------------------------------------------------------------------------------------------- %>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td><strong>��ǰ �ڵ�</strong></td>
		<td><strong>�귣��</strong></td>
		<td><strong>��ǰ��</strong></td>
		<td>��౸��</td>
		<td>���λ���</td>
		<td>�Һ��ڰ�</td>
		<td>�����԰�</td>
		<td>������</td> 
		<td>�ǸŰ�</td>
		<td>���԰�</td>
		<td>����</td> 
		<td>������</td>
		<td>�������԰�</td>
		<td>��������</td> 
		<td>������</td> 
		<td><strong>�Ǹſ���</strong></td>
		<td><strong>��������</strong></td>
		<td><strong>��뿩��</strong></td>
		<td><strong>ī�װ�(��Ͻ�)</strong></td>
		<td><strong>ī�װ�(����)</strong></td>
	</tr>
	<% if oclear.FResultCount > 0 then %>
		<% for i = 0 to oclear.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 

			<%''��ǰ�ڵ�%>
			<td><%= oclear.FItemList(i).Fitemid %></td>

			<%''�귣��� %>
			<td><%= oclear.FItemList(i).Fmakerid %></td>
			
			<%''��ǰ�� %>
			<td><%= oclear.FItemList(i).Fitemname %></td>

			<%''��౸�� %>
			<td><%=fnColor(oclear.FItemList(i).FmwDiv,"mw") %></td>
			<%''���λ��� %>
			<td><%=fnColor(oclear.FItemList(i).Fsaleyn,"yn") %></td>
			<%''�Һ��ڰ� %>
			<td><%=FormatNumber(oclear.FItemList(i).ForgPrice,0)%></td>
				<%''�����԰� %>
			<td><%=FormatNumber(oclear.FItemList(i).ForgSuplyCash,0)%></td>
				<%''�������� %>
			<td><%=fnPercent(oclear.FItemList(i).ForgSuplyCash,oclear.FItemList(i).ForgPrice,1)%></td>
			<%''�ǸŰ� %>
			<td>
			<% 		'���ΰ�(������=(�Һ��ڰ�-���ΰ�)/�Һ��ڰ�*100) 
				if oclear.FItemList(i).Fsaleyn ="Y" then
			%>
				<font color=#F08050>(<%=CLng((oclear.FItemList(i).ForgPrice-oclear.FItemList(i).FsellCash)/oclear.FItemList(i).ForgPrice*100) %>%��)<%=FormatNumber(oclear.FItemList(i).FsellCash,0)%></font>
			<%	else
				Response.Write FormatNumber(oclear.FItemList(i).FsellCash,0)
				end if
			%>
			</td>
			<td>
				<% '���θ��԰�
					if oclear.FItemList(i).Fsaleyn ="Y" then
				%>		
					 <font color=#F08050><%=FormatNumber(oclear.FItemList(i).FsailSuplyCash,0) %></font> 
				<%	else
						Response.Write FormatNumber(oclear.FItemList(i).FbuyCash,0)
					end if %>
			</td>
			<td>
				<%
					'���θ���
					if oclear.FItemList(i).Fsaleyn ="Y"  then
						Response.Write "<font color=#F08050>" & fnPercent(oclear.FItemList(i).FsailSuplyCash,oclear.FItemList(i).FsailPrice,1) & "</font>"
					else
						Response.Write "<font color=#F08050>" & fnPercent(oclear.FItemList(i).FbuyCash,oclear.FItemList(i).FsellCash,1) & "</font>"
					end if
				%>
			</td>
			<%''������ %>
			<td>
				<%
				if oclear.FItemList(i).FitemcouponYn="Y" then
					
					Select Case oclear.FItemList(i).FitemcouponType
						Case "1" '% ����
				%>
					<font color=#5080F0>(��)<%=FormatNumber(oclear.FItemList(i).FsellCash-(CLng(oclear.FItemList(i).FsellCash*oclear.FItemList(i).FitemcouponValue/100)),0)%></font>  
				<%
						Case "2" '�� ����
				%>		
					<font color=#5080F0>(��)<%=FormatNumber(oclear.FItemList(i).FsellCash-oclear.FItemList(i).FitemcouponValue,0)%></font>
				<%			
					end Select
				end if
				%>
			</td>
			<%
				'������
				iSalePercent = (1-(clng(oclear.FItemList(i).FsellCash)/clng(oclear.FItemList(i).ForgPrice)))*100
			%> 
			</td>
			<td>
				<%
					'��������
					if  oclear.FItemList(i).FitemcouponYn="Y" then
						if oclear.FItemList(i).FitemcouponType="1" or oclear.FItemList(i).FitemcouponType="2" then
							if  oclear.FItemList(i).FitemcouponBuyPrice=0 or isNull(oclear.FItemList(i).FitemcouponBuyPrice) then
								Response.Write "<font color=#5080F0>" & FormatNumber(oclear.FItemList(i).FbuyCash,0) & "</font>"
							else
								Response.Write "<font color=#5080F0>" & FormatNumber(oclear.FItemList(i).FitemcouponBuyPrice,0) & "</font>"
							end if
						end if
					end if
				%>
			</td>
			<td>
				<%
					'��������
					if oclear.FItemList(i).FitemcouponYn="Y" then
						Select Case  oclear.FItemList(i).FitemcouponType
							Case "1"
								if oclear.FItemList(i).FitemcouponBuyPrice=0 or isNull(oclear.FItemList(i).FitemcouponBuyPrice) then
									Response.Write "<font color=#5080F0>" & fnPercent(oclear.FItemList(i).FbuyCash,oclear.FItemList(i).FsellCash-(CLng(oclear.FItemList(i).FitemcouponValue*oclear.FItemList(i).FsellCash/100)),1) & "</font>"
								else
									Response.Write "<font color=#5080F0>" & fnPercent(oclear.FItemList(i).FitemcouponBuyPrice,oclear.FItemList(i).FsellCash-(CLng(oclear.FItemList(i).FitemcouponValue*oclear.FItemList(i).FsellCash/100)),1) & "</font>"
								end if
							Case "2"
								if oclear.FItemList(i).FitemcouponBuyPrice=0 or isNull(oclear.FItemList(i).FitemcouponBuyPrice) then
									Response.Write "<font color=#5080F0>" & fnPercent(oclear.FItemList(i).FbuyCash,oclear.FItemList(i).FsellCash,1) & "</font>"
								else
									Response.Write "<font color=#5080F0>" & fnPercent(oclear.FItemList(i).FitemcouponBuyPrice,oclear.FItemList(i).FsellCash,1) & "</font>"
								end if
						end Select 
				end if
			%>
			</td> 
			<%''������ %>
			<td style="<%=chkIIF(iSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=formatnumber(iSalePercent,0)%> %</td>

			<%''�Ǹſ��� %>
			<td><%= oclear.FItemList(i).Fsellyn %></td> <% '�Ǹſ���%>
			
			<%''�������� %>
			<td><%= oclear.FItemList(i).Flimityn %></td> <% '�������� %>
			
			<%''��뿩�� %>
			<td><%=chkIIF(oclear.FItemList(i).FIsusing="Y","�����","������")%></td>

			<%''ī�װ� %>
			<td><%= oclear.FItemList(i).FdispCateName %></td>
			<td><%= oclear.FItemList(i).FdispCateNameReal %></td>
		</tr>

		<%
				if (i mod 100)=0 then Response.Flush
			next
		%>
	<% else %>	
		<tr>
			<td colspan=14 align="center">
				��� ����
			</td>
		</tr>
	<% end if %>
</table>
<% ''����Ʈ ��------------------------------------------------%>
<%
set oclear = nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->