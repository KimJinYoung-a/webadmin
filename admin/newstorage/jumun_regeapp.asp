<%@  language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/tenmember/lib/header.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
Dim  iscmlinkno  ,purchasetype , ieidx, purchaseNm
dim ojumunmaster,clseapp, tContents
iscmlinkno		=  requestCheckvar(Request("iSL"),10)
purchasetype		=  requestCheckvar(Request("purchasetype"),1)

 if purchasetype="1" or  purchasetype="4" or  purchasetype="5" then '매입
 	ieidx = 65
 	purchaseNm ="<span style='color:red;'>상품사입</span>"
 elseif purchasetype="7" or purchasetype="6" or purchasetype="9" then '수입
 	ieidx = 69
 	purchaseNm ="<span style='color:red;'>상품수입</span>"
elseif purchasetype="8" or purchasetype="3" then '제작
 	ieidx = 68
 	purchaseNm ="<span style='color:red;'>상품제작</span>"
end if

set clseapp = new CEApproval
	clseapp.Fedmsidx = ieidx
	clseapp.fnGetEAppForm
	tContents		= clseapp.FedmsForm
set clseapp = nothing

set ojumunmaster = new COrderSheet
ojumunmaster.FRectIdx = iscmlinkno
ojumunmaster.GetOneOrderSheetMaster

 %>
 <meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
	<input type="hidden" name="tC" value="">
	<input type="hidden" name="ieidx" value="<%=ieidx%>">
	<input type="hidden" name="iSL" value="<%=iscmlinkno%>">
	</form>
	<div id="divEapp" style="display:none;">
	<p style="padding-bottom:30px;">다음과 같이 <%=purchaseNm%>을 진행하고자 하오니 검토 후 재가 바랍니다.</p>
	<p style="padding-bottom:30px;text-align:center;">- 다 음 -</p>
	<p style="padding-bottom:10px;"><strong>■ 내용 </strong>: <%=ojumunmaster.FOneItem.Ftargetid%>&nbsp;<%=purchaseNm%></p>
	<p><strong>■ 주문내역 </strong></p>
	<p style="padding-bottom:10px;">
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>주문코드</td>
			<td>브랜드ID</td>
			<td>주문일</td>
			<td>금액(VAT포함)</td>
			<td>입고예정일</td>
			<td>비고</td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td><%=ojumunmaster.FOneItem.Fbaljucode%></td>
		 <td><%=ojumunmaster.FOneItem.Ftargetid%></td>
		 <td><%=Left(ojumunmaster.FOneItem.Fregdate,10)%></td>
		 <td><%=formatnumber(ojumunmaster.FOneItem.Ftotalbuycash,0)%></td>
		 <td><%=ojumunmaster.FOneItem.Fscheduledate%></td>
		 <td><a href="/admin/newstorage/jumuninputedit.asp?menupos=537&idx=<%=ojumunmaster.FOneItem.Fidx%>" target="_blank">[상세]</a></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" colspan="6" height="50"><%=ojumunmaster.FOneItem.FComment%></td>
		</tr>
	</table>
	</p>

	<%=tContents%>
	</div>

 <%set ojumunmaster = nothing %>
	<script type="text/javascript">
		document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");
	 	document.frmEapp.submit();
		</script>
