<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : 결제요청서 등록
' History : 2011.03.14 정윤정  생성
' 0 요청/1 진행중/ 5 반려/7 승인/ 9 완료
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim yyyy1, mm1, tmpdate

yyyy1 = requestCheckvar(Request("yyyy1"),32)
mm1 = requestCheckvar(Request("mm1"),32)

if yyyy1="" then
	tmpdate = CStr(Now)

	tmpdate = DateAdd("m", -1, tmpdate)

	yyyy1 = Left(tmpdate, 4)
	mm1 = Mid(tmpdate, 6, 2)
end if

%>
<script language="javascript">

function jsRegInsertShopChulgo(frm) {
	if (confirm("일괄생성 하시겠습니까?\n\n생성에 시간이 소요됩니다.(5~10초)") == true) {
		frm.mode.value = "reginsertshopchulgo";
		frm.submit();
	}
}

function jsRegInsertUpcheShopMaeip(frm) {
	if (confirm("일괄생성 하시겠습니까?\n\n생성에 시간이 소요됩니다.(5~10초)") == true) {
		frm.mode.value = "reginsertupcheshopmaeip";
		frm.submit();
	}
}

function jsRegInsertUpcheShopWitak(frm) {
	if (confirm("일괄생성 하시겠습니까?\n\n생성에 시간이 소요됩니다.(5~10초)") == true) {
		frm.mode.value = "reginsertupcheshopwitak";
		frm.submit();
	}
}

function jsRegInsertShopWitakSell(frm) {
	if (confirm("일괄생성 하시겠습니까?\n\n생성에 시간이 소요됩니다.(5~10초)") == true) {
		frm.mode.value = "reginsertshopwitaksell";
		frm.submit();
	}
}

function jsRegInsertPartToOnline(frm) {
	if (confirm("일괄생성 하시겠습니까?") == true) {
		frm.mode.value = "reginsertparttoonline";
		frm.submit();
	}
}

function jsRegInsertPartToOffline(frm) {
	if (confirm("일괄생성 하시겠습니까?") == true) {
		frm.mode.value = "reginsertparttooffline";
		frm.submit();
	}
}

function jsRegInsertAll(frm, target) {
	if (confirm("일괄생성 하시겠습니까?") == true) {
		frm.mode.value = "reginsertall";
		frm.target.value = target;
		frm.submit();
	}
}

</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<form name="frm" method="post" action="popRegInnerOrderByMonth_process.asp">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="target" value="">
		<tr>
			<td width="100%">
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30><b>온/오프/내부부서간 내부거래 일괄생성</b></td>
				</tr>
				<tr>
					<td>
						거래월 : <% Call DrawYMBox(yyyy1, mm1) %>
						&nbsp;
						(내부부서 = 직영점 or 아이띵소 등)
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="100%">
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 01. 온라인판매(아이띵소)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '01');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 02. 온라인매입(아이띵소)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '02');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 03. 출고매입(ON상품)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '03');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 04. 기타매입(ON상품)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '04');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 05. 출고매입(OFF상품)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '05');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 06. 기타매입(OFF상품)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '06');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 07. 출고매입(위탁상품)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '07');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 08. 매장매입
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '08');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 09. 업체위탁
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '09');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 10. 기타정산
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '10');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 11. 출고매입(띵소상품)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '11');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 12. 기타매입(띵소상품)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '12');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 13. 매장판매(띵소상품)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '13');"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="35">
						* 14. 기타판매(띵소상품)
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center><input type="button" class="button" value="생성" onClick="jsRegInsertAll(frm, '14');"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="100%">
				<br>* 위 내부거래 이외의 내부거래는 "<font color=red>수익율분석>>오프라인수익서머리</font>" 에서 일괄생성됩니다.
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
