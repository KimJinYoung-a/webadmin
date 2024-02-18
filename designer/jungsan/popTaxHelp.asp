<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- 사용안함 헤더에 포함할 예정 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td colspan="2">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="400" style="padding:5; border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" background="/images/menubar_1px.gif">
						<font color="#333333"><b>세금계산서 발행방법 안내</></font>
					</td>
					<td align="right" style="border-bottom:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
						&nbsp;
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>;border-top:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
	    방법1. 더존 위하고 (구 bill36524) 연동 발행
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td width="20" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">&nbsp;</td>
		<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
			<img src="/images/icon_num01.gif" border="0"> <b><font color="red">공인인증서 준비</font></b><br>
				사업자범용 인증서 혹은 전자세금계산서용 인증서를 준비하시면 됩니다.<br>
				1.사업자범용 인증서 : 모든 전자상거래에 이용가능한 인증서이며, 발행수수료는 <b>110,000원/년</b> 입니다.<br>
				2.세금계산서용 인증서 : 거래은행 방문하셔서 <b>[세금계산서용 인증서]</b>를 발행 받으실 수 있습니다. 발생수수료 발행수수료는 <b>약4,400원/년</b> 입니다.<br>
				3.위하고 전용 인증서 : 위하고 및 국세청 e세로에서만 이용가능한 인증서이며, 발행수수료는 <b>11,000원/년</b> 입니다.<br>
				
				
				<font color="purple">*기존에 범용인증서를 가지고 계신 사업자는 그 인증서를 사용하여 회원가입 및 세금계산서 발행이 가능합니다.</font><br>
				<font color="purple">*범용인증서가 없으신분은 가능한 거래은행 방문하셔서 [세금계산서용 인증서]를 발행후 사용하시기 바랍니다.</font><br>
				* 위하고 전용 인증서는 <strong>추천 하지 않습니다.</strong>(발행 비용 비싸며, 발행 기간 오래 걸림. 타 전자계산서 업체에서 사용불가)<br>
				<br>
			
			
			<img src="/images/icon_num02.gif" border="0"> <b><font color="red">위하고 회원가입</font></b><br>
				공인인증서가 준비되셨으면, 위하고 에 회원가입을 하시면 됩니다.(<a href="https://www.wehago.com" target="_blank">https://www.wehago.com</a>)<br>
				회원가입시에는 <b>사업자(법인/개인)회원</b>으로 가입하시면 됩니다.<br>
				<font color="purple">회원가입시에 공인인증서로 사업자 확인을 진행하므로, 공인인증서를 먼저 준비하시기 바랍니다.</font><br>
				<br>
				
			
			<img src="/images/icon_num03.gif" border="0"> <b><font color="red">로그인 후, 인증서 등록</font></b><br>
				회원가입 완료 후, 로그인 하시면 왼쪽 세로메뉴에 <b>[사용자환경설정]</b>이라는 아이콘이 있습니다.<br>
				[사용자환경설정]에서 4번째 항목에 있는 <b>인증서 등록</b>을 해주시기 바랍니다.<br>
				<font color="purple">인증서 등록이 안되어 있을 경우, 텐바이텐SCM과의 연동발행이 되지 않습니다.</font><br>
				<br>
				
			
			<img src="/images/icon_num04.gif" border="0"> <b><font color="red">발행수수료 포인트 충전</font></b><br>
				로그인 정보가 표시되는 오른쪽 상단에 보시면, [충전]버튼이 있습니다.<br>
				전자세금계산서의 경우, 공급자가 발행수수료를 지불하게 됩니다. 건당 200원의 발행수수료가 부과됩니다.<br>
				예를들어, 1만원을 충전하시면, 50건의 세금계산서 발행이 가능합니다.<br>
				<br>
				
			
			<img src="/images/icon_num05.gif" border="0"> <b><font color="red">텐바이텐SCM에서 세금계산서 발행</font></b><br>
				위 4가지 사항이 모두 준비되었다면, 텐바이텐SCM(<a href="https://scm.10x10.co.kr" target="_blank">scm.10x10.co.kr</a>) 에서 세금계산서를 발행하시면 됩니다.<br>
				<br>

		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>;border-top:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
	    방법2. 국세청 이세로 또는 자체 발행업체이용
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td width="20" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">&nbsp;</td>
		<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
		    1. 이세로 또는 타 발행업체를 이용할경우도 인증서는 필요합니다. 사업자범용인증서 또는 세금계산서용 인증서를 발급받으세요. <br><br>
		    2. 공급받는자 정보 (텐바이텐 사업자등록증 <a href="http://scm.10x10.co.kr/images/10x10lic.jpg" target="_blank"><font color="blue">[보기]</font></a>) <br><br>
		    3. 총 정산액 = 공급액+세액 = 합계금액 (총 정산액과 계산서 발행금액 합계가 일치해야 정산확정됩니다.)<br><br>
		    4. 작성일자(발행일) : <b>해당 정산월 말일</b>(ex] 2013년1월정산 : 2013-01-31), (발행 기한이 지난경우 발행하시는달 1일 ex] 2013년1월정산 : <%= LEFT(now(),7)&"-01" %> (금일기준))<br><br>
		    5. 발행해주실 이메일 주소 : etax@10x10.co.kr<br><br>
			6. 발행후 그달 14일까지 정산상태가 <font color="blue">정산 확정</font>으로 되어 있지 않은경우 이메일 수신을 못하였을수 있으니 전화 주시기 바랍니다.<br>
			
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>;border-top:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
	    자주 문의 하시는 내용
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td width="20" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">&nbsp;</td>
		<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
		    Q : 꼭 위하고 로 연동발행해야 하나요? <br>
		    &nbsp;&nbsp; A : 이세로 또는 자체 연동 프로그램으로 발행 해 주셔도 됩니다. 다만 연동발행이 아닌 경우 텐바이텐에서 수기로 전자 계산서를 확인해야 되므로 발행 후 정산확정까지 2~5일 소요됩니다. <br><br>
		    
		    Q : 수수료에 대한 매입세금계산서를 발행 해 주나요? <br>
		    &nbsp;&nbsp; A : 총 매출에서 수수료 제한 금액을 텐바이텐에 공급하는 방식을 취하고 있습니다. 수수료에 대한 계산서는 따로 발행하지 않습니다.<br><br>
		    
		    Q : 면세 사업자인데 꼭 전자계산서로 발행 해야 되나요? <br>
		    &nbsp;&nbsp; A : 면세인 경우 종이 계산서 발행 가능합니다. 캡쳐해서 이메일(etax@10x10.co.kr) 로 먼저 보내주신후 원본은 우편으로 보내주세요.<br><br>
		    
		    Q : 위하고 로 발행 후 출력버튼이 나오지 않습니다.<br>
		    &nbsp;&nbsp; A : 면세인경우 텐바이텐에서 승인후 출력 가능하며, 과세인경우 국세청 전송후(익일) 출력 가능합니다.<br><br>
		    
		    Q : 결제일은 언제인가요?<br>
		    &nbsp;&nbsp; A : 업체정보 수정 > 브랜드 계약관련 정보 정산일에 보시면 나와 있습니다(보통 익월 말일). <!-- 이월 발행하신경우 작성일자 익월 15일날 결제됩니다. -->정산일이 토/일요일,공휴일인 경우, 익일(이후 첫영업일)에 결제됩니다. 
		    <br><br>
		    
		    Q : 위하고 연동발행시 오류<br>
		    &nbsp;&nbsp; 1. 정산 담당자 핸드폰 번호가 올바르지 않습니다. 업체정보수정에서 정산담당자 핸드폰을 000-000-0000 대시 형태로 수정후 사용하세요.<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; 업체정보 수정 > 정산담당자 핸드폰, 이메일 정보를 기입해 주시기 바랍니다.<br><br>
		    
		    &nbsp;&nbsp; 2. 위하고 사이트에 가입된 사업자번호와 텐바이텐에 등록된 사업자번호가 일치하지 않습니다. 위하고 에 등록된 사업자번호: 000-00-00000 텐바이텐에 등록된 사업자번호:XXX-XX-XXXXX<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; 텐바이텐에 등록된 사업자 번호와 위하고 에 등록된 사업자 번호가 일치해야 연동 발행 됩니다. 사업자번호가 변경된경우 사업자변경 신청후 가능합니다.(사업자 변경 신청은 담당엠디에게 신청하세요. 사업자등록증사본, 통장사본)<br><br>
		    
		    &nbsp;&nbsp; 3. API 기발행 세금계산서<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; 통신중 오류등의 이유로 정상 발행 되었으나, 정산 확정상태로 변경되지 않은경우 입니다. 1~2일후 정산확정상태로 변경되지 않으셨으면 연락바랍니다.  <br><br>
		    
		    &nbsp;&nbsp; 4. 위하고 에서 사용자환경설정 => 인증서 등록에서 인증서 등록후 사용하시기 바랍니다.<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; 인증서가 만료 되었거나, 인증서 등록이 되지 않은경우 발생하는 메세지입니다. 위하고 로그인후 왼쪽 세로메뉴의 사용자환경설정 버튼 클릭후 인증서 탭에서 인증서 등록후 다시 시도해 주세요. <br><br>
		     <img src="/images/Snap_bill_set1.jpg" width="560">
		     <img src="/images/Snap_bill_set2.jpg" width="560">
		    <br><br>
		    &nbsp;&nbsp; 5. 포인트가 부족합니다.<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; 위하고 로그인 후 상단 [충전] 클릭 후 과금후 사용하시기 바랍니다. 건당 200원 발행비용 발생 <br><br>
		    <img src="/images/Snap_bill_charge1.jpg" width="560">
		    <img src="/images/Snap_bill_charge2.jpg" width="560">
		    <br>
		    
		</td>
	</tr>
</table>
<!-- 사용안함 헤더에 포함 -->
<!-- #include virtual="/designer/lib/poptail.asp"-->