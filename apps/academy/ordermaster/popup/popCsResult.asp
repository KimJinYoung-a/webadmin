<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - CS 처리결과 작성"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/ordermaster/ordercls.asp"-->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<%
dim MakerID, idx

MakerID = requestCheckVar(request.cookies("partner")("userid"),32)
idx = RequestCheckVar(request("idx"),10)

If (idx="" Or MakerID="") Then
	Response.Write "<script>alert('CS 정보가 없습니다.');fnAPPclosePopup();</script>"
	Response.End
End If

dim ioneas,i
set ioneas = new CCSASList
ioneas.FRectMakerID=MakerID
ioneas.FRectCsAsID=idx
ioneas.GetOneCSASMaster

'' 택배사 일괄적용
Sub drawSelectBoxDeliverCompanyAssign(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>택배사를 선택해 주세요</option><%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("divcd")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub
%>
</head>
<script>
<!--
function fnAppCallWinConfirm(){
	if($("#finishmemo").val()==""){
		alert("처리 내용을 입력해 주세요.");
		return false;
	}else{
		document.sfrm.action="/apps/academy/ordermaster/popup/docsresult.asp";
		document.sfrm.target="FrameCKP";
		document.sfrm.submit();
	}
}

function fnCSInputEnd(msg,OrderStateNum){
	fnAPPParentsWinReLoad();
	alert(msg);
	setTimeout(function(){
		fnAPPChangeBadgeCount("ordercount",OrderStateNum);
	}, 300);
	setTimeout(function(){
		fnAPPclosePopup();
	}, 600);
}
//-->
</script>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<form name="sfrm" method="post">
			<input type="hidden" name="orderserial" value="<%= ioneas.FOneItem.FOrderSerial %>">
			<input type="hidden" name="finishuser" value="<%=MakerID%>">
			<input type="hidden" name="id" value="<%= ioneas.FOneItem.FID %>">
			<h1 class="hidden">CS 처리결과 작성</h1>
			<div class="invoiceWrite csResult">
				<ul class="artList">
					<li>
						<textarea rows="5" name="finishmemo" id="finishmemo"><% If ioneas.FOneItem.Fcontents_finish<>"" Then %><%= ioneas.FOneItem.Fcontents_finish %><% Else %><% if ioneas.FOneItem.Fdivcd="A000" then %>
출고일 : 
기타내용 : <% elseif ioneas.FOneItem.Fdivcd="A001" then %>
출고일 : 
기타내용 : <% elseif ioneas.FOneItem.Fdivcd="A004" then %>
반품방법 : 
반품사유 : 
환불계좌 : 
기타내용 : 
<% End If %>
<% End If %>
						</textarea>
					</li>
					<li class="selectBtn">
						<% drawSelectBoxDeliverCompanyAssign "songjangdiv", ioneas.FOneItem.FSongjangdiv %>
					</li>
					<li class="list">
						<dfn><b>운송장 번호</b></dfn>
						<div><input type="text" name="songjangno" placeholder="운송장 번호를 입력해주세요" value="<%= ioneas.FOneItem.Fsongjangno %>" size="14" maxlength="14" /></div>
					</li>
				</ul>
				<div class="csHelp">
					<div class="boxUnit bdrTRtGry">
						<% if ioneas.FOneItem.Fdivcd="A000" then %> <!-- 맞교환 설명 -->
							<p class="tPad1r">*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.<br />(고객님께 오픈되는 정보가 아닙니다.)</p>
							<p class="tPad1r">맞교환상품 출고후, 택배정보를 꼭 입력 부탁드립니다.</p>
							<p class="tPad1r">- 출고일 :<br />- 기타내용 :</p>
							<p class="tPad1r">위 내용을 작성하여 처리내용에 남겨주시면 감사하겠습니다.</p>
						<% elseif ioneas.FOneItem.Fdivcd="A001" then %> <!-- 누락재발송 설명 -->
							<p class="tPad1r">*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.<br />(고객님께 오픈되는 정보가 아닙니다.)</p>
							<p class="tPad1r">맞교환상품 출고후, 택배정보를 꼭 입력 부탁드립니다.</p>
							<p class="tPad1r">- 출고일 :<br />- 기타내용 :</p>
							<p class="tPad1r">위 내용을 작성하여 처리내용에 남겨주시면 감사하겠습니다.</p>
						<% elseif ioneas.FOneItem.Fdivcd="A004" then %> <!-- 반품 설명 -->
							<p class="tPad1r">*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.<br />(고객님께 오픈되는 정보가 아닙니다.)</p>
							<p class="tPad1r">반품상품 입고 완료 후, 처리내용 입력과 함께<br />완료처리 부탁드립니다.</p>
							<p class="tPad1r">- 반품방법 : 고객선불 / 착불<br />- 반품사유 : 불량반품 / 고객반품<br />- 계좌 : 은행명 + 계좌번호 + 예금주명(고객님이 첨부한 경우)<br />- 기타내용 :</p>
							<p class="tPad1r">위 내용을 작성하여 처리내용에 남겨주시면 감사하겠습니다.</p>
						<% elseif ioneas.FOneItem.Fdivcd="A006" then %> <!-- 출고시 유의사항 설명 -->
							<p class="tPad1r">*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.<br />(고객님께 오픈되는 정보가 아닙니다.)</p>
							<p class="tPad1r">고객센터에서 요청한 출고유의사항에 대한 처리유무를 알려주시기 바랍니다.</p>
							<p class="tPad1r">발송 후, 이 내용을 확인하셨을 경우에도, 미반영 출고로 완료처리 부탁드립니다.</p>
						<% else %>
							<p>반품상품 입고 완료 후, 입력 부탁드립니다.</p>
							<p class="tPad1r">* 처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.<br />(고객님께 오픈되는 정보가 아닙니다.)</p>
							<p class="tPad1r">- 반품방법 : 고객선불 / 착불<br />- 반품사유 : 불량반품 / 고객반품<br />- 계좌 : 은행명 + 계좌번호 + 예금주명</p>
							<p class="tPad1r">위 내용을 작성하여 처리내용에 남겨주시면 감사하겠습니다.</p>
						<% end if %>
					</div>
				</div>
			</div>
			</form>
		</div>
		<!--// content -->
	</div>
</div>
</body>
</html>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
fnAPPShowRightConfirmBtns();
});
//-->
</script>
<%
Set ioneas = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->