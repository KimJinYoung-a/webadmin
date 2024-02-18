<%
Dim page : page = 1
Dim oEval , iTotCnt

set oEval = new CEvaluateSearcher
	oEval.FPageSize = 3
	oEval.FCurrpage = page
	oEval.FRectItemID = itemid
	
	'상품 후기가 있을때만 쿼리.
	if isEvaluateCnt > 0 then
		oEval.getItemEvalList()
		iTotCnt = oEval.FResultCount
	end If

	'//구매 상품인지 확인
	Dim cEval, vArrEval, intLoop
	If IsUserLoginOK() Then
		Set cEval = New CEvaluateSearcher
		cEval.FRectUserID = getEncLoginUserID
		cEval.FRectItemID = itemid
		vArrEval = cEval.getMyEvalListLightVer
		Set cEval = Nothing
	End If
%>
<script type="text/javascript">
	//상품 후기
	function diyReviewWrite(mode,itemid,idx,orderserial,optionCD){
		<% If LoginUserid <> "" Then %>
		location.href = "/diyshop/diyreviewwrite.asp?mode="+mode+"&itemid="+itemid+"&idx="+idx+"&orderserial="+orderserial+"&optionCD="+optionCD;
		<% Else %>
		jsChklogin_mobile('','<%= Server.URLEncode(CurrURLQ()) %>');
		<% End If %>
	}

	//리스트 ajax
	function rwonlyNumber(event,itemid){
		var val = $("#rwtextpage").val();
		if(val > <%= CInt(oEval.FtotalPage) %>){
			val=<%= CInt(oEval.FtotalPage) %>;
		}else if(val < 1){
			val=1;
		}

		event = event || window.event;
		var keyID = (event.which) ? event.which : event.keyCode;
		if( (keyID >= 48 && keyID <= 57) || (keyID >= 96 && keyID <= 105) || keyID == 8 || keyID == 46 || keyID == 37 || keyID == 39 ){
			return;
		}else if(keyID == 13){
			$.ajax({
				url : "/diyshop/inc/ajax_shop_prd_tabs_review.asp?cpg="+val+"&itemid="+itemid,
				dataType : "html",
				type : "get",
				success : function(result){
					$("#reviewList").empty().html(result);
				}
			});
		}else{
			return false;
		}
	}

	//상품 review 페이징
	function rwgopage(page,bfgubun,itemid){
		var vPg = "1";
		var urlgubun
		if (bfgubun=="b"){
			if(page > 1){
				vPg = page-1;
			}else{
				alert('이전 페이지가 없습니다');
				return;
			}
		}else if(bfgubun=="f"){
			if(page < <%= CInt(oEval.FtotalPage)+1 %>){
				vPg = page++;
			}else{
				alert('다음 페이지가 없습니다');
				return;
			}
		}else{
			alert("aa");
			//parent.location.reload();
		}
		$.ajax({
			url: "/diyshop/inc/ajax_shop_prd_tabs_review.asp?cpg="+vPg+"&itemid="+itemid,
			dataType : "html",
			type : "get",
			success : function(result){
				$("#reviewList").empty().html(result);
			}
		});
	}

	//qna 전체보기,삭제
	function diyReviewDel(mode,itemid,idx,orderserial,optionCD){
		if(confirm("삭제하시겠습니까?")){
			$.ajax({
				url: "/diyshop/inc/ajax_shop_prd_tabs_review.asp?mode="+mode+"&itemid="+itemid+"&idx="+idx+"&orderserial="+orderserial+"&optionCD="+optionCD,
				dataType : "html",
				type : "get",
				success : function(result){
					$("#reviewList").empty().html(result);
				}
			});
		}
	}
</script>
	<% if fnMyEvalCheck( vArreval,itemid ) then %>
		<div class="btnGroup">
			<a href="" onclick="diyReviewWrite('add','<%=itemid%>','','<%= fnMyEvalorderserial( vArreval,itemid ) %>','<%= fnMyEvalitemoption( vArreval,itemid ) %>');return false;" class="btn btnB1 btnYgn">후기 쓰기</a>
		</div>
	<% end if %>
	<div class="reviewList" id="reviewList">
		<ul>
			<%
				IF oEval.FResultCount>0 then
					FOR i =0 to oEval.FResultCount-1
			%>
			<li>
				<div class="reviewCont">
					<div class="star score0<%= oEval.FItemList(i).FTotalPoint %>"><span></span></div>
					<p class="txt"><%= nl2br(oEval.FItemList(i).FUesdContents) %></p>
					<div class="reviewImg">
						<% if oEval.FItemList(i).Flinkimg1 <> "" then %>
						<img name="image_fix_1" id="image_fix_1" src="<%= oEval.FItemList(i).getLinkImage1 %>"/><br>
						<% End if %>
						<% if oEval.FItemList(i).Flinkimg2 <>"" then %>
							<img name="image_fix_2" id="image_fix_2" src="<%= oEval.FItemList(i).getLinkImage2 %>"/>
						<% End if %>
					</div>
					<p class="txtInfo"><span><%= printUserId(oEval.FItemList(i).FUserID,2,"*") %></span><span><%= FormatDate(oEval.FItemList(i).FRegdate, "0000.00.00") %></span></p>
					<div class="btnGroup">
						<!--<button type="button" class="btn btnM2 btnWht">신고</button>-->
						<% If LoginUserid = oEval.FItemList(i).FUserID Then %>
						<button type="button" class="btn btnM2 btnWht" onclick="diyReviewWrite('edit','<%=itemid%>','<%=oEval.FItemList(i).Fidx%>','<%=oEval.FItemList(i).Forderserial%>','<%=oEval.FItemList(i).Fitemoption%>');return false;">수정</button>
						<button type="button" class="btn btnM2 btnWht" onclick="diyReviewDel('del','<%=itemid%>','<%=oEval.FItemList(i).Fidx%>','','');return false;">삭제</button>
						<% End If %>
					</div>
				</div>
			</li>
			<% Next %>
			<% Else %>
			<li class="noData"><span>등록된 상품후기가 없습니다.</span></li>
			<% End If %>			
		</ul>
		<%' pagination %>
		<% IF oEval.FResultCount>0 Then %>
		<div class="pagination">
			<a href="" onclick="rwgopage('<%= oEval.FCurrPage %>','b','<%= itemid %>');return false;" class="btnPrev"><span>이전 페이지</span></a>
			<span><input type="text" id="rwtextpage" class="pageNum" maxlength = "4" onkeydown="rwonlyNumber(event,'<%= itemid %>')" style='ime-mode:disabled;' value="<%= CInt(oEval.FCurrPage) %>" /> / <%= CInt(oEval.FtotalPage) %></span>
			<a href="" onclick="rwgopage('<%= oEval.FCurrPage+1 %>','f','<%= itemid %>');return false;" class="btnNext"><span>다음 페이지</span></a>
		</div>
		<% End If %>
		<%' pagination %>
	</div>
<% set oEval = nothing %>