
<!-- #include virtual="/apps/academy/preview/diyshopCls.asp" -->
<%
''		http://testm.thefingers.co.kr/diyshop/shop_prd_tabs_qna.asp?mode=add&itemid=1008

''	i, itemid,
Dim oDiyItemQnAList
Dim PageSize, CurrPage
'Dim  loginuserid
	itemid  = requestCheckVar(request("itemid"),10)	''매거진 idx
	PageSize = getNumeric(requestCheckVar(request("psz"),9))
	CurrPage = getNumeric(requestCheckVar(request("cpg"),9))

IF itemid = "" THEN
	Response.Write "<script language='javascript'>alert('잘못된 경로입니다.3');</script>"
	Response.Write "<script language='javascript'>location.href = '/diyshop/diyList.asp?itemid="&itemid&"';</script>"
	response.end
END IF
IF IsNumeric(itemid) = False THEN
	Response.Write "<script language='javascript'>alert('잘못된 경로입니다.4');</script>"
	Response.Write "<script language='javascript'>location.href = '/diyshop/diyList.asp?itemid="&itemid&"';</script>"
END If

	if CurrPage="" or CurrPage=0 then CurrPage=1
	if PageSize ="" then PageSize =12

	loginuserid = GetencLoginUserID()

	'상품 QnA 리스트
	set oDiyItemQnAList = new DiyItemCls
		oDiyItemQnAList.FPageSize = PageSize
		oDiyItemQnAList.FCurrPage = CurrPage
		oDiyItemQnAList.FRectuserid = loginuserid
		oDiyItemQnAList.FRectitemid = itemid				'상품코드
		oDiyItemQnAList.FRectmode = "list"
		oDiyItemQnAList.GetDiyQnaList()

%>
<script type="text/javascript">

//새 QNA 쓰기
function diyQnaWrite(md,itemid,ridx,mkid,rowidx,qnagb){
	location.href = "/diyshop/diyQnaWrite.asp?mode="+md+"&itemid="+itemid+"&ridx="+ridx+"&makerid="+mkid+"&rowidx="+rowidx+"&qnagb="+qnagb;
}

//코멘트 삭제
//function MagazineCommentdel(md,itemid,ridx,cidx,dp){
//	if(confirm("삭제하시겠습니까?")){
//		document.delfrm.idx.value = cidx;
//		document.delfrm.itemid.value = itemid;
//		document.delfrm.ridx.value = ridx;
//		document.delfrm.depth.value = dp;
//   		document.delfrm.submit();
//	}
//}

function onlyNumber(event,aaa,itemid){
	var val = $("#textpage").val();
	if(val > <%= CInt(oDiyItemQnAList.FtotalPage) %>){
		val=<%= CInt(oDiyItemQnAList.FtotalPage) %>;
	}else if(val < 1){
		val=1;
	}

	event = event || window.event;
	var keyID = (event.which) ? event.which : event.keyCode;
	if( (keyID >= 48 && keyID <= 57) || (keyID >= 96 && keyID <= 105) || keyID == 8 || keyID == 46 || keyID == 37 || keyID == 39 ){
		return;
	}else if(keyID == 13){
		$.ajax({
		    url : "/diyshop/inc/ajax_shop_prd_tabs_QnA.asp?cpg="+val+"&itemid="+itemid,
		    dataType : "html",
		    type : "get",
		    success : function(result){
		        $("#tab03").empty().html(result);
		    }
		});
	}else{
		return false;
	}
}

//상품 QNA 페이징
function fngopage(page,bfgubun,itemid,mkid){
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
		if(page < <%= CInt(oDiyItemQnAList.FtotalPage)+1 %>){
			vPg = page++;
		}else{
			alert('다음 페이지가 없습니다');
			return;
		}
	}else{
		parent.location.reload();
	}
	$.ajax({
		url: "/diyshop/inc/ajax_shop_prd_tabs_QnA.asp?cpg="+vPg+"&itemid="+itemid+"&makerid="+mkid,
		dataType : "html",
		type : "get",
		success : function(result){
		    $("#tab03").empty().html(result);
		}
	});
}

//qna 전체보기,삭제
function fnlistdel(md,ridx,itemid,mkid,idx){
	if (md=="del"){
		if(confirm("삭제하시겠습니까?")){
			$.ajax({
				url: "/diyshop/inc/ajax_shop_prd_tabs_QnA.asp?mode="+md+"&ridx="+ridx+"&itemid="+itemid+"&makerid="+mkid,
				dataType : "html",
				type : "get",
				success : function(result){
				    $("#tab03").empty().html(result);
				}
			});
		}
	}else if (md=="adel"){
		if(confirm("삭제하시겠습니까??")){
			$.ajax({
				url: "/diyshop/inc/ajax_shop_prd_tabs_QnA_detail.asp?mode="+md+"&itemid="+itemid+"&idx="+idx+"&makerid="+mkid,
				dataType : "html",
				type : "get",
				success : function(result){
				    $("#tab03").empty().html(result);
				}
			});
		}
	}else{
		$.ajax({
			url: "/diyshop/inc/ajax_shop_prd_tabs_QnA.asp?mode="+md+"&ridx="+ridx+"&itemid="+itemid+"&makerid="+mkid,
			dataType : "html",
			type : "get",
			success : function(result){
			    $("#tab03").empty().html(result);
			}
		});
	}
}

//상품 QNA 상세
function fnQnaDetail(itemid,gridx,mkid){
	var vPg = "1";

	$.ajax({
		url: "/diyshop/inc/ajax_shop_prd_tabs_QnA_detail.asp?cpg="+vPg+"&itemid="+itemid+"&gridx="+gridx+"&makerid="+mkid,
		dataType : "html",
		type : "get",
		success : function(result){
		    $("#tab03").empty().html(result);
		}
	});
}
</script>
<!--
					FItemList(i).Fidx				= rsget("idx")
					FItemList(i).Fitemid				= rsget("itemid")
					FItemList(i).Fuserlevel			= rsget("userlevel")
					FItemList(i).Freplyuserid		= rsget("replyuserid")
					FItemList(i).Ftitle				= rsget("title")
					FItemList(i).Fdevice				= rsget("device")
					FItemList(i).Fuserid				= rsget("userid")	
					FItemList(i).Fcomment			= rsget("comment")
					FItemList(i).Fisusing			= rsget("isusing")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).Freply_num			= rsget("reply_num")			'리플 순서
					FItemList(i).Freply_depth		= rsget("reply_depth")		'원글0,리플1 뎁스
					FItemList(i).Freply_group_idx	= rsget("reply_group_idx")	'리플 그룹 idx
-->
		<!-- Q&A -->
			<div class="btnGroup bMar1-5r">
				<a href="" onclick="diyQnaWrite('add','<%= itemid %>','','<%= oItem.Prd.FMakerid %>','','Q'); return false;" class="btn btnB1 btnYgn">문의하기</a>
			</div>

			<!-- 전체 질문 리스트 -->
			<div class="fingerQna qnaList">
			<% if oDiyItemQnAList.FresultCount > 0 then %>
				<ul>
					<% for i=0 to oDiyItemQnAList.FresultCount-1 %>
						<!-- for dev msg : 내가 쓴 글 일경우 클래스 myQ 넣어주세요 -->
						<li <%=chkIIF(loginuserid = oDiyItemQnAList.FItemList(i).Fuserid,"class='myQ'","")%>>
							<a href="" onclick="fnQnaDetail('<%= itemid %>','<%= oDiyItemQnAList.FItemList(i).Freply_group_idx %>','<%= oItem.Prd.FMakerid %>'); return false;" >
								<div class="titleCont">
									<div class="qnaProcess">
										<% if oDiyItemQnAList.FItemList(i).FanswerYN = "Y" then %>
											<span class="finish">답변완료</span>
										<% else %>
											<span class="ing">답변중</span>
										<% end if %>

										<% if loginuserid = oDiyItemQnAList.FItemList(i).Fuserid then %>
											<span class="my">나의 문의글</span>
										<% end if %>
									</div>
									<p class="title"><%= oDiyItemQnAList.FItemList(i).Ftitle %></p>
									<p class="txtInfo"><span><%= oDiyItemQnAList.FItemList(i).Fuserid %></span><span><%= FormatDate(oDiyItemQnAList.FItemList(i).FRegdate,"0000.00.00") %></span></p>
								</div>
							</a>
						</li>
					<% next %>
				</ul>
				<!-- pagination -->
				<div class="pagination">
					<a href="" onclick="fngopage('<%= oDiyItemQnAList.FCurrPage %>','b','<%= itemid %>','<%= oItem.Prd.FMakerid %>'); return false;" class="btnPrev"><span>이전 페이지</span></a>
					<span><input type="text" id="textpage" class="pageNum" maxlength = "4" onkeydown="return onlyNumber(event,'','<%= itemid %>')" style='ime-mode:disabled;' value="<%= CInt(oDiyItemQnAList.FCurrPage) %>" /> / <%= CInt(oDiyItemQnAList.FtotalPage) %></span>
					<a href="" onclick="fngopage('<%= oDiyItemQnAList.FCurrPage+1 %>','f','<%= itemid %>','<%= oItem.Prd.FMakerid %>'); return false;" class="btnNext"><span>다음 페이지</span></a>
				</div>
				<!--// pagination -->
			<% else %>
				<ul>
					<li class="noData"><span>등록된 질문글이 없습니다.</span></li>
				</ul>
			<% end if %>
			</div>
			<!--// 전체 질문 리스트 -->
		<!--// Q&A -->
<% set oDiyItemQnAList = nothing %>