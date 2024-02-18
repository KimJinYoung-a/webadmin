<%
'#############################################
' PageName : /diyshop/shop_prd_tabs.asp	
' Description : DIY Shop_tabs 상품상세
' History : 2016.07.11 이종화 생성
'#############################################
%>
<div id="detailView" class="detailView fingerTab">
	<ul class="tab1 tabNav">
		<li class="current" name="tab01" id="mtab01">작품정보</li>
		<li name="tab02" id="mtab02">구매후기(<%=isEvaluateCnt%>)</li>
		<li name="tab03" id="mtab03">Q&amp;A(<%=isQnACnt%>)</li>
	</ul>
	<div class="tabContainer">
		<%' <!-- 작품정보 --> %>
		<div id="tab01" class="tabCont itemInfo">
			<div class="box1">
				<%' <!-- 작품 설명 입력 영역 --> %>
				<div class="detailCont">
					<% FOR i= 0 to oAdd.FResultCount-1  %>
						<% IF oAdd.FADD(i).FAddImageType=2 THEN %>
						<div class="image"><img src="<%= oAdd.FADD(i).FAddimage %>" alt="<%= oItem.Prd.FItemName %>" /></div>
						<% If oAdd.FADD(i).FAddimgText <> "" Then %>
						<div class="txt"><%=nl2br(oAdd.FADD(i).FAddimgText)%></div>
						<% End If %>
						<% End IF %>
					<% NEXT %>
				</div>
				<%' <!--// 작품 설명 입력 영역 --> %>

				<%' <!-- 기타 정보 --> %>
				<div class="restInfo">
					<dl class="viewPop btnDelivery">
						<dt>배송비 안내</dt>
						<dd><span><% = oItem.Prd.GetDeliveryName %></span><button type="button" class="openInfo btnDelivery"><span>배송비 정보 더보기</span></button></dd>
					</dl>
					<% if (oItem.Prd.FItemDiv = "06") then %>
					<dl class="viewPop btnCustom">
						<dt>제작기간</dt>
						<dd><span><%=oItem.Prd.Frequiremakeday%>일 이내</span><button type="button" class="openInfo btnCustom"><span>제작기간 정보 더보기</span></button></dd>
					</dl>
					<% End If %>
					<dl>
						<dt>상품 필수 정보</dt>
						<dd>
							<p>전자상거래 상품정보 제공 고시에 따라 작성 되었습니다.</p>
							<ul>
								<%
								IF addEx.FResultCount > 0 THEN
									FOR i= 0 to addEx.FResultCount-1
										If addEx.FItemList(i).FinfoCode = "35005" Then
											If tempsource <> "" then
											response.write "<li><p><span>재질 :</span>"&tempsource&"</p></li>"
											End If
											If tempsize <> "" then
											response.write "<li><p><span>사이즈 :</span>"&tempsize&"</p></li>"
											End If
										End If
								%>
									<li style="display:<%=chkiif(addEx.FItemList(i).FInfoContent="" And addEx.FItemList(i).FinfoCode ="02004" ,"none","")%>;"><p><span><%=addEx.FItemList(i).FInfoname%> :</span> <%=addEx.FItemList(i).FInfoContent%></p></li>
								<%
										Next
									End If
								%>
							</ul>
						</dd>
					</dl>
					<dl>
						<dt>교환/환불 정책</dt>
						<dd>
							<%=db2html(nl2br(oItem.Prd.Frefundpolicy))%>
						</dd>
					</dl>
				</div>
				<%' <!--// 기타 정보 --> %>

				<%' <!-- 작품 키워드 --> %>
				<%
					if Not(oItem.Prd.FKeyWords="" or isNull(oItem.Prd.FKeyWords)) then
						Dim ArrTag
						ArrTag = Split(oItem.Prd.FKeyWords,",")
				%>
				<div class="keyword">
					<h2>작품 키워드</h2>
					<div class="tag">
						<%
							For lp=0 to Ubound(ArrTag)
								if Trim(ArrTag(lp))<>"" then
									Response.Write "<a href='javascript:void(0);'>#" & Trim(ArrTag(lp)) & "</a>"
									'if lp<Ubound(ArrTag) then Response.Write ", "
								end if
							Next
						%>
					</div>
				</div>
				<% end if %>
				<%' <!--// 작품 키워드 --> %>
			</div>
		</div>
		<%' 작품정보 %>
	</div>
</div>