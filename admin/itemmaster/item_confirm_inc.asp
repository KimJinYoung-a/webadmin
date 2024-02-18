<script type="text/javascript">
	function jsChkValue(currstate){ 
		if(currstate==0){
		document.all.chkV0[6].checked = true;
	}else{ 
		document.all.chkV2[36].checked = true;
	}
	}
</script> 
<div id="dv2" style="display:none;"><!-- currstate=2, 승인보류 사유---------------------------->
<table border=0 cellpadding=5 cellspacing=5   class="a" width="650">  
	<tr>
		<td><b>승인 거부 사유:</b> <font color="#FF0000">승인보류(재등록요청)</font></td>
		<td align="right"><!--초기화--></td>
	</tr>
	<tr>
		<td colspan="2">
			<table border=0 cellpadding=5 cellspacing=1 bgcolor="<%=adminColor("tablebg")%>" class="a" width="100%">
				<tr bgcolor="#FFFFFF">
					<td>
							<table border=0 cellpadding=3 cellspacing=0  class="a">
								<tr>
									<td><b>이미지 정보 오류</b></td>
								</tr>
								<tr>
									<td>
									    <table border=0 cellpadding=3 cellspacing=0  class="a">
									    <tr>
									        <td><input type="checkbox" name="chkV2" value="1"><span id="sp21">기본이미지불량</span></td>
										    <td><input type="checkbox" name="chkV2" value="2"><span id="sp22">기본이미지텍스트삭제</span></td>
										    <td><input type="checkbox" name="chkV2" value="3"><span id="sp23">기본이미지배경불량(흰바탕)</span></td>
								        </tr>
								        <tr>
        									<td><input type="checkbox" name="chkV2" value="4"><span id="sp24">기본이미지사이즈 불량</span></td>
        									<td><input type="checkbox" name="chkV2" value="5"><span id="sp25">선명한 이미지로 수정</span></td>    
        									<td><input type="checkbox" name="chkV2" value="6"><span id="sp26">상세페이지불량</span> </td>
        								</tr>
        								<tr>
        								    <td><input type="checkbox" name="chkV2" value="7"><span id="sp27">선명한 상세페이지로 수정</span></td>
        									<td><input type="checkbox" name="chkV2" value="16"><span id="sp216">추가이미지 등록요망</span></td>
        									<td><input type="checkbox" name="chkV2" value="17"><span id="sp217">기본이미지 추가등록(다른 이미지로 2개이상등록)</span> </td> 
        								</tr>
        							    </table>
        							</td>
        						</tr>	
								<tr>
									<td><hr width="100%"></td>
								</tr>
								<tr>
									<td><b>상품 정보 오류</b></td>
								</tr>
								<tr>
									<td>
									    <table border=0 cellpadding=3 cellspacing=0  class="a"> 
									    <tr>
									        <td><input type="checkbox" name="chkV2" value="10"><span id="sp210">품목상세정보누락</span></td>
										    <td><input type="checkbox" name="chkV2" value="11"><span id="sp211">안전인증정보누락</span></td>
										   
								        </tr>
								        <tr>
								             <td><input type="checkbox" name="chkV2" value="12"><span id="sp212">상품명수정(사용불가단어 사용)</span></td>  
								            <td><input type="checkbox" name="chkV2" value="13"><span id="sp213">상세페이지 자사몰 주소삭제</span></td>
										   
										 </tr>
										 <tr>   
										     <td><input type="checkbox" name="chkV2" value="18"><span id="sp218">기본수수료 확인</span></td>
										    <td><input type="checkbox" name="chkV2" value="19"><span id="sp219">판매가 등록</span></td> 
								        </tr>
								        <tr>
									        <td><input type="checkbox" name="chkV2" value="8"><span id="sp28">전시카테고리누락</span></td>
										    <td><input type="checkbox" name="chkV2" value="9"><span id="sp29">전시카테고리수정</span></td>   
										 </tR>
										 <tr>   
										    <td><input type="checkbox" name="chkV2" value="15"><span id="sp215">관리카테고리, 전시카테고리 동일 카테고리로 등록 요망</span></td>     
										    <Td><input type="checkbox" name="chkV2" value="20"><span id="sp220">전시카테고리 [추가]삭제</span></td>  
									       
									    </tR>
								             <tr> 
										        <td><input type="checkbox" name="chkV2" value="21"><span id="sp221">전시카테고리 [기본]'디자인문구'로 등록</span></td>
										        <Td><input type="checkbox" name="chkV2" value="22"><span id="sp222">전시카테고리 [기본]'디지털/핸드폰'로 등록</span></td>
										    </tr>
								            <tr> 
										        <td><input type="checkbox" name="chkV2" value="23"><span id="sp223">전시카테고리 [기본]'캠핑/트래블'로 등록</span></td>
										        <td> <input type="checkbox" name="chkV2" value="24"><span id="sp224">전시카테고리 [기본]'토이'로 등록</span> </td> 
										    </tr>
										     <tr> 
            									<td><input type="checkbox" name="chkV2" value="25"><span id="sp225">전시카테고리 [기본]'디자인가전'로 등록</span> </td>
            									<td><input type="checkbox" name="chkV2" value="26"><span id="sp226">전시카테고리 [기본]'가구/수납'로 등록</span></td>
            								</tr>
            								<tr> 
            									<td><input type="checkbox" name="chkV2" value="27"><span id="sp227">전시카테고리 [기본]'데코/조명'로 등록</span></td>
            									 <td><input type="checkbox" name="chkV2" value="28"><span id="sp228">전시카테고리 [기본]'패브릭/생활'로 등록</span></td> 
            								</tr>
            								<tr> 
            									<td><input type="checkbox" name="chkV2" value="29"><span id="sp229">전시카테고리 [기본]'키친'로 등록</span></td>
            									<td><input type="checkbox" name="chkV2" value="30"><span id="sp230">전시카테고리 [기본]'푸드'로 등록</span></td>
            								</tr>
            								<tr> 
            									<td><input type="checkbox" name="chkV2" value="31"><span id="sp231">전시카테고리 [기본]'패션의류'로 등록</span></td>
            									<td><input type="checkbox" name="chkV2" value="32"><span id="sp232">전시카테고리 [기본]'패션잡화'로 등록</span></td>
            								</tr>
            								<tr>
            									<td><input type="checkbox" name="chkV2" value="33"><span id="sp233">전시카테고리 [기본]'주얼리/시계'로 등록</span> </td>
            									<td><input type="checkbox" name="chkV2" value="34"><span id="sp234">전시카테고리 [기본]'뷰티'로 등록</span></td>
            								</tr>  
            								<tr>
            									<td><input type="checkbox" name="chkV2" value="35"><span id="sp235">전시카테고리 [기본]'베이비/키즈'로 등록</span> </td>
            									<td><input type="checkbox" name="chkV2" value="36"><span id="sp236">전시카테고리 [기본]'Cat & Dog'로 등록</span></td>
            								</tr>
								        </table>
								    </td>
								</tr> 
								<tr>
									<td><hr width="100%"></td>
								</tr>
								<tr>
									<td><b>협의되지 않은 상품 등록</b></td>
								</tr>
									<tr>
									<td>
										<input type="checkbox" name="chkV2" value="14"><span id="sp214">해당카테고리 MD에게 상품제안 요망</span> 
									</td>
								</tr>
								<tr>
									<td><hr width="100%"></td>
								</tr>
								<tr>
									<td><b>기타</b></td>
								</tr>
									<tr>
									<td>
										<input type="checkbox" name="chkV2" value="999">추가 코멘트 : <input type="text" name="sM2" class="input" size="50" onFocus="jsChkValue(2);">
									</td>
								</tr>
							</table>
						</td>
				</tr> 
			</table>
		</td>
	</tr> 
	<tr>
		<Td colspan="2" align="center">
			<input type="button" class="button" value="취소"  onClick="jsCancel();"> 
			<input type="button" class="button" value="확인" onClick="jsConfirm(2);">
		</td>
	</tr>
</table>
</div><!-- //currstate=2, 승인보류 사유---------------------------->
<div id="dv0" style="display:none;overflow:scroll;"><!-- currstate=0, 승인반려 사유---------------------------->
	<table border=0 cellpadding=5 cellspacing=5   class="a" width="100%">  
		<tr>
		<td>승인 거부 사유: <font color="#FF0000">승인반려(재등록불가)</font> </td>
	</tr>
	<tr>
		<td>
			<table border=0 cellpadding=5 cellspacing=1 bgcolor="<%=adminColor("tablebg")%>" class="a" width="100%">
				<tr bgcolor="#FFFFFF">
					<td>
							<table border=0 cellpadding=3 cellspacing=0  class="a">
						<tr>
							<td>
								<input type="checkbox" name="chkV0" value="1"><span id="sp01">디자인상품을 지향하는 텐바이텐의 컨셉과 맞지 않아 진행이 어려울 것 같습니다.</span> 
							</td>
						</tr>
						<tr>
							<td>
								<input type="checkbox" name="chkV0" value="2"><span id="sp02">해당카테고리 진행방향과 맞지 않아 진행이 어려울 것 같습니다.</span> 
							</td>
						</tr>
						<tr>
							<td>
								<input type="checkbox" name="chkV0" value="3"><span id="sp03">이미지 퀄리티가 떨어져 진행이 어려울 것 같습니다. 다시 제작하여 연락 부탁 드립니다.</span> 
							</td>
						</tr>
						<tr>
							<td>
								<input type="checkbox" name="chkV0" value="4"><span id="sp04">상품의 퀄리티 문제로 진행이 어려울 것 같습니다.</span> 
							</td>
						</tr>
						<tr>
							<td><input type="checkbox" name="chkV0" value="5"><span id="sp05">동일한 상품이 이미 판매되고 있습니다.</span> </td>
						</tr>
						<tr>
							<td><input type="checkbox" name="chkV0" value="6"><span id="sp06">옵션이 누락되어 있습니다.옵션 추가해주세요</span> </td>
						</tr>
						<tr>
							<td>
							 	<input type="checkbox" name="chkV0" value="999">추가 코멘트 : <input type="text" name="sM0" class="input" size="50" onFocus="jsChkValue(0);">
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
	<tr>
		<Td colspan="2" align="center">
			<input type="button" class="button" value="취소" onClick="jsCancel();"> 
			<input type="button" class="button" value="확인" onClick="jsConfirm(0);">
		</td>
	</tr>
	</table>
</div><!-- //currstate=0, 승인반려 사유----------------------------> 