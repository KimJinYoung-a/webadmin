<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : WEBADMIN 메인
' Hieditor : 서동석 생성
'			 2022.07.08 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<!-- #include virtual="/lib/classes/board/surveyCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim sWorkerGubun
	sWorkerGubun = NullFillWith(Request("workergubun"),session("ssBctId"))
%>

<!--
<a href="/cscenter/cscenter_main.asp?menupos=757">CsMain.asp</a>
<a href="/admin/notice/MdMain.asp">MdMain.asp</a>
-->

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <!-- 왼쪽메뉴 시작 -->
	<td width="66%" valign="top">
	    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	    <!--
        <tr valign="top">
            <td>
				<form name="frmSearch" action="MdMain.asp" style="margin:0px;">
				<input type="hidden" name="mode">
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr height="30" bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	        	<img src="/images/icon_star.gif" align="absbottom">
							선택 :
            	    </td>
            	</tr>
            	</table>
				</form>
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
        -->
		<!-- 공지사항 시작-->
        <tr valign="top">
            <td>
				<%
					Dim lBoard, page
					Set lBoard = new board
						lBoard.FAdminlsn = session("ssAdminLsn")
						lBoard.FPartpsn = session("ssAdminPsn")
						lBoard.FPositsn = session("ssAdminPOSITsn")
						lBoard.FJob_sn = session("ssAdminPOsn")
						lBoard.Fdepartment_id =  GetUserDepartmentID("", session("ssBctId"))
						lBoard.fnmain_notice_list
						If lBoard.fresultcount > 0 Then
				%>
            	<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        		<tr bgcolor="<%= adminColor("menubar") %>">
        			<td>
        				<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        				<tr height="25">
        					<td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>공지사항</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="/admin/member_board/board_list.asp?menupos=1304">
        				        바로가기<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
        				</tr>
        				<tr height="25">
            			    <td colspan="2">
        						<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" bgcolor="#FFFFFF" >
								<tr height="25" align="center" bgcolor="#FFFFFF">
									<td width="30" style="border-bottom:1px solid #BABABA"><b>번호</b></td>
									<td width="60" style="border-bottom:1px solid #BABABA"><b>글쓴이</b></td>
									<td style="border-bottom:1px solid #BABABA"><b>제목</b></td>
									<td width="200" style="border-bottom:1px solid #BABABA"><b>열람부서</b></td>
									<td width="90" style="border-bottom:1px solid #BABABA"><b>등록일</b></td>
									<td width="40" style="border-bottom:1px solid #BABABA"><b>조회수</b></td>
								</tr>
								<script type='text/javascript'>
								function goView(bsn){
									location.href = "/admin/member_board/board_proc.asp?mode=count&brd_sn="+bsn;
								}
								</script>
								<%
									Dim Fteam_name, arrTN
									For i = 0 to lBoard.fresultcount -1
										arrTN="": Fteam_name=""
										If lboard.FbrdList(i).Fbrd_team <> "" Then
											arrTN = split(lboard.FbrdList(i).Fbrd_team,",")
											if ubound(arrTN)>1 then
												Fteam_name = arrTN(0) & " 외 " & ubound(arrTN)-1 & "건"
											else
												Fteam_name = arrTN(0)
											end if
										End If
								%>

								<tr height="25" bgcolor="FFFFFF" onClick="goView('<%=lboard.FbrdList(i).Fbrd_sn%>')" style="cursor:pointer" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" >
									<td align="center" width="30" style="border-bottom:1px solid #BABABA"><%=lboard.FbrdList(i).Fbrd_sn%></td>
									<td align="center" width="60" style="border-bottom:1px solid #BABABA"><%=lboard.FbrdList(i).Fbrd_username%></td>
									<td style="border-bottom:1px solid #BABABA">
										<%
											If lboard.FbrdList(i).Fbrd_fixed = "1" Then
												response.write "<b>"&lboard.FbrdList(i).Fbrd_subject
												If lboard.FbrdList(i).Fcnt = "0" Then
													response.write ""
												Else
													response.write "&nbsp;<b><font color='RED'>["&lboard.FbrdList(i).Fcnt&"]</font></b>"
												End If
											Else
												response.write lboard.FbrdList(i).Fbrd_subject
												If lboard.FbrdList(i).Fcnt = "0" Then
													response.write ""
												Else
													response.write "&nbsp;<b><font color='RED'>["&lboard.FbrdList(i).Fcnt&"]</font></b>"
												End If
											End If
										%>
									</td>
									<td width="150" style="border-bottom:1px solid #BABABA"><%=Fteam_name%></td>
									<td align="center" width="70" style="border-bottom:1px solid #BABABA"><%=left(lboard.FbrdList(i).Fbrd_regdate,10)%></td>
									<td align="center" width="40" style="border-bottom:1px solid #BABABA"><%=lboard.FbrdList(i).Fbrd_hit%></td>
								</tr>
								<%
									Next
								%>
								</table>
							</td>
						</tr>
        				</table>
        			</td>
        		</tr>
            	</table>
            	<table><tr height="7"><td></td></tr></table>
	           	<% End If %>
            	<br>
        	    <!-- 설문조사 시작-->
				<%
					Dim oSurvey
					Set oSurvey = new CSurvey
					oSurvey.FPagesize = 15
					oSurvey.FCurrPage = 1
					oSurvey.FRectUsing = "Y"
					oSurvey.FRectDiv = "2"						'직원설문
					oSurvey.FRectState = "2"					'진행중인 설문
					oSurvey.FRectUserid = session("ssBctId")	'내설문상태 파악
					oSurvey.GetSurveyList

					If oSurvey.FResultCount > 0 Then
				%>
				<script type='text/javascript'>
				<!--
					function fnSurveyPopup(sno) {
						var popSurvey = window.open("/admin/board/popup_survey.asp?sn="+sno,"popSurvey","width=1400,height=768,scrollbars=yes");
						popSurvey.focus();
					}
				//-->
				</script>
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA"><img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>설문조사</b></td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">&nbsp;</td>
            			</tr>
            			<tr height="25">
							<td colspan="2">
								<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" bgcolor="#FFFFFF">
								<tr height="25" align="center" bgcolor="#FFFFFF">
									<td width="40" style="border-bottom:1px solid #BABABA"><b>번호</b></td>
									<td style="border-bottom:1px solid #BABABA"><b>제목</b></td>
									<td width="150" style="border-bottom:1px solid #BABABA"><b>설문기간</b></td>
									<td width="80" style="border-bottom:1px solid #BABABA"><b>상태</b></td>
								</tr>
							<% for i=0 to oSurvey.FResultCount-1 %>
								<tr height="20" align="center" bgcolor="#FFFFFF">
									<td style="border-bottom:1px solid #BABABA"><%= i+1 %></td>
									<td align="left" style="border-bottom:1px solid #BABABA">
									<% if oSurvey.FItemList(i).getSurveyStateCD="1" then %>
										<a href="javascript:fnSurveyPopup(<%= oSurvey.FItemList(i).Fsrv_sn %>)"><%= ReplaceBracket(oSurvey.FItemList(i).Fsrv_subject) %></a>
									<% else %>
										<%= ReplaceBracket(oSurvey.FItemList(i).Fsrv_subject) %>
									<% end if %>
									</td>
									<td style="border-bottom:1px solid #BABABA"><%= left(oSurvey.FitemList(i).Fsrv_startDt,10) & "~" & left(oSurvey.FitemList(i).Fsrv_endDt,10) %></td>
									<td style="border-bottom:1px solid #BABABA"><%= oSurvey.FitemList(i).getSurveyState %></td>
								</tr>
							<% next %>
								</table>
            			    </td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
            	<p>&nbsp;</p>

            	<%
            		end if
            		Set oSurvey = Nothing
            	%>
        	    <!-- 업체게시판 시작-->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>업체게시판&nbsp;[미답변]</b>
            			    	&nbsp;
            			    	<select class="select" name="workergubun" onChange="location.href='scmmain.asp?workergubun='+this.value+'';">
            			    		<option value="all_" <% If sWorkerGubun = "all_" Then %>selected<% End If %>>전체보기</option>
            			    		<option value="<%=session("ssBctId")%>" <% If sWorkerGubun = session("ssBctId") Then %>selected<% End If %>>내가받은문의보기</option>
            			    	</select>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">

            			        <a href="/admin/board/upche_qna_board_list.asp?menupos=402&workergubun=<%=session("ssBctId")%>">
        				        바로가기<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td colspan="2">


            			    	<% if (session("ssBctDiv") < 10) then %>
								<%
								dim itemqanotinclude, research, i

								'==============================================================================
								dim boardqna
								set boardqna = New CUpcheQnA

								boardqna.FPageSize = 200
								boardqna.FCurrPage = 1
								boardqna.FRectRelpy = "N"
								boardqna.FWorkerGubun = Replace(sWorkerGubun,"all_","")
								boardqna.list

								%>

            			    	<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" bgcolor="#FFFFFF" >
								  <tr height="25" align="center" bgcolor="#FFFFFF">
								    <td width="100" style="border-bottom:1px solid #BABABA"><b>업체명</b></td>
								    <td style="border-bottom:1px solid #BABABA"><b>제목</b></td>
								    <td width="100" style="border-bottom:1px solid #BABABA"><b>구분</b></td>
								    <td width="100" style="border-bottom:1px solid #BABABA"><b>업체구분</b></td>
								    <td width="70" style="border-bottom:1px solid #BABABA"><b>담당자</b></td>
								    <td width="100" style="border-bottom:1px solid #BABABA"><b>작성일</b></td>
								  </tr>
								<% for i = 0 to (boardqna.FResultCount - 1) %>
								  <tr height="20" align="center" bgcolor="#FFFFFF">
								    <td align="left" style="border-bottom:1px solid #BABABA"><%= boardqna.FItemList(i).Fusername %>(<%= boardqna.FItemList(i).Fuserid %>)</td>
								    <td align="left" style="border-bottom:1px solid #BABABA">
								    	<a href="/admin/board/upche_qna_board_reply.asp?idx=<%= boardqna.FItemList(i).Fidx %>"><%= (boardqna.FItemList(i).Ftitle) %></a>
								    	<% if datediff("d",boardqna.FItemList(i).Fregdate,now())<6 then %>
										&nbsp;&nbsp;&nbsp;<img src="/images/new.gif">
										<% end if %>
								    </td>
								    <td style="border-bottom:1px solid #BABABA"><%= boardqna.FItemList(i).GubunName %></td>
								    <td style="border-bottom:1px solid #BABABA"><%= boardqna.FItemList(i).UpcheGubun %></td>
								    <td style="border-bottom:1px solid #BABABA"><%= boardqna.FItemList(i).Fworker %></td>
								    <td style="border-bottom:1px solid #BABABA"><%= FormatDate(boardqna.FItemList(i).Fregdate, "0000.00.00") %></td>
								  </tr>
								<% next %>
								</table>
								<% set boardqna = Nothing %>

								<% end if %>


            			   	</td>
            			</tr>
            	        </table>



            	    </td>
            	</tr>
            	</table>
        	    <!-- 업체게시판 관리 끝-->
        	</td>
        </tr>

        </table>
    </td>
    <!-- 왼쪽메뉴 끝 -->

    <td width="10"></td>

    <!-- 오픈쪽메뉴 시작 -->
    <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        <tr valign="top">
            <td>
                <!-- 새로고침 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
                        	<td>
            			    	<img src="/images/icon_star.gif" align="absbottom">
								<b>로그인 ID : </b>
								<%=session("ssBctId")%>
								<!-- 초기로그인시 로그인 아이디로 설정 -->
            			    </td>
            			    <td align="right">
            			    <!--
            			    	<a href="javascript:document.location.reload();">
        				        새로고침
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
        				        -->
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
            	<!-- 새로고침 끝 -->
            </td>
        </tr>

        <tr valign="top"><td height="10"></td></tr>

        <%
        	Dim NewCoop
        	Set NewCoop = new CCooperate
        	NewCoop.FDoc_Id = session("ssBctId")
        	NewCoop.fnGetCooperateCount
        %>
        <tr valign="top">
            <td>
                <!-- 협조문 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>사내협조문</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="/admin/notice/cooperate/?menupos=1167" target="_blank">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
				                </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>받은 협조문 (미처리)</td>
            			    <td align="right">
            			        <a href="/admin/notice/cooperate/?menupos=1167&doc_status=0&onlymine=o" target="_blank">
            			        <%
            			        	If NewCoop.FComeCnt = 0 Then
            			        		Response.Write "[" & NewCoop.FComeCnt & "] 건"
            			        	Else
            			        		Response.Write "[<b>" & NewCoop.FComeCnt & "</b>] 건"
            			        		if NewCoop.FComeNewCnt>0 then Response.Write".<img src='http://fiximage.10x10.co.kr/web2009/main/news_icon_new.gif' border='0'>"
            			        	End If
            			        %>
            			        </a>
        				    	<a href="/admin/notice/cooperate/?menupos=1167" target="_blank">
                    		    <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
                    		    </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>보낸 협조문 (미처리)</td>
            			    <td align="right">
            			        <a href="/admin/notice/cooperate/my_cooperate.asp?menupos=1167&doc_status=0" target="_blank">
            			        <%
            			        	If NewCoop.FSendCnt = 0 Then
            			        		Response.Write "[" & NewCoop.FSendCnt & "] 건"
            			        	Else
            			        		Response.Write "[<b>" & NewCoop.FSendCnt & "</b>] 건"
            			        		if NewCoop.FSendNewCnt>0 then Response.Write".<img src='http://fiximage.10x10.co.kr/web2009/main/news_icon_new.gif' border='0'>"
            			        	End If
            			        %>
            			        </a>
        				    	<a href="/admin/notice/cooperate/my_cooperate.asp?menupos=1167" target="_blank">
                    		    <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
                    		    </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  협조문 끝-->
            </td>
        </tr>
        <!--  직원 생일자 시작 		'// 4시간에 한번 돌아감 -->
        <tr valign="top"><td height="10"></td></tr>
        <tr valign="top">
            <td><!-- #include virtual="/admin/member/inc_member_birthInfo.asp" --></td>
        </tr>
        <!--  직원 생일자 끝 -->

		<!-- 담당 MD만 보이는 당첨자 리스트 -->
		<% If session("ssAdminPsn") = "11" or session("ssAdminPsn") = "21" or session("ssBctId") ="hrkang97" Then %>
        <tr valign="top"><td height="10"></td></tr>
        <tr valign="top">
            <td><!-- #include virtual="/admin/member/inc_member_MD.asp" --></td>
        </tr>
		<% End If %>
		<!-- 담당 MD만 보이는 당첨자 리스트 끝-->

		<!-- 사내일정공지 -->
        <tr valign="top"><td height="10"></td></tr>
        <tr valign="top">
            <td><!-- #include virtual="/admin/member/inc_member_notice.asp" --></td>
        </tr>
		<!-- 사내일정공지 -->

        <% if ((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or (session("ssAdminPOsn") = "5") or session("ssAdminPsn")=7 or session("ssAdminPsn")=30) then %>
        <!--  직원 휴가신청 시작 -->
        <tr valign="top"><td height="10"></td></tr>
        <tr valign="top">
            <td><!-- #include virtual="/admin/member/inc_member_vacation.asp" --></td>
        </tr>
        <!--  직원 휴가신청 끝 -->
        <% end if %>

<!--
        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            내용
            	    </td>
            	</tr>
            	</table>
            </td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            내용
            	    </td>
            	</tr>
            	</table>
            </td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr>
            <td>
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            내용
            	    </td>
            	</tr>
            	</table>

            </td>
        </tr>
        </table>
    </td>
-->
    <!-- 오픈쪽메뉴 끝 -->

</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
