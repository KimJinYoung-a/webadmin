<%
''제휴몰 selectbox 정보
 function fnGetOptOutMall(byVal sellsite)
%>
<select class="select" name="sellsite">
			<option></option>
			<option value="interpark" <% if (sellsite = "interpark") then %>selected<% end if %> >인터파크</option>
			<option value="lotteimall" <% if (sellsite = "lotteimall") then %>selected<% end if %> >롯데아이몰</option>
			<option value="lotteCom" <% if (sellsite = "lotteCom") then %>selected<% end if %> >롯데닷컴</option>
			<option value="11st1010" <% if (sellsite = "11st1010") then %>selected<% end if %> >11번가</option>
			<option value="auction1010" <% if (sellsite = "auction1010") then %>selected<% end if %> >옥션</option>
			<option value="gmarket1010" <% if (sellsite = "gmarket1010") then %>selected<% end if %> >지마켓(NEW)</option>
			<!-- option value="lotteComM" <% if (sellsite = "lotteComM") then %>selected<% end if %> >롯데닷컴(직매출)</option -->
			<option value="gseshop" <% if (sellsite = "gseshop") then %>selected<% end if %> >GS샵</option>
			<!-- option value="dnshop" <% if (sellsite = "dnshop") then %>selected<% end if %> >디앤샵</option -->
			<option value="cjmall" <% if (sellsite = "cjmall") then %>selected<% end if %> >CJ몰</option>
			<!-- option value="wizwid" <% if (sellsite = "wizwid") then %>selected<% end if %> >위즈위드</option -->
			<!-- option value="gabangpop" <% if (sellsite = "gabangpop") then %>selected<% end if %> >패션팝(가방팝)</option -->
			<!-- option value="wconcept" <% if (sellsite = "wconcept") then %>selected<% end if %> >더블유컨셉</option -->
			<!-- option value="privia" <% if (sellsite = "privia") then %>selected<% end if %> >현대프리비아</option -->
			<!-- option value="player" <% if (sellsite = "player") then %>selected<% end if %> >플레이어</option -->
			<option value="homeplus" <% if (sellsite = "homeplus") then %>selected<% end if %> >홈플러스</option>
			<option value="ssg" <% if (sellsite = "ssg") then %>selected<% end if %> >SSG</option>
			<option value="ssg6006" <% if (sellsite = "ssg6006") then %>selected<% end if %> >SSG-이마트</option>
			<option value="ssg6007" <% if (sellsite = "ssg6007") then %>selected<% end if %> >SSG-ssg</option>
			<option value="nvstorefarm" <% if (sellsite = "nvstorefarm") then %>selected<% end if %> >스토어팜</option>
			<option value="ezwel" <% if (sellsite = "ezwel") then %>selected<% end if %> >이지웰페어</option>
			<option value="kakaogift" <% if (sellsite = "kakaogift") then %>selected<% end if %> >카카오기프트</option>
			<option value="coupang" <% if (sellsite = "coupang") then %>selected<% end if %> >쿠팡</option>
			<option value="halfclub" <% if (sellsite = "halfclub") then %>selected<% end if %> >하프클럽</option>
			<option value="hmall" <% if (sellsite = "hmall") then %>selected<% end if %> >Hmall</option>
		</select>
<%
 end function

%>