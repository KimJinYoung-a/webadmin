<%
''���޸� selectbox ����
 function fnGetOptOutMall(byVal sellsite)
%>
<select class="select" name="sellsite">
			<option></option>
			<option value="interpark" <% if (sellsite = "interpark") then %>selected<% end if %> >������ũ</option>
			<option value="lotteimall" <% if (sellsite = "lotteimall") then %>selected<% end if %> >�Ե����̸�</option>
			<option value="lotteCom" <% if (sellsite = "lotteCom") then %>selected<% end if %> >�Ե�����</option>
			<option value="11st1010" <% if (sellsite = "11st1010") then %>selected<% end if %> >11����</option>
			<option value="auction1010" <% if (sellsite = "auction1010") then %>selected<% end if %> >����</option>
			<option value="gmarket1010" <% if (sellsite = "gmarket1010") then %>selected<% end if %> >������(NEW)</option>
			<!-- option value="lotteComM" <% if (sellsite = "lotteComM") then %>selected<% end if %> >�Ե�����(������)</option -->
			<option value="gseshop" <% if (sellsite = "gseshop") then %>selected<% end if %> >GS��</option>
			<!-- option value="dnshop" <% if (sellsite = "dnshop") then %>selected<% end if %> >��ؼ�</option -->
			<option value="cjmall" <% if (sellsite = "cjmall") then %>selected<% end if %> >CJ��</option>
			<!-- option value="wizwid" <% if (sellsite = "wizwid") then %>selected<% end if %> >��������</option -->
			<!-- option value="gabangpop" <% if (sellsite = "gabangpop") then %>selected<% end if %> >�м���(������)</option -->
			<!-- option value="wconcept" <% if (sellsite = "wconcept") then %>selected<% end if %> >����������</option -->
			<!-- option value="privia" <% if (sellsite = "privia") then %>selected<% end if %> >�����������</option -->
			<!-- option value="player" <% if (sellsite = "player") then %>selected<% end if %> >�÷��̾�</option -->
			<option value="homeplus" <% if (sellsite = "homeplus") then %>selected<% end if %> >Ȩ�÷���</option>
			<option value="ssg" <% if (sellsite = "ssg") then %>selected<% end if %> >SSG</option>
			<option value="ssg6006" <% if (sellsite = "ssg6006") then %>selected<% end if %> >SSG-�̸�Ʈ</option>
			<option value="ssg6007" <% if (sellsite = "ssg6007") then %>selected<% end if %> >SSG-ssg</option>
			<option value="nvstorefarm" <% if (sellsite = "nvstorefarm") then %>selected<% end if %> >�������</option>
			<option value="ezwel" <% if (sellsite = "ezwel") then %>selected<% end if %> >���������</option>
			<option value="kakaogift" <% if (sellsite = "kakaogift") then %>selected<% end if %> >īī������Ʈ</option>
			<option value="coupang" <% if (sellsite = "coupang") then %>selected<% end if %> >����</option>
			<option value="halfclub" <% if (sellsite = "halfclub") then %>selected<% end if %> >����Ŭ��</option>
			<option value="hmall" <% if (sellsite = "hmall") then %>selected<% end if %> >Hmall</option>
		</select>
<%
 end function

%>