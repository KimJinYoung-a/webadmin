<%@ language=vbscript %>
    <% option explicit %>
        <% '#############################################################
' Description : ���� ����ǰ ���� '	History		: 2022.01.05 ������ ����
' ############################################################# %>
            <!-- #include virtual="/lib/function.asp"-->
            <!-- #include virtual="/lib/db/dbopen.asp" -->
            <!-- #include virtual="/lib/util/htmllib.asp" -->
            <!-- #include virtual="/admin/incSessionAdmin.asp" -->
            <!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
            </p>
            <link rel="stylesheet" type="text/css" href="/css/commonV20.css" />
            <div>
                <!-- �������� ����� -->
                <div class="content paidGift detail">
                    <div class="paidGift_top">
                        <div class="search_wrap">
                            <select name="" id="select01">
                                <option value="">����</option>
                                <option value="">�÷�������</option>
                                <option value="">�������ǰ</option>
                            </select>
                            <select name="" id="select02">
                                <option value="">�������</option>
                                <option value="">������</option>
                                <option value="">���࿹��</option>
                                <option value="">����</option>
                                <option value="">������</option>
                            </select>
                            <div class="input_wrap">
                                <select name="" id="select03">
                                    <option value="">���μ�</option>
                                    <option value="">�����</option>
                                </select>
                                <span></span>
                                <button class="btn_select">�����ϱ�</button>
                                <li class="selected" style="display:none;">������<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                            </div>
                             <div class="input_wrap">
                                <select name="" id="select04">
                                    <option value="">��������</option>
                                    <option value="">���� ù ����</option>
                                    <option value="">�ֱ� �����ݾ�</option>
                                    <option value="">ȸ�����</option>
                                    <option value="">���űݾ�</option>
                                    <option value="">����Ƚ��</option>
                                    <option value="">��ǰ</option>
                                    <option value="">ī�װ�</option>
                                    <option value="">�귣��</option>
                                    <option value="">��ȹ��/�̺�Ʈ</option>
                                </select>
                                <li class="selected" style="display:none;">��ȹ��/�̺�Ʈ<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                                <li class="selected" style="display:none;">��ǰ<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                            </div>
                            <div class="input_wrap">
                                <select name="" id="select05">
                                    <option value="">��������/��ǰ/����ǰ��</option>
                                    <option value="">�������� ��ȣ</option>
                                    <option value="">����ǰ�ڵ�</option>
                                    <option value="">��ǰ�ڵ�</option>
                                </select>
                                <span></span>
                                <input type="text" placeholder="�˻�� �Է����ּ���">
                            </div>
                            <button class="btn_search"><img src="https://webadmin.10x10.co.kr/images/icon/search.png">�˻��ϱ�</button>
                        </div>
                        <div class="tgl_wrap">
                            <span>���� ����� �������ø� ����</span>
                            <div class="tgl_btn">
                                <input type="checkbox" id="tgl_btn_my">
                                <label for="tgl_btn_my"></label>
                            </div>
                        </div>
                    </div>
                    <div class="paidGift_aside">
                        <div class="list_wrap">
                            <div class="add_btn">
                                <button>+ �������� �����</button>
                            </div>
                            <div class="list_top">
                                <li>�� <span>103</span>��</li>
                            </div>
                            <div class="list_cont">
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st01">���¿���</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st02">������</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st02">������</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st02">������</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st01">���¿���</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st03">����</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st04">������</li>
                                    </ul>
                                </div></a>
                            </div>
                            <div class="list_bottom">
                                <ul class="pagination">
                                    <li class="on"><a>1</a></li>
                                    <li class=""><a>2</a></li>
                                    <li class=""><a>3</a></li>
                                    <li class=""><a>4</a></li>
                                    <li class=""><a>5</a></li> 
                                    <li class=""><a>></a></li>
                                    <li class=""><a>>></a></li>
                                </ul>
                            </div>
                        </div>
                    </div>
                    <div class="paidGift_section">
                        <!-- ���ǳ��� : on -->
                        <div class="section_top">
                            <li>�������� ��ȣ : <span>1234</span></li>
                            <div>
                                <button class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png" alt="">���ļ��� �����ϱ�</button>
                                <button><img src="https://webadmin.10x10.co.kr/images/icon/trash_gray.png" alt="">�����ϱ�</button>
                            </div>
                        </div>
                        <div class="step_wrap">
                            <div class="step_cont detail on">
                                <li class="dt_type">�÷�������</li>
                                <button class="btn_edit">�����ϱ�</button>
                                <div class="dt_list_wrap">
                                    <textarea class="title01" rows="1">1�� ù���� �÷��� ����</textarea>
                                    <textarea class="title02" rows="1">1������ ������ ã���ֽ� ù ���� ȸ���鿡�Ը� �帮�� Ư���� ����</textarea>
                                </div>
                                <div class="dt_list_wrap">
                                    <li>��Ⱓ</li>
                                     <div class="date_wrap">
                                        <!-- ���� �� date start on / date end on -->
                                        <ul class="date start">
                                        <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png"></button>
                                        <!-- ��¥ ���� -->
                                            <div class="cal_month t02" style="display:none;">
                                                <div class="arrow"></div>
                                                    <table class="table-condensed table-bordered table-striped">
                                                            <thead>
                                                                <tr>
                                                                    <th colspan="7">
                                                                        <ul class="btn_group">
                                                                            <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></a>
                                                                            <li class="btn active">2�� 2022</li>
                                                                            <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                        </ul>
                                                                    </th>
                                                                </tr>
                                                            </thead>
                                                            <tbody>
                                                                <tr>
                                                                    <td class="gray">30</td>
                                                                    <td class="gray">31</td>
                                                                    <td class="on">1</td>
                                                                    <td>2</td>
                                                                    <td>3</td>
                                                                    <td>4</td>
                                                                    <td>5</td>
                                                                </tr>
                                                                <tr>
                                                                    <td>6</td>
                                                                    <td>7</td>
                                                                    <td>8</td>
                                                                    <td>9</td>
                                                                    <td>10</td>
                                                                    <td>11</td>
                                                                    <td>12</td>
                                                                </tr>
                                                                <tr>
                                                                    <td>13</td>
                                                                    <td>14</td>
                                                                    <td>15</td>
                                                                    <td>16</td>
                                                                    <td>17</td>
                                                                    <td>18</td>
                                                                    <td>19</td>
                                                                </tr>
                                                                <tr>
                                                                    <td>20</td>
                                                                    <td>21</td>
                                                                    <td>22</td>
                                                                    <td>23</td>
                                                                    <td>24</td>
                                                                    <td>25</td>
                                                                    <td>26</td>
                                                                </tr>
                                                                <tr>
                                                                    <td>27</td>
                                                                    <td>28</td>
                                                                    <td class="gray">1</td>
                                                                    <td class="gray">2</td>
                                                                    <td class="gray">3</td>
                                                                    <td class="gray">4</td>
                                                                    <td class="gray">5</td>
                                                                </tr>
                                                            </tbody>
                                                </table>
                                            </div>
                                            <input type="text" placeholder="00" class="time">:<input type="text" placeholder="00" class="time"><img src="https://webadmin.10x10.co.kr/images/icon/clock.png">
                                        </ul>
                                        <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png">
                                        <ul class="date end">
                                            <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png"></button>
                                        <!-- ��¥ ���� -->
                                            <div class="cal_month t02" style="display:none;">
                                                <div class="arrow"></div>
                                                    <table class="table-condensed table-bordered table-striped">
                                                            <thead>
                                                                <tr>
                                                                    <th colspan="7">
                                                                        <ul class="btn_group">
                                                                            <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                            <li class="btn active">2�� 2022</li>
                                                                            <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                        </ul>
                                                                    </th>
                                                                </tr>
                                                            </thead>
                                                            <tbody>
                                                                <tr>
                                                                    <td class="gray">30</td>
                                                                    <td class="gray">31</td>
                                                                    <td class="on">1</td>
                                                                    <td>2</td>
                                                                    <td>3</td>
                                                                    <td>4</td>
                                                                    <td>5</td>
                                                                </tr>
                                                                <tr>
                                                                    <td>6</td>
                                                                    <td>7</td>
                                                                    <td>8</td>
                                                                    <td>9</td>
                                                                    <td>10</td>
                                                                    <td>11</td>
                                                                    <td>12</td>
                                                                </tr>
                                                                <tr>
                                                                    <td>13</td>
                                                                    <td>14</td>
                                                                    <td>15</td>
                                                                    <td>16</td>
                                                                    <td>17</td>
                                                                    <td>18</td>
                                                                    <td>19</td>
                                                                </tr>
                                                                <tr>
                                                                    <td>20</td>
                                                                    <td>21</td>
                                                                    <td>22</td>
                                                                    <td>23</td>
                                                                    <td>24</td>
                                                                    <td>25</td>
                                                                    <td>26</td>
                                                                </tr>
                                                                <tr>
                                                                    <td>27</td>
                                                                    <td>28</td>
                                                                    <td class="gray">1</td>
                                                                    <td class="gray">2</td>
                                                                    <td class="gray">3</td>
                                                                    <td class="gray">4</td>
                                                                    <td class="gray">5</td>
                                                                </tr>
                                                            </tbody>
                                                </table>
                                            </div>
                                            <input type="text" placeholder="00" class="time">:<input type="text" placeholder="00" class="time"><img src="https://webadmin.10x10.co.kr/images/icon/clock.png">
                                        </ul>
                                    </div> 
                                    <span class="line"></span>
                                    <li>��뿩��</li>
                                    <select>
                                        <option>�����</option>
                                        <option>������</option>
                                    </select>
                                    <span class="line"></span>
                                    <li>�����</li>
                                    <li>
                                        <input type="text" value="�������� - ������" class="name">
                                    </li>
                                </div>
                            </div> 
                        </div>
                        <div class="step_wrap">
                            <div class="step_cont detail on">
                                <h3 class="dt_list_top">���� ���� �����<span class="noti blue">�׸��� �����ϸ� ���� �Ǵ� �߰��� �� �־��</span></h3>
                                <div class="dt_list_wrap">
                                    <div class="info_wrap_top">
                                        <button>
                                            <li>��� ������</li>
                                            <li class="check">&#10003;</li>
                                        </button>
                                         <button class="on">
                                            <li>Ư�� ������</li>
                                            <li class="check">&#10003;</li>
                                            <span class="arrow"></span>
                                        </button>
                                    </div>
                                    <div class="info_wrap_list">
                                        <li class="bold">���� �������� <span>2</span></li>
                                        <div class="btn_wrap">
                                           <button class="type02 on">
                                                <h3>���� ù ����</h3>
                                                <li>���� �� ù ������ ���</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                            <button class="type02">
                                                <h3>�ֱ� �����ݾ� ����</h3>
                                                <li>�ֱ� 5���� �� ��ۿϷ�� �ֹ����� �� �����ݾ�</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                            <button class="type02">
                                                <h3>ȸ����� �ش�</h3>
                                                <li>�ſ� ���ŵǴ� ȸ����� ����</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                            <button class="type02">
                                                <h3>���űݾ� ����</h3>
                                                <li>���� �ֹ� ���� ���űݾ� ����</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                            <button class="type02">
                                                <h3>����Ƚ�� ����</h3>
                                                <li>�ֱ� 5���� �� ��ۿϷ�� �ֹ�Ƚ�� ����</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                            <button class="type02">
                                                <h3>������ ��ǰ ����</h3>
                                                <li>Ư�� ��ǰ ���� ��</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                            <button class="type02">
                                                <h3>������ ī�װ� ��ǰ ����</h3>
                                                <li>���� �Ǵ� ���� �ֹ� �ǿ��� Ư�� ī�װ� ��ǰ ���� ��</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                            <button class="type02">
                                                <h3>������ �귣�� ��ǰ ����</h3>
                                                <li>���� �Ǵ� ���� �ֹ� �ǿ��� Ư�� �귣�� ��ǰ ���Ž�</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                            <button class="type02">
                                                <h3>������ ��ȹ��/�̺�Ʈ ��ǰ ����</h3>
                                                <li>Ư�� ��ȹ��/�̺�Ʈ ��ǰ ���� ��</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                        </div>
                                    </div>
                                    <!-- ����� �̺�Ʈ -->
                                    <div class="info_wrap_list" style="display:block;">
                                        <li class="bold">����� �̺�Ʈ</li>
                                        <div class="btn_wrap">
                                            <button class="type02 on">
                                                <h3>���̾ ���丮</h3>
                                                <li>���̾ ���丮 ����ǰ ���� ���</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                            <button class="type02">
                                                <h3>�ֳ��̺�Ʈ</h3>
                                                <li>N �ֳ��̺�Ʈ ����ǰ ���� ���</li>
                                                 <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            </div> 
                        </div>
                        <div class="step_wrap">
                            <div class="step_cont detail on">
                                <h3 class="dt_list_top">���� ����</h3>
                                <!-- �ܼ��׷켳�� �ÿ��� ���� -->
                                <button class="btn_edit">�׷� ���� �����ϱ�</button>
                                <div class="dt_list_wrap">
                                    <div class="info_wrap_top">
                                        <div class="btn_wrap t02">
                                            <button>�׷���� ��ǰ ���</button>
                                            <button class="on">�ܼ� �׷� ����</button>
                                            <button>�ݾ׺� �׷� ����</button>
                                        </div>
                                    </div>
                                    <div class="info_wrap_list t02">
                                        <!-- �÷������� -->
                                        <!-- �׷���� ��ǰ ��� -->
                                        <div class="step_detail group group01_01" style="display:block;">
                                            <div class="table_wrap">
                                                <div class="ttop">
                                                    <span>�� 0��</span><span class="selected">���õ� ��ǰ 1��</span>
                                                    <button>���û�ǰ ����</button>
                                                </div>
                                                <div class="table">
                                                    <ul class="thead">
                                                        <li><input type="checkbox"></li>
                                                        <li>��ǰ�ڵ�</li>
                                                        <li>��ǥ�̹���/�귣��/��ǰ��</li>
                                                        <li>�߰�����</li>
                                                        <li>��� ������<span>��������/��ü����</span></li>
                                                        <li>�ǸŰ�</li>
                                                        <li>���԰�</li>
                                                        <li>����</li>
                                                        <li>��౸��</li>
                                                        <li></li>
                                                    </ul>
                                                    <div class="tbody_wrap add">
                                                        <ul class="tbody">
                                                            <span>
                                                                + ��ǰ �߰��ϱ�
                                                            </span>
                                                        </ul>
                                                    </div>
                                                    <div class="tbody_wrap">
                                                        <ul class="tbody">
                                                            <li><input type="checkbox"></li>
                                                            <li>12345678</li>
                                                            <li>
                                                                <img class="prd_img">
                                                                <ul><p class="prd_brand">dailylike</p>
                                                                <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul><p class="badge">����[MDƯ��]</p>
                                                                <p class="info">���ǻ���</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="percent">99%</p>
                                                                    <ul class="bar">
                                                                        <p class="bar_gray"></p>
                                                                        <p class="bar_black"></p>
                                                                    </ul>
                                                                    <p class="">99/100</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="price p_type01"><span>38,000</span></p>
                                                                    <p class="price p_type02">7% ����<span>38,000</span></p>
                                                                    <p class="price p_type03">7% ����<span>38,000</span></p>
                                                                    <p class="price p_type04">15% �÷�������<span>38,000</span></p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="p_type01">38,000</p>
                                                                    <p class="p_type02">23,902</p>
                                                                    <p class="p_type03">23,902</p>
                                                                    <p class="p_type04">23,902</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="p_type01">27%</p>
                                                                    <p class="p_type02">26%</p>
                                                                    <p class="p_type03">26%</p>
                                                                    <p class="p_type04">28%</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                ����
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <button>����</button>
                                                                    <button>����</button>
                                                                </ul>
                                                            </li>
                                                        </ul>
                                                    </div>
                                                    <div class="tbody_wrap">
                                                        <ul class="tbody">
                                                            <li><input type="checkbox"></li>
                                                            <li>12345678</li>
                                                            <li>
                                                                <img class="prd_img">
                                                                <ul><p class="prd_brand">dailylike</p>
                                                                <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul><p class="badge"></p>
                                                                <p class="info"></p></ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="percent"></p>
                                                                    <ul class="bar" style="display:none;">
                                                                        <p class="bar_gray"></p>
                                                                        <p class="bar_black"></p>
                                                                    </ul>
                                                                    <p class="">99</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="price p_type01"><span>38,000</span></p>
                                                                    <p class="price p_type02">7% ����<span>38,000</span></p>
                                                                    <p class="price p_type03">7% ����<span>38,000</span></p>
                                                                    <p class="price p_type04">15% �÷�������<span>38,000</span></p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="p_type01">38,000</p>
                                                                    <p class="p_type02">23,902</p>
                                                                    <p class="p_type03">23,902</p>
                                                                    <p class="p_type04">23,902</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="p_type01">27%</p>
                                                                    <p class="p_type02">26%</p>
                                                                    <p class="p_type03">26%</p>
                                                                    <p class="p_type04">28%</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                ����
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <button>����</button>
                                                                    <button>����</button>
                                                                </ul>
                                                            </li>
                                                        </ul>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                        <!-- �ܼ� �׷� ���� -->
                                        <div class="step_detail group group02_01" style="display:none;">
                                            <div class="tab">
                                                <button class="added on"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>���̾���� ���Ϲ����ص� ��ƼĿ��</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                                <button class="added"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>ġƮŰ�� ���޸���</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                                <button><li>+�׷� �߰��ϱ�</li></button>
                                            </div>
                                            <!-- �׷Ӿ��� ��ǰ ��ϰ� ����<div class="table_wrap"></div> -->
                                        </div>
                                        <!-- �ݾ׺� �׷� ���� -->
                                        <div class="step_detail group group03_01" style="display:none;">
                                            <div class="tab">
                                                <button class="added on"><li>10,000��~<span>&#65372;</span><span class="gr_cond">���űݾ�</span><span class="gr_name">10,000�� �̻� �����ߴٸ�</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                                <button class="added"><li>10,000��~<span>&#65372;</span><span class="gr_cond">���űݾ�</span><span class="gr_name">10,000�� �̻� �����ߴٸ�</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                                <button><li>+�ݾ׺� �׷� �߰��ϱ�</li></button>
                                            </div>
                                            <!-- �׷Ӿ��� ��ǰ ��ϰ� ����<div class="table_wrap"></div> -->
                                        </div>
                                        <!-- �������ǰ -->
                                        <!-- �׷���� ��ǰ ��� -->
                                        <div class="step_detail group group01_02" style="display:none;">
                                            <div class="table_wrap">
                                                <div class="ttop">
                                                    <span>�� 0��</span><span class="selected">���õ� ��ǰ 1��</span>
                                                    <button>���û�ǰ ����</button>
                                                </div>
                                                <div class="table">
                                                    <ul class="thead">
                                                        <li><input type="checkbox"></li>
                                                        <li>����ǰ�ڵ�</li>
                                                        <li>��ǥ�̹���/�귣��/��ǰ��</li>
                                                        <li>�߰�����</li>
                                                        <li>��� ������<span>��������/��ü����</span></li>
                                                        <li>�ǸŰ�</li>
                                                        <li>���԰�</li>
                                                        <li>����</li>
                                                        <li>��౸��</li>
                                                        <li></li>
                                                    </ul>
                                                    <div class="tbody_wrap add">
                                                        <ul class="tbody">
                                                            <div>
                                                                + ����ǰ �߰��ϱ�
                                                                <div class="add_btn">
                                                                    <button>����ǰ �űԵ��</button><span></span>
                                                                    <button>���� ����ǰ �ҷ�����</button>
                                                                </div>
                                                            </div>
                                                        </ul>
                                                    </div>
                                                    <div class="tbody_wrap">
                                                        <ul class="tbody">
                                                            <li><input type="checkbox"></li>
                                                            <li>12345678</li>
                                                            <li>
                                                                <img class="prd_img">
                                                                <ul><p class="prd_brand">dailylike</p>
                                                                <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul><p class="badge">����[MDƯ��]</p>
                                                                <p class="info">���ǻ���</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="percent">99%</p>
                                                                    <ul class="bar">
                                                                        <p class="bar_gray"></p>
                                                                        <p class="bar_black"></p>
                                                                    </ul>
                                                                    <p class="">99/100</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="price p_type01"><span>38,000</span></p>
                                                                    <p class="price p_type02">7% ����<span>38,000</span></p>
                                                                    <p class="price p_type03">7% ����<span>38,000</span></p>
                                                                    <p class="price p_type04">15% �÷�������<span>38,000</span></p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="p_type01">38,000</p>
                                                                    <p class="p_type02">23,902</p>
                                                                    <p class="p_type03">23,902</p>
                                                                    <p class="p_type04">23,902</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="p_type01">27%</p>
                                                                    <p class="p_type02">26%</p>
                                                                    <p class="p_type03">26%</p>
                                                                    <p class="p_type04">28%</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                ����
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <button>����</button>
                                                                    <button>����</button>
                                                                </ul>
                                                            </li>
                                                        </ul>
                                                    </div>
                                                    <div class="tbody_wrap">
                                                        <ul class="tbody">
                                                            <li><input type="checkbox"></li>
                                                            <li>12345678</li>
                                                            <li>
                                                                <img class="prd_img">
                                                                <ul><p class="prd_brand">dailylike</p>
                                                                <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul><p class="badge"></p>
                                                                <p class="info"></p></ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="percent"></p>
                                                                    <ul class="bar" style="display:none;">
                                                                        <p class="bar_gray"></p>
                                                                        <p class="bar_black"></p>
                                                                    </ul>
                                                                    <p class="">99</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="price p_type01"><span>38,000</span></p>
                                                                    <p class="price p_type02">7% ����<span>38,000</span></p>
                                                                    <p class="price p_type03">7% ����<span>38,000</span></p>
                                                                    <p class="price p_type04">15% �÷�������<span>38,000</span></p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="p_type01">38,000</p>
                                                                    <p class="p_type02">23,902</p>
                                                                    <p class="p_type03">23,902</p>
                                                                    <p class="p_type04">23,902</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="p_type01">27%</p>
                                                                    <p class="p_type02">26%</p>
                                                                    <p class="p_type03">26%</p>
                                                                    <p class="p_type04">28%</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                ����
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <button>����</button>
                                                                    <button>����</button>
                                                                </ul>
                                                            </li>
                                                        </ul>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                        <!-- �ܼ� �׷� ���� -->
                                        <div class="step_detail group group02_02" style="display:none;">
                                            <div class="tab">
                                                <button class="added on"><span class="sort">�ӤӤ�</span><li>���̾���� ���Ϲ����ص� ��ƼĿ��</li><img src="https://fiximage.10x10.co.kr/web2021/anniv2021/icon_edit.png" alt=""></button>
                                                <button class="added"><span class="sort">�ӤӤ�</span><li>ġƮŰ�� ���޸���</li><img src="https://fiximage.10x10.co.kr/web2021/anniv2021/icon_edit.png" alt=""></button>
                                                <button><li>+�׷� �߰��ϱ�</li></button>
                                            </div>
                                            <!-- �׷Ӿ��� ��ǰ ��ϰ� ����<div class="table_wrap"></div> -->
                                        </div>
                                        <!-- �ݾ׺� �׷� ���� -->
                                        <div class="step_detail group group03_02" style="display:none;">
                                            <div class="tab">
                                                <button class="added on"><li>10,000��~<span>&#65372;</span><span class="gr_cond">���űݾ�</span><span class="gr_name">10,000�� �̻� �����ߴٸ�</span></li><img src="https://fiximage.10x10.co.kr/web2021/anniv2021/icon_edit.png" alt=""></button>
                                                <button class="added"><li>10,000��~<span>&#65372;</span><span class="gr_cond">���űݾ�</span><span class="gr_name">10,000�� �̻� �����ߴٸ�</span></li><img src="https://fiximage.10x10.co.kr/web2021/anniv2021/icon_edit.png" alt=""></button>
                                                <button><li>+�ݾ׺� �׷� �߰��ϱ�</li></button>
                                            </div>
                                            <div class="table_wrap t02">
                                                <div class="ttop">
                                                    <span>�� 0��</span><span class="selected">���õ� ��ǰ 1��</span>
                                                    <button>���û�ǰ ����</button>
                                                </div>
                                                <div class="table">
                                                    <ul class="thead">
                                                        <li><input type="checkbox"></li>
                                                        <li>��ǰ�ڵ�</li>
                                                        <li>��ǥ�̹���/�귣��/��ǰ��</li>
                                                        <li>�߰�����</li>
                                                        <li>��� ������<span>��������/��ü����</span></li>
                                                        <li>��� QR ����/���ϸ��� ����</li>
                                                        <li></li>
                                                    </ul>
                                                    <!-- �׷� �߰� �� -->
                                                    <div class="tbody_wrap none">
                                                        <ul class="tbody">
                                                            �׷��� ���� �߰��ϸ� ����ǰ�� �߰��� �� �־��!
                                                        </ul>
                                                    </div>
                                                    <div class="tbody_wrap add">
                                                        <ul class="tbody">
                                                            <div>
                                                                + ����ǰ �߰��ϱ�
                                                                <div class="add_btn">
                                                                    <button>����ǰ �űԵ��</button><span></span>
                                                                    <button>���� ����ǰ �ҷ�����</button>
                                                                </div>
                                                            </div>
                                                        </ul>
                                                    </div>
                                                    <div class="tbody_wrap">
                                                        <ul class="tbody">
                                                            <li><input type="checkbox"></li>
                                                            <li>12345678</li>
                                                            <li>
                                                                <img class="prd_img">
                                                                <ul><p class="prd_brand">dailylike</p>
                                                                <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul><p class="badge"></p>
                                                                <p class="info">���ǻ���</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="percent">100%</p>
                                                                    <ul class="bar">
                                                                        <p class="bar_gray"></p>
                                                                        <p class="bar_black"></p>
                                                                    </ul>
                                                                    <p class="">100/100</p>
                                                                </ul>
                                                            </li>
                                                        <li>
                                                                <ul>
                                                                    <p class="info">�ٹ�����</p>
                                                                    <p class="coupon"></p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <button>����</button>
                                                                    <button>����</button>
                                                                </ul>
                                                            </li>
                                                        </ul>
                                                    </div>
                                                    <div class="tbody_wrap">
                                                        <ul class="tbody">
                                                            <li><input type="checkbox"></li>
                                                            <li>12345678</li>
                                                            <li>
                                                                <img class="prd_img">
                                                                <ul><p class="prd_brand">dailylike</p>
                                                                <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul><p class="badge"></p>
                                                                <p class="info">���ǻ���</p></ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <p class="percent"></p>
                                                                    <ul class="bar">
                                                                        <p class="bar_gray"></p>
                                                                        <p class="bar_black"></p>
                                                                    </ul>
                                                                    <p class="">50,000/50,000</p>
                                                                </ul>
                                                            </li>
                                                        <li>
                                                                <ul>
                                                                    <p class="info"></p>
                                                                    <p class="coupon">���ʽ����� ���� 3,000</p>
                                                                </ul>
                                                            </li>
                                                            <li>
                                                                <ul>
                                                                    <button>����</button>
                                                                    <button>����</button>
                                                                </ul>
                                                            </li>
                                                        </ul>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div> 
                        </div>
                        <div class="step_wrap">
                            <div class="step_cont detail on">
                                <h3 class="dt_list_top">�ڼ��� ����</h3>
                                <button class="btn_edit">�����ϱ�</button>
                                <li class="noti gray">�ڼ��� �ȳ��ϰ��� �ϴ� ������ �ִٸ� �ۼ����ּ���!</li>
                                <div class="dt_list_wrap">
                                    <div class="info_wrap_list t03">
                                        <textarea>�� 2022 ���̾ ���丮 ��ǰ ���� 20,000�� �̻� ���� �� �����˴ϴ�.&#10;�� ����, ����ī�� �� ��� �� ����Ȯ�� �ݾ� �����Դϴ�.&#10;�� 2022 ���̾ ���丮 ��ǰ ���� 20,000�� �̻� ���� �� �����˴ϴ�.&#10;�� ����, ����ī�� �� ��� �� ����Ȯ�� �ݾ� �����Դϴ�.&#10;�� 2022 ���̾ ���丮 ��ǰ ���� 20,000�� �̻� ���� �� �����˴ϴ�.&#10;�� ����, ����ī�� �� ��� �� ����Ȯ�� �ݾ� �����Դϴ�.&#10;
                                        </textarea>
                                    </div>
                                </div>
                            </div> 
                        </div>
                    </div>
                    <!-- ���̾��˾� ���� : paidGift_lyr on -->
                    <div class="lyr paidGift_lyr">
                        <div class="lyr_overlay"></div>
                     <!-- lyr_wrap lyr01 on -->
                        <!-- �������� ���� ���� �����ϱ� -->
                        <div class="lyr_wrap lyr17 on">
                            <div class="lyr_top">
                                <li>�������� ���� ���� �����ϱ�</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <li class="noti"><span>2022.01.03 - 2022.01.20</span> ����Ǵ� ���������� ������ �������ּ���</li>
                                    <div class="sort_wrap">
                                        <div class="box">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>�긮�� ��Ī��� ���</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                        <div class="box on">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>�������ð����� ��Ī ��� �������� �׽�Ʈ</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>�긮�� ��Ī��� ���</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>�긮�� ��Ī��� ���</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>�긮�� ��Ī��� ���</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
                <script>
                </script>
            </div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->