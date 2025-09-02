import time, requests, json
from datetime import datetime


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

ITSM_TYPE = {
    '故障': 2,
    '配置调整': 4,
    '日常工作登记': 7,
}


class Execl:
    def __init__(self):
        """
            用户token获取地址：
            https://open.feishu.cn/document/uAjLw4CM/ukTMukTMukTM/reference/bitable-v1/app-table-record/search
            推荐获取user_access_token，防止删除或修改数据权限不足
        """
        self.token = 'Bearer'  # 用户token
        self.app_token = ''  # 表格token
        self.table_id = ''  # 表格ID
        self.view_id = ''  # 视图标识符，唯一

    def request_datas(self, request_mode):
        url = f'https://open.feishu.cn/open-apis/bitable/v1/apps/{self.app_token}/tables/{self.table_id}/records/search'
        data = {
            'view_id': self.view_id,
            'field_names': ['发生时间', '归属区域', '巡检', '业务组一级', '业务组二级', '事件内容', '故障原因',
                            '处理过程', '处理结果', '是否已录入']
        }

        header = {
            'Authorization': self.token,
            'Content-Type': 'application/json'
        }
        response = requests.post(url, headers=header, data=json.dumps(data))
        execl_data = response.json()['data']
        data_list = []  # 存储表格信息
        for i in execl_data['items']:
            if i['fields']['是否已录入'] == request_mode:  # 获取录入或未录入的工单
                region = 1  # 归属区域，默认为南基
                inspect = 2  # 是否为巡检，默认为巡检
                # if i['fields']:
                date = int(i['fields']['发生时间']) // 1000
                if i['fields']['归属区域'] == '南沙':
                    region = 2
                if i['fields']['巡检'] == '否':
                    inspect = 1
                data_list.append({'group1': i['fields']['业务组一级'][0]['text'],
                                  'group2': i['fields']['业务组二级'][0]['text'],
                                  'title': self.formatdata(i['fields']['事件内容']),
                                  'process': self.formatdata(i['fields']['处理过程']),
                                  'reason': self.formatdata(i['fields']['故障原因']),
                                  'result': self.formatdata(i['fields']['处理结果']),
                                  'region': region,
                                  'date': datetime.fromtimestamp(date).strftime('%Y-%m-%d'),
                                  'record_id': i['record_id'],
                                  'inspect': inspect,
                                  # 'order_type': ITSM_TYPE[i['fields']['事件类型']]
                                  })
        # print(data_list)
        return data_list

    def formatdata(self, data):  # 格式化内容信息
        foramt = ''
        if len(data) != 1:
            for value in data:
                foramt += value['text']
        else:
            foramt = data[0]['text']
        return foramt

    def del_data(self, record_id):  # 删除飞书表格数据
        url = f'https://open.feishu.cn/open-apis/bitable/v1/apps/{self.app_token}/tables/{self.table_id}/records/{record_id}'

        header = {'Authorization': self.token}

        response = requests.delete(url, headers=header)
        if response.json()['code'] == 0:
            print('数据清理完成')
        else:
            print('清理失败', response.json())

    def update_data(self, record_id):  # 更新表格信息
        url = f'https://open.feishu.cn/open-apis/bitable/v1/apps/{self.app_token}/tables/{self.table_id}/records/{record_id}'
        header = {
            'Authorization': self.token,
            'Content-Type': 'application/json; charset=utf-8'
        }
        data = {
            "fields": {
                "是否已录入": "是"
            }
        }
        response = requests.put(url, headers=header, json=data)
        if response.json()['code'] == 0:
            print('数据更新完成')
        else:
            print('更新失败', response.json())


class Itsm:
    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_experimental_option('detach', True)
        self.chrom = webdriver.Chrome(options=options)
        self.chrom.get('https://dap.gaccloud.com.cn/itsm/sys/main/main.jsp#')
        # 显式等待，直到元素可以点击或交互，最大等待时间为10秒
        self.wait = WebDriverWait(self.chrom, 10)
        self.if_dialog = 1  # 弹窗ID

    def order_mode(self, mode=1):
        self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'if_servicedesk-service-enter')))  # 进入创单html嵌套
        if mode == 1:  # 记录选择工单的选项，事件登记选项
            # self.chrom.find_element(by='xpath', value='//*[@id="regionTree_3_span"]').click()
            self.chrom.find_element(by='xpath', value='//*[@id="regionTree_1_ul"]/li/a[@title="事件登记"]').click()
            time.sleep(1)
            self.chrom.find_element(by='xpath', value='//*[@id="img_0_25"]').click()
        elif mode == 2:  # 巡检登记选项
            # self.chrom.find_element(by='xpath', value='//*[@id="regionTree_4_span"]').click()
            self.chrom.find_element(by='xpath', value='//*[@id="regionTree_1_ul"]/li/a[@title="运维内部流程"]').click()
            time.sleep(1)
            self.chrom.find_element(by='xpath', value='//*[@id="img_0_70"]').click()
        self.chrom.switch_to.default_content()  # 退出html嵌套

    def dialog_box(self, num=1):  # 对话框处理
        for __ in range(0, num):
            self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, f'if_dialog_{self.if_dialog}')))  # 进入html嵌套
            self.if_dialog += 1
            self.wait.until(EC.element_to_be_clickable((By.ID, 'btn_confirm'))).click()  # 确认点击
            self.chrom.switch_to.default_content()  # 退出html嵌套

    def time_click(self, close_time):  # 时间选择
        if close_time.split('-')[-1][0] == '0':
            close_time = close_time.split('-')[-1][1]
        else:
            close_time = close_time.split('-')[-1]
        # 获取时间选择嵌套
        iframe_elent = self.chrom.execute_script("""
                    let iframeId
                    document.querySelectorAll('iframe').forEach(item => {
                            if (!item.id) {
                                iframeId = item
                            }
                        })
                    return iframeId
                    """)
        self.chrom.switch_to.frame(iframe_elent)  # 进入日期选择html嵌套
        for i in range(2, 8):
            try:
                # self.chrom.find_element(by='xpath',value=f'/html/body/div/div[3]/table/tbody/tr[{str(i)}]/td[text()="{close_time}"]').click()  # 点击日期
                if_date = self.chrom.find_element(by='xpath',
                                                  value=f'/html/body/div/div[3]/table/tbody/tr[{str(i)}]/td[text()="{close_time}"]')
                # 判断是否为本月的日期
                if if_date.get_attribute('class') != 'WotherDay':
                    actions = ActionChains(self.chrom)
                    actions.double_click(if_date).perform()  # 双击日期
                    break
            except:
                continue
        self.chrom.switch_to.default_content()  # 退出html嵌套
        time.sleep(1)

    def sign_for(self):  # 获取代签收工单详细
        self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'if_mywork-todo-incident')))  # 进入接单，结单嵌套
        js = "return document.querySelectorAll('#list tbody .ui-widget-content')"
        order_list = self.chrom.execute_script(js)
        if order_list:  # 判断是否有代签收的工单
            for __ in order_list:
                js = "return document.querySelectorAll('.ui-widget-content a')[0]"
                itsm_data = self.chrom.execute_script(js)  # 获取创建的待处理的工单列表
                itsm_data.click()
                self.chrom.switch_to.default_content()  # 退出html嵌套
                time.sleep(3)
                js = """
                    let iframeId
                    document.querySelectorAll('iframe').forEach(item => {
                        if (item.id.includes('if_dealWorkorder')) {
                            iframeId = item.id
                        }
                    })
                    return iframeId
                """
                iframe_id = self.chrom.execute_script(js)  # 获取当前html嵌套的ID
                self.receive(iframe_id)  # 接单
                time.sleep(5)
                self.chrom.switch_to.frame(iframe_id)  # 进入html嵌套
                js = """
                        while (1) {
                            if (document.querySelector('#subject')) {
                                return document.querySelector('#subject').value
                                break
                            }
                        }
                    """
                title = self.chrom.execute_script(js)  # 获取工单的主题
                title_list = title.split('-')
                title = '-'.join([title_list[0], title_list[1], title_list[2]])
                js = """
                    while (1) {
                        if (document.querySelector('#description')) {
                            return document.querySelector('#description').value
                            break
                        }
                    }
                """
                text = self.chrom.execute_script(js)  # 获取工单的描述信息
                self.chrom.switch_to.default_content()  # 退出html嵌套
                text_list = data_dist[title + '&' + text.replace('\n', '')]
                self.complete(text_list[0], text_list[1], text_list[2], text_list[3], iframe_id)  # 结单
                self.chrom.switch_to.default_content()  # 退出html嵌套
                Execl().del_data(text_list[-1])
                time.sleep(3)
                break
            self.sign_for()  # 调用自己

    def add_itsm(self, group1, group2, title, region, date, inspect):  # 创建工单
        self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'if_createWorkorder_25')))  # 进入html嵌套
        self.wait.until(EC.element_to_be_clickable((By.ID, 'Belonging_region-makeupCo'))).click()  # 归属地区元素定位
        # self.chrom.find_element(by='xpath', value='//*[@id="Belonging_region-makeupCo"]').click()  # 归属地区元素定位
        time.sleep(0.5)
        self.chrom.find_element(by='xpath', value=f'//*[@id="Belonging_region"]/option[{region}]').click()  # 归属地区点击
        time.sleep(0.5)
        # 一级业务组定位
        html_group1 = self.chrom.find_element(by='xpath', value='//*[@id="First_level_business_group-makeupCo"]')
        html_group1.click()  # 点击
        html_group1.send_keys(group1)
        time.sleep(1)
        try:
            check1 = self.chrom.find_element(by='xpath',
                                             value=f'//*[@id="First_level_business_group"]/option[text()="{group1}"]')
        except Exception:
            check1 = self.chrom.find_element(by='xpath', value='//*[@id="First_level_business_group"]/option')
        check1.click()  # 选择一级
        time.sleep(1)
        # 二级业务组定位
        html_group2 = self.chrom.find_element(by='xpath', value='//*[@id="Secondary_business_group-makeupCo"]')
        html_group2.click()
        html_group2.send_keys(group2)
        time.sleep(1)
        try:
            check2 = self.chrom.find_element(by='xpath',
                                             value=f'//*[@id="Secondary_business_group"]/option[text()="{group2}"]')
        except Exception:
            check2 = self.chrom.find_element(by='xpath', value='//*[@id="Secondary_business_group"]/option')
        check2.click()  # 选择二级
        self.chrom.find_element(by='xpath', value='//*[@id="Project_administrator"]').click()  # 项目管理员下拉菜单点击
        time.sleep(0.7)
        for i in range(5, 1, -1):
            try:
                self.chrom.find_element(by='xpath',
                                        value=f'//*[@id="Project_administrator"]/option[{i}]').click()  # 选择项目管理员
                time.sleep(0.5)
                # 获取节点管理员
                admin_id = self.chrom.execute_script('''return document.querySelector("#shijiActor_4343 > select").value''')
                if admin_id:  # 判断节点中是否有管理员
                    break
                else:
                    continue
            except Exception:
                # print('event error:', i)
                continue
        self.chrom.find_element(by='xpath', value='//*[@id="FM_1201"]').click()  # 是否服务台转派
        time.sleep(0.7)
        self.chrom.find_element(by='xpath', value='//*[@id="FM_1201"]/option[2]').click()  # 选择服务台转派
        itsm_title = self.chrom.find_element(by='xpath', value='//*[@id="subject"]')  # 工单标题定位
        itsm_title.clear()  # 清空标题内容
        if inspect == 2 or '其他' in group2:
            itsm_title.send_keys(f'{date}-{title}')
        else:
            itsm_title.send_keys(f'{date}-{group2}-{title}')
        itsm_content = self.chrom.find_element(by='xpath', value='//*[@id="description"]')  # 工单描述定位
        itsm_content.clear()  # 清空描述信息
        itsm_content.send_keys(title)
        self.chrom.find_element(by='xpath', value='//*[@id="button1"]').click()  # 提单
        self.chrom.switch_to.default_content()  # 退出html嵌套
        self.dialog_box()  # 创单对话框

    def add_itsm_inspection(self, title, region, date, result):  # 巡检单登记
        self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'if_createWorkorder_70')))  # 进入html嵌套
        self.wait.until(EC.element_to_be_clickable((By.ID, 'Belonging_region-makeupCo'))).click()  # 归属地区元素定位
        time.sleep(1)
        self.chrom.find_element(by='xpath', value=f'//*[@id="Belonging_region"]/option[{region}]').click()  # 归属地区点击
        time.sleep(1)
        title_text = self.chrom.find_element(by='xpath', value='//*[@id="subject"]')  # 标题定位
        title_text.clear()
        title_text.send_keys(f'{date}-{title}')  # 标题输入
        time.sleep(1)
        content_text = self.chrom.find_element(by='xpath', value='//*[@id="description"]')  # 描述定位
        content_text.clear()
        content_text.send_keys(title)  # 描述输入
        result_text = self.chrom.find_element(by='xpath', value='//*[@id="solution"]')
        result_text.send_keys(result)  # 解决方案输入
        time.sleep(1)
        self.chrom.find_element(by='xpath', value='//*[@id="soluttime"]').click()  # 时间选择点击
        self.chrom.switch_to.default_content()  # 退出html嵌套
        time.sleep(1)
        self.time_click(date)  # 巡检单时间选择
        self.chrom.switch_to.frame('if_createWorkorder_70')  # 进入html嵌套
        self.chrom.execute_script('window.scrollTo(0,document.body.scrollHeight)')  # 滚动页面
        self.chrom.find_element(by='xpath', value='//*[@id="button1"]').click()  # 点击提交
        self.chrom.switch_to.default_content()  # 退出html嵌套
        self.dialog_box()  # 巡检对话框
        # self.receive('if_createWorkorder_70')  # 巡检接单
        self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'if_createWorkorder_70')))  # 巡检结单嵌套
        time.sleep(1)
        try:
            self.wait.until(EC.element_to_be_clickable((By.ID, 'incidentsource'))).click()
        except:
            self.wait.until(EC.element_to_be_clickable((By.ID, 'incidentsource'))).click()
        time.sleep(2)
        self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="incidentsource"]/option[3]'))).click()  # 选择事件来源
        # self.chrom.find_element(by='xpath', value='//*[@id="incidentsource"]/option[3]').click()
        self.chrom.execute_script('window.scrollTo(0,document.body.scrollHeight)')  # 滚动页面
        self.wait.until(EC.element_to_be_clickable((By.ID, 'button1'))).click()  # 点击提交
        self.chrom.switch_to.default_content()  # 退出html嵌套
        self.dialog_box(2)  # 巡检结单对话框

    def receive(self, iframe='if_createWorkorder_25'):  # 接单
        self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, iframe)))  # 进入html嵌套
        try:
            btn = self.wait.until(EC.element_to_be_clickable((By.ID, 'button1')))
        except:
            btn = self.wait.until(EC.element_to_be_clickable((By.ID, 'button1')))
        self.chrom.execute_script('window.scrollTo(0,document.body.scrollHeight)')  # 滚动页面
        time.sleep(2)
        btn.click()  # 签收
        self.chrom.switch_to.default_content()  # 退出html嵌套
        self.dialog_box(2)  # 巡检结单对话框

    def complete(self, process, reason, result, close_time, iframe='if_createWorkorder_25'):  # 结单
        self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, iframe)))
        # if close_time.split('-')[-1][0] == '0':
        #     close_time = close_time.split('-')[-1][1]
        # else:
        #     close_time = close_time.split('-')[-1]
        itsm_reason = self.chrom.find_element(by='xpath', value='//*[@id="FM_1093"]')  # 定位故障原因输入框
        itsm_reason.send_keys(reason)  # 故障原因输入
        time.sleep(1)
        itsm_process = self.chrom.find_element(by='xpath', value='//*[@id="processing_process"]')  # 定位处理过程
        itsm_process.send_keys(process)  # 处理过程输入
        time.sleep(1)
        itsm_result = self.chrom.find_element(by='xpath', value='//*[@id="FM_1195"]')  # 定位处理结果
        itsm_result.send_keys(result)  # 处理结果输入
        self.chrom.find_element(by='xpath', value='//*[@id="FM_1094"]').click()  # 点击处理时间
        time.sleep(1)
        self.chrom.switch_to.default_content()  # 退出html嵌套
        self.time_click(close_time)  # 结单时间选择
        # # 获取时间选择嵌套
        # iframe_elent = self.chrom.execute_script("""
        #         let iframeId
        #         document.querySelectorAll('iframe').forEach(item => {
        #                 if (!item.id) {
        #                     iframeId = item
        #                 }
        #             })
        #         return iframeId
        #         """)
        # self.chrom.switch_to.frame(iframe_elent)  # 进入日期选择html嵌套
        # for i in range(2, 8):
        #     try:
        #         # self.chrom.find_element(by='xpath',value=f'/html/body/div/div[3]/table/tbody/tr[{str(i)}]/td[text()="{close_time}"]').click()  # 点击日期
        #         if_date = self.chrom.find_element(by='xpath', value=f'/html/body/div/div[3]/table/tbody/tr[{str(i)}]/td[text()="{close_time}"]')
        #         actions = ActionChains(self.chrom)
        #         actions.double_click(if_date).perform()  # 双击日期
        #         break
        #     except:
        #         continue
        # time.sleep(1)
        # self.chrom.switch_to.default_content()  # 退出html嵌套
        # time.sleep(1)
        self.chrom.switch_to.frame(iframe)  # 进入html嵌套
        itsm_reason.find_element(by='xpath', value='//*[@id="button1"]').click()  # 提交
        self.chrom.switch_to.default_content()  # 退出html嵌套
        self.dialog_box(2)  # 巡检结单对话框


orders_dist = {}  # 保存工单的字典


if __name__ == '__main__':
    # if_dialog = 1  # 弹窗ID
    itsm = Itsm()  # 创单
    itsm_id = input('依次点击：服务登录-->事件登记后回车开始录入工单，或进入待办事件后输入“1”处理待办工单：')
    if not itsm_id:
        itsm_data = Execl().request_datas('否')  # 获取飞书数据
        for orders in itsm_data:
            # self.chrom.switch_to.frame('if_createWorkorder_25')  # 进入html嵌套
            itsm.order_mode(orders['inspect'])  # 创单
            if orders['inspect'] == 1:  # 判断是否为巡检单
                itsm.add_itsm(orders['group1'], orders['group2'], orders['title'], orders['region'], orders['date'],
                              orders['inspect'])
                Execl().update_data(orders['record_id'])  # 更新记录
            else:
                itsm.add_itsm_inspection(title=orders['title'], region=orders['region'], date=orders['date'], result=orders['result'])
                Execl().del_data(orders['record_id'])  # 删除记录
    else:
        itsm_data = Execl().request_datas('是')  # 获取飞书数据
        data_dist = {}
        for i in itsm_data:
            # data_dist[i['title']] = [i['reason'], i['process'], i['result'], i['date'], i['record_id']]
            data_dist[f"{i['date']}&{i['title']}"] = [i['reason'], i['process'], i['result'], i['date'], i['record_id']]
        itsm.if_dialog = 2
        itsm.sign_for()  # 代办处理
    itsm.chrom.close()
