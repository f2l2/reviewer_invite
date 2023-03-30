import sys
import os
import io
import time
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import warnings
import re

warnings.simplefilter('ignore')

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')  # 改变标准输出的默认编码


class AutoReviewersFind:
    def __init__(self):

        # print('启动中，请稍等')

        # # 不打开浏览器
        option = webdriver.ChromeOptions()
        option.add_argument('headless')
        service = Service(path)
        self.browser = webdriver.Chrome(service=service, chrome_options=option)

        # # # 打开浏览器
        # self.browser = webdriver.Chrome()


    def run_all_function(self):
        def before_add_reviewers():
            self.open_browser()
            self.all_paper_statis()
            # 输入要处理的文章序号
            print('请输入您要处理的文章序号{}-{}:'.format(1, self.paper_all_count))
            sys.stdout.flush()
            paper_get_index = input()
            paper_get_index = int(paper_get_index)

            self.select_process_paper(paper_get_index)

        def add_reviewers():
            self.keywords_search()
            self.reviewers_display()
            sys.stdin.flush()
            add_reviewer_list = list(
                map(int, input('请输入选择的审稿人序号,返回上层菜单输入n:').split()))
            self.reviewers_add(add_reviewer_list)

        self.open_browser()

        # 输入要处理的文章序号




        while True:
            mail_send_state = 'y'
            self.all_paper_statis()
            print('请输入您要处理的文章序号{}-{}:'.format(1, self.paper_all_count))
            sys.stdout.flush()

            sys.stdin.flush()
            paper_get_index = input()
            paper_get_index = int(paper_get_index)
            self.select_process_paper(paper_get_index)

            sys.stdin.flush()
            add_reviewer_flag = input('要增加审稿人吗？y/n,n返回上一菜单,输入end结束程序:')


            if add_reviewer_flag == 'y':
                self.turn_to_page_search()
                while True:

                    search_flag = self.keywords_search()
                    if search_flag == 0:
                        break
                    else:
                        sys.stdin.flush()
                        keywords = input('请输入文章关键词,n返回上一菜单:')
                        if keywords == 'n':
                            break
                        else:
                            while True:
                                reviewers_flag = self.reviewers_display(keywords)
                                if reviewers_flag == 1:
                                    while True:
                                        sys.stdin.flush()
                                        add_reviewer_list = input('请输入选择的审稿人序号,空格分割,返回上层菜单输入n:')

                                        if add_reviewer_list == 'n':
                                            break
                                        else:
                                            try:
                                                add_reviewer_list = add_reviewer_list.split()
                                                add_reviewer_list = [int(num) for num in add_reviewer_list]
                                            except AttributeError:
                                                add_reviewer_list = [int(add_reviewer_list)]
                                            mail_send_state = self.reviewers_add(add_reviewer_list)
                                            break
                                    break
                                else:
                                    break
                    break

                if mail_send_state == 'n':
                    self.browser.find_element(by='xpath', value='//*[@id="table3"]/tbody/tr/td[1]/a[4]').click()
                    self.browser.find_element(by='xpath', value='/html/body/blockquote/form/p[3]/a[3]').click()

                else:
                    self.browser.close()
                    handles = self.browser.window_handles  # 切换回窗口2
                    new_handle = handles[1]
                    self.browser.switch_to.window(new_handle)

                    self.browser.back()
                    self.browser.back()
                    self.browser.back()

            elif add_reviewer_flag == 'end':
                sys.exit(0)
            else:
                self.browser.back()
                self.all_paper_statis()





    def remove_newline(self, my_list):
        i = 0
        if type(my_list) == list:
            while i < len(my_list):
                if my_list[i] == '\n':
                    my_list.pop(i)
                else:
                    i += 1
        elif type(my_list) == str:
            my_list = my_list.replace('\n', '')
        return my_list

    def open_browser(self):

        # 读取网页


        link1 = 'https://css.paperplaza.net/conferences/scripts/start.pl'
        self.browser.get(link1)
        time.sleep(1)
        self.browser.execute_script("window.scrollTo(0, 500);")
        self.browser.find_element(by='xpath', value='//*[@id="conftable"]/tbody/tr[5]/td[4]/ul/li/a').click()
        sys.stdin.flush()
        pin = input('Please input your pin:')
        self.browser.find_element(by='xpath', value='//*[@id="pin"]').send_keys(pin)
        sys.stdin.flush()
        psw = input('please input your password:')
        self.browser.find_element(by='xpath', value='/html/body/table[2]/tbody/tr[1]/td[3]/form/div[2]/div[3]/div[2]/table/tbody/tr[3]/td[2]/input').send_keys(psw)
        time.sleep(1)
        self.browser.find_element(by='xpath', value='/html/body/table[2]/tbody/tr[1]/td[3]/form/div[2]/div[3]/div[2]/table/tbody/tr[4]/td[2]/a').click()
        self.browser.find_element(by='xpath', value='//*[@id="table1"]/tbody/tr[4]/td[2]/a').click()

        handles = self.browser.window_handles  # 弹出新窗口
        new_handle = handles[1]
        self.browser.switch_to.window(new_handle)
        self.browser.find_element(by='xpath', value='//*[@id="x1"]/a').click()
        self.browser.find_element(by='xpath', value='/html/body/blockquote/form/p[3]/a[3]').click()




    def all_paper_statis(self):

        # 解析网页
        html_doc = self.browser.page_source
        soup = BeautifulSoup(html_doc, 'lxml')
        head = soup.find_all('tbody')[2]
        self.paper_all_count = len(head.select('.line'))
        print('现在总共有{}篇文章在处理'.format(self.paper_all_count))
        head_list = head.contents
        paper_list = self.remove_newline(head_list)
        # 获取文章标题
        print('获取文章标题中')
        for i in range(self.paper_all_count):
            paper_title = paper_list[i].get_text().split('\xa0')[-1][1:-2]
            print('文章{}的标题是{}\n'.format(i+1, paper_title))


    def select_process_paper(self, paper_get_index):
        # 进入处理文章


        # 测试用
        # paper_get_index = 1
        print('处理中，请等待.....')
        paper_reviewer_xpath = '//*[@id="tblSort"]/tbody[1]/tr[{}]/td[3]/a'.format(paper_get_index)
        self.browser.find_element(by='xpath', value=paper_reviewer_xpath).click()
        reviewer_html_doc = self.browser.page_source
        reviewer_soup = BeautifulSoup(reviewer_html_doc, 'lxml')
        reviewer_head = reviewer_soup.find_all('tbody')[3]
        reviewer_head_list = reviewer_head.contents
        reviewer_head_list = self.remove_newline(reviewer_head_list)
        # 读取目前审稿人列表
        reviewer_count_all = int(reviewer_head_list[-1].select(('.c'))[2].get_text())
        print('这篇文章现在总共有{}名审稿人'.format(reviewer_count_all))
        confirmed_reviewers_count = 0
        for i in range(reviewer_count_all):
            statue = reviewer_head_list[i+3].get_text().replace('\n', '')[-53:-2]
            if '+' not in statue:
                statue = statue[1:-1]
            print('第{}名审稿人的状态是：{}\n'.format(i, statue))
            if statue.find('Confirmed') != -1:
                confirmed_reviewers_count = confirmed_reviewers_count + 1
        print('目前共有{}名审稿人Confirmed'.format(confirmed_reviewers_count))


    def turn_to_page_search(self):
        print('搜索审稿人中....')
        wait = WebDriverWait(self.browser, 5)
        element = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/blockquote/form/p[4]/a[1]")))
        element.click()

        wait = WebDriverWait(self.browser, 5)
        element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tblSort"]/thead/tr[2]/td/a[1]')))
        element.click()

        self.browser.find_element(by='xpath', value='//*[@id="legendhandle"]/img').click()

        wait = WebDriverWait(self.browser, 10)
        element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tblSort"]/thead/tr[2]/td/ul/li[1]/span')))

        element.click()

        handles = self.browser.window_handles  # 弹出新窗口
        new_handle = handles[2]
        self.browser.switch_to.window(new_handle)

    def keywords_search(self):
        try:
            # 根据关键词搜索审稿人
            sel = self.browser.find_element(by='xpath', value="/html/body/blockquote/form/table/tbody/tr[2]/td[2]/select")
            Select(sel).select_by_visible_text('Keywords')
            search_flag = 1
        except NoSuchElementException:
            # 如果找不到指定元素，则执行以下操作
            print("指定摘要没用审稿人,请重新输入:")
            search_flag = 0
        return search_flag


    def reviewers_display(self, keywords):

        self.browser.find_element(by='xpath', value='//*[@id="sfor"]').clear()
        self.browser.find_element(by='xpath', value='//*[@id="sfor"]').send_keys(keywords)
        self.browser.find_element(by='xpath', value='/html/body/blockquote/form/table/tbody/tr[4]/td[2]/input').click()
        # 一页显示200个审稿人
        sel = self.browser.find_element(by='xpath', value="/html/body/blockquote/form/table/tbody/tr[6]/td[2]/span/select")
        Select(sel).select_by_visible_text('200')
        # 爬取审稿人
        add_reviewer_html_doc = self.browser.page_source
        add_reviewer_soup = BeautifulSoup(add_reviewer_html_doc, 'lxml')
        add_reviewer_head = add_reviewer_soup.find_all('tbody')[5]
        add_reviewer_head_list = add_reviewer_head.contents
        add_reviewer_head_list = self.remove_newline(add_reviewer_head_list)
        add_reviewer_output_title = ['序号', 'Name', 'Title', 'Affiliation', 'E-mail']
        print("{:<3}\t{:<35}\t{:<5}\t{:<70}\t{}".format(*add_reviewer_output_title))

        for i in range(len(add_reviewer_head_list)):
            try:
                add_reviewer_head_text = add_reviewer_head_list[i].get_text()
                add_reviewer_head_text = self.remove_newline(add_reviewer_head_text).split('\xa0')

                if self.browser.find_element(by='xpath', value='//*[@id="tblSort"]/tbody[1]/tr[{}]/td[2]/input'.format(i+1)).is_enabled():
                    if len(add_reviewer_head_text) == 8:
                        print("{:<3}\t{:<35}\t{:<5}\t{:<70}\t{}".format(*[i+1, add_reviewer_head_text[3], add_reviewer_head_text[4], add_reviewer_head_text[5], add_reviewer_head_text[6]]))
                        sys.stdout.flush()
                    else:
                        print("{:<3}\t{:<35}\t{:<5}\t{:<70}\t{}".format(
                            *[i + 1, add_reviewer_head_text[1], add_reviewer_head_text[2], add_reviewer_head_text[3],
                              add_reviewer_head_text[4]]))
                        sys.stdout.flush()

                # 添加审稿人

            except NoSuchElementException:
                # 如果找不到指定元素，则执行以下操作
                print("审稿人不可用")
                search_flag = 0

        print('以上为可选择的审稿人,如果没有请输入n')
        sys.stdout.flush()
        search_flag = 1

        return search_flag

    def reviewers_add(self, add_reviewer_list):
        judge_final = input('对选择的审稿人确认吗？y/n:')
        if judge_final == 'y':
            if len(add_reviewer_list) > 1:
                for i in range(len(add_reviewer_list)):
                    self.browser.find_element(by='xpath', value='//*[@id="tblSort"]/tbody[1]/tr[{}]/td[2]/input'.format(add_reviewer_list[i])).click()  # 勾选审稿人
            else:
                    self.browser.find_element(by='xpath', value='//*[@id="tblSort"]/tbody[1]/tr{}/td[2]/input'.format(
                        add_reviewer_list)).click()
            self.browser.find_element(by='xpath', value='/html/body/blockquote/form/p[5]/table[2]/tbody/tr/td[2]/input').click()  # 确定添加到列表里
            time.sleep(1)
            add_feedback_str = self.browser.find_element(by='xpath', value='//*[@id="msg"]/div[2]').text[:-3]
            print(add_feedback_str, flush=True)  # 添加完审稿人,获取反馈
            print('添加完毕', flush=True)


            reviewer_id = []
            reviewer_name_list = []
            number_id = []
            name_id = []

            if '\n' in add_feedback_str:
                add_feedback_list = add_feedback_str.split('\n')
                for i in range(len(add_feedback_list)):
                    number_id = []
                    name_id = []
                    for j in range(len(add_feedback_list[i].split())):

                        if len(number_id) == 0:
                            number_id = re.findall(r'\d+', add_feedback_list[i].split()[j])
                            name_id.append(add_feedback_list[i].split()[j])
                        else:
                            break
                    reviewer_name = ' '.join(name_id[:-1])
                    reviewer_name_list.append(reviewer_name)
                    reviewer_id.append('include' + str(number_id[0]))
            else:

                for j in range(len(add_feedback_str.split())):

                    if len(number_id) == 0:
                        number_id = re.findall(r'\d+', add_feedback_str.split()[j])
                        name_id.append(add_feedback_str.split()[j])
                    else:
                        break
                reviewer_id.append('include' + str(number_id[0]))
                reviewer_name = ' '.join(name_id[:-1])
                reviewer_name_list.append(reviewer_name)
                # reviewer_id = 'include' + add_feedback_str.split()[2][1:-1]
            # value = 'include' + str[-53:-47]
            #


            # 将添加后的审稿人选中,进行审稿工作
            # self.browser.close()
            handles = self.browser.window_handles  # 切换回窗口2
            new_handle = handles[1]
            self.browser.switch_to.window(new_handle)
            self.browser.back()
            self.browser.refresh()


            if len(reviewer_id) > 1:
                for i in range(len(reviewer_id)):
                    self.browser.find_element(by='id', value=reviewer_id[i]).click()
            else:
                self.browser.find_element(by='id', value=reviewer_id[0]).click()

            element = self.browser.find_element_by_xpath('//*[@id="requestlistwindow"]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/input')  # 定位需要点击的元素
            self.browser.execute_script("arguments[0].click();", element)  # 执行点击操作

            element = self.browser.find_element_by_xpath('/html/body/blockquote/form/table/tbody/tr[4]/td[2]/input')  # 定位需要点击的元素
            self.browser.execute_script("arguments[0].click();", element)  # 执行点击操作

            sys.stdin.flush()
            judge_final = input('对选择的审稿人确认吗？确认后即发送邮件邀请y/n:')
            if judge_final == 'y':
                self.browser.find_element(by='xpath', value='//*[@id="snd"]').click()
                print('审稿邀请邮件已发送')
                sys.stdout.flush()
                mail_send_state = 'y'
                return mail_send_state
            else:

                print('返回开始菜单', flush=True)
                mail_send_state = 'n'
                self.browser.back()
                self.browser.find_element(by='xpath', value='//*[@id="tblSort"]/thead/tr[2]/td/a[1]').click()
                self.browser.find_element(by='xpath', value='//*[@id="legendhandle"]/img').click()


                # 将刚才添加的审稿人删除掉
                for i in range(len(reviewer_id)):

                    self.browser.find_element(by='link text', value=reviewer_name_list[i]).click()
                    self.browser.find_element(by='xpath', value='//*[@id="updatewindow"]/table/tbody/tr[8]/td[2]/input').click()
                    # self.browser.find_element(by='xpath', value='//*[@id="updatewindow"]/table/tbody/tr[8]/td[2]/input').click()

                    alert = self.browser.switch_to.alert
                    alert.accept()
                    time.sleep(1)

                return mail_send_state





if hasattr(sys, '_MEIPASS'):
    path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
else:
    path = 'chromedriver.exe'


print('启动中')
sys.stdout.flush()
AutoReviewersFind().run_all_function()
# switched linear systems
