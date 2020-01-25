# coding=utf-8
# !/usr/bin/python3.7

from selenium import webdriver
from selenium.webdriver.firefox.options import Options
import time
import datetime
import os
import requests
import schedule
import ConfigParser


class nCov_2019(object):
    '''
    class
    '''

    def get_ini(self, ini_path, section, key):
        '''
        用于返回ini档中key的value
        :param ini_path: ini档案位置，默认与程式放在一起
        :param section: ini档中[]字符串
        :param key: ini档中[]下字符串
        :return: value--ini档中[section]key=后面的字符
        '''

        conf = ConfigParser.ConfigParser()
        conf.readfp(open(ini_path, 'r'))
        value = conf.get(section, key)
        return value

    def load_ini(self):
        '''
        导入ini数据
        :return:
        '''
        ini_path = 'nCov_2019.ini'
        self.robot_url = self.get_ini(ini_path, 'Common', 'Roboturl')
        self.web_url = self.get_ini(ini_path, 'Common', 'weburl')
        self.is_head_less = self.get_ini(ini_path, 'Common', 'headless')
        self.start_time = self.get_ini(ini_path, 'skd', 'start_time')
        self.inter = int(self.get_ini(ini_path, 'skd', 'interval'))

    def init_web_driver(self):
        '''
        初始化网页
        :return:
        '''
        ############################################################
        # 浏览器参数
        # 加载配置
        self.profile = webdriver.FirefoxProfile()
        self.profile.set_preference('browser.download.manager.showWhenStarting', False)
        # browser.download.folderList 设置Firefox的默认 下载 文件夹。
        # 0是桌面；1是“我的下载”；2是自定义
        self.profile.set_preference("browser.download.folderList", 2)
        self.profile.set_preference("browser.helperApps.neverAsk.saveToDisk",
                                    "application/octet-stream, application/vnd.ms-excel, text/csv, application/zip")
        # 无头参数添加
        self.option = Options()
        self.option.add_argument('-headless')

        # 启动时不显示浏览器窗口
        if self.is_head_less == '1':
            self.driver = webdriver.Firefox(firefox_profile=self.profile, firefox_options=self.option)
        else:
            self.driver = webdriver.Firefox(firefox_profile=self.profile)

    def get_infomation(self):
        '''
        获取最新网页图片连接和人数
        :return:
        '''
        self.driver.get(self.web_url)
        time.sleep(2)
        self.img_url = self.driver.find_element_by_class_name('mapImg___3LuBG').get_attribute('src')
        self.update_time = self.driver.find_element_by_class_name('mapTitle___2QtRg').text.split(u'（')[0]
        self.person_count = self.driver.find_element_by_class_name('content___2hIPS').text

    def wx_work(self):
        '''
        发送企业微信图文信息
        :return:
        '''
        self.headers = {"Content-Type": "application/json"}
        # requests.adapters.DEFAULT_RETRIES = 5
        s = requests.session()
        s.keep_alive = False
        self.data = {
            "msgtype": "news",
            "news": {
                "articles": [
                    {
                        "title": "nCov-2019最新进度",
                        "description": "",
                        "url": self.web_url,
                        "picurl": self.img_url
                    },
                    {
                        "title": self.update_time,
                        "description": "",
                        "url": self.web_url,
                        "picurl": self.img_url
                    },
                    {
                        "title": self.person_count,
                        "description": "",
                        "url": self.web_url,
                        "picurl": self.img_url
                    },
                    {
                        "title": '请做好防护措施，戴口罩，勤洗手',
                        "description": "",
                        "url": self.web_url,
                        "picurl": self.img_url
                    }
                ]
            }
        }

        while True:
            now_time = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S.%f')
            try:
                r = requests.post(url=self.robot_url,
                                  headers=self.headers,
                                  json=self.data)
                # print(r.raise_for_status())
                # print('Status:' + str(r.status_code))
                if 'ok' in r.text:
                    r.close()
                    print(now_time + ' Send OK')
                    break
                else:
                    r.close()
                    print(r.text)
                    break
            except (IOError) as ex:
                print(ex.message)
                print(now_time + ' Send Error,process will retry 10s later')
                time.sleep(10)
            finally:
                pass

    def start(self):
        self.init_web_driver()
        self.get_infomation()
        print(self.update_time)
        print(self.person_count)
        self.wx_work()
        self.driver.close()
        self.driver.quit()

    def pre_start(self):
        self.load_ini()
        print('Add schedule...')
        # 添加定时任务
        # 开始时间如果为N，则只运行一次
        if self.start_time == 'N':
            self.start()
        else:
            # 开始时间如果不为N，则看间隔时间
            start_time_hr, start_time_min = self.start_time.split(':')
            time_list = [self.start_time]
            temp_hr = int(start_time_hr)
            str_time_log = ''
            # 将间隔时间加入skd
            while len(time_list) <= int(24 / self.inter):
                temp_time = '{:0>2s}'.format(str(temp_hr)) + ':' + '{:0>2s}'.format(start_time_min)
                time_list.append(temp_time)
                str_time_log += temp_time
                schedule.every().day.at(temp_time).do(self.start)
                temp_hr += self.inter
                if temp_hr >= 24:
                    temp_hr = temp_hr - 24
                print(temp_hr)
            print('min:' + start_time_min)
            # 新建skd
            while True:
                schedule.run_pending()
                time.sleep(1)


if __name__ == '__main__':
    nCov_2019 = nCov_2019()
    nCov_2019.pre_start()
