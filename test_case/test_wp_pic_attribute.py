import logging
import time
from businessView.generalView import GeneralView
from businessView.openView import OpenView
from businessView.wpView import WpView
from common.myunit import StartEnd


class TestWordPictureAttrbute(StartEnd):
    def choose_pic_setup(self):
        ov = OpenView(self.driver)
        ov.open_file('欢迎使用永中Office.docx')
        gv = GeneralView(self.driver)
        gv.switch_write_read()
        wp = WpView(self.driver)
        wp.swipeup()
        wp.choose_pic()

    def test_wp_pic_fixed_rotate(self):
        # 图片旋转
        logging.info('==========test_wp_pic_fixed_rotation==========')
        self.choose_pic_setup()
        gv = GeneralView(self.driver)
        gv.fold_expand()
        wp = WpView(self.driver)
        wp.pic_fixed_rotate()

    def test_wp_pic_change_size(self):
        # 设置图片宽高
        logging.info('==========test_wp_pic_change_size==========')
        self.choose_pic_setup()
        gv = GeneralView(self.driver)
        gv.fold_expand()
        wp = WpView(self.driver)
        wp.pic_change_size()

    def test_wp_pic_shadow(self):
        # 图片阴影
        logging.info('==========test_wp_pic_shadow==========')
        self.choose_pic_setup()
        wp = WpView(self.driver)
        wp.pic_effect()

    def test_wp_pic_broad(self):
        # 图片轮廓
        logging.info('==========test_wp_pic_broad==========')
        self.choose_pic_setup()
        wp = WpView(self.driver)
        wp.pic_broad()

    def test_wp_pic_broad_type(self):
        # 图片轮廓类型
        logging.info('==========test_wp_pic_broad_type==========')
        self.choose_pic_setup()
        wp = WpView(self.driver)
        wp.pic_broad_type()

    def test_wp_pic_broad_width(self):
        # 设置图片轮廓粗细
        logging.info('==========test_wp_pic_broad_width==========')
        self.choose_pic_setup()
        wp = WpView(self.driver)
        wp.pic_broad_width()

    def test_wp_pic_surround(self):
        # 设置图片环绕
        logging.info('==========test_wp_pic_surround==========')
        self.choose_pic_setup()
        wp = WpView(self.driver)
        wp.surround('picture')
        time.sleep(10)

    def test_wp_pic_layer(self):
        # 设置图片叠放次序
        logging.info('==========test_wp_pic_surround==========')
        self.choose_pic_setup()
        wp = WpView(self.driver)
        wp.pic_layer()
        # pass



