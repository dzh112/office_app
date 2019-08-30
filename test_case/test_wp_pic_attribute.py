import logging
import time
import unittest

from appium.webdriver.common.touch_action import TouchAction

from businessView.generalView import GeneralView
from businessView.openView import OpenView
from businessView.wpView import WpView
from common.myunit import StartEnd
from airtest.core.api import *


class TestWordPictureAttrbute(StartEnd):
    def choose_pic_setup(self):
        ov = OpenView(self.driver)
        ov.open_file('欢迎使用永中Office.docx')
        time.sleep(3)
        gv = GeneralView(self.driver)
        gv.switch_write_read()
        wp = WpView(self.driver)
        wp.swipeup()
        wp.choose_pic()

    @unittest.skip('skip test_undo_redo')
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
        time.sleep(5)

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

    def test_wp_pic_move(self):
        # 设置图片叠放次序
        logging.info('==========test_wp_pic_move==========')
        ov = OpenView(self.driver)
        ov.open_file('欢迎使用永中Office.docx')
        time.sleep(3)
        gv = GeneralView(self.driver)
        gv.switch_write_read()
        wp = WpView(self.driver)
        wp.swipeup()
        a = wp.choose_pic()
        wp.surround_peripheral()
        self.driver.keyevent(4)
        self.driver.keyevent(4)
        time.sleep(2)
        # 移动图片位置
        self.driver.swipe(a[0], a[1], 0, 0)
        time.sleep(10)

    @staticmethod
    def adjust_pic_place():
        # 调整图片位置
        connect_device('Android:///?cap_method=JAVACAP&&ori_method=ADBORI&&touch_method=ADBTOUCH')
        while not exists(Template(r'../Res/res_object.png', resolution=(1080, 1920))):
            swipe([20, 500], [20, 800])

    def test_wp_pic_free_rotate(self):
        # 图片自由旋转
        logging.info('==========test_wp_pic_free_rotate==========')
        self.choose_pic_setup()
        res_object = r'../Res/res_object.png'
        res_pic_option = r'../Res/res_pic_option.png'
        self.adjust_pic_place()
        swipe(Template(res_object, resolution=(1080, 1920)), Template(res_pic_option, resolution=(1080, 1920)))
        time.sleep(10)

    def test_wp_pic_save_album(self):
        # 保存图片至相册
        logging.info('==========test_wp_pic_save_album==========')
        self.choose_pic_setup()
        res_object = r'../Res/res_object.png'
        res_save_to_album = r'../Res/res_save_to_album.png'
        self.adjust_pic_place()
        touch(Template(res_object, resolution=(1080, 1920)))

        touch(Template(res_save_to_album, resolution=(1080, 1920)))
        self.assertTrue(WpView(self.driver).get_toast_message('图片保存成功'), 'picture save to album fail')

    def test_wp_pic_rotate_90(self):
        # 图片旋转90度
        logging.info('==========test_wp_pic_rotate_90==========')
        self.choose_pic_setup()
        res_object = r'../Res/res_object.png'
        res_rotate_90 = r'../Res/res_rotate_90.png'
        self.adjust_pic_place()
        touch(Template(res_object, resolution=(1080, 1920)))

        touch(Template(res_rotate_90, resolution=(1080, 1920)))
        time.sleep(10)

    def test_wp_pic_cut_paste(self):
        logging.info('==========test_wp_pic_cut_paste==========')
        self.choose_pic_setup()
        res_object = r'../Res/res_object.png'
        res_cut = r'../Res/res_cut.png'
        res_paste = r'../Res/res_paste.png'
        self.adjust_pic_place()
        touch(Template(res_object, resolution=(1080, 1920)))

        # 剪切 粘贴
        touch(Template(res_cut, resolution=(1080, 1920)))
        wp = WpView(self.driver)
        wp.choose_pic()
        touch(Template(res_object, resolution=(1080, 1920)))
        touch(Template(res_paste, resolution=(1080, 1920)))
        time.sleep(10)

    def test_wp_pic_copy_paste(self):
        logging.info('==========test_wp_pic_rotate_90==========')
        self.choose_pic_setup()
        res_object = r'../Res/res_object.png'
        res_copy = r'../Res/res_copy.png'
        res_paste = r'../Res/res_paste.png'
        self.adjust_pic_place()
        touch(Template(res_object, resolution=(1080, 1920)))

        # 复制 粘贴
        touch(Template(res_copy, resolution=(1080, 1920)))
        touch(Template(res_object, resolution=(1080, 1920)))
        touch(Template(res_paste, resolution=(1080, 1920)))
        time.sleep(10)

    def test_wp_pic_delete(self):
        logging.info('==========test_wp_pic_rotate_90==========')
        self.choose_pic_setup()
        res_object = r'../Res/res_object.png'
        res_delete = r'../Res/res_delete.png'
        self.adjust_pic_place()
        touch(Template(res_object, resolution=(1080, 1920)))

        # 删除图片
        touch(Template(res_delete, resolution=(1080, 1920)))
        time.sleep(10)

    def test_wp_pic_control_point(self):
        logging.info('==========test_wp_pic_control_point==========')
        self.choose_pic_setup()
        res_control_point = r'../Res/res_control_point.png'
        res_pic_option = r'../Res/res_pic_option.png'
        self.adjust_pic_place()
        # 手势拖拉大小控制点
        swipe(Template(res_control_point, resolution=(1080, 1920)), Template(res_pic_option, resolution=(1080, 1920)))
        time.sleep(10)
