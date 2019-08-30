import logging
import time
from businessView.generalView import GeneralView
from businessView.openView import OpenView
from businessView.wpView import WpView
from common.myunit import StartEnd
from airtest.core.api import *


class TestWordShapeAttrbute(StartEnd):
    def shapeatt_setup(self):
        ov = OpenView(self.driver)
        ov.open_file('欢迎使用永中Office.docx')
        gv = GeneralView(self.driver)
        gv.switch_write_read()
        wp = WpView(self.driver)
        wp.switch_option('插入')
        wp.insert_text_box()

    def test_wp_shape_copy_paste(self):
        # 形状复制粘贴
        logging.info('==========test_wp_shape_copy_paste==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        connect_device('Android:///?cap_method=JAVACAP&&ori_method=ADBORI&&touch_method=ADBTOUCH')

        wp.object_copy_paste()
        time.sleep(10)

    def test_wp_shape_cut_paste(self):
        # 形状剪切粘贴
        logging.info('==========test_wp_shape_cut_paste==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        connect_device('Android:///?cap_method=JAVACAP&&ori_method=ADBORI&&touch_method=ADBTOUCH')

        wp.object_cut_paste()
        time.sleep(10)

    def test_wp_shape_delete(self):
        # 形状 删除
        logging.info('==========test_wp_shape_delete==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        connect_device('Android:///?cap_method=JAVACAP&&ori_method=ADBORI&&touch_method=ADBTOUCH')

        wp.object_delete()
        time.sleep(10)

    def test_wp_shape_rotate_90(self):
        # 形状旋转90度
        logging.info('==========test_wp_shape_rotate_90==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        connect_device('Android:///?cap_method=JAVACAP&&ori_method=ADBORI&&touch_method=ADBTOUCH')

        wp.object_rotate_90()
        time.sleep(10)

    def test_wp_shape_free_rotate(self):
        # 形状自由旋转
        logging.info('==========test_wp_shape_free_rotate==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        connect_device('Android:///?cap_method=JAVACAP&&ori_method=ADBORI&&touch_method=ADBTOUCH')

        wp.object_free_rotate()
        time.sleep(10)

    def test_wp_shape_control_point(self):
        # 手势拖拉形状控制点
        logging.info('==========test_wp_shape_control_point==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        connect_device('Android:///?cap_method=JAVACAP&&ori_method=ADBORI&&touch_method=ADBTOUCH')

        wp.shape_control_point()
        time.sleep(10)

    def test_wp_shape_text_select(self):
        # 文本框内容选取
        logging.info('==========test_wp_shape_text_select==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        connect_device('Android:///?cap_method=JAVACAP&&ori_method=ADBORI&&touch_method=ADBTOUCH')
        wp.text_box_text_select()
        self.assertTrue(exists(Template(r'../Res/res_delete.png', resolution=(1080, 1920))))

    def test_wp_shape_move(self):
        # 文本框移动
        logging.info('==========test_wp_shape_move==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        connect_device('Android:///?cap_method=JAVACAP&&ori_method=ADBORI&&touch_method=ADBTOUCH')
        wp.text_box_move()
        time.sleep(10)

    def test_wp_shape_fixed_rotate(self):
        # 形状旋转
        logging.info('==========test_wp_shape_fixed_rotate==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        wp.shape_fixed_rotate()

    def test_wp_shape_change_size(self):
        # 设置形状宽高
        logging.info('==========test_wp_shape_change_size==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        wp.shape_chang_size()

    def test_wp_shape_text_box_margin(self):
        # 设置文本框边距
        logging.info('==========test_wp_shape_margin_text_box==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        wp.text_box_margin()

    def test_wp_shape_fill_color(self):
        # 设置形状填充色及透明度
        logging.info('==========test_wp_shape_fill_color==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        self.assertTrue(wp.shape_fill_color(), msg='Filling color transparency fail')

    def test_wp_shape_broad(self):
        # 形状轮廓
        logging.info('==========test_wp_shape_broad==========')
        self.shapeatt_setup()
        gv = GeneralView(self.driver)
        gv.fold_expand()
        wp = WpView(self.driver)
        wp.shape_broad()

    def test_wp_shape_type(self):
        # 形状轮廓类型
        logging.info('==========test_wp_shape_broad==========')
        self.shapeatt_setup()
        gv = GeneralView(self.driver)
        gv.fold_expand()
        wp = WpView(self.driver)
        wp.shape_broad_type()

    def test_wp_shape_broad_width(self):
        # 设置形状轮廓粗细
        logging.info('==========test_wp_shape_broad==========')
        self.shapeatt_setup()
        gv = GeneralView(self.driver)
        gv.fold_expand()
        wp = WpView(self.driver)
        wp.shape_broad_width()

    def test_wp_shape_shadow(self):
        # 设置形状阴影和三维效果
        logging.info('==========test_wp_shape_broad==========')
        self.shapeatt_setup()
        gv = GeneralView(self.driver)
        gv.fold_expand()
        wp = WpView(self.driver)
        wp.shape_effect()

    def test_wp_shape_surround(self):
        # 设置形状文字环绕效果
        logging.info('==========test_wp_shape_surround==========')
        self.shapeatt_setup()
        gv = GeneralView(self.driver)
        gv.fold_expand()
        wp = WpView(self.driver)
        wp.surround('shape')

    def test_wp_shape_layer(self):
        # 设置形状叠放次序
        logging.info('==========test_wp_shape_surround==========')
        self.shapeatt_setup()
        wp = WpView(self.driver)
        gv = GeneralView(self.driver)
        gv.switch_write_read()
        gv.switch_write_read()
        wp.switch_option('插入')
        wp.insert_text_box()
        gv.fold_expand()
        wp.shape_layer()
