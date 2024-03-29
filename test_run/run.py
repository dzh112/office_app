import HTMLTestRunner
import os
import unittest
# from  BSTestRunner import BSTestRunner
import time, logging
import sys


path = 'C:/PycharmProjects/office_test_project'
sys.path.append(path)

test_dir = '../test_case'
report_dir = '../reports'

discover = unittest.defaultTestLoader.discover(test_dir, pattern='test_homepage.py')
now = time.strftime('%Y-%m-%d %H_%M_%S')
report_name = report_dir + '/' + now + 'Mobile_Office_Report.html'
with open(r'%s' % report_name, 'wb') as f:
    runner = HTMLTestRunner.HTMLTestRunner(stream=f, title='YOZO_Mobile_Office_Report',
                                           description='yozo Android app test report')
    logging.info('start run test case...')
    runner.run(discover)
