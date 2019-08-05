import unittest
from common.desired_caps import appium_desired
import logging
from time import sleep

from common.start import start_server


class StartEnd(unittest.TestCase):

    def setUpClass(self):
        logging.info('=====setUpClass=====')
        start_server()

    def setUp(self):
        logging.info('=====setUp====')
        self.driver = appium_desired()

    def tearDown(self):
        logging.info('====tearDown====')
        sleep(5)
        self.driver.close_app()
