#!/usr/bin/python
# _*_ coding: utf-8 _*_

"""
Created on 2020/6/17
Springer has released 65 Machine Learning and Data Books for Free.
The script downloads the 65 books automatically.
Note: Introduction to Time Series and Forecasting (The 27th and 28th books are the same)
      Robotics, Vision and Control (The 18th book is an updated version of the 54th book)
      Introduction to Partial Differential Equations (The 7th book is not the same as the 10th book)
@author: Dingquan Li
"""

from selenium import webdriver
import urllib
# import re
import os

import ssl
ssl._create_default_https_context = ssl._create_unverified_context

path = './SpringerFreeBooks'
if not os.path.exists(path):
    os.makedirs(path)

browser = webdriver.Firefox()
browser.maximize_window()
browser.get('https://ifors.org/developing_countries/index.php/Springer_has_released_65_Machine_Learning_and_Data_Books_for_Free')

# target's xpath
xpath1 = '//a[@class="external free"]'
xpath2 = '//a[@data-track-action="Book download - pdf"]'
# reg = r'http://link.springer.com/openurl*'
# bookre = re.compile(reg)
for k, element in enumerate(browser.find_elements_by_xpath(xpath1)):
	book_url = element.get_attribute('href')
	if 'http://link.springer.com/openurl' in book_url:
		browser1 = webdriver.Firefox()
		browser1.maximize_window()
		browser1.get(book_url)
		pdf_url = browser1.find_elements_by_xpath(xpath2)[0].get_attribute('href')
		# print(pdf_url)
		name = browser1.title[:-15].encode('utf8')
		print(name)
		if not os.path.exists('{}/{}.{}.pdf'.format(path, k+1, name)):
			urllib.urlretrieve(pdf_url, '{}/{}.{}.pdf'.format(path, k+1, name))
		browser1.quit()

browser.quit()
