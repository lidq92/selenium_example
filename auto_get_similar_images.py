# -*- coding: utf-8 -*-
"""
Created on Thu Nov 19 21:21:24 2015

Automatically search and download similar images

@author: Dingquan Li
"""

from selenium import webdriver
import time
import urllib
import re

# Function to create new folder
def mkdir(path):
    import os
    path=path.strip()
    path=path.rstrip("\\")
    isExists=os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        return True
    else:
        return False

browser = webdriver.Firefox()
browser.maximize_window()
browser.get('https://www.google.com.hk/imghp?hl=en&tab=wi')
assert 'Google Images' in browser.title

# Total images
n = 590
# download M simular images
M = 100

# pattern
reg = r'imgurl=(.*?)&imgrefurl'
imgre = re.compile(reg)

# target's xpath
xpath = '//div[@id="rg_s"]/div/a'
				
for k in range(6,n):
    # 1.Search Step
    browser.find_element_by_id('qbi').click()
    image_url = 'http://www02.smt.ufrj.br/~eduardo/ImageDatabase/DatabaseImage%04d.JPG'%(k+1)    
    browser.find_element_by_name('image_url').send_keys(image_url)   
    browser.find_element_by_id('qbbtc').click()
    time.sleep(2)
    browser.find_element_by_link_text('Visually similar images').click()
    # Size larger than 400*300
    browser.find_element_by_xpath('//div[@id="hdtbMenus"]/div/div/div').click()
    browser.find_element_by_id('isz_lt').click()
    browser.find_element_by_id('iszlt_qsvga').click()

    
    # 2.Download Step    
    # Create Folder to storage images
    path = 'G:\TDDOWNLOAD\\prior_images\\image%04d' %(k+1)
    mkdir(path)
    
				
    # record the downloaded images
    img_url_dic = {}
    
    # Simulated rolling window to view and download more pictures
    pos = 0
    m = 0 # figuure number
    i = 0
				
    while(i<20):
		pos += i*500 # Roll step: 500
		js = "document.documentElement.scrollTop=%d" % pos
		browser.execute_script(js)
		time.sleep(1)   
    	
		for element in browser.find_elements_by_xpath(xpath):
			img_url = element.get_attribute('href')
			img_url = imgre.findall(img_url)[0]
			# Save Images
			if img_url != None and not img_url_dic.has_key(img_url):
				img_url_dic[img_url] = ''
				m += 1
				if m <= M:
					try:
						# Download and save image data
						urllib.urlretrieve(img_url,path + '\\image_%04d_%03d.jpg' % (k+1,m))
						print '%04d_%03d' %(k+1,m)
					except:
						m -= 1
		else:
				break
        
		if m>=M:
			break
  
		i += 1
   
browser.quit()
