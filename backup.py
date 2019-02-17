from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
import time 
from selenium.webdriver.common.action_chains import ActionChains
import xlwt
import pandas
from openpyxl import load_workbook
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf
import xlsxwriter
from selenium.webdriver.common.keys import Keys

facebook =" "
twitter =" "
linkedin =" "
googleplus =" "
instagram=" "
skype=" "
category_count=1
area_count = 1
area_count1 =2
profile_count = 1
profile_count1 = 1
category=""
f_name=""
l_name=""
social = True
tagline_count =1
category_l_part = ')>a'
lu = '#practice > div > div.col-lg-7 > div.global.clearfix > div > div.social-icons > a:nth-child('
ll = ')'
ll_count = 1
area_ff_part = ') > li:nth-child('

tagline = ' '
area_f_part = '#fbl-locality-explorer > ul:nth-child('
category_f_part = '#therapy-reference>li:nth-child('
area_l_part = ') > a'

while(True):
	category_boolean = True
	path_to_chromedriver = 'C:\Python27\chromedriver'
	browser = webdriver.Chrome(executable_path = path_to_chromedriver)
	url = 'https://www.annuaire-therapeutes.com/index-des-therapies'
	browser.get(url)

	try:
		category_f_part = '#therapy-reference>li:nth-child('
		category_l_part = ')>a'
		category = browser.find_element_by_css_selector(category_f_part + str(category_count) + category_l_part).text
		browser.find_element_by_css_selector(category_f_part + str(category_count) + category_l_part).click()
		#area_part_ul = True
		#categoryurl = browser.current_url
		try:
			try:
					#browser.get(categoryurl)
				if profile_count < 2:
					try:
						profile_boolean = True
						area_f_part = '#fbl-locality-explorer > ul:nth-child('
						area_ff_part = ') > li:nth-child('
						area_l_part = ') > a'
						area = browser.find_element_by_css_selector(area_f_part + str(area_count1) + area_ff_part + str(area_count) + area_l_part).text
						browser.find_element_by_css_selector(area_f_part + str(area_count1) + area_ff_part + str(area_count) + area_l_part).click()
						areaurl = browser.current_url
					except Exception as e:
						raise e
				else:

					profile_boolean = True
					#loop1 = 0
					area_f_part = '#fbl-locality-explorer > ul:nth-child('
					area_ff_part = ') > li:nth-child('
					area_l_part = ') > a'
					area = browser.find_element_by_css_selector(area_f_part + str(area_count1) + area_ff_part + str(area_count) + area_l_part).text
					browser.find_element_by_css_selector(area_f_part + str(area_count1) + area_ff_part + str(area_count) + area_l_part).click()
					areaurl = browser.current_url
				while(profile_boolean):
					try:


						profile_f_part = '#therapists-listing > div.results.clearfix > div:nth-child('
						profile_l_part = ') > div > a.button.ui-default2-inset'
						element = browser.find_element_by_css_selector(profile_f_part + str(profile_count) + profile_l_part)
						actions = ActionChains(browser)
						actions.move_to_element(element).perform()
						browser.find_element_by_css_selector(profile_f_part + str(profile_count) + profile_l_part).click()
						print category
						tagline_boolean = True
						social_boolean = True

						try:
							name = browser.find_element_by_css_selector('#practice > div > div.col-lg-7 > div.global.clearfix > div > h1').text
							f_name,l_name = name.split(' ')
							print f_name + '*' + l_name + '*'
						except Exception as e:
							print '***'

						try:
							browser.find_element_by_css_selector('#practice > div > div.col-lg-7 > div.global.clearfix > div > div.user-controls > a:nth-child(2)').click()
							time.sleep(2)
							phone_number = browser.find_element_by_css_selector('#practice > div > div.col-lg-7 > div.global.clearfix > div > div.phone-number > div').text
							phone_number.strip(" ")
							print phone_number
						except Exception as e:
							print '***'

						try:
							address = browser.find_element_by_css_selector('#practice > div > div.col-lg-7 > div:nth-child(7) > div').text
							#print address + '*'
							street_address,cityfull,country = address.split(',')
							pin,city = cityfull.split(' ',1)
							print pin +'*' + city + '*' + street_address + 'country' + '*'
						except Exception as e:
							print '***'
							
						while(True):
							try:
								tagline_f_part = '#practice-therapies > li:nth-child('
								tagline_l_part = ') > span'
								tagline = tagline + browser.find_element_by_css_selector(tagline_f_part + str(tagline_count) + tagline_l_part).text
								tagline = tagline + ','

								tagline_count=tagline_count+1

							except Exception as e:
								tagline_count = 1
								#tagline_boolean = False
								print '***'
								break
						print tagline + '*'

						try:
							price =  browser.find_element_by_css_selector('#practice > div > div.col-lg-7 > div.part.pricings > ul > li').text
							print price + '*'
						except Exception as e:
							print '***'

						try:
							user_link = browser.find_element_by_css_selector('#practice > div > div.col-lg-7 > div.global.clearfix > div > div.user-controls > a:nth-child(2)').get_attribute("href")
							print user_link + '*'
						except Exception as e:
							print '***'	

						while(social_boolean):
							try:
								(main_window) = browser.current_window_handle
								userurl = browser.current_url
								browser.find_element_by_css_selector(lu+str(ll_count)+ll).send_keys(Keys.CONTROL + Keys.RETURN)
								time.sleep(2)
								browser.switch_to.window(browser.window_handles[-1])
								temp =  browser.current_url
								browser.switch_to.window(main_window)
								browser.get(userurl)

								#browser..send_keys(Keys.CONTROL + 'w')
								#browser.switch_to.window(browser.window_handles[main_window])
								ll_count = ll_count +1
								print temp

								if temp.find('facebook') >= 0:
									facebook = temp
								elif temp.find('twitter') >= 0:
									twitter = temp
								elif temp.find('google') >= 0:
									googleplus = temp
								elif temp.find('linkedin') >= 0:
									linkedin = temp
								elif temp.find('instagram') >= 0:
									instagram = temp
								elif temp.find('skype') >= 0:
									skype = temp
								else:
									print ''
								print facebook
								#browser.refresh()
							except Exception as e:
								print e
								social_boolean = False
								ll_count = 1
								

						try:
							element = browser.find_element_by_css_selector('#display-description')
							actions = ActionChains(browser)
							actions.move_to_element(element).perform()
							browser.find_element_by_css_selector('#display-description').click()
							description = browser.find_element_by_css_selector('#practice > div > div.col-lg-7 > div.part.description > p').text
							print description + '*'
						except Exception as e:
							print '***'

						try:
							book_ro = open_workbook('Data2.xls')
							book = copy(book_ro)
							sheet1 = book.get_sheet(0)
							sheet1.write(profile_count1, 0, category)
							sheet1.write(profile_count1, 1, f_name)
							sheet1.write(profile_count1, 2, l_name)
							sheet1.write(profile_count1, 3, tagline)
							sheet1.write(profile_count1, 4, price)
							sheet1.write(profile_count1, 5, phone_number)
							sheet1.write(profile_count1, 6, user_link)
							sheet1.write(profile_count1, 7, country)
							sheet1.write(profile_count1, 8, city)
							sheet1.write(profile_count1, 9, address)
							sheet1.write(profile_count1, 10, pin)
							sheet1.write(profile_count1, 11, description)
							sheet1.write(profile_count1, 12, facebook)
							sheet1.write(profile_count1, 13, twitter)
							sheet1.write(profile_count1, 14, linkedin)
							sheet1.write(profile_count1, 15, googleplus)
							sheet1.write(profile_count1, 16, instagram)
							sheet1.write(profile_count1, 17, skype)
							book.save("Data2.xls")
						except:
							print '***'
						f_name = " "
						#social = True
						l_name = " "
						tagline = " "
						price = " "
						phone_number =" "
						user_link = " "
						description = " "
						profile_count = profile_count +1
						profile_count1 = profile_count1 +1
						facebook = ""
						twitter = ""
						google = ""
						linkedin = ""
						skype =""
						instagram =""
						browser.get(areaurl)
					except Exception as e:
							profile_boolean = False
							profile_count = 1
							area_count = area_count +1	
							print e

			except Exception as e:
				profile_count =1
				area_part_li = False
				area_count1 = area_count1 + 3
				area_count = 1
				browser.close()
				print e

		except Exception as e:
			profile_count =1
			browser.close()
			area_count = 1
			area_count1 = 1
			area_part_ul = False
			category_count = category_count +1
			print e

	except:
		category_boolean = False
		category_count = category_count +1
