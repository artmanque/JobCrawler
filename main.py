import sys,os,time
import optparse, openpyxl, datetime
import Tkinter as tk
from lxml.html import fromstring 
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC

def get_element_from_xpath(xpath):
	global driver
	return fromstring(driver.page_source).xpath(xpath)[0]

def get_elements_from_xpath(xpath):
	global driver
	return fromstring(driver.page_source).xpath(xpath)

def savexl(filename,n=0):
	global wb
	try:
		if(n == 0):
			wb.save(filename+'.xlsx')
		else:
			wb.save(filename+"({})".format(str(n))+'.xlsx')
	except:
		return savexl(filename,n+1)
	else:
		if(n == 0):
			return filename+'.xlsx'
		else:
			return filename+"({})".format(str(n))+'.xlsx'

def get_jobs(driver):
	global delay
	try:
		wait(driver, delay/5).until(EC.presence_of_element_located((By.XPATH, "//div[@data-tn-component='organicJob']")))
	except:
		try:
			wait(driver, delay/5).until(EC.presence_of_element_located((By.XPATH, "//div[@data-tn-component='organicJob']")))
		except:
			return [("N/A","N/A","N/A","N/A")]
	jobs = driver.find_elements_by_xpath("//div[@data-tn-component='organicJob']")
	try:
		titl = [job.find_element_by_class_name("jobtitle").text for job in jobs]
	except:
		title = ["N/A"]
	try:
		stat = [job.find_element_by_class_name("location").text for job in jobs]
	except:
		stat = ["N/A"]*len(jobs)
	try:
		comp = [job.find_element_by_class_name("company").text for job in jobs]
	except:
		comp = ["N/A"]*len(jobs)
	try:
		date = [job.find_element_by_class_name("date").text for job in jobs]
	except:
		date = ["N/A"]*len(jobs)
	try:
		aply = [job.find_element_by_xpath(".//h2/a").get_attribute('href') for job in jobs]
	except:
		aply = ["N/A"]*len(jobs)

	# for i,job in enumerate(jobs):
		# print '\tcrawling job#',i+1
		# driver.find_element_by_xpath("//div[@data-tn-component='organicJob'][{}]".format(i+1)).click()
		# print driver.find_element_by_xpath("//div[@data-tn-component='organicJob'][{}]/h2/a".format(i+1)).text
		# wait(driver, delay).until(EC.presence_of_element_located((By.ID, "vjs-container")))
		# time.sleep(delay)
		# try:
		# 	driver.find_element_by_xpath("//button[@id='popover-x']").click()
		# except:
		# 	pass
		# try:
		# 	titl.append(driver.find_element_by_id("vjs-jobtitle").text)
		# except:
		# 	titl.append("N/A")
		# try:
		# 	stat.append(driver.find_element_by_id("vjs-loc").text.replace('- ',''))
		# except:
		# 	stat.append("N/A")
		# try:
		# 	comp.append(job.find_element_by_id("vjs-cn").text)
		# except:
		# 	comp.append("N/A")
		# try:
		# 	date.append(job.find_element_by_xpath("//div[@id='vjs-footer']/div/div/span[@class='date']").text)
		# except:
		# 	date.append("N/A")
		# aply.append('n/a')
		# print titl[i],stat[i],comp[i],date[i],aply[i]



		# 		# driver.find_elements_by_xpath("//div[@data-tn-component='organicJob']")[i].click()
		# 		# time.sleep(delay/2)
		# try:
		# 	wait(driver, delay/4).until(EC.presence_of_element_located((By.XPATH, "//div[@class='job-footer-button-row']/a")))
		# 	driver.find_element_by_xpath("//div[@class='job-footer-button-row']/a").click()
		# 	time.sleep(delay)
		# 	driver.switch_to_window(driver.window_handles[1])
		# 	aply.append(driver.current_url)
		# 	driver.close()
		# 	driver.switch_to_window(driver.window_handles[0])
		# except Exception as err:
		# 	print err
		# 	aply.append(driver.find_element_by_xpath("//div[@data-tn-component='organicJob'][{}]//h2/a".format(i+1)).get_attribute('href'))
		# try:
		# 	driver.find_element_by_id('vjs-x').click()
		# except:
		# 	pass
		# time.sleep(delay/3)
		# print aply[i]
	return zip(titl,stat,comp,date,aply)

def get_angel_jobs(driver):
	jobs = driver.find_elements_by_xpath("//div[@data-_tn = 'job_listings/browse_startups_table_row']")
	data = []
	for i in range(len(jobs)):
		if i != 0:
			jobs[i].click()
		try:
			comp = driver.find_elements_by_xpath("//a[@class='startup-link']")[i].text
		except:
			comp = "N/A"
		try:
			titl = driver.find_elements_by_xpath("//div[@class='title']/a")[i].text
		except:
			titl = "N/A"
		try:
			stat = driver.find_elements_by_xpath("//div[@class='tag locations tiptip']")[i].text
		except:
			stat = "N/A"
		try:
			link = driver.find_elements_by_xpath("//a[@class='website-link']")[i].text
			if link == "none":
				print "link is none"
				link = "N/A"
		except:
			link = "N/A"
		try:
			size = driver.find_elements_by_xpath("//div[@class='tag employees']")[i].text
		except:
			size = "N/A"
		try:
			fndr = driver.find_elements_by_xpath("//div[@class='name']/a[@class='profile-link']")[i].text
		except:
			fndr = "N/A"
		data.append((titl,stat,comp,link,size,fndr))
	return data


def login_angel(driver,uid,pwd):
	global delay
	url = "https://angel.co/login"
	if driver.current_url != url:
		driver.get(url)
	time.sleep(delay/2)
	driver.find_element_by_xpath("//input[@type='email']").send_keys(uid)
	driver.find_element_by_xpath("//input[@type='password']").send_keys(pwd)
	driver.find_element_by_xpath("//input[@type='password']").send_keys(Keys.RETURN)
	time.sleep(delay/2)

def remove_tags(driver):
	tags = driver.find_elements_by_xpath("//div[@class='search-box']/div[@class='currently-showing']/div")
	for i in range(len(tags)):
		driver.find_element_by_xpath("//img[@title='Remove']").click()

def add_location(driver,location):
	global delay
	menu = driver.find_element_by_xpath("//div[@data-menu = 'locations']")
	elem = driver.find_element_by_xpath("//input[@placeholder='Enter a location']")
	action = Acts(driver)
	action.move_to_element(menu)
	action.send_keys_to_element(elem,location)
	action.wait(delay)
	action.send_keys_to_element(elem,Keys.RETURN)
	action.perform()
	# action = Acts(driver)
	# action.move_to_element(menu)
	# action.send_keys_to_element(elem,Keys.RETURN)
	# action.perform()

delay = 10
start = time.time()
filename = str(datetime.date.today())
fields = {'keyword': None,'location': None,'pages': None,'posted': None,'ver': 1}

#################driver###################

# //Initialize firefox driver
firefox_profile = webdriver.FirefoxProfile()
firefox_profile.set_preference('permissions.default.image', 2)
firefox_profile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', 'false')
driver = webdriver.Firefox(firefox_profile=firefox_profile)

##########################################

# chrome_options = Options()
# # chrome_options.add_extension("clearbit.crx")
# chrome_options.add_argument('--silent')
# # chrome_options.add_argument("user-data-dir=C:\Users\AbdulRehmanAslam\AppData\Local\Google\Chrome\User Data");
# chrome_options.add_argument("--start-maximized");
# chrome_options.add_argument('--disable-logging')
# # chrome_options.add_experimental_option("prefs",{"profile.managed_default_content_settings.images":2})
# driver = webdriver.Chrome(chrome_options=chrome_options,service_log_path='NUL')
# # extra_windows = len(driver.window_handles)
# # time.sleep(delay/2)
# # for i in range(extra_windows):
# driver.close()
# driver.switch_to_window(driver.window_handles[0])
###################################
################GUI################
class Application(tk.Frame):
	def __init__(self):
		self.root = tk.Tk()
		self.root.geometry("400x300")
		self.root.resizable(width=False, height=False)
		self.root.bind('<Return>', self.destruct)

		tk.Frame.__init__(self, self.root)
		global fields
		self.versions = [("GlassDoor",1),("Indeed",2),("AngelList",3)]
		self.create_widgets()
		self.pack()

	def create_widgets(self):
		self.row1 = tk.Frame(self.root)
		self.row2 = tk.Frame(self.root)
		self.row3 = tk.Frame(self.root)
		self.row4 = tk.Frame(self.root)
		self.row5 = tk.Frame(self.root)
		self.row6 = tk.Frame(self.root)
		self.row7 = tk.Frame(self.root)

		self.labl_key = tk.Label(self.row1, width=15, text="Keyword", anchor='w')
		self.labl_loc = tk.Label(self.row2, width=15, text="Location", anchor='w')
		self.labl_pgs = tk.Label(self.row3, width=15, text="No. of Pages", anchor='w')
		self.labl_tim = tk.Label(self.row4, width=15, text="Date Posted", anchor='w')
		self.labl_ver = tk.Label(self.row5, width=15, text="Choose Version:",anchor='w')

		self.sv_key = tk.StringVar()
		self.sv_key.trace("w", lambda name, index, mode, sv=self.sv_key:self.save(self.sv_key,'keyword'))

		self.sv_loc = tk.StringVar()
		self.sv_loc.trace("w", lambda name, index, mode, sv=self.sv_loc:self.save(self.sv_loc,'location'))

		self.sv_pgs = tk.StringVar(value='1')
		self.sv_pgs.trace("w", lambda name, index, mode, sv=self.sv_pgs:self.save(self.sv_pgs,'pages'))

		self.sv_tim = tk.StringVar()
		self.sv_tim.trace("w", lambda name, index, mode, sv=self.sv_tim:self.save(self.sv_tim,'posted'))

		self.ent_key = tk.Entry(self.row1,textvariable=self.sv_key)
		self.ent_loc = tk.Entry(self.row2,textvariable=self.sv_loc)
		self.ent_pgs = tk.Entry(self.row3,textvariable=self.sv_pgs)
		self.ent_tim = tk.Entry(self.row4,textvariable=self.sv_tim)

		self.ver = tk.IntVar(value=fields['ver'])
		self.tgl = {}
		for (version, val) in self.versions:
			self.tgl[version] = tk.Radiobutton(self.row6,text=version,padx = 110,variable=self.ver,command=lambda ver=self.ver:self.save(self.ver,'ver'),value=val,justify=tk.RIGHT)

		self.btn = tk.Button(self.row7, text="Run", command=lambda root=self.root:self.destruct(),width=5)

	def pack(self):
		self.row1.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
		self.row2.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
		self.row3.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
		self.row4.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
		self.row5.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
		self.row6.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
		self.row7.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

		self.labl_key.pack(side=tk.LEFT)
		self.labl_loc.pack(side=tk.LEFT)
		self.labl_tim.pack(side=tk.LEFT)
		self.labl_pgs.pack(side=tk.LEFT)
		self.labl_ver.pack(side=tk.LEFT)

		self.ent_key.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
		self.ent_loc.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
		self.ent_pgs.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
		self.ent_tim.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
		for (version, _) in self.versions:
			self.tgl[version].pack(anchor=tk.W)
		self.btn.pack(pady=10)

	def destruct(self):
		self.root.destroy()

	def save(self,sv,field):
		fields[field] = sv.get()

	def start(self):
		self.root.title("SOME Tech")
		img = tk.PhotoImage(os.path.join(os.getcwd(),'assets','logo.ico'))
		self.root.tk.call('wm', 'iconphoto', self.root._w, img)
		# self.root.iconbitmap('/home/toxic/logo.ico')
		# self.root.iconbitmap(os.path.join(os.path.join(os.getcwd(),'assets'),'logo.ico'))
		self.root.mainloop()

print 'Launching...\n'
Application().start()

###################################
keyword = fields['keyword']
if keyword is None:
	keyword = ''

location = fields['location']
if location is None:
	location = ''

pages = fields['pages']
if pages is not None:
	pages = int(fields['pages'])
else:
	pages = 1

posted = fields['posted']
ver = fields['ver']
###################################

print 'The Server is up and running.\n',str(datetime.date.today()),str(datetime.datetime.now().time())
if __name__ == '__main__':
	#  //Init_XL
	wb = openpyxl.Workbook()
	wb.remove(wb.worksheets[0])
	#####################################################################################
	if ver == 1:
		#########################################################
		driver.get("https://www.glassdoor.com/")
		driver.find_element_by_id("KeywordSearch").send_keys(keyword)
		elem = driver.find_element_by_id("LocationSearch")
		elem.clear()
		elem.send_keys(location)
		driver.find_element_by_id("LocationSearch").send_keys(Keys.RETURN)
		time.sleep(delay/2)
		#########################__init__########################
		try:
			if posted is not None:
				if posted.lower() == 'last day' or posted == '0':  
					driver.find_element_by_xpath("//*[@id='DKFilters']/div/div/div[2]").click()
					driver.find_element_by_xpath("//ul[@class='flyout']/li[@value='1']").click()
				elif posted.lower() == 'last week' or posted == '1':  
					driver.find_element_by_xpath("//*[@id='DKFilters']/div/div/div[2]").click()
					driver.find_element_by_xpath("//ul[@class='flyout']/li[@value='7']").click()
				elif posted.lower() == 'last 2 weeks' or posted == '2':  
					driver.find_element_by_xpath("//*[@id='DKFilters']/div/div/div[2]").click()
					driver.find_element_by_xpath("//ul[@class='flyout']/li[@value='14']").click()
				elif posted.lower() == 'last month' or posted == '4':
					driver.find_elements_by_xpath("//*[@id='DKFilters']/div/div/div[2]").click()
					driver.find_element_by_xpath("//ul[@class='flyout']/li[@value='30']").click()
		except Exception as err:
			print err
		###########################################################
		ws1 = wb.create_sheet()
		ws1.append(('Sr#','Company','Title','Location','Posted','Rating','Company Link','Apply at'))
		time.sleep(delay/5)

		jobs = driver.find_elements_by_xpath("//*[@id='MainCol']/div/ul/li")
		for p in range(pages):
			for n,job in enumerate(jobs):
				try:
					driver.find_elements_by_xpath("//*[@id='MainCol']/div/ul/li")[n].click()	
					time.sleep(delay/5)
				except:
					pass
				try:
					wait(driver, delay).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='plain strong empDetailsLink']")))
				except:
					pass
				try:
					while True:
						close = driver.find_element_by_xpath("//*[@id='JAModal']/div/div[2]/div[@class='xBtn']")
						try:
							close.click()
						except:
							break
				except:
					pass
				try:
					comp = driver.find_element_by_xpath("//a[@class='plain strong empDetailsLink']").text
				except:
					comp = 'N/A'
				try:
					titl = driver.find_element_by_xpath("//h1[@class='jobTitle h2 strong']").text
				except:
					title = 'N/A'
				try:
					loc = driver.find_element_by_xpath("//div[@class='compInfo']/span[2]").text.replace('- ','')
				except:
					loc = 'N/A'
				try:
					rate = driver.find_element_by_xpath("//span[@class='compactRating lg margRtSm']").text
				except:
					rate = 'N/A'
				try:
					dago = driver.find_element_by_xpath("//span[@class='hideHH nowrap']/span[@class='minor']").text
				except:
					dago = 'N/A'
				try:
					driver.find_element_by_xpath("//ul[@class='pageTabs']/li[2]").click()
				except:
					pass
				try:
					link  = driver.find_element_by_xpath("//span[@class='value website']/a").get_attribute("href")
				except:
					link = 'N/A'
				try:
					aply  = driver.find_element_by_xpath("//div[@class='applyCTA']/a").get_attribute("href")
					if aply == '':
						aply  = jobs[n].find_element_by_xpath(".//a[@class='jobLink']").get_attribute("href")
				except:
					aply = 'N/A'
				###########################################################
				ws1.append((n+1,comp,titl,loc,dago,rate,link,aply))
				wb.save(filename+'.xlsx')
				###########################################################
			try:
				driver.find_element_by_xpath("//div[@class='pagingControls cell middle']/ul/li[@class='next']").click()
			except:
				break
		wb.save(filename+'.xlsx')
		print "Output file saved as: ",filename
		###################################
	elif ver == 2:
		url = "https://www.indeed.com/"
		driver.get(url)
		wait(driver, delay/10).until(EC.presence_of_element_located((By.ID, "text-input-what")))
		driver.find_element_by_id("text-input-what").send_keys(keyword)
		elem = driver.find_element_by_id("text-input-where")
		elem.send_keys(Keys.CONTROL + "a");
		elem.send_keys(Keys.DELETE)
		elem.send_keys(location)
		elem.send_keys(Keys.RETURN)

		ws1 = wb.create_sheet()
		ws1.append(('Sr#','Company','Title','Location','Posted','Apply at'))
		j_ind = 0
		for p in range(pages):
			try:
				if j_ind == 0:
					try:
						wait(driver, 0.1).until(EC.presence_of_element_located((By.XPATH, "//div[@class='invalid_location']/p[@class='oocs']/a")))
						driver.find_element_by_xpath("//div[@class='invalid_location']/p[@class='oocs']/a").click()
						time.sleep(delay/5)
					except:
						pass
				wait(driver, 0.1).until(EC.presence_of_element_located((By.XPATH, "//*[@id='prime-popover-close-button']")))
				driver.find_element_by_xpath("//*[@id='prime-popover-close-button']").click()
			except:
				try:
					driver.find_element_by_xpath("//button[@id='popover-x-button']").click()
				except:
					pass
			else:
				try:
					driver.find_element_by_xpath("//button[@id='popover-x-button']").click()
				except:
					wait(driver, 0.1).until(EC.presence_of_element_located((By.XPATH, "//*[@id='prime-popover-close-button']")))
					driver.find_element_by_xpath("//*[@id='prime-popover-close-button']").click()
			if j_ind == 0:
				pgs = [x.get_attribute("href") for x in driver.find_elements_by_xpath("//div[@class='pagination']/a")[:-1] ]
			jobs = get_jobs(driver)
			for n,(titl,stat,comp,dago,aply) in enumerate(jobs):
				j_ind = j_ind + 1
				ws1.append((j_ind,comp,titl,stat,dago,aply))
				filename = str(datetime.date.today())
				wb.save(filename+'.xlsx')
			try:
				driver.get(pgs[p])
			except Exception as e:
				print "\tNo more pages to crawl"
				break
		print "Output file saved as: ",filename
	elif ver == 3:
		url = "https://www.angel.co/login"
		driver.execute_script("window.open('{}')".format(url))
		driver.switch_to_window(driver.window_handles[1])
		login_angel(driver,uid,pwd)
		url = "https://www.angel.co/jobs"
		driver.get(url)
		wait(driver, delay*5).until(EC.element_to_be_clickable((By.XPATH, "//div[@data-_tn = 'job_listings/browse_startups_table_row']")))
		time.sleep(delay)
		remove_tags(driver)

		elem = driver.find_element_by_xpath("//div[@class='search-box']")
		elmn = driver.find_element_by_xpath("//input[@class='input keyword-input']")
		elem.click()
		elmn.send_keys(keyword)
		elmn.send_keys(Keys.RETURN)

		add_location(driver,location)
		# raw_input("Select Location then press Enter to continue")

		ws1 = wb.create_sheet()
		ws1.append(('Sr#','Company','Title','Location','Size','Link','Founder','Employee Name','Designation','Email','Lives in','Social Links'))
		filename = str(datetime.date.today())
		jobs = get_angel_jobs(driver)
		try:
			for n,(titl,stat,comp,link,size,fndr) in enumerate(jobs):
				driver.switch_to_window(driver.window_handles[0])
				emps = get_clearbit_data(driver,comp)
				if emps == []:
					ws1.append((n+1,comp,titl,stat,size,link,fndr,"N/A","N/A","N/A","N/A","N/A"))
				else:
					for (name,post,email,loct,socials) in emps:
						ws1.append((n+1,comp,titl,stat,size,link,fndr,name,post,email,loct,socials))
				driver.switch_to_window(driver.window_handles[1])
				wb.save(filename+'.xlsx')
		except:
			wb.save(filename+'.xlsx')
		finally:
			wb.save(filename+'.xlsx')
			print "Output file saved as: ",filename

	else:
		print "Version Under Construction!"
	#####################################################################################
# driver.quit()