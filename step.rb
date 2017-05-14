#########################################################################
### browser profile initialization
### Reference: https://github.com/SeleniumHQ/selenium/wiki/Ruby-Bindings
### Require: geckodriver-v0.16.1
### Require: Selenium webdriver 3.4.0
##########################################################################

def open_firefox_using_profile

  # setup_watir_webdriver
  require 'watir-webdriver'
  download_directory = "C:\\Users\\Yatin\\Downloads"
  profile = Selenium::WebDriver::Firefox::Profile.new

  profile['credentials_enable_service'] = false
  ### file will be downloaded to custom directory
  profile['browser.download.folderList'] = 2
  profile['browser.download.dir'] = download_directory
  ### By Default file will be downloaded to Download directory
  # profile['browser.download.folderList'] = 1 # 0 = save to user's desktop, 1 = save to Downloads, 2 = save to custom location
  # profile["browser.download.useDownloadDir"] = true
  profile['browser.download.manager.showWhenStarting'] = false
  profile['browser.helperApps.neverAsk.saveToDisk'] = "text/csv, text/plain, application/vnd.ms-excel, text/comma-separated-values, application/pdf, application/zip"
  # driver = Selenium::WebDriver.for :firefox, :profile => profile
  # @browser = Watir::Browser.new(driver)
  @browser = Watir::Browser.new :firefox,  :profile => profile

end

#browser
#:firefox (default)
#:chrome
#agent
#:iphone (default)
#:ipad
#:android_phone
#:android_tablet
#:random
#orientation
#:portrait (default)
#:landscape

def browser_user_agent(browser, agent, orientation)
setup_watir_webdriver
browser ="firefox"
agent = "iphone"
orientation="portrait"
require 'watir-webdriver'
require 'webdriver-user-agent'
driver = Webdriver::UserAgent.driver(:browser => :"#{browser}", :agent => :"#{agent}", :orientation => :"#{orientation}")
browser = Watir::Browser.new(driver)
end

#iphone
#width=667, height=375 - landscape
#width=375, height=559 - portrait

def window_size()
	@browser.window.resize_to(x, y)
end

def open_chrome_using_profile

  # setup_watir_webdriver
  require 'watir-webdriver'
  download_directory = "C:\\Users\\Yatin\\Downloads"
  profile = Selenium::WebDriver::Chrome::Profile.new
  profile['credentials_enable_service'] = false
  profile['download.prompt_for_download'] = false
  profile['download.default_directory'] = download_directory
  profile['browser.helperApps.neverAsk.saveToDisk'] = "text/csv, text/plain, application/vnd.ms-excel, text/comma-separated-values, application/pdf, application/zip"

  @browser = Watir::Browser.new :chrome,  :profile => profile , :switches => %w[--ignore-certificate-errors --disable-popup-blocking --disable-translate]
end


def open_internet_explorer_using_profile
  require 'watir-webdriver'

  setup_watir_webdriver

  client         = Selenium::WebDriver::Remote::Http::Default.new
  client.timeout = 300 # seconds – default is 60
  caps           = Selenium::WebDriver::Remote::Capabilities.internet_explorer(
      #:nativeEvents => false,
      #:requireWindowFocus => true,
      :initialBrowserUrl              => 'about:blank',
      :enablePersistentHover          => false,
      :ignoreProtectedModeSettings    => true,
      'ie.ensureCleanSession'         => true,
      :unexpectedAlertBehaviour       => 'ignore'
  )
  @browser       = Watir::Browser.new :ie, :http_client => client, :desired_capabilities => caps




end



def open_edge_using_profile

  require 'watir-webdriver'
  download_directory = "C:\\Users\\Yatin\\Downloads"
  profile = Selenium::WebDriver::Chrome::Profile.new
  profile['credentials_enable_service'] = false
  profile['download.prompt_for_download'] = false
  profile['download.default_directory'] = download_directory
  profile['browser.helperApps.neverAsk.saveToDisk'] = "text/csv, text/plain, application/vnd.ms-excel, text/comma-separated-values, application/pdf, application/zip"

  @browser = Watir::Browser.new :edge,  :profile => profile , :switches => %w[--ignore-certificate-errors --disable-popup-blocking --disable-translate]

end


###################
#### Load Data ####
###################

def read_excel(name)
  require 'roo'
  require 'spreadsheet'
  @myRoot = File.join(File.dirname(__FILE__), '/')
  book = Roo::Spreadsheet.open("#{@myRoot}/#{name}.xlsx", extension: :xlsx)

  obj_repo = book.sheet("object")
  # Loading the data from "object" spreedsheet
  @obj_repo_row = {}
  obj_repo.each do |row|
    row.each do |x|
      @obj_repo_row[row[0]] = row[1..11]
    end
  end

  user_data = book.sheet("data")
  # Loading the data from "data" spreedsheet
  @user_data_row = {}
  user_data.each do |row|
    row.each do |x|
      @user_data_row[row[0]] = row[1]
    end
  end

  # Removing the first row data from the excel(column name).
  @obj_repo_row.delete('Key')
  @user_data_row.delete('Key')

end

######################################################
### Below method Used for create the unique Name
### Using the timestamp
### Read the value during run time for the validation
### For Example generate_uniq_timestamp ("MyName")
### Return "MyANme201704251554"
######################################################

def generate_uniq_timestamp (text)
  cur_time_stmp = Time.now.strftime("%Y%m%d%H%M").to_s
  case text.downcase + "_stamp"
    when "firstname_stamp"
      $title = "#{arg1.downcase}#{cur_time_stmp}"
      uniq_text = $title
    when "lastname_stamp"
      $description = "#{arg1.downcase}#{cur_time_stmp}"
      uniq_text = $description
    else
      uniq_text = text
  end
  return uniq_text
end

def read_temp_val(text)
  case text.downcase + "_stamp"
    when "firstname_stamp"
      data = $title
    when "lastname_stamp"
      data = $description
    else
      data = text
  end
end

#####################################################
### generating the time Stamp, with 15 min interval
### For Example: time_interval(20)
### current Time: 2017-04-26 10:45:05 -0700
### Return MM:SS => "05:00"
#####################################################

def time_interval(inter_val_num=15)

  require 'rubygems'
  require 'active_support'
  require 'active_support/time'

  cur_time = Time.now
  comp_time = cur_time - cur_time.sec - cur_time.min%15*60
  base_time = comp_time + 1.hour

  if comp_time < base_time
    comp_time = comp_time + inter_val_num.minutes
  end

  return comp_time.strftime("%M:%S")

end

#############################################################################################
### generating Date based on the argument passed
### By default offset=0, gives current Date with all argument nil
###offset = -1, gives past date with atleast one argument need to be passed along offset
###offset = 1, gives future date with atleast one argument need to be passed along offset
#############################################################################################
### example: get_date() --2017/04/28 --- 04/28/2017
#### Future Date below
### example: get_date(2,2,2,1) ---  06/30/2019
### example: get_date(2,2,nil,1) --- 06/30/2017
### example: get_date(2,nil,2,1) --- 04/30/2019
### example: get_date(2,nil,nil,1) --- 04/30/2017
### example: get_date(nil,2,2,1) --- 06/28/2019
### example: get_date(nil,2,nil,1) --- 06/28/2017
### example: get_date(nil,nil,2,1) --- 04/28/2019
#### Past Date below
### example: get_date(2,2,2,-1) --- 02/26/2015
### example: get_date(2,2,nil,-1)--- 02/26/2017
### example: get_date(2,nil,2,-1)--- 04/26/2015
### example: get_date(2,nil,nil,-1)--- 04/26/2017
### example: get_date(nil,2,2,-1)--- 02/28/2015
### example: get_date(nil,2,nil,-1)--- 02/28/2017
### example: get_date(nil,nil,2,-1)--- 04/28/2015
################################################

def get_date(day=nil, month=nil, year=nil, offset=0)

  require 'rubygems'
  require 'active_support'
  require 'active_support/time'

  cur_date = Date.today

  if offset < 0
    if day != nil &&	month != nil &&	year != nil
      @com_date = cur_date - day
      @com_date = @com_date << month
      @com_date = @com_date << 12 * year
    end
    if day != nil && month != nil && year == nil
      @com_date = cur_date - day
      @com_date = @com_date << month
    end

    if day != nil && month == nil && year != nil
      @com_date = cur_date - day
      @com_date = @com_date << 12 * year
    end
    if day != nil && month == nil && year == nil
      @com_date = cur_date - day
    end
    if day == nil && month != nil && year != nil
      @com_date = cur_date << month
      @com_date = @com_date << 12 * year
    end
    if day == nil && month != nil && year == nil
      @com_date = cur_date << month
    end
    if day == nil && month == nil && year != nil
      @com_date = cur_date << 12 * year
    end

  elsif offset > 0
    if day != nil &&	month != nil &&	year != nil
      @com_date = cur_date + day
      @com_date = @com_date >> month
      @com_date = @com_date >> 12 * year

    end
    if day != nil && month != nil && year == nil
      @com_date = cur_date + day
      @com_date = @com_date >> month
    end

    if day != nil && month == nil && year != nil
      @com_date = cur_date + day
      @com_date = @com_date >> 12 * year
    end
    if day != nil && month == nil && year == nil
      @com_date = cur_date + day
    end
    if day == nil && month != nil && year != nil
      @com_date = cur_date >> month
      @com_date = @com_date >> 12 * year
    end
    if day == nil && month != nil && year == nil
      @com_date = cur_date >> month
    end
    if day == nil && month == nil && year != nil
      @com_date = cur_date >> 12 * year
    end
  elsif day == nil &&	month == nil &&	year == nil && offset == 0
    @com_date = cur_date
  else
    fail("Invalid Argument Passed to get date method. Please enter correct arguments")
  end
return @com_date.strftime("%m/%d/%Y").to_s
end

#######################################################
### generating Date Range based on the argument passed with current date

### example:date_range(nil,nil,nil,nil,nil,nil) --- ["04/28/2017", "04/28/2017"]
### example:date_range(2,2,2,2,2,2) --- ["02/26/2015", "06/30/2019"]
### example:date_range(2,2,2,nil,nil,nil) --- ["02/26/2015", "04/28/2017"]
### example:date_range(nil,nil,nil,2,2,2) --- ["04/28/2017", "06/30/2019"]
### example:date_range(nil,nil,nil,nil,2,2) --- ["04/28/2017", "06/28/2019"]
### example:date_range(nil,2,2,nil,nil,nil) --- ["02/28/2015", "04/28/2017"]
### example:date_range(2,nil,nil,2,nil,nil) --- ["04/26/2017", "04/30/2017"]

#######################################################

def date_range(past_day=nil, past_month=nil, past_year=nil, future_day=nil, future_month=nil,future_year=nil)
  if (past_day != nil || past_month != nil || past_year != nil) && (future_day != nil || future_month != nil || future_year != nil)
  start_date = get_date(past_day,past_month,past_year,-1)
  end_date = get_date(future_day,future_month,future_year, 1)
  elsif past_day == nil && past_month == nil && past_year == nil && future_day == nil && future_month == nil && future_year == nil
    start_date = get_date(nil,nil,nil,0)
    end_date = get_date(nil,nil,nil,0)
  elsif (past_day == nil && past_month == nil && past_year == nil) && (future_day != nil || future_month != nil || future_year != nil)
    start_date = get_date(nil,nil,nil,0)
    end_date = get_date(future_day,future_month,future_year,1)
  elsif (past_day != nil || past_month != nil || past_year == nil) && (future_day == nil && future_month == nil && future_year == nil)
    start_date = get_date(past_day,past_month,past_year, -1)
    end_date = get_date(nil,nil,nil,0)
  else
    fail("Invalid Argument Passed to date range method. Please enter correct arguments")
  end
  return start_date,end_date
end


############################################################
### validate the length of text field or textarea
############################################################
def val_element_length(element, exp_size)

  o = [('a'..'z'), ('A'..'Z')].map(&:to_a).flatten
  exp_len_str = (0...exp_size).map { o[rand(o.length)] }.join
  puts "generate random string => #{exp_len_str}"

  act_len_str = "#{exp_len_str}extra"

  if element.attribute_value('maxlength').to_i != nil
    if element.attribute_value('maxlength').to_i == exp_size.to_i
      element.send_keys act_len_str
    end
  end

  if element.value == ""
    element.send_keys act_len_str
  end

  #### comparing the Value with maxlength
  if element.attribute_value('maxlength').to_i != nil
    ele_max_len = element.attribute_value('maxlength').to_i
    if ele_max_len != exp_size
      puts ("Actual Element Maxlength: #{ele_max_len} != Expected Element Maxlength: #{exp_size}")

    end
  end

  run_time_str =element.value
  if run_time_str.size == exp_size
    if run_time_str != exp_len_str
      fail("Actual Element Value #{run_time_str} != Expected Element Value #{exp_len_str}")
    end
  else
    fail("Actual Element Value Count: #{run_time_str.size} != Expected Element Value Count: #{exp_size}")
  end
end

#######################################################
### Generating Random String
#######################################################

def generate_random_alphabetical_string(len)
  o = [('a'..'z'), ('A'..'Z')].map(&:to_a).flatten
  str = (0...len).map { o[rand(o.length)] }.join
  puts "generate random string => #{str}"
  return str
end

def generate_random_num_string(len)
  o = [('0'..'9'), ('0'..'9')].map(&:to_a).flatten
  str = (0...len).map { o[rand(o.length)] }.join
  puts "generate random string => #{str}"
  return str
end

def generate_random_alphanumeric_string(len)
  o = [('0'..'9'),('a'..'z'),('0'..'9')].map(&:to_a).flatten
  str = (0...len).map { o[rand(o.length)] }.join
  puts "generate random string => #{str}"
  return str
end


###################################################
### Encoding and Decoding string
###################################################

def encode_str(str)
  require "base64"
  encoded_str=Base64.encode64(str)
  return encoded_str
end

def decode_str(str)
  require "base64"
  decoded_str = Base64.decode64(str)
  return decoded_str
end

#####################################################
### File Upload Using Rautomation
### Refernce: http://itreallymatters.net/post/2352350743/automating-windows-and-their-controls-with-ruby#.WQjSTFHyvIU
### filepath = C:\\Users\\Yatin\\Downloads\\images.jpg
### specify title (other then "Open" "File Upload" or "Choose File to Uplaod")
#####################################################

def file_upload_using_r_automation (filepath,window_title=nil)

  if window_title.nil?
    if @params and @params['browser'] != nil
      browser_type = @params['browser']
    else
      fail("Failed unrecognized browser type, Please check the Params under the manifest.json")
    end

    case browser_type.downcase
      when "gc","chrome","edge"
        title = "Open"
      when "ff","firefox"
        title = "File Upload"
      when "ie"
        title = "Choose File to Uplaod"
      else
        fail("Failed unrecognized browser type, Please check the Params under the manifest.json")
    end
  else
    title = window_title
  end
  require "rautomation"
  window = RAutomation::Window.new :title => title

  if window.exists?
    ### validate the File Name Text_field exists,
    ### if exists then enter the Path
    if window.text_field(:class => "Edit", :index => 0).exists?
    textField=window.text_field(:class => "Edit", :index => 0)
    textField.set filepath
    else
      fail("Failed to Find the FileName Text Field")
    end

    ### Validate the Text_field is not empty after filepath added
    if textField.value == ""
      fail("Edit Text Field is Empty")
    end

    ### Validate the Open button exists
    ### if exists then perform click on Open button
    if window.button(:value => "&Open").exists?
      window.button(:value => "&Open").click
    else
      fail "Failed to Find Open button on window dialog box"
    end

    ### Validate After Clicking Open button the window dialog box should not exists
    if window.exists?
      fail "Window dialog box still exists even after clicking on Open button"
    end
  else
    fail "Failed to Initialize the Window"
  end
end


### Reference https://help4qa.wordpress.com/2012/03/17/selenium-easy-way-to-handeling-file-download-dialog-box/
### https://www.autoitscript.com/site/autoit/downloads/

def downlaod_file_and_save
  win = RAutomation::Window.new :title => /Opening rautomation/
  p win.present?
  p win.controls.length
  p win.text
  win.button(:value => "&Save").click
  win.close
end



def download_file_using_chrome (filepath)

  title = "Save As"
  require "rautomation"
  window = RAutomation::Window.new :title => title

  if window.exists?
    ### validate the File Name Text_field exists,
    ### if exists then enter the Path
    if window.text_field(:class => "Edit", :index => 0).exists?
      textField=window.text_field(:class => "Edit", :index => 0)
      textField.set filepath
    else
      fail("Failed to Find the FileName Text Field")
    end

    ### Validate the Text_field is not empty after filepath added
    if textField.value == ""
      fail("Edit Text Field is Empty")
    end

    ### Validate the Open button exists
    ### if exists then perform click on Open button
    if window.button(:value => "&Save").exists?
      window.button(:value => "&Save").click
    else
      fail "Failed to Find Open button on window dialog box"
    end

    ### Validate After Clicking Open button the window dialog box should not exists
    if window.exists?
      fail "Window dialog box still exists even after clicking on Open button"
    end
  else
    fail "Failed to Initialize the Window"
  end
end



client = Selenium::WebDriver::Remote::Http::Default.new
client.timeout = 2000 # seconds – default is 60
caps = Selenium::WebDriver::Remote::Capabilities.internet_explorer(
    #:nativeEvents => false,
    #:requireWindowFocus => true,
    :initialBrowserUrl              => 'about:blank',
    :enablePersistentHover          => false,
    :ignoreProtectedModeSettings    => true,
    'ie.ensureCleanSession'         => true,
    :unexpectedAlertBehaviour       => 'ignore'
)
@browser = Watir::Browser.new :ie, :http_client => client, :desired_capabilities => caps











$hResult = WinActivate("File Download - Security Warning")
If($hResult == 0) Then
_Log("Unable to find Download Window from IE")
Else
$IETitle=WinGetTitle("File Download - Security Warning")
_Log("Download Window activated"&$IETitle)
WinActivate($IETitle)
ControlClick($IETitle, "","[CLASS:Button; INSTANCE:2]")
_Log("FileChooser Window opend")
_Log("CommandLine Parameter Found and Value is:"&$CmdLine[2])
WinActivate("Save As")
_Log("FileChooser Window opend"&WinGetTitle("Save As"))
ControlSetText(WinGetTitle("Save As"),"","Edit1",$CmdLine[2])
Send("!s")
EndIf
EndFunc













require 'win32ole'
$ai = ::WIN32OLE.new('AutoItX3.Control')
$ai.WinActivate("File Download")
$ai.WinWait("[DirectUIHWND]", "",1)

