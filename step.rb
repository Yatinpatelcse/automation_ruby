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

################################################
### generating Date based on the argument passed
### example: get_date() --2017/04/28 --current date
#### Future Date below
### example: get_date(2,2,2,1) --- 2019/06/30
### example: get_date(2,2,nil,1) --- 2017/06/30
### example: get_date(2,nil,2,1) --- 2019/04/30
### example: get_date(2,nil,nil,1) --- 2017/04/30
### example: get_date(nil,2,2,1) --- 2019/06/28
### example: get_date(nil,2,nil,1) --- 2017/06/28
### example: get_date(nil,nil,2,1) --- 2019/04/28
#### Past Date below
### example: get_date(2,2,2,-1) --- 2015/02/26
### example: get_date(2,2,nil,-1)--- 2017/02/26
### example: get_date(2,nil,2,-1)--- 2015/04/26
### example: get_date(2,nil,nil,-1)--- 2017/04/26
### example: get_date(nil,2,2,-1)--- 2015/02/28
### example: get_date(nil,2,nil,-1)--- 2017/02/28
### example: get_date(nil,nil,2,-1)--- 2015/04/28

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
  end
return @com_date.strftime("%Y/%m/%d").to_s
end

################################################
### generating Date Range based on the argument passed



def date_range()

end







