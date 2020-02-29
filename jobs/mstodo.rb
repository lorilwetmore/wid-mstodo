require 'rest-client'
require 'rufus-scheduler'

MS_GRANTTYPE = 'urn:ietf:params:oauth:grant-type:device_code'

$msAuth_TokenRequest = ""
$msAuth_RefreshToken = ""
$msAuth_CurrentToken = ""
$msAuth_UnsuccessfulRefreshCounter = 0
$devauthjob = ""
$msAuth_Authenticated=0
$taskloader = ""
#############################################################################################################################
## This job/class uses the refresh-token in order to receive a new token without repeated authentication. It will be scheduled once a valid refresh-token has been received
#############################################################################################################################
class DO_msAuth_RefreshToken
  def initialize(relevant_info)
    @ri = relevant_info
    $msAuth_UnsuccessfulRefreshCounter
  end
  def call(job)
    #here comes the actual job content
    msAuth_TokenRequestResponse = ""
    begin
      msAuth_TokenRequestResponse = \
      RestClient.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',
         {
           grant_type: 'refresh_token',
           refresh_token: $msAuth_RefreshToken,
           client_id: MS_CLIENTID
         }
       )
      puts "successfully refreshed token"
      $msAuth_Authenticated = 1
      $msAuth_TokenRequest = JSON.parse(msAuth_TokenRequestResponse)
      msAuth_RefreshTime=$msAuth_TokenRequest["expires_in"].to_i
      $msAuth_RefreshToken=$msAuth_TokenRequest["refresh_token"]
      $msAuth_CurrentToken=$msAuth_TokenRequest["access_token"]
      puts $msAuth_CurrentToken
      tokenFile = File.open('refresh_token','w')
      tokenFile.puts $msAuth_RefreshToken
      tokenFile.close
      SCHEDULER.in 1800.to_s+'s', self
    rescue RestClient::ExceptionWithResponse => e
      # exception handling if refresh has not been successful
      puts "token not refreshed"
      if $msAuth_UnsuccessfulRefreshCounter < 3
        $msAuth_UnsuccessfulRefreshCounter+=1
        SCHEDULER.in '5s', self
      else
        $msAuth_UnsuccessfulRefreshCounter = 0
        $devauthjob.resume
      end      
    end
  end
end


#############################################################################################################################
## This job triggers a device authorization request
#############################################################################################################################
class DO_msAuth_DevAuthJob
  attr_accessor :userCode, :deviceAuth
  def initialize(relevant_info)
    @ri = relevant_info
    @userCode = ""
    @deviceAuth = ""
  end
  def call(job)
    puts "devauthjob"
    msAuth_DeviceAuthResponse = \
        RestClient.post('https://login.microsoftonline.com/common/oauth2/v2.0/devicecode',
          {
            client_id: MS_CLIENTID,
            scope: 'https://outlook.office.com/user.read https://outlook.office.com/tasks.read https://outlook.office.com/tasks.read.shared Tasks.Read.Shared Tasks.Read offline_access'
          }
        )
    @deviceAuth = JSON.parse(msAuth_DeviceAuthResponse)
    @userCode = @deviceAuth["user_code"]
    $devtokenjob.resume if $devtokenjob.paused?
    
  end
end

###################################################################################
## This job checks if the user has already authenticated the request and if yes,  #
## retrieves the token. If the token has been received, the job will pause itself #
###################################################################################
class DO_msAuth_DevTokenJob
  def initialize(relevant_info, refAuthJob, refRefreshJob)
    @ri = relevant_info
    @refAuthJob = refAuthJob
    @refRefreshJob = refRefreshJob
  end
  def call(job)    
    puts "devtokenjob"
    begin
      msAuth_TokenRequestResponse = \
            RestClient.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',
              {
                grant_type: MS_GRANTTYPE,
                code: @refAuthJob.deviceAuth["device_code"],
                client_id: MS_CLIENTID
              }
            )
    rescue RestClient::ExceptionWithResponse => e
      puts "exception in token retrieval"
      if JSON.parse(e.response)["error"]="expired_token"
        if $devauthjob.paused?
          $devauthjob.resume
          puts "devauthjob resumed because of expired token"
        end
      end
    end
    if msAuth_TokenRequestResponse
      puts "successfully received token"
      $msAuth_Authenticated = 1
      $devauthjob.pause
      puts "devauthjob paused"
      $msAuth_TokenRequest = JSON.parse(msAuth_TokenRequestResponse)
      msAuth_RefreshTime=$msAuth_TokenRequest["expires_in"]
      $msAuth_RefreshToken=$msAuth_TokenRequest["refresh_token"]
      msAuth_CurrentToken=$msAuth_TokenRequest["access_token"]
      puts msAuth_CurrentToken
      SCHEDULER.in 1800.to_s+'s', @refRefreshJob
      tokenFile = File.open('refresh_token','w')
      tokenFile.puts $msAuth_RefreshToken
      tokenFile.close
      $devtokenjob.pause
    else
#      @refAuthJob.authMessage="Authentication pending"
    end
  end
end

#############################################################################################################################
## This job loads the tasks from MS To-Do if authentication is available
#############################################################################################################################
class DO_msTasks_LoadTasks
  def initialize(relevant_info)
    @ri = relevant_info
    @tasklist = []
  end
  def call(job)
    @tasklist = []
    if $msAuth_Authenticated
      puts "Loading Tasks"
      begin
        msTasks_Response = \
            RestClient.get("https://outlook.office.com/api/v2.0/me/taskfolders('AQMkADAwATM0MDAAMS1iNmI0LTZkYTgtMDACLTAwCgAuAAADn7XOtKava0a6j0LGSnyvAwEAQOKQCG_6GEmlj14G760_1QADL07kYgAAAA==')/tasks?%24top=500", {:'Authorization' => "Bearer " + $msAuth_CurrentToken })
        msTasks_ResponseParsed = JSON.parse(msTasks_Response)
        msTasks_ResponseParsed['value'].map do |task|
          @tasklist.push(subject: task['Subject'].to_str) unless task['Status'].to_str == "Completed"
        end
      rescue

      end
      if tasklist.size == 0
        @tasklist = {}
        @tasklist['notification'] = 'No open tasks'
      end
    end
  end
  def tasklist
    @tasklist
  end
end

handler_msAuth_RefreshToken = DO_msAuth_RefreshToken.new('myRefreshJob')
handler_msAuth_DevAuthJob = DO_msAuth_DevAuthJob.new('myDevAuthJob')
handler_msAuth_DevTokenJob = DO_msAuth_DevTokenJob.new('myDevTokenJob', handler_msAuth_DevAuthJob, handler_msAuth_RefreshToken)
handler_msTasks_LoadTasks = DO_msTasks_LoadTasks.new('myLoadTasksJob')

########################
##Start of program flow
########################

SETTINGS_FILE = "assets/config/msauth_settings.json"
str = IO.read(SETTINGS_FILE)
if not str or str.empty?
  puts "Problem reading clientid and clientsecret"
end
settings = JSON.parse(str)
MS_CLIENTID = settings['clientid']  
MS_CLIENTSECRET = settings ['clientsecret']


## Read token from file, if possible
begin
  tokenFile = File.open('refresh_token','r')
  $msAuth_RefreshToken = tokenFile.read
  tokenFile.close
rescue
  puts "no tokenFile found"
end




## IF token read from file THEN refresh token, schedule and pause the other jobs
## ELSE start authentication flow
if $msAuth_RefreshToken
  SCHEDULER.in '0s', handler_msAuth_RefreshToken
  $devauthjob = SCHEDULER.schedule_every '15m', handler_msAuth_DevAuthJob, :mutex => 'msauthmutex', :tags => 'devauth', :first_in => 5, allow_overlapping:false
  $devauthjob.pause
  $devtokenjob = SCHEDULER.schedule_every '60s', handler_msAuth_DevTokenJob, :mutex => 'msauthmutex', :tags => 'devtoken', :first_in =>10,  allow_overlapping:false
  $devtokenjob.pause
else
  $devauthjob = SCHEDULER.schedule_every '15m', handler_msAuth_DevAuthJob, :mutex => 'msauthmutex', :tags => 'devauth', :first_in => 5, allow_overlapping:false
  $devtokenjob = SCHEDULER.schedule_every '60s', handler_msAuth_DevTokenJob, :mutex => 'msauthmutex', :tags => 'devtoken', :first_in =>10,  allow_overlapping:false
end

$taskloader = SCHEDULER.schedule_every '1m', handler_msTasks_LoadTasks, :mutex => 'msauthmutex', :tags => 'loadtasks', :first_in => 5, allow_overlapping:false

#############################################################################################################################
## This job collects infos via the API and posts them to the widget
#############################################################################################################################
SCHEDULER.every '60s', :first_in => 20, allow_overlapping:false do |displayjob|
puts "displayjob"
  if $msAuth_Authenticated == 0
    send_event('mstodo',  tasks: [], notification: handler_msAuth_DevAuthJob.userCode.to_s , cornersymbol: "fa fa-clipboard fw"  )
    puts "data sent with error"
  else
    send_event('mstodo', { tasks: handler_msTasks_LoadTasks.tasklist, cornersymbol: "fa fa-clipboard fw" } )
    puts "data sent without error"
  end
end
