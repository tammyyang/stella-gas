var __DEBUG__ = false
var req_prefix = "https://api.launchpad.net/devel/"

function msg(msg) {  
  if( __DEBUG__ ) {
    Logger.log(msg)
  }
}
/**
* Calculate week number
*
* @param {Date} given date object, will use the date time now if this is not given.
* @return {number} the week number
*/
function calWeekNumber(d){ 
  if (d == null) {d = new Date()}
  var target  = new Date(d.valueOf());  
  var dayNr   = (d.getDay() + 6) % 7;   
  target.setDate(target.getDate() - dayNr + 3);  
  // ISO 8601 states that week 1 is the week  
  // with january 4th in it  
  var jan4 = new Date(target.getFullYear(), 0, 4);  
  var dayDiff = (target - jan4) / 86400000;     
  var weekNr = 1 + Math.ceil(dayDiff / 7);    
  return weekNr;
}
/**
* Find the first day of the week
*
* @return {string} YYYY-M-D
*/
function getFirstDayString(){
  var today = new Date()
  today.setDate(today.getDate()-(today.getDay()-1))
  var date_array = new Array()
  date_array.push(today.getYear())
  date_array.push(today.getMonth()+1)
  date_array.push(today.getDate())
  return date_array.join('-')
}

/**
* Convert number to string
*
* @param {number} input number
* @return {string} convert to string without decimal
*/
function num2string(value) {
    var new_value = value | 0
    return new_value.toString();
}

/**
* Get today's date as string
*
* @return {string} today's date
*/
function getToday() {
    var d = new Date()
    var dateArray = new Array();
    dateArray.push(this.num2string(d.getFullYear()))
    dateArray.push(this.num2string(d.getMonth()+1))
    dateArray.push(this.num2string(d.getDate()))
    return dateArray.join("-")
}

/**
* Create a hyperlink content for a cell
*
* @param {string} input string
* @return {string} hyperlink
*/
function getHyperlink(url){
  var cell_content = "=HYPERLINK(\"" + url + "\",\"" + 
                     url.split('/').pop() + "\")";
  return cell_content
}

/**
* Clean rows after a give index
*
* @param {sheet} sheet to clean
* @param {number} starting row index
*/
function clean_sheet_rows(sheet, starting_row){
  data_range = this.getDataRange(sheet)
  msg(data_range[0]-starting_row+1)
  var rows_to_clean = data_range[0] - starting_row + 1
  if(rows_to_clean > 0){
    sheet.deleteRows(starting_row, rows_to_clean)
  }
  return
}

/**
* Fetch content from the given url
*
* @param {string} input string
* @return {Array} json data 
*/
function fetchURL(url){
  this.msg("[Stella Library] Enter fetchURL")
  var options ={headers : this.create_dynamic_customHeaders()};    
  var res = UrlFetchApp.fetch(url,options)    
  var res_con = res.getContentText()
  var res_json = JSON.parse(res_con);
  return res_json
}

/**
* Return the max range of a sheet
*
* @param {Sheet} input sheet
* @return {Array} [Nrow, Ncol]
*/
function getDataRange(sheet){
   this.msg("[Stella Library] Enter getDataRange")
   var range = sheet.getDataRange();
   var values = range.getValues();
   return [values.length, values[0].length] //NROW, NCOL
}

/**
* Return all values of a sheet
*
* @param {Sheet} input sheet
* @return {Array} all data
*/
function getSheetData(sheet){
   var range = sheet.getDataRange();
   return range.getValues();
}


/**
* Write content to the sheet
*
* @param {sheet} Sheet to write
* @param {Array} [starting_row, starting_column]
* @param {Array} Data array
* @param {number} optional: clear all row content after the starting point
* @param {number} optional: clear all column content after the starting point
*/
function writeContent(sheet, starting_pt, data_array, clear_all_rows, clear_all_cols) {
    var lock = LockService.getPublicLock()
    var range = sheet.getRange(starting_pt[0], starting_pt[1], data_array.length, data_array[0].length)

    if( lock.tryLock(60000) ) {  
      if (clear_all_rows){
        var all_range = getDataRange(sheet);
        var range_to_clear =  sheet.getRange(starting_pt[0], starting_pt[1], all_range[0]-starting_pt[0]+1, data_array[0].length)
      }
      else if (clear_all_cols){
        var all_range = getDataRange(sheet);
        var range_to_clear =  sheet.getRange(starting_pt[0], starting_pt[1], data_array.length, all_range[1]-starting_pt[1]+1)
      }
      else {
        var range_to_clear = range;
      }
      range_to_clear.clearContent();
      range.setValues(data_array)
    } else {
      var err = "failed to lock"
      throw err
      msg(err)
    }
    lock.releaseLock()  
    return  
}

/**
* Get input data from prompt dialog
*
* @param {string} message to ask
* @param {number} 1 to show the confirm dialog
* @param {string} message to return
* @return {string} input data string
*/
function showPrompt(message, show_confirm, msg_confirm) {
    this.msg("[Stella Library] Enter showPrompt")
    var result = Browser.inputBox(
        message,
        Browser.Buttons.OK_CANCEL);

  // Process the user's response.
    if (show_confirm) {
      if (result != 'cancel') {
        // User clicked "OK".
        Browser.msgBox(msg_confirm + result + '.');
      } else {
        // User clicked "Cancel" or X in the title bar.
        Browser.msgBox('You did not set the value properly, might cause errors.');
      }
    }
    return result
}
/**
* Compare if the day, month and year of two Date objects are the same
*
* @param {Date} first date object
* @param {Date} second date object
* @return {boolen} true if the two dates match
*/
function compareDates(d1, d2){
  if (d1.getDate() != d2.getDate()) {return false}
  if (d1.getMonth() != d2.getMonth()) {return false}
  if (d1.getYear() != d2.getYear()) {return false}
  return true
}
/**
* Match two strings
*
* @param {string} base string (longer)
* @param {string} sub-string to be compared (shorter)
* @param {boolen} true if case sensitive
* @return {number} 0
*/
function matchString(s1, s2, case_sensitive){
  if (typeof s1 == "string" && typeof s2 == "string"){
    if ( !case_sensitive){
      s1 = s1.toLowerCase()
      s2 = s2.toLowerCase()
    }
    if (s1.indexOf(s2) > -1){
      return 1
    }
  }
  return 0
}
/**
* Convert a 2D column array contains only one column to 1D array
*
* @param {Array} 2D column array
* @return {Array} 1D array
*/
function convert1DCol(two_dimension_array){
  var return_array = new Array()
  for (var i = 0; i < two_dimension_array.length; i++) {
    return_array.push(two_dimension_array[i][0])
  }
  return return_array
}
/**
* Calculate day difference of two date objects
*
* @param {Date} d1 (older)
* @param {Date} d2 (newer)
* @return {number} Day difference
*/
function calDaysDiff(d1, d2){
  return Math.floor((d2 - d1) / (1000*60*60*24))
}

/**
* Calculate hour difference of two date objects
*
* @param {Date} d1 (older)
* @param {Date} d2 (newer)
* @return {number} Hour difference
*/
function calHoursDiff(d1, d2){
  return Math.floor((d2 - d1) / (1000*60*60))
}
/**
* [Launchpad]: Convert from the string returned by Launchpad to Date object
*
* @param {string} launchpad date string
* @return {Date} Date object
*/
function _convertLPTime(launchpad_time_st){
  if(launchpad_time_st == null){
    return new Date()}
  var day = launchpad_time_st.split('T')[0].replace(/-/g, "/")
  var time = launchpad_time_st.split('T')[1].split('.')[0]
  return new Date(day + " " + time)
}

/**
* [Launchpad][Stella]: Filter out platform tags
*
* @param {Array} List of tags
* @return {Array} Filtered list of tags
*/
function _filterTags(tags_list){
  tags_list = tags_list.sort().filter( function(v,i,o){ return !i||v&&!RegExp(o[i-1],'i').test(v)});
  var filter_tags = new Array()
  var string_to_check = ["stella", "hwe", "cert", "cqa-verified"]
  for(var i in tags_list) {
    var flag = 0
    for(var j in string_to_check) {
      if (this.matchString(tags_list[i], string_to_check[j])) {
        flag = 1
        break
      }
    }
    if(flag == 0) {
        filter_tags.push(tags_list[i])
    }
  }  
  return filter_tags
}

/**
* [Launchpad][Stella]: Get link of a given milestone name
*
* @param {string} Milestron name
* @return {string} api url of the milestone
*/
function _getMilestoneLink(milestone_name) {
  return req_prefix + "stella/%2bmilestone/" + milestone_name
}

/**
* [Launchpad][Stella]: Get link of a given person name
*
* @param {string} Person name
* @return {string} api url of the person
*/
function _getPersonLink(name){
  return req_prefix + "~" + name
}

/**
* [Launchpad]: Return the target series of the given milestone
*
* @param {string} Milestron name
* @return {string} Name of the target series
*/
function _findSeriesFromMilestone(milestone_name) {
  milestone = this._getMilestoneLink(milestone_name)
  lp_milestone = this.fetchURL(milestone)
  return lp_milestone.series_target_link.split('/').pop()
}

/**
* [Launchpad][Stella]: Return a list of bt of a given assignee name
*
* @param {string} Assignee name
* @param {boolen} True to calculate HPS
* @return {Array} A list of bugs
*/
function _getBugsByAssignee(assignee, cal_hps) { 
  var person = this._getPersonLink(assignee)
  var suffix = "?ws.op=searchTasks&ws.size=300&" + "assignee=" + person
  var req_url = req_prefix + "stella" + suffix
  msg("req_url="+req_url)
  var stella_bugs = _getBugs(req_prefix + "stella" + suffix, cal_hps)
  var stella_base_bugs = _getBugs(req_prefix + "stella-base" + suffix, cal_hps)
  return stella_bugs.concat(stella_base_bugs)
}

/**
* [Stella]: Calculate HPS (rules see http://goo.gl/YualcF)
*
* @param {Object} bug
* @return {number} HPS value
*/
function __calHPS(bug){
  if (bug.status == "Fix Committed" || bug.status == "Fix Released" || bug.status == "Invalid" || bug.status == "Won't Fix"){
    return 0
  }
  var hps_main = 0
  var tags = bug.tags.split(",")
  function hps_main_by_tag(search_str, hps_tag){
    var index = tags.indexOf(search_str)
    if (index > -1){
      hps_main += hps_tag
      tags.splice(index, 1);
    }
  }
  hps_main_by_tag('stella-tools', 0.3);
  hps_main_by_tag('stella-oem-highlight', 0.7);
  hps_main_by_tag('hwe-cert-risk', 0.6);
  
  if (matchString(bug.title, 'SIO', true) || matchString(bug.title, 'OEM', true)){hps_main += 0.1}

  var hps_sub = {"Critical": 1.1,
                 "High": 1.0,
                 "Medium": 0.3,
                 "Low": -0.2,
                 "Wishlist": -0.3,
                 "Undecided": 0.0 }[bug.importance]
  var now = new Date()
  //Calculate how many days has the bug been opened
  var date_to_cal = (bug.date_confirmed!=null)?bug.date_confirmed:bug.date_created
  var days_confirmed_or_opened = calDaysDiff(date_to_cal, now)
  function hps_sub_open_days(ndays){
    if(ndays > 21){return 0.3} 
    else {return (ndays/7)*0.1}
  }
  hps_sub += hps_sub_open_days(days_confirmed_or_opened)
  //Calculate how many days it is close to GM
  if (bug.target != null && bug.target.indexOf('stella-base') > -1) {
    return hps_main + hps_sub
  }
  var cutoff_cycle = '3C14'
  var shortest = 1000
  for (var i = 0; i < tags.length; i++) {
    if (tags[i].indexOf('stella') < 0) {continue}
    var platform_gm = _getPlatformGMDate(tags[i].replace(/stella-/g, ""), cutoff_cycle)
    if (platform_gm != null ){
      var days_to_gm = calDaysDiff(now, platform_gm)
      if(days_to_gm < shortest) {shortest = days_to_gm}
    } 
  }
  var targets = bug.target
  if (shortest == 1000 && targets != null){
    for (var i = 0; i < targets.length; i++) {
      if (matchString(targets[i], 'stella')) {continue}
      var target_gm = _getTargetGMDate(targets[i])
      if (target_gm != null ){
        var days_to_gm = calDaysDiff(now, target_gm)
        if(days_to_gm < shortest) {shortest = days_to_gm}
      }
    }
  }
  if (shortest <= 14) {hps_sub += 0.2}
  if (shortest <= 7) {hps_sub += 0.1}
  return Number(hps_main + hps_sub).toFixed(2)
}

/**
* [Stella]: Get platform GM from Project Book
*
* @param {string} platform name
* @param {string} search for no earlier than this cycle
* @return {Date} date object, null if not found
*/
function _getPlatformGMDate(platform_name, cutoff_cycle){
  msg(platform_name)
  var project_book_key = "0AlgJx7XzUn6hdEVrY2d3bWhrMXpWdUxhZGlLVEJZNUE";
  var project_book = SpreadsheetApp.openById(project_book_key);
  var all_sheets = project_book.getSheets()
  var col_cutoff = 41 //skip milestone columns
  var found = false
  for (iSheet in all_sheets){
    if (iSheet < 3) { continue } //Skip the first two sheets
    if (matchString(all_sheets[iSheet].getSheetName(), cutoff_cycle)) {break} //Skip old sheets
    var title_row = all_sheets[iSheet].getSheetValues(1, 1, 1, col_cutoff)[0]
    var sheet_range = getDataRange(all_sheets[iSheet]);
    var platform_tag = null
    var gm_date = null
    msg(all_sheets[iSheet].getSheetName())
    for (var i = 0; i < title_row.length; i++) {
      if (matchString(title_row[i], "platform tag")){ 
        platform_tag = all_sheets[iSheet].getSheetValues(1, i+1, sheet_range[0], 1)
      }
      else if (matchString(title_row[i], "GM")){ 
        gm_date = all_sheets[iSheet].getSheetValues(1, i+1, sheet_range[0], 1)
      }
      if (platform_tag != null && gm_date != null) {break}
    }
    var gm_date_of_platform = null
    for (var i = 0; i < platform_tag.length; i++) {
      if (matchString(platform_tag[i][0], platform_name)){ return gm_date[i][0] }
    }
  }
  return null
}
/**
* [Stella]: Get series GM from Project Overview
*
* @param {string} series(target) name
* @return {Date} date object, null if not found
*/
function _getTargetGMDate(target_name){
  var project_overview = __getSheetInProjectOverview("Project Overview")
  var series_ix = -1
  var gm_ix = -1
  for (var i = 0; i < project_overview[0].length; i++) {
    if (matchString(project_overview[0][i], 'Series')) {series_ix = i}
    else if (matchString(project_overview[0][i], 'GM')) {gm_ix = i}
    if (series_ix > -1 && gm_ix > -1) {break}
  }
  var target_row = -1
  for (var i = 0; i < project_overview.length; i++) {
    if (matchString(project_overview[i][series_ix], target_name)) {
      target_row = i;
      break
    }
  }
  if (target_row < 0) {return new Date()}
  return project_overview[target_row][gm_ix]
}
/**
* [Stella] Get a sheet data in project overview spreadsheet
*
* @param {string} Name of the sheet
* @return {Sheet} Sheet with the given name
*/
function __getSheetInProjectOverview(sheet_name){
  var project_overview_key = __getProjectOverviewKey();
  var project_overview = SpreadsheetApp.openById(project_overview_key);
  var ss = project_overview.getSheetByName(sheet_name);
  return getSheetData(ss)
}
/**
* [Stella] Get the spreadsheet object of project overview spreadsheet
*
* @return {Spreadsheet} Project Overview
*/
function __getProjectOverview(){
  var project_overview_key = __getProjectOverviewKey();
  var project_overview = SpreadsheetApp.openById(project_overview_key);
  return project_overview
}
/**
* [Stella] Get the spreadsheet key of project overview spreadsheet
*
* @return {string} Key of the Project Overview
*/
function __getProjectOverviewKey(){
  return "0ApMQND2fzJEVdE0ycU5RNGw4Znl3WE5KT3k0a3FGbVE"
}
/**
* [Stella]: Find HW tracking table for a given platform or module name
*
* @param {string} name or keyword of a platform or a module
* @param {string} specify the row to search
* @param {string} specify the HP cycle to search
* @return {Array} A list results found
*/
function _searchHWTracking(row_label, search_kwd, hp_cycle){
  var project_overview = __getProjectOverview()
  var hw_tracking_range = project_overview.getRangeByName("hwtracking");
  var doc_array = hw_tracking_range.getValues()
  var found_array = new Array()
  var cycle_to_search = new Array()
  for (var k = 0; k < doc_array.length; k++) {
    cycle_to_search.push(doc_array[k][1])
  }
 
  if (cycle_to_search.indexOf(hp_cycle) > -1) {cycle_to_search = [hp_cycle]}  
  found_array.push(["Found Item", "HP cycle","Link to table", "Sheet Name", "Platform Label"])
  
  for (var i = 0; i < doc_array.length; i++) {

    var cycle = doc_array[i][1]
    if(cycle_to_search.indexOf(cycle) < 0){ continue; }

    var sheet_key = doc_array[i][3]
    var link = doc_array[i][2]
    var hw_tracking_table = SpreadsheetApp.openById(sheet_key);
    var all_sheets = hw_tracking_table.getSheets()
    for (var iSheet in all_sheets){
      var sheet_range = getDataRange(all_sheets[iSheet]);
      var first_col = all_sheets[iSheet].getSheetValues(1, 1, sheet_range[0], 1)
      var platform_row = all_sheets[iSheet].getSheetValues(2, 1, 1, sheet_range[0])[0]
      var row_ix_array = new Array()
      for (var j = 0; j < first_col.length; j++) {
        if (matchString(first_col[j][0], row_label)){row_ix_array.push(j + 1);}
      }
      if (row_ix_array.length < 1) {continue;}
      for (var ix = 0; ix < row_ix_array.length; ix++) {
        var row_ix = row_ix_array[ix]
        var row_to_search = all_sheets[iSheet].getSheetValues(row_ix, 1, 1, sheet_range[1])[0]
        for (var irow in row_to_search){
          if (matchString(row_to_search[irow], search_kwd)){
            var sheet_name = all_sheets[iSheet].getSheetName()
            found_array.push([row_to_search[irow], cycle, link, sheet_name, platform_row[irow]])
            break
          }//End if
        }//End loop of platforms in one sheet
      }//End looping of all matched rows
    }//End loop of all sheets
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  }//End loop of documents
  return found_array
}
/**
To be finished
*/
function _getBugByID(bugid){
  __DEBUG__ = true
  bugid = '1323153'
  var req_url = req_prefix + 'bugs/' + bugid
  var bug_content = this.fetchURL(req_url)
  bug_content.description = 'test'
  var req_url_modified = req_prefix + 'bugs/' + bugid 
  var options ={
    headers : this.create_dynamic_customHeaders(), 
    method : "put", 
    payload: bug_content, 
    "Content-Type":"application/json"
  };
  var res = UrlFetchApp.fetch(req_url_modified, options)
  msg(res)
}
/**
* [Launchpad]: Return a list of bt for a given url
*
* @param {string} Full url
* @param {boolen} HPS is only calculated if this is true and returned bug list is sorted by HPS.
* @return {Array} A list of bugs
*/
function _getBugs(req_url, cal_hps){  
  var entries = this.fetchURL(req_url).entries
  var return_bugs = new Array()
  for( var i in entries ) {
    var owner = /.*~(.*)/.exec(entries[i].owner_link)
    var assignee = /.*~(.*)/.exec(entries[i].assignee_link)
    var title = /"(.*)"/.exec(entries[i].title)
    var id = /(\d+)$/.exec(entries[i].bug_link)
    var date_confirmed = _convertLPTime(entries[i].date_confirmed)
    var date_created = _convertLPTime(entries[i].date_created)
    var date_triaged = _convertLPTime(entries[i].date_triaged)
    var tags = entries[i].tags
    var _bug = this.fetchURL(entries[i].bug_link)
    var tags = _bug.tags

    var bug_tasks = this.fetchURL(_bug.bug_tasks_collection_link)
    var target = new Array()
    for (var itask in bug_tasks.entries){
      target.push(bug_tasks.entries[itask].bug_target_name.split('/').pop())
    }

    var bt = {owner:(owner!=null)?owner[1]:"",
               assignee:(assignee!=null)?assignee[1]:"",
               importance:entries[i].importance,
               hps:0,
               target:(target != null)?target:null,
               date_confirmed:(date_confirmed != null)?date_confirmed:"",
               date_created:(date_created != null)?date_created:"",
               date_triaged:(date_triaged != null)?date_triaged:"",
               package:(target!=null)?target[1]:"",
               status:entries[i].status,
               link:entries[i].web_link,
               title:(title!=null)?title[1]:"",      
               id:(id!=null)?id[1]:"",
               tags:(tags!=null)?tags.join():"",
              }
    if (cal_hps) {bt.hps = __calHPS(bt)}
    return_bugs.push(bt)      
  }   
  if (cal_hps) {return_bugs.sort(function(b, a) {return a.hps - b.hps})}
  else {return_bugs.sort(function(b, a) {return a.id - b.id})}

  return return_bugs
}

/**
* [Launchpad]: Combine tags for searching bugtasks
*
* @param {Array} A list of tags
* @return {string} Combined tag string
*/
function _combineTags(tags){
  var tag_parameter = new Array()

  for( var i in tags ) {
    tag_parameter.push("tags="+tags[i])
  }
  
  tag_parameter = tag_parameter.join("&")
  return tag_parameter
}


/**
* [Stella]: Return the spreadsheet object of the engineering dashboard
*
* @return {Array} Data array in the Data sheet
*/
function _getDashboardSS(){
  var dashboard_key = "1pMiFUMgdlTD0cBadYL-JH3pzAOHTawdKcniyn8oucUE";
  return SpreadsheetApp.openById(dashboard_key)
}
/**
* [Stella]: Return data array in the Data sheet of the dashboard spreadsheet
*
* @return {Array} Data array in the Data sheet
*/
function _getEngInfoFromDataSheet(){
  var dashboard = _getDashboardSS()
  return dashboard.getRangeByName('enginfo').getValues()
}

/**
* [Stella]: Return the defined names of the engineers
*
* @return {Array} the name array
*/
function _getDefinedEngNames(){
  var dashboard = _getDashboardSS()
  var _values = dashboard.getRangeByName('engnames').getValues()
  return convert1DCol(_values)
}

/**
* [Stella]: Return the defined IDs of the engineers
*
* @return {Array} the ID array
*/
function _getDefinedEngIDs(){
  var dashboard = _getDashboardSS()
  var _values = dashboard.getRangeByName('engids').getValues()
  return convert1DCol(_values)
}
/**
* [Stella]: Return the release plan array
*
* @return {Array} Release plan of the week
*/
function _getReleasePlan(){
  var first_day_of_week = getFirstDayString();
  var release_plan = __getSheetInProjectOverview(first_day_of_week)
  var col_to_find = {"Actual Time":-1, "Expect Time":-1, 
                     "IBS Project":-1, "LP Series":-1, "LP Milestone":-1}
  for (var i = 0; i < release_plan[0].length; i++) {
    for (var title in col_to_find){
      if (col_to_find[title] > -1) {continue}
      if (release_plan[0][i].indexOf(title) > -1) {col_to_find[title] = i}
    }
  }
  var return_plan = new Array()
  for (var i = 1; i < release_plan.length; i++) {
    var tmp = new Array()
    if (typeof release_plan[i][col_to_find["Expect Time"]] == 'string') {break}
    tmp.push(release_plan[i][col_to_find["Actual Time"]])
    tmp.push(release_plan[i][col_to_find["Expect Time"]])
    tmp.push(release_plan[i][col_to_find["IBS Project"]])
    tmp.push(release_plan[i][col_to_find["LP Series"]])
    tmp.push(release_plan[i][col_to_find["LP Milestone"]])
    return_plan.push(tmp)
  }
  return return_plan
}
/**
* [Stella]: Return the launchpad id of the engineer on duty this week
*
* @param {boolen} True to convert launchpad ID to the name
* @return {string} Engineer on duty this week
*/
function _getDutyOfTheWeek(convert_id){
  var project_overview = __getProjectOverview()
  var range = project_overview.getRangeByName("weeklyduty");
  var assignee = range.getValue()
  if (convert_id){
    assignee = _convertIDName(assignee)
  }
  return assignee
}
/**
* [Stella]: Return the launchpad id of the engineer on duty this week
*
* @param {boolen} True to convert launchpad ID to the name
* @return {string} Engineer on duty this week
*/
function _convertIDName(eng){
  var id_array = _getDefinedEngIDs()  
  var name_array = _getDefinedEngNames()
  var ix = id_array.indexOf(eng)
  if (ix > -1){
    return name_array[ix]
  }
  ix = name_array.indexOf(eng)
  if (ix > -1){
    return id_array[ix]
  }
}
/**
* [Stella]: Sending emails to the whole Stella Mainstream team
*
* @param {string} Email title
* @param {string} Email body
* @param {boolen} True to also send a backup to the line manager
* @param {boolen} True to send to TL only
*/
function _emailStellaMainstream(title, body, send_to_manager, test){
  body = body.replace(/\n/g, " <br>")
  body = body.split(" ")
  for (var ibody = 0; ibody < body.length; ibody++){
    if (matchString(body[ibody], 'lp:')) {
      var bugid = body[ibody].replace('lp:','').replace(',','').replace('.','')
      body[ibody] = '<a href="' + __getBugLink(bugid) + '">' + bugid + '</a>'
    }
  }
  body = body.join(" ")
  var manager_email = ""
  var data = _getEngInfoFromDataSheet()
  var name_col_ix = 1
  var tl_col_ix = 4
  if (__checkAuth()){
    for (var idata = 1; idata < data.length; idata++){
      if (test && data[idata][tl_col_ix] != 'Y') { continue; }
      GmailApp.sendEmail(data[idata][3], title, '\n', {htmlBody: body});
    }
    if ( !test && send_to_manager){
      GmailApp.sendEmail(manager_email, title, '\n', {htmlBody: body});
    }
  }
  return
}
/**
* [Launchpad]: Return the full link of the bug
*
* @param {string} bugid 
* @return {string} full link
*/
function __getBugLink(bugid){
  return 'http://launchpad.net/bugs/' + bugid.toString()
}

function __checkAuth(e) {
  var addonTitle = 'My Add-on Title';
  var props = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    var lastAuthEmailDate = props.getProperty('lastAuthEmailDate');
    var today = new Date().toDateString();
    if (lastAuthEmailDate != today) {
      if (MailApp.getRemainingDailyQuota() > 0) {
        var html = HtmlService.createTemplateFromFile('AuthorizationEmail');
        html.url = authInfo.getAuthorizationUrl();
        html.addonTitle = addonTitle;
        var message = html.evaluate();
        MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
            'Authorization Required',
            message.getContent(), {
                name: addonTitle,
                htmlBody: message.getContent()
            }
        );
      }
      props.setProperty('lastAuthEmailDate', today);
    }
    return false
  } 
  else {
    return true
  }
}

