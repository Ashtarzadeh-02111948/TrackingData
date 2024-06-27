const PROCESS_SHEET_NAME = "Main"
const MAX_NO_TRACKING_NO_PER_PROCESS = 5000
const white_color = "#b7e1cd"
const light_green_color = "#19a737"
const light_grey_color = '#58cbf5'
const orange_color = '#ed7d31'
const yellow_color = '#bded55'
const red_color = '#ff0000'
const USPS_CHUNK_SIZE = 25
const USPS_USER_ID = "[REDACTED]"


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Initialize for coloring', 'initialize')
      .addItem('Check Tracking Status', 'processMonitorTracking')
      .addToUi();
}

function processMonitorTracking() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() < 2) {
    return;
  }
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var name_col = findColNoByName("Tracking", header);
  var track_no_col = findColNoByName("Tracking", header);
  var management_col = findColNoByName("Manual update", header);
  var track_status_col = findColNoByName("Tracking Status", header);
  if ((name_col == 0) || (track_no_col == 0) || (track_status_col == 0)) {
    console.log("Invalid column name");
    return;
  }
  var script_properties = PropertiesService.getScriptProperties();
  var current_last_row = sheet.getLastRow();
  var start_process_row = parseInt(script_properties.getProperty("LastProcessRow")) + 1;
  if (start_process_row > current_last_row) {
    start_process_row = 2;
  }

  var input_track_no_data = sheet.getRange(start_process_row, track_no_col, current_last_row - start_process_row + 1, 1).getValues();
  var input_track_status_data = sheet.getRange(start_process_row, track_status_col, current_last_row - start_process_row + 1, 1).getValues();
  var current_color_data = sheet.getRange(start_process_row, name_col, current_last_row - start_process_row + 1, 1).getBackgrounds();
  var management_data = sheet.getRange(start_process_row, management_col, current_last_row - start_process_row + 1, 1).getValues();
  var output_track_status_data = [];
  var color_update_data = [];
  var track_no_list = [];
  var row_info_dict = {};
  for (var index = 0; index < input_track_no_data.length; index++) {
    var current_input_track_status_data = input_track_status_data[index][0];
    if (management_data[index][0] != true) {
      var current_track_no_data = input_track_no_data[index][0];
      if (current_track_no_data.length > 0) {
        var is_delivered = false;
        if (current_input_track_status_data.length > 0) {
          var track_status_list = current_input_track_status_data.split("|");
          var no_index = 0;
          do {
            var current_status = track_status_list[no_index].toLowerCase();
            is_delivered = current_status.includes("delivered");
            no_index++;
          } while (is_delivered && (no_index < track_status_list.length))
        }
        if (is_delivered) {
          color_update_data.push([light_green_color]);
          output_track_status_data.push([current_input_track_status_data]);
        } else {
          var row_track_no_list = current_track_no_data.split(",");
          var index_key = index.toString();
          row_info_dict[index_key] = [];
          for (var track_no_index = 0; track_no_index < row_track_no_list.length; track_no_index++) {
            var track_no = row_track_no_list[track_no_index].trim();
            row_info_dict[index_key].push(track_no);
            track_no_list.push(track_no);
          }
          // var color_code = evaluate_tracking_status(track_status_data_list);
          color_update_data.push("");
          output_track_status_data.push("");
        }
      } else {
        color_update_data.push([white_color]);
        output_track_status_data.push([""]);
      }    
    } else {
      color_update_data.push([current_color_data[index][0]]);
      output_track_status_data.push([current_input_track_status_data]);
    }
  }
  var track_no_dict = {}
  var job_no_list = [];
  track_no_list.forEach(track_no => {
    var carrier = guessCarrier(track_no);
    if (carrier == "usps") {
      job_no_list.push(track_no);
    } else {
      track_no_dict[track_no] = {
        "status": "Invalid Carrier: " + carrier
      }
    }
  })
  var job_list = split_array(job_no_list, USPS_CHUNK_SIZE);
  job_list.forEach(job_id_list => {
    var result_dict = getTrackingStatus(job_id_list);
    track_no_dict = Object.assign({}, track_no_dict, result_dict);
  })

  for (var row_index in row_info_dict) {
    var id_list = row_info_dict[row_index];
    var track_status_data_list = [];
    id_list.forEach(track_no => {
      track_status_data_list.push(track_no_dict[track_no]["status"])
    })
    var color_code = evaluate_tracking_status(track_status_data_list);
    color_update_data[parseInt(row_index)] = [color_code];
    output_track_status_data[parseInt(row_index)] = [track_status_data_list.join("|")];
  }

  sheet.getRange(start_process_row, track_status_col, index, 1).setValues(output_track_status_data);
  sheet.getRange(start_process_row, track_no_col, index, 1).setBackgrounds(color_update_data);
  script_properties.setProperty("LastProcessRow", start_process_row + index - 1);
}


function getTrackingStatus(trackingNoList) {
  var result_dict = {};
  var params = {
    "method": "get",
    "headers": {
      "Accept": "text/xml"
    }
  }

  var request = `API=TrackV2&XML=<TrackFieldRequest USERID="${USPS_USER_ID}">`;
  
  for (var i = 0; i < trackingNoList.length; i++) {
    request += `<TrackID ID="${trackingNoList[i]}"></TrackID>`;
  }
  request += "</TrackFieldRequest>";
  var url = `https://secure.shippingapis.com/shippingapi.dll?${encodeURI(request)}`;
  var response = UrlFetchApp.fetch(url, params);
  var text = response.getContentText();
  var document = XmlService.parse(text);
  var root = document.getRootElement();
  var trackResponse = root.getChildren("TrackInfo");
  for (var i = 0; i < trackResponse.length; i++) {
    var track_result = trackResponse[i];
    var trackingNumber = track_result.getAttribute("ID").getValue();
    var tracking_status = "";
    if (track_result.getChild("Error")) {
      tracking_status = "Invalid Tracking Number"
    } else {
      var trackingSumary = track_result.getChild("TrackSummary");
      tracking_status = trackingSumary.getChild("Event").getText().split(",")[0];
    }
    result_dict[trackingNumber] = {
      "status": tracking_status
    }
  }
  return result_dict;
}


function evaluate_tracking_status(track_status_list) {
  var is_delivered = true;
  var is_delivering = true;
  var is_pick_up = true;
  var color_code = "";
  for (var track_index = 0; track_index < track_status_list.length; track_index++) {
    current_status = track_status_list[track_index].toLowerCase();
    if (current_status.includes("invalid")) {
      color_code = yellow_color;
      break;
    }

    if (current_status.includes("alert")) {
      color_code = red_color;
      break;
    }       
    var temp_is_delivered = current_status.includes("delivered") || current_status.includes("out for delivery");
    is_delivered &= temp_is_delivered;
    var temp_is_delivering = temp_is_delivered || current_status.includes("arrived shipping partner facility") || current_status.includes("departed usps") || current_status.includes("arrived at usps") || current_status.includes("departed shipping partner facility") || current_status.includes("transit") || current_status.includes("awaiting delivery scan");
    is_delivering &= temp_is_delivering;
    var temp_is_pick_up = temp_is_delivering || current_status.includes("available for pickup") || current_status.includes("agent pickup") || current_status.includes("action needed") || current_status.includes("delivery attempt");
    is_pick_up &= temp_is_pick_up;
    var is_pending = current_status.includes("pre-shipment") || current_status.includes("on its way to usps") || current_status.includes("created");
    if ((!temp_is_delivered) && (!temp_is_pick_up) && (!temp_is_delivering) && (!is_pending)) {
      is_delivering = true;
    }
  } 
  if (color_code.length == 0) {
    if (is_delivered) {
      color_code = light_green_color;
    } else if (is_delivering) {
      color_code = light_grey_color;
    } else if (is_pick_up) {
      color_code = orange_color;
    } else if (is_pending){
      color_code = yellow_color;
    } else {
      color_code = white_color;
    }
  }
  return color_code;
}


function initialize() {
  var script_properties = PropertiesService.getScriptProperties();
  script_properties.setProperty("LastProcessRow", 1);
}


function findColNoByName(col_name, header) {
  for (var index = 0; index < header.length; index++) {
    if (col_name == header[index]) {
      return index + 1;
    }
  }
  return 0;
}


function guessCarrier(trackingNo) {
  const usps_regex = /\b([A-Z]{2}\d{9}[A-Z]{2}|(420\d{9}(9[2345])?)?\d{20}|(420\d{5})?(9[12345])?(\d{24}|\d{20})|82\d{8})\b/;
  const fedex_regex = /\b([0-9]{12}|100\d{31}|\d{15}|\d{18}|96\d{20}|96\d{32})\b/;
  const ups_regex = /\b1Z[A-Z0-9]{16}\b/;
  if (usps_regex.test(trackingNo)) {
    return "usps";
  } else if (fedex_regex.test(trackingNo)) {
    return "fedex";
  } else if (ups_regex.test(trackingNo)) {
    return "ups";
  } else {
    return "not supported";
  }
}


function split_array(array, chunkSize) {
  var result = [];
  for (let i = 0; i < array.length; i += chunkSize) {
      result.push(array.slice(i, i + chunkSize));
  }
  return result;
}

