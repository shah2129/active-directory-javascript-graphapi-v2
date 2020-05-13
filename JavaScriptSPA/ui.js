// Select DOM elements to work with
const welcomeDiv = document.getElementById("WelcomeMessage");
const signInButton = document.getElementById("SignIn");
const cardDiv = document.getElementById("card-div");
const mailButton = document.getElementById("readMail");
const profileButton = document.getElementById("seeProfile");
const profileDiv = document.getElementById("profile-div");
const readBookingsButton = document.getElementById("readBookings");

var apptCnt = 0;
var staffCnt = 0;
var BusinessArray = [];
var AppointmentArray = [];
var StaffArray = [];
var StaffNames = [];
var StaffMembers = [];
var totalBusinessCnt = 0;
var businessesProcessedCnt = 0;
var appointmentTable = "";

$(document).ready(function () {
  $("#datepicker").datepicker();
  $("#datepicker2").datepicker();
});

function showWelcomeMessage(account) {
  // Reconfiguring DOM elements
  cardDiv.style.display = 'initial';
  welcomeDiv.innerHTML = `Welcome ${account.name}`;
  signInButton.nextElementSibling.style.display = "none";
  signInButton.setAttribute("onclick", "signOut();");
  signInButton.setAttribute("class", "btn btn-success");
  signInButton.innerHTML = "Sign Out";
  document.getElementById("signIn").innerHTML = "Sign Out";
  document.getElementById("signIn").setAttribute("onclick", "signOut();");

  document.getElementById("inputGroup2").style.display = "block";
  //document.getElementById("statsGroup2").style.display = "block";
}

function updateUI(data, endpoint, token) {
  console.log("Graph API responded at: " + new Date().toString());

  if (endpoint === graphConfig.graphMeEndpoint) {
    const title = document.createElement("p");
    title.innerHTML = "<strong>Title: </strong>" + data.jobTitle;
    const email = document.createElement("p");
    email.innerHTML = "<strong>Mail: </strong>" + data.mail;
    const phone = document.createElement("p");
    phone.innerHTML = "<strong>Phone: </strong>" + data.businessPhones[0];
    const address = document.createElement("p");
    address.innerHTML = "<strong>Location: </strong>" + data.officeLocation;
    profileDiv.appendChild(title);
    profileDiv.appendChild(email);
    profileDiv.appendChild(phone);
    profileDiv.appendChild(address);
  } else if (endpoint === graphConfig.graphMailEndpoint) {
    if (data.value.length < 1) {
      alert("Your mailbox is empty!");
    } else {
      const tabList = document.getElementById("list-tab");
      const tabContent = document.getElementById("nav-tabContent");

      data.value.map((d, i) => {
        // Keeping it simple
        if (i < 10) {
          const listItem = document.createElement("a");
          listItem.setAttribute(
            "class",
            "list-group-item list-group-item-action"
          );
          listItem.setAttribute("id", "list" + i + "list");
          listItem.setAttribute("data-toggle", "list");
          listItem.setAttribute("href", "#list" + i);
          listItem.setAttribute("role", "tab");
          listItem.setAttribute("aria-controls", i);
          listItem.innerHTML = d.subject;
          tabList.appendChild(listItem);

          const contentItem = document.createElement("div");
          contentItem.setAttribute("class", "tab-pane fade");
          contentItem.setAttribute("id", "list" + i);
          contentItem.setAttribute("role", "tabpanel");
          contentItem.setAttribute("aria-labelledby", "list" + i + "list");
          contentItem.innerHTML =
            "<strong> from: " +
            d.from.emailAddress.address +
            "</strong><br><br>" +
            d.bodyPreview +
            "...";
          tabContent.appendChild(contentItem);
        }
      });
    }
  } else { // MS Booking Testing related logic. 


    // loop through Businesses to get appointments
    $.blockUI({
      message: "Obtaining Bookings."
    });
    document.getElementById("ProcessingID").style = "none";
    StaffMembers.splice(0, staffCnt);
    businessesProcessedCnt = 0;
    apptCnt = 0;
    staffCnt = 0;
    AppointmentArray.splice(0, AppointmentArray.length);
    document.getElementById("appointmentTable").innerHTML = "<pre></pre>";
    document.getElementById("staffCnt").innerHTML = "";


    appointmentTable = `<table class="table .table-striped">
  <thead>
  <tr>
    <th>Business Name</th>
    <th>Customer Name</th>
    <th>Customer Email</th>  
    <th>Booking Start Date</th>
    <th>Booking Start Time</th>
    <th>Booking End Date</th>
    <th>Booking End Time</th>
    <th>Viewing Staff</th>
  </tr>`;
    document.getElementById("csv_dl").style = "visibility:hidden;";


    var currentBusiness = "";
    totalBusinessCnt = data.value.length;
    for (var b = 0; b < data.value.length; b++) {
      // currentBusiness = data.value[b].displayName
      //console.log(currentBusiness)
      BusinessArray.push(data.value[b]);

      let staffPayload = getStaffMembersForSpecificBusiness(
        data.value[b],
        token
      );
      staffPayload.then(function (staffData) {
        if (staffData.value != null) {
          staffData.value.forEach(function (arrayItem) {
            var x = arrayItem;
            // console.log(x);

            if (StaffMembers[x.id] == undefined) {
              var staffMember = new Object();
              staffMember.id = x.id;
              staffMember.displayName = x.displayName;
              staffMember.role = x.role;
              StaffMembers[x.id] = staffMember;
              staffCnt += 1;
            }
            document.getElementById("staffCnt").innerHTML = staffCnt;
          });
        }
      });

      let appointmentPayload = getAppointmentsForSpecificBusiness(
        data.value[b],
        token
      );

      appointmentPayload.then(function (appointmentData) {
        businessesProcessedCnt += 1;
        if (appointmentData.businessName != null) {
          // console.log(businessesProcessedCnt)

          appointmentData.forEach(function (arrayItem) {
            var x = arrayItem;
 
            currerentAppointment = parseAppointmentData(
              x,
              appointmentData.businessName
            );


            document.getElementById("businessID").innerHTML = appointmentData.businessName;

            if (currerentAppointment.inRange ||
              ((currerentAppointment.inRangeE.getDate() == currerentAppointment.inRangeD.getDate()) &&
                (currerentAppointment.inRangeE.getMonth() == currerentAppointment.inRangeD.getMonth())
              )) {
              apptCnt += 1;
              document.getElementById("apptCnt").innerHTML = apptCnt;
              AppointmentArray.push(currerentAppointment);
              appointmentTable += appointmentTableRow(currerentAppointment);
            }
          });
        }

        // console.log(appointmentData.length+"~~"+JSON.stringify(appointmentData, undefined, 4))

        if (businessesProcessedCnt == totalBusinessCnt) {
          appointmentTable += "</table>";
          document.getElementById(
            "appointmentTable"
          ).innerHTML = appointmentTable;
          document.getElementById("ProcessingID").style = "visibility:hidden;";
          show_csv_download();
          $.unblockUI();
          $(document).ready(function () {});
        }
      }).catch(function (err) {
        console.log(businessesProcessedCnt)
        if (businessesProcessedCnt == totalBusinessCnt) {
          appointmentTable += "</table>";
          document.getElementById(
            "appointmentTable"
          ).innerHTML = appointmentTable;

          show_csv_download();
          $.unblockUI();
          document.getElementById("ProcessingID").style = "visibility:hidden;";
          $(document).ready(function () {});
        }

      });
    } // business loop
  }
} //function updateUI(data, endpoint,token) {

function appointmentTableRow(a) {
  var t_html_row = "";

  t_html_row = `<tr>       
              <td>${a.business}</td> 
              <td>${a.customerName}</td> 
              <td>${a.customerEmailAddress}</td>               
              <td>${a.start_date_formatted}</td> 
              <td>${a.start_time}</td> 
              <td>${a.end_date_formatted}</td> 
              <td>${a.end_time}</td>    
              <td>${a.viewer}</td>         
              </tr>`;

  return t_html_row;
}

 

async function getAppointmentsForSpecificBusiness(b, token) {
  let appointmentDataArray = [];
  appointmentDataArray.splice(0, appointmentDataArray.length);
  const headers = new Headers();
  const bearer = `Bearer ${token}`;
  headers.append("Authorization", bearer);

  const options = {
    method: "GET",
    headers: headers,
  };
  // console.log(b.displayName)
 
  var startDate = $("#from").datepicker("getDate");
  var endDate = $("#to").datepicker("getDate");

  endDate.setHours(23);
  endDate.setMinutes(59);
  endDate.setSeconds(59);
  endDateUTC = endDate.toISOString();
  startDateUTC = startDate.toISOString();
  console.log(startDate.toString());
  console.log(startDateUTC);
  console.log(endDate.toString());
  console.log(endDateUTC);
  var count = 0;

  //  information/code taken from this stack overflow post
  //  https://stackoverflow.com/questions/54091374/get-all-data-from-rest-api-odata-response-with-node-recursively-via-odata-nex

  let finished = false;

  try {
    let response = await fetch(
      graphConfig.graphbookingBusinessesAppointments +
      "/" +
      b.id +
      `/calendarView?start=${startDateUTC}&end=${endDateUTC}`,
      options
    );
    let status = response.status;
    let appointmentData = await response.json();

    if (status == 200) {
      finished = true;
    } else {
      finished = false;
    }

    if (status == 416) { // 416 encountered because /calendarView has a limit to the quanity of bookings that are returned.
                         //    Use /appointments instead if 416 is encountered. 
      throw 416; 
    }

    appointmentDataArray = appointmentDataArray.concat(appointmentData.value);
  } catch (err) {
    // catches errors both in fetch and response.json
    console.log("status" + status + "::" + b.displayName);
    var url = `https://graph.microsoft.com/beta/bookingBusinesses/${b.id}/appointments`;

    while (!finished) {
      count++;
      console.log(count + url);
      let response = await fetch(url, options);
      let appointmentData = await response.json();
      appointmentDataArray = appointmentDataArray.concat(appointmentData.value);
      finished = appointmentData["@odata.nextLink"] === undefined;
      url = appointmentData["@odata.nextLink"];
    }

    console.log(err);
  }
  console.log(status + "status");
  appointmentDataArray.businessName = b.displayName;
  return appointmentDataArray;
}

async function getStaffMembersForSpecificBusiness(b, token) {
  const headers = new Headers();
  const bearer = `Bearer ${token}`;
  headers.append("Authorization", bearer);

  const options = {
    method: "GET",
    headers: headers,
  };
  // console.log(b.displayName)
  var staffEndpoint = `https://graph.microsoft.com/beta/bookingBusinesses/${b.id}/staffMembers/`;
  let response = await fetch(staffEndpoint, options);

  let StaffData = await response.json();

  return StaffData;
}

async function download_csv() {
  //var currentDate = $( ".selector" ).datepicker( "getDate" );
  var csvRows = [];
  var csvRow = [];
  var csv = "";
  csvRows.push("Name,Title\n");
  csvRow.splice(0, csvRow.length);
  csvRow.push("Booking_business");
  csvRow.push("CustomerName");
  csvRow.push("StartDate");
  csvRow.push("StartTime");
  csvRow.push("EndDate");
  csvRow.push("EndTime");

  csvRow.push("CustomerEmailAddress");
  /*
  csvRow.push("Appointment_duration")
  csvRow.push("PreBuffer")
  csvRow.push("PostBuffer")
  */

  csvRow.push("Staff_Viewer");

  csv += csvRow.join(",") + "\n";

  var t_staffMembers = "";
  AppointmentArray.forEach(function (row) {
    t_staffMembers = "";
    t_StaffMemberIDs = row.staffMemberIds.join(";");
    for (var l = 0; l < row.staffMemberIds.length; l++) {
      //console.log(row.staffMemberIds[l] + "Staff");
      if (StaffMembers[row.staffMemberIds[l]] != undefined) {
        if (l == 0) {
          t_staffMembers += StaffMembers[row.staffMemberIds[l]].displayName;
        } else {
          t_staffMembers +=
            ";" + StaffMembers[row.staffMemberIds[l]].displayName;
        }
      }
    }

    //csv += row.join(',');
    csvRow.splice(0, csvRow.length);
    csvRow.push(row.business);
    csvRow.push('"' + row.customerName + '"');
    csvRow.push(row.start_date_formatted);
    csvRow.push(row.start_time);
    csvRow.push(row.end_date_formatted);
    csvRow.push(row.end_time);
    csvRow.push(row.customerEmailAddress);

    //csvRow.push(row.duration)
    //csvRow.push(row.preBuffer)
    //csvRow.push(row.postBuffer)
    //csvRow.push(JSON.stringify(row.serviceLocation))

    csvRow.push('"' + t_staffMembers + '"');
    csv += csvRow.join(",") + "\n";
  });

  var hiddenElement = document.createElement("a");
  hiddenElement.href = "data:text/csv;charset=utf-8," + encodeURI(csv);
  hiddenElement.target = "_blank";
  hiddenElement.download = "bookingAppointments.csv";
  hiddenElement.click();
}

function parseAppointmentData(appointment, businessName) {

  //console.log(businessName)
  Appointment = new Object();
  var staffMemberNames = [];

  start_date_time_object = new Date(appointment.start.dateTime);
  start_dt_year = start_date_time_object.getFullYear();
  start_dt_month = start_date_time_object.getMonth() + 1;
  start_dt = start_date_time_object.getDate();

  customerName = "";
  customerEmailAddress = "";
  customerPhone = "";
  duration = "";
  preBuffer = "";
  postBuffer = "";
  serviceLocation = "";
  t_csv_line = "";

  if (start_dt < 10) {
    start_dt = "0" + start_dt;
  }

  if (start_dt_month < 10) {
    start_dt_month = "0" + start_dt_month;
  }

  end_date_time_object = new Date(appointment.end.dateTime);
  end_dt_year = start_date_time_object.getFullYear();
  end_dt_month = start_date_time_object.getMonth() + 1;
  end_dt = start_date_time_object.getDate();

  if (end_dt < 10) {
    end__dt = "0" + end_dt;
  }
  if (end_dt_month < 10) {
    end_dt_month = "0" + end_dt_month;
  }

  start_date_formatted = start_dt_month + "/" + start_dt + "/" + start_dt_year;
  end_date_formatted = end_dt_month + "/" + end_dt + "/" + end_dt_year;
  start_time = appointment.start.dateTime.substring(11, 16);
  end_time = appointment.end.dateTime.substring(11, 16);
  Appointment = new Object();
  Appointment.start_date_formatted = start_date_formatted;
  Appointment.id = appointment.id;
  Appointment.start_time = start_time;
  Appointment.end_time = end_time;
  Appointment.end_date_formatted = end_date_formatted;
  Appointment.business = businessName; //d.displayName
  Appointment.customerName = appointment.customerName;
  Appointment.customerEmailAddress = appointment.customerEmailAddress;
  Appointment.customerPhone = appointment.customerPhone;
  /*
  Appointment. duration=  appointment.duration  
  Appointment. preBuffer =appointment.preBuffer 
  Appointment.postBuffer = appointment.postBuffer 
  */

  Appointment.inRange = dates.inRange(
    start_date_time_object,
    $("#from").datepicker("getDate"),
    $("#to").datepicker("getDate")
  );

  Appointment.inRangeS = $("#from").datepicker("getDate");

  Appointment.inRangeE = $("#to").datepicker("getDate");

  Appointment.inRangeD = start_date_time_object;

  Appointment.serviceLocation = appointment.serviceLocation;
  Appointment.staffMemberIds = appointment.staffMemberIds;

  for (s = 0; s < appointment.staffMemberIds.length; s++) {
    try {
      var curStaff = appointment.staffMemberIds[s];
      var testrole = StaffMembers[appointment.staffMemberIds[s]].role;
      if (StaffMembers[appointment.staffMemberIds[s]].role == "viewer") {
        Appointment.admin =
          StaffMembers[appointment.staffMemberIds[s]].displayName;
        staffMemberNames.push(
          StaffMembers[appointment.staffMemberIds[s]].displayName
        );
        Appointment.viewer =
          StaffMembers[appointment.staffMemberIds[s]].displayName; //+ StaffMembers[appointment.staffMemberIds[s]].role
      }
    } catch (err) {}
  }

  Appointment.staffMemberNames = staffMemberNames;
  return Appointment;
}

async function show_csv_download() {
  document.getElementById("csv_dl").style = "none";

  return 0;
}

$(function () {
  var dateFormat = "mm/dd/yy",
    from = $("#from")
    .datepicker({
      defaultDate: "+1w",
      changeMonth: true,
      numberOfMonths: 1,
      dateFormat: "yy-mm-dd",
    })
    .on("change", function () {
      to.datepicker("option", "minDate", getDate(this));
    }),
    to = $("#to")
    .datepicker({
      defaultDate: "+1w",
      changeMonth: true,
      numberOfMonths: 1,
      dateFormat: "yy-mm-dd",
    })
    .on("change", function () {
      from.datepicker("option", "maxDate", getDate(this));
    });

  function getDate(element) {
    var date;
    try {
      date = $.datepicker.parseDate(dateFormat, element.value);
    } catch (error) {
      date = null;
    }

    return date;
  }
});

$(document).ready(function () {
  var date = new Date();
  date.setDate(date.getDate() - 1);

  document.getElementById("from").value = date.toISOString();
  document.getElementById("to").value = date.toISOString();

  var lastDate = new Date();
  lastDate.setDate(lastDate.getDate() - 7); //any date you want
  $("#from").datepicker("setDate", lastDate);

  var curdate = new Date();
  curdate.setDate(curdate.getDate()); //any date you want
  $("#to").datepicker("setDate", curdate);
});

// Source: https://stackoverflow.com/questions/492994/compare-two-dates-with-javascript
// Source: http://stackoverflow.com/questions/497790
var dates = {
  convert: function (d) {
    // Converts the date in d to a date-object. The input can be:
    //   a date object: returned without modification
    //  an array      : Interpreted as [year,month,day]. NOTE: month is 0-11.
    //   a number     : Interpreted as number of milliseconds
    //                  since 1 Jan 1970 (a timestamp)
    //   a string     : Any format supported by the javascript engine, like
    //                  "YYYY/MM/DD", "MM/DD/YYYY", "Jan 31 2009" etc.
    //  an object     : Interpreted as an object with year, month and date
    //                  attributes.  **NOTE** month is 0-11.
    return d.constructor === Date ?
      d :
      d.constructor === Array ?
      new Date(d[0], d[1], d[2]) :
      d.constructor === Number ?
      new Date(d) :
      d.constructor === String ?
      new Date(d) :
      typeof d === "object" ?
      new Date(d.year, d.month, d.date) :
      NaN;
  },
  compare: function (a, b) {
    // Compare two dates (could be of any type supported by the convert
    // function above) and returns:
    //  -1 : if a < b
    //   0 : if a = b
    //   1 : if a > b
    // NaN : if a or b is an illegal date
    // NOTE: The code inside isFinite does an assignment (=).
    return isFinite((a = this.convert(a).valueOf())) &&
      isFinite((b = this.convert(b).valueOf())) ?
      (a > b) - (a < b) :
      NaN;
  },
  inRange: function (d, start, end) {
    // Checks if date in d is between dates in start and end.
    // Returns a boolean or NaN:
    //    true  : if d is between start and end (inclusive)
    //    false : if d is before start or after end
    //    NaN   : if one or more of the dates is illegal.
    // NOTE: The code inside isFinite does an assignment (=).
    return isFinite((d = this.convert(d).valueOf())) &&
      isFinite((start = this.convert(start).valueOf())) &&
      isFinite((end = this.convert(end).valueOf())) ?
      start <= d && d <= end :
      NaN;
  },
};