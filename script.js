window.onload = () => {
  const dropContainer = document.getElementById('drop-area');
  const btn = document.getElementById('saveBtn');
  const fileName = document.getElementById('filename');
  const progressLog = document.getElementById('progress');


  function handleFile(e) {
    let files = e.target.files, f = files[0];
    fileName.innerHTML = `<p>Name: <span>${f.name}</span></p><p>Size <span>${Math.round(f.size / 1024)}KB</span> </p><p>Modified <span>${f.lastModifiedDate}</span></p>`;

    let reader = new FileReader();
    reader.onload = function (e) {
      progressLog.innerText = `Processing in progress!`;
      btn.innerHTML = `<span class="ml-2"><i class="fa fa-spinner fa-spin"></i></span>`;

      let data = new Uint8Array(e.target.result);
      let workbook = XLSX.read(data, { type: 'array' });
      let wb = XLSX.utils.book_new();
      let first_sheet_name = workbook.SheetNames[0];

      /* Get worksheet */
      let worksheet = workbook.Sheets[first_sheet_name];
      // Convert the worksheet into JSON
      let dataJson = XLSX.utils.sheet_to_json(worksheet);
      let newArr = [];
      newArr.push(Object.keys(dataJson[0]));
      dataJson.forEach(element => {
        newArr.push(Object.values(element));
      });
      // Remove the column titles from the other array
      let headers = newArr.splice(0, 1);
      // Change the first column name to 'No'
      headers[0][0] = "No";

      // sheet names
      let ws_names = ["All Customers", "Connected Customers", "Online Customers", "Blocked Customers", "Inactive Customers", "New Customers", "Daily Report", "Pending Tickets"];
      let ws_functions = [allCustomers, connectedCustomers, onlineCustomers, blockedCustomers, inactiveCustomers, newCustomers, dailyReport, pendingTickets];

      // loop seven times too create the seven sheets
      for (let i = 0; i <= 7; i++) {
        /* make worksheet */
        let ws_data = [];
        if (i == 6) {
          // On the sixth loop use the generated title instead of the sheet name
          ws_data.push([`${formatDateString('title')}`]);
        } else {
          ws_data.push([ws_names[i]]);
        };
        // Push all the headers into the ws_data array
        ws_data.push(...headers);
        let resData = ws_functions[i](dataJson);

        // Increment the numbering in the first column
        count = 1;
        resData.forEach(el => {
          el['__EMPTY'] = count;
          count++;
        });
        let filteredData = [];
        // Get the object values from the filtered response and put them in the filteredData array
        resData.forEach(element => {
          filteredData.push(Object.values(element));
        });
        ws_data.push(...filteredData);

        var ws = XLSX.utils.aoa_to_sheet(ws_data);
        /* Add the worksheet to the workbook */
        XLSX.utils.book_append_sheet(wb, ws, ws_names[i]);
        console.log(ws_names[i] + " sheet added to wb");
      }
      progressLog.innerText = `All sheets added to the workbook!`;
      console.log('Modified wb:', wb);
      var wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

      // un disable the download button
      btn.innerHTML = `Download Report <span class="ml-2"><i class="fa fa-download"></i></span>`;
      btn.disabled = false;

      // Function to save and download the processed excel report file
      function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;

      }
      // set report
      btn.addEventListener('click', (e) => {
        saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), `${formatDateString()}.xlsx`);
      }, false);
    };
    reader.readAsArrayBuffer(f);
  };
  // Add change even listener on the droparea
  dropContainer.addEventListener('change', handleFile, false);
};

// Function to get all customers
const allCustomers = (fileContent) => {
  let list = fileContent.filter(
    (item) =>
      item.Status !== "New" ||
      item["Account balance"].match(/\d/g).join("") > 1999
  );
  return list;
};
// Function to get all connnected customers
const connectedCustomers = (fileContent) => {
  let list = fileContent.filter((item) => item.Status !== "New");
  return list;
};
// Function to get all customers who are online
const onlineCustomers = (fileContent) => {
  let list = fileContent.filter(
    (item) => item.Status !== "New" && item.Status !== "Blocked" && item.Status !== "Inactive"
  );
  return list;
};
// Function to get all blocked customers
const blockedCustomers = (fileContent) => {
  let list = fileContent.filter((item) => item.Status == "Blocked");
  return list;
};
// Function to get a list of all inactive customers
const inactiveCustomers = (fileContent) => {
  let list = fileContent.filter((item) => item.Status == "Inactive");
  return list;
};
// Function to get an array of all new customers both non-connected and connected
const newCustomers = (fileContent) => {
  let today = new Date();
  today = today.toLocaleDateString("en-US");
  let list = fileContent.filter(
    (item) =>
      (item.Status == "New" && item["Account balance"].match(/\d/g).join("") > 1999) || excelDateToJSDate(item["Internet services start date"]) == today
  );
  return list;
};
// Function to get the daily report - only addes the headline at the moment
const dailyReport = (fileContent) => {
  let list = fileContent.filter(
    (item) =>
    (item.Status == "New" &&
      item["Account balance"].match(/\d/g).join("") > 999999)
  );
  return list;
};
// Function to get new but not yet connected customers
const pendingTickets = (fileContent) => {
  let list = fileContent.filter(
    (item) =>
    (item.Status == "New" &&
      item["Account balance"].match(/\d/g).join("") > 1999)
  );
  return list;
};

// Function to format both the excel report name and the daily report heading title
const formatDateString = (type) => {
  let dayOfReport, dateOfReport, monthOfReport, yearOfReport, dateNotation;
  const date = new Date(Date.now());
  const Days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  console.log(Days[date.getDay()]);
  dayOfReport = Days[date.getDay()];
  dateOfReport = date.getDate();
  dateNotation = formatDay(date.getDate());
  monthOfReport = date.toLocaleString('en-us', { month: 'long' });
  yearOfReport = date.getFullYear();

  if (type && typeof type == 'string' && type == 'title') {
    formattedDate = `Daily Report (${dayOfReport}, ${dateOfReport}${dateNotation} ${monthOfReport} ${yearOfReport})`;
  } else {
    formattedDate = `Report ${dateOfReport}${dateNotation} ${monthOfReport} ${yearOfReport}`;

  };
  return formattedDate;
};


// Function to get the correct date notation
const formatDay = (day) => {
  let formattedDay;
  if (day == 11 || day == 12 || day == 13) {
    formattedDay = 'th';
  }
  else if (day % 10 === 1) {
    formattedDay = "st";
  } else if (day % 10 === 2) {
    formattedDay = "nd";
  } else if (day % 10 === 3) {
    formattedDay = "rd";
  } else {
    formattedDay = "th";
  }
  return formattedDay;
};

// Function to convert a date from the excel date format to normal JS date format
const excelDateToJSDate = (excelDate) => {
  if (typeof excelDate == "number") {

    let date = new Date(Math.round((excelDate - (25567 + 2)) * 86400 * 1000));
    let converted_date = date.toISOString().split('T')[0];
    converted_date = new Date(date);
    converted_date = date.toLocaleDateString("en-US");
    return converted_date;
  }
  return excelDate;
};
