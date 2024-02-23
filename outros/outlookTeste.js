var outlook = require('node-outlook');

var jsonBody = {
    "Subject": "test event",
    "Body": {
        "ContentType": "HTML",
        "Content": "hello world"
    },
    "Start": {
        "DateTime": "2024-02-02T17:00:00",
        "TimeZone": "India Standard Time"
    },
    "End": {
        "DateTime": "2024-02-02T17:36:24",
        "TimeZone": "India Standard Time"
    },
    "location": {
        "displayName": "Noida"
    },
    "Attendees": []
};

let createEventParameters = {
    token: "9012fe60-2379-4e76-9324-26070288c45b",
    event: jsonBody
};
outlook.calendar.createEvent(createEventParameters, function (error, event) {
    if(error) {
        console.log(error);                 
    } else {
        console.log(event);                         
    }
});